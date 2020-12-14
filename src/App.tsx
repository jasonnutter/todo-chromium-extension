import React, { useState, useEffect, useRef } from 'react';
import './App.css';
import { AccountInfo, AuthenticationResult, InteractionRequiredAuthError, PublicClientApplication } from '@azure/msal-browser';
import {
    PrimaryButton, DefaultButton,
    Spinner, SpinnerSize,
    MessageBar, MessageBarType,
    Stack,
    CompoundButton,
    Persona, PersonaSize,
    Pivot, PivotItem,
    TextField, ComboBox, IComboBoxOption, IComboBox, Label
} from 'office-ui-fabric-react';

import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

initializeIcons();

// Set the redirect URI to the chromiumapp.com provided by Chromium
const redirectUri = typeof chrome !== "undefined" && chrome.identity ?
    chrome.identity.getRedirectURL() : 
    `${window.location.origin}/index.html`;

const msal = new PublicClientApplication({
    auth: {
        authority: "https://login.microsoftonline.com/common/",
        clientId: "36cb3b59-915a-424e-bc06-f8f557baa72f",
        redirectUri,
        postLogoutRedirectUri: redirectUri
    },
    cache: {
        cacheLocation: "localStorage"
    }
});

const TASKS_SCOPES = [
    "Tasks.ReadWrite",
    "User.Read",
    "openid",
    "profile"
];

const DEFAULT_TASKS_FOLDER_NAME = "Links";
const TASK_FOLDER_SYNC_KEY = "taskFolder";
const NEW_FOLDER_SUFFIX = " (New)";

type ChromeTab = Partial<chrome.tabs.Tab>;
type ChromeUser = chrome.identity.UserInfo;

type OutlookTask = {
    subject: string,
    id?: string,
    body?: {
        contentType: string,
        content: string
    },
    parentFolderId?: string
}

type OutlookTaskFolder = {
    id: string,
    name: string
}

type OutlookTaskFolders = {
    value: OutlookTaskFolder[]
}

type GraphProfile = {
    displayName: string,
    userPrincipalName: string
}

async function getAccessToken(scopes: string[]): Promise<string> {
    try {
        const { accessToken } = await msal.acquireTokenSilent({ scopes, account: msal.getAllAccounts()[0] });

        return accessToken;
    } catch (e) {
        if (InteractionRequiredAuthError.isInteractionRequiredError(e.errorCode, e.errorMessage)) {
            const { accessToken } = await msal.acquireTokenPopup({ scopes });

            return accessToken;
        }

        throw e;
    }
}

async function getCachedAccessToken(scopes: string[]): Promise<string | null> {
    // acquireTokenSilent will throw an error in Chrome extensions if a network request is made.
    const response = await msal.acquireTokenSilent({ scopes, account: msal.getAllAccounts()[0] }).catch(() => null);

    return response && response.accessToken;
}

async function getSignedInUser(): Promise<ChromeUser> {
    return new Promise((resolve, reject) => {
        if (chrome && chrome.identity) {
            // Running in extension popup
            chrome.identity.getProfileUserInfo((user: ChromeUser) => {
                if (user) {
                    resolve(user);
                } else {
                    reject();
                }
            });
        } else {
            // Running on localhost
            reject();
        }
    })
}

async function login(loginHint?: string): Promise<AuthenticationResult | null> {
    const redirectUri = chrome.identity.getRedirectURL();
    const loginUrl = await getLoginUrl(redirectUri, loginHint);

    return launchWebAuthFlow(loginUrl);
}

async function getLoginUrl(redirectUri: string, loginHint?: string): Promise<string> {
    return new Promise((resolve) => {
        msal.loginRedirect({
            redirectUri,
            scopes: TASKS_SCOPES,
            loginHint,
            onRedirectNavigate: (url) => {
                resolve(url);
                return false;
            }
        })
    });
}

async function logout(): Promise<void> {
    const logoutUrl = await getLogoutUrl();

    await launchWebAuthFlow(logoutUrl);
}

async function getLogoutUrl(): Promise<string> {
    return new Promise(resolve => {
        msal.logout({
            onRedirectNavigate: (url: string) => {
                resolve(url);
                return false;
            }
        })
    })
}

/**
 * Launch the Chromium web auth UI.
 * @param {*} url AAD url to navigate to.
 * @param {*} interactive Whether or not the flow is interactive
 */
async function launchWebAuthFlow(url: string): Promise<AuthenticationResult | null> {
    return new Promise<AuthenticationResult | null>((resolve, reject) => {
        chrome.identity.launchWebAuthFlow({
            interactive: true,
            url
        }, (responseUrl) => {
            // Response urls includes a hash (login, acquire token calls)
            if (responseUrl && responseUrl.includes("#")) {
                msal.handleRedirectPromise(`#${responseUrl.split("#")[1]}`)
                    .then(resolve)
                    .catch(reject)
            } else {
                // Logout calls
                resolve(null);
            }
        })
    })
}

async function fetchJson<T>(url: string, method: string, headers: object, body?: object): Promise<T> {
    const request = await fetch(url, {
        method,
        headers: new Headers({
            "Content-Type": "application/json",
            ...headers
        }),
        body: body ? JSON.stringify(body) : undefined
    });

    const response = await request.json();

    if (request.status >= 400) {
        throw new Error(response.error.message);
    }

    return response as T;
}

async function getGraphJson<T>(url: string, scopes: string[]): Promise<T> {
    const accessToken = await getAccessToken(scopes);

    const headers = {
        Authorization: `Bearer ${accessToken}`
    };

    return fetchJson<T>(url, "GET", headers);
}

async function postGraphJson<T>(url: string, scopes: string[], body: object): Promise<T> {
    const accessToken = await getAccessToken(scopes);

    const headers = {
        Authorization: `Bearer ${accessToken}`
    };

    return fetchJson<T>(url, "POST", headers, body);
}

async function getCurrentTab(): Promise<ChromeTab> {
    return new Promise((resolve, reject) => {
        if (chrome && chrome.tabs) {
            // Running in extension popup
            chrome.tabs.query({ active: true }, (tab: ChromeTab[]) => {
                if (tab)  {
                    resolve(tab[0]);
                } else {
                    reject();
                }
            })
        } else {
            // Running on localhost
            resolve({
                title: document.title,
                url: window.location.href
            })
        }
    });
}

async function readSyncedValue<T>(key: string, defaultValue: T): Promise<T> {
    return new Promise((resolve, reject) => {
        if (chrome && chrome.storage) {
            chrome.storage.sync.get(key, (result) => {
                resolve(result[key] || defaultValue);
            });
        } else {
            try {
                const value = localStorage.getItem(key);
                const result = value ? JSON.parse(value) as T : defaultValue;

                resolve(result);
            } catch (e) {
                reject(e);
            }
        }
    });
}

async function saveSyncedValue<T>(key: string, value: T): Promise<T> {
    return new Promise((resolve, reject) => {
        if (chrome && chrome.storage) {
            chrome.storage.sync.set({ [key]: value }, () => {
                resolve(value);
            })
        } else {
            try {
                localStorage.setItem(key, JSON.stringify(value));
                resolve(value);
            } catch (e) {
                reject(e);
            }
        }
    });
}

async function getGraphProfile(): Promise<GraphProfile> {
    const profile = await getGraphJson<GraphProfile>("https://graph.microsoft.com/beta/me", TASKS_SCOPES);

    return profile;
}

async function getTaskFolders(name?: string): Promise<OutlookTaskFolders> {
    const url = name ? `https://graph.microsoft.com/beta/me/outlook/taskFolders?$filter=startswith(name, '${name}')` : `https://graph.microsoft.com/beta/me/outlook/taskFolders`;

    return getGraphJson<OutlookTaskFolders>(url, TASKS_SCOPES);
}

async function createTaskFolder(name: string): Promise<OutlookTaskFolder> {
    return postGraphJson<OutlookTaskFolder>('https://graph.microsoft.com/beta/me/outlook/taskFolders', TASKS_SCOPES, { name });
}

async function getOrCreateTaskFolder(name: string): Promise<OutlookTaskFolder> {
    const existingFolders = await getTaskFolders(name);

    if (existingFolders.value.length > 0) {
        return existingFolders.value[0];
    }

    return createTaskFolder(name);
}

async function addTask(title: string, url: string, folderName: string): Promise<OutlookTask> {
    const { id: folderId } = await getOrCreateTaskFolder(folderName);

    const newTask = await postGraphJson<OutlookTask>(`https://graph.microsoft.com/beta/me/outlook/taskFolders/${folderId}/tasks`, TASKS_SCOPES, {
        subject: title,
        body: {
            contentType: "Text",
            content: url
        }
    } as OutlookTask);

    return newTask;
}

function useSignedInUser(defaultUser: ChromeUser): [ ChromeUser, React.Dispatch<React.SetStateAction<ChromeUser>>] {
    const [ signedInUser, setSignedInUser ] = useState<ChromeUser>(defaultUser);

    useEffect(() => {
        getSignedInUser()
            .then(user => setSignedInUser(user))
            .catch(() => setSignedInUser(defaultUser))
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ signedInUser, setSignedInUser ];
}

function useCurrentTab(defaultTab: ChromeTab): [ ChromeTab, React.Dispatch<React.SetStateAction<ChromeTab>>] {
    const [ currentTab, setCurrentTab ] = useState<ChromeTab>(defaultTab);

    useEffect(() => {
        getCurrentTab()
            .then(tab => setCurrentTab(tab))
            .catch(() => setCurrentTab(defaultTab))
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ currentTab, setCurrentTab ];
}

function useSyncedValue<T>(key: string, defaultValue: T, callback?: (result: T) => void): [T | null, (value: T) => Promise<T> ] {
    const [ syncedValue, setSyncedValue ] = useState<T | null>(defaultValue);

    useEffect(() => {
        readSyncedValue(key, defaultValue)
            .then((result: T) => {
                if (result) {
                    setSyncedValue(result as T);
                }

                if (callback) {
                    callback(result as T);
                }
            })
            .catch(() => setSyncedValue(null));
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    const saveAndSetSyncedValue = (value: T): Promise<T> => {
        setSyncedValue(value);

        return saveSyncedValue(key, value);
    };

    return [ syncedValue, saveAndSetSyncedValue ];
}

function useCachedAccessToken(defaultCachedAccessToken: string | null): [ string | null, React.Dispatch<React.SetStateAction<string | null>>] {
    const [ cachedAccessToken, setCachedAccessToken ] = useState<string | null>(defaultCachedAccessToken);

    useEffect(() => {
        getCachedAccessToken(TASKS_SCOPES)
            .then(accessToken => setCachedAccessToken(accessToken))
            .catch(() => setCachedAccessToken(defaultCachedAccessToken));
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ cachedAccessToken, setCachedAccessToken ];
}

function useGraphProfile(defaultProfile: GraphProfile): [ GraphProfile, React.Dispatch<React.SetStateAction<GraphProfile>> ] {
    const [ graphProfile, setGraphProfile ] = useState<GraphProfile>(defaultProfile);

    useEffect(() => {
        getGraphProfile()
            .then(profile => setGraphProfile(profile))
            .catch(() => setGraphProfile(defaultProfile));
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ graphProfile, setGraphProfile ];
}

const GraphProfile: React.FC = () => {
    const [ graphProfile ] = useGraphProfile({
        userPrincipalName: '',
        displayName: ''
    });

    if (!graphProfile.userPrincipalName) {
        return (
            <Spinner size={SpinnerSize.large} />
        );
    }

    return (
        <Persona
            size={PersonaSize.regular}
            text={graphProfile.displayName}
            secondaryText={graphProfile.userPrincipalName}
            imageInitials={graphProfile.displayName[0]}
        />
    )
};

const App: React.FC = () => {
    const [ account, setAccount ] = useState<AccountInfo | null>(msal.getAllAccounts()[0]);
    const [ success, setSuccess ] = useState<boolean>(false);
    const [ inProgress, setInProgress ] = useState<boolean>(false);
    const [ latestTask, setLatestTask ] = useState<string | undefined>('');
    const [ selectedTaskFolderIndex, setSelectedTaskFolderIndex] = useState<number>(0);

    // Selected folder (in progress)
    const [ selectedTaskFolder, setSelectedTaskFolder ] = useState<IComboBoxOption | null>(null);

    // Folder name (saved)
    const [ , setSavedTaskFolder ] = useState<string>(DEFAULT_TASKS_FOLDER_NAME);

    // Folders fetched from API
    const [ taskFolders, setTaskFolders ] = useState<IComboBoxOption[]>([
        { key: DEFAULT_TASKS_FOLDER_NAME, text: DEFAULT_TASKS_FOLDER_NAME }
    ]);

    // Load synced folder name
    const [ ,setSyncedFolderName ] = useSyncedValue<string>(TASK_FOLDER_SYNC_KEY, DEFAULT_TASKS_FOLDER_NAME, (folderName => {
        setTaskFolders([
            { key: folderName, text: folderName}
        ]);
        setSavedTaskFolder(folderName);
    }));

    const comboBoxRef = useRef<IComboBox | null>(null);

    const [ cachedAccessToken, setCachedAccessToken ] = useCachedAccessToken(null);

    const [ currentTab, setCurrentTab ] = useCurrentTab({
        title: '',
        url: ''
    });

    const [ signedInUser ] = useSignedInUser({
        email: '',
        id: ''
    });

    return (
        <div className="wrapper">
            {account ? (
                <Pivot>
                    <PivotItem headerText="Save Link">
                        <Stack tokens={{ childrenGap: 15 }}>
                            <form
                                onSubmit={async (e) => {
                                    e.preventDefault();

                                    setSuccess(false);
                                    setInProgress(true);
                                    setLatestTask('');

                                    if (selectedTaskFolder && selectedTaskFolder.text) {
                                        const newFolderName = selectedTaskFolder.text.split(NEW_FOLDER_SUFFIX)[0].trim();

                                        setSavedTaskFolder(newFolderName);
                                        setSyncedFolderName(newFolderName);

                                        const { id } = await addTask(currentTab.title || "", currentTab.url || "", newFolderName);
                                        setLatestTask(id);
                                        setSuccess(true);
                                    }

                                    setInProgress(false);
                                }}
                            >
                                <ul>
                                    {!cachedAccessToken && (
                                        <li>
                                            <PrimaryButton
                                                onClick={async () => {
                                                    const accessToken = await getAccessToken(TASKS_SCOPES);
                                                    setCachedAccessToken(accessToken);
                                                }}
                                            >
                                                Get Access Token
                                            </PrimaryButton>
                                        </li>
                                    )}
                                    <li>
                                        <TextField
                                            label="Title"
                                            value={currentTab.title}
                                            onChange={e => {
                                                const target = e.target as HTMLTextAreaElement;

                                                setCurrentTab({
                                                    ...currentTab,
                                                    title: target.value
                                                });
                                            }}
                                        />
                                    </li>
                                    <li>
                                        <TextField
                                            label="URL"
                                            multiline={true}
                                            value={currentTab.url}
                                            onChange={e => {
                                                const target = e.target as HTMLInputElement;

                                                setCurrentTab({
                                                    ...currentTab,
                                                    url: target.value
                                                });
                                            }}
                                        />
                                    </li>
                                    <li>
                                        <ComboBox
                                            allowFreeform={true}
                                            selectedKey={selectedTaskFolder ? selectedTaskFolder.key : taskFolders[0].key}
                                            label="Task Folder"
                                            onPendingValueChanged={async (option, index, value) => {
                                                if (value) {
                                                    const folders = await getTaskFolders(value);
                                                    const folderExists = folders.value.find(folder => folder.name === value);
                                                    const newFolders: IComboBoxOption[] = folderExists ? [] : [ { key: value, text: `${value}${NEW_FOLDER_SUFFIX}` }]

                                                    if (comboBoxRef.current) {
                                                        comboBoxRef.current.focus(true);
                                                    }

                                                    if (folders.value.length) {
                                                        setTaskFolders(newFolders.concat(folders.value.map(folder => ({
                                                            key: folder.id,
                                                            text: folder.name
                                                        }))));
                                                    } else {
                                                        setTaskFolders(newFolders);
                                                    }
                                                } else {
                                                    setSelectedTaskFolder(null);
                                                }
                                            }}
                                            onItemClick={(e, option, index) => {
                                                if (option && option.text) {
                                                    setSelectedTaskFolder(option);
                                                    setSelectedTaskFolderIndex(index || 0);
                                                }
                                            }}
                                            onChange={(e, option, index) => {
                                                const target = e.target as HTMLInputElement;

                                                if (option && option.text) {
                                                    setSelectedTaskFolder(option);
                                                    setSelectedTaskFolderIndex(index || 0);
                                                } else {
                                                    const folderIndex = taskFolders.findIndex(folder => folder.text.split(NEW_FOLDER_SUFFIX)[0] === target.value);
                                                    setSelectedTaskFolder(taskFolders[folderIndex]);
                                                    setSelectedTaskFolderIndex(folderIndex);
                                                }
                                            }}
                                            onBlur={(e => {
                                                if (taskFolders.length) {
                                                    setSelectedTaskFolder(taskFolders[selectedTaskFolderIndex])
                                                }
                                            })}
                                            onScrollToItem={(index) => {
                                                setSelectedTaskFolderIndex(index > -1 ? index : 0);
                                            }}
                                            componentRef={comboBoxRef}
                                            options={taskFolders}
                                        />
                                    </li>
                                    <li>
                                        <PrimaryButton
                                            disabled={(
                                                inProgress ||
                                                !cachedAccessToken ||
                                                !currentTab.title
                                            )}
                                            type="submit"
                                        >
                                            Save Link
                                        </PrimaryButton>
                                    </li>
                                </ul>
                            </form>

                            {(inProgress || success) && (
                                <Stack tokens={{ childrenGap: 15 }}>
                                    {inProgress && (
                                        <Spinner size={SpinnerSize.medium} />
                                    )}
                                    {success && (
                                        <MessageBar
                                            messageBarType={MessageBarType.success}
                                        >
                                            Link saved successfully.
                                        </MessageBar>
                                    )}
                                    {latestTask && (
                                        <DefaultButton
                                            href={`https://to-do.microsoft.com/tasks/id/${latestTask}/details`}
                                            target="_blank"
                                        >
                                            View Task
                                        </DefaultButton>
                                    )}
                                </Stack>
                            )}
                        </Stack>
                    </PivotItem>
                    <PivotItem headerText="Account">
                        <Stack tokens={{ childrenGap: 15 }}>
                            <Label>Account</Label>

                            {account && (
                                <GraphProfile />
                            )}

                            <DefaultButton
                                onClick={async () => {
                                    setAccount(null);
                                    await logout();
                                }}
                            >
                                Logout
                            </DefaultButton>
                        </Stack>
                    </PivotItem>
                </Pivot>
            ) : (
                <Pivot>
                    <PivotItem headerText="Account" style={{ paddingTop: '15px'}}>
                        <Stack horizontal={true} tokens={{ childrenGap: 15 }}>
                            {signedInUser.email && (
                                <CompoundButton
                                    primary={true}
                                    onClick={async () => {
                                        await login(signedInUser.email);
                                        setAccount(msal.getAllAccounts()[0]);
                                    }}
                                    secondaryText={`(w/ ${signedInUser.email})`}
                                    style={{
                                        width: '200px'
                                    }}
                                >
                                    Login
                                </CompoundButton>
                            )}
                            <CompoundButton
                                onClick={async () => {
                                    await login();
                                    setAccount(msal.getAllAccounts()[0]);
                                }}
                                secondaryText="(w/ your Microsoft account)"
                                style={{
                                    width: '200px'
                                }}
                            >
                                Login
                            </CompoundButton>
                        </Stack>
                    </PivotItem>
                </Pivot>
            )}
        </div>
    );
}

export default App;
