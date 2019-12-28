import React, { useState, useEffect } from 'react';
import './App.css';
import * as Msal from 'msal';
import {
    PrimaryButton, DefaultButton,
    Spinner, SpinnerSize,
    MessageBar, MessageBarType,
    Stack,
    CompoundButton,
    Persona, PersonaSize,
    Pivot, PivotItem,
    TextField
} from 'office-ui-fabric-react';

import { initializeIcons } from 'office-ui-fabric-react/lib/Icons';

initializeIcons();

const msal = new Msal.UserAgentApplication({
    auth: {
        authority: "https://login.microsoftonline.com/common/",
        clientId: "36cb3b59-915a-424e-bc06-f8f557baa72f",
        redirectUri: `${window.location.origin}/redirect.html`
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

const TASKS_FOLDER_NAME = "Links";

declare const chrome: {
    tabs: {
        getSelected: (windowId: Number | null, callback: (tab: ChromeTab) => void) => void
    },
    identity: {
        getProfileUserInfo: (callback: (user: ChromeUser) => void) => void
    }
}

type ChromeTab = {
    title: string,
    url: string
}

type ChromeUser = {
    email: string,
    id: string
}

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
        const { accessToken } = await msal.acquireTokenSilent({ scopes });

        return accessToken;
    } catch (e) {
        if (Msal.InteractionRequiredAuthError.isInteractionRequiredError(e.errorCode)) {
            const { accessToken } = await msal.acquireTokenPopup({ scopes });

            return accessToken;
        }

        throw e;
    }
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

async function login(loginHint?: string): Promise<Msal.AuthResponse> {
    return msal.loginPopup({
        scopes: TASKS_SCOPES,
        loginHint
    });
}

function logout(): void {
    msal.logout();
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
            chrome.tabs.getSelected(null, (tab: ChromeTab) => {
                if (tab)  {
                    resolve(tab);
                } else {
                    reject();
                }
            });
        } else {
            // Running on localhost
            resolve({
                title: document.title,
                url: window.location.href
            })
        }
    });
}

async function getGraphProfile(): Promise<GraphProfile> {
    const profile = await getGraphJson<GraphProfile>("https://graph.microsoft.com/beta/me", TASKS_SCOPES);

    return profile;
}

async function getOrCreateTaskFolder(name: string): Promise<OutlookTaskFolder> {
    const existingFolders = await getGraphJson<OutlookTaskFolders>(`https://graph.microsoft.com/beta/me/outlook/taskFolders?$filter=startswith(name, '${name}')`, TASKS_SCOPES);

    if (existingFolders.value.length > 0) {
        return existingFolders.value[0];
    }

    const newFolder = await postGraphJson<OutlookTaskFolder>('https://graph.microsoft.com/beta/me/outlook/taskFolders', TASKS_SCOPES, { name });

    return newFolder;
}

async function addTask(title: string, url: string): Promise<OutlookTask> {
    const { id: folderId } = await getOrCreateTaskFolder(TASKS_FOLDER_NAME);

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
    const [ account, setAccount ] = useState<Msal.Account | null>(msal.getAccount());
    const [ success, setSuccess ] = useState<boolean>(false);
    const [ inProgress, setInProgress ] = useState<boolean>(false);
    const [ latestTask, setLatestTask ] = useState<string | undefined>('');

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

                                    const { id } = await addTask(currentTab.title, currentTab.url);

                                    setLatestTask(id);
                                    setInProgress(false);
                                    setSuccess(true);
                                }}
                            >
                                <ul>
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
                                        <PrimaryButton
                                            disabled={inProgress}
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
                    <PivotItem headerText="Account" style={{ paddingTop: '15px'}}>
                        <Stack tokens={{ childrenGap: 15 }}>
                            {account && (
                                <GraphProfile />
                            )}

                            <DefaultButton
                                onClick={() => {
                                    logout();
                                    setAccount(null);
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
                                        setAccount(msal.getAccount());
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
                                    setAccount(msal.getAccount());
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
