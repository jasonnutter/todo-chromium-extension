import { useState, useEffect } from "react";
import { ChromeTab, ChromeUser } from "./types";

export async function getSignedInUser(): Promise<ChromeUser> {
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

export function useSignedInUser(defaultUser: ChromeUser): [ ChromeUser, React.Dispatch<React.SetStateAction<ChromeUser>>] {
    const [ signedInUser, setSignedInUser ] = useState<ChromeUser>(defaultUser);

    useEffect(() => {
        getSignedInUser()
            .then(user => setSignedInUser(user))
            .catch(() => setSignedInUser(defaultUser))
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ signedInUser, setSignedInUser ];
}

export async function getCurrentTab(): Promise<ChromeTab> {
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

export function useCurrentTab(defaultTab: ChromeTab): [ ChromeTab, React.Dispatch<React.SetStateAction<ChromeTab>>] {
    const [ currentTab, setCurrentTab ] = useState<ChromeTab>(defaultTab);

    useEffect(() => {
        getCurrentTab()
            .then(tab => setCurrentTab(tab))
            .catch(() => setCurrentTab(defaultTab))
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ currentTab, setCurrentTab ];
}

export async function readSyncedValue<T>(key: string, defaultValue: T): Promise<T> {
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

export async function saveSyncedValue<T>(key: string, value: T): Promise<T> {
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

export function useSyncedValue<T>(key: string, defaultValue: T, callback?: (result: T) => void): [T | null, (value: T) => Promise<T> ] {
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
