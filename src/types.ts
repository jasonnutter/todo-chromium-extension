export type ChromeTab = Partial<chrome.tabs.Tab>;
export type ChromeUser = chrome.identity.UserInfo;

export type OutlookTask = {
    subject: string,
    id?: string,
    body?: {
        contentType: string,
        content: string
    },
    parentFolderId?: string
}

export type OutlookTaskFolder = {
    id: string,
    name: string
}

export type OutlookTaskFolders = {
    value: OutlookTaskFolder[]
}

export type GraphProfile = {
    displayName: string,
    userPrincipalName: string
}
