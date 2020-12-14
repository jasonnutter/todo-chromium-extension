import { useState, useEffect } from "react";
import { getAccessToken, TASKS_SCOPES } from "./auth";
import { GraphProfile, OutlookTask, OutlookTaskFolder, OutlookTaskFolders } from "./types";

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

async function getGraphProfile(): Promise<GraphProfile> {
    const profile = await getGraphJson<GraphProfile>("https://graph.microsoft.com/beta/me", TASKS_SCOPES);

    return profile;
}

export function useGraphProfile(defaultProfile: GraphProfile): [ GraphProfile, React.Dispatch<React.SetStateAction<GraphProfile>> ] {
    const [ graphProfile, setGraphProfile ] = useState<GraphProfile>(defaultProfile);

    useEffect(() => {
        getGraphProfile()
            .then(profile => setGraphProfile(profile))
            .catch(() => setGraphProfile(defaultProfile));
    }, []); // eslint-disable-line react-hooks/exhaustive-deps

    return [ graphProfile, setGraphProfile ];
}

export async function getTaskFolders(name?: string): Promise<OutlookTaskFolders> {
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

export async function addTask(title: string, url: string, folderName: string): Promise<OutlookTask> {
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
