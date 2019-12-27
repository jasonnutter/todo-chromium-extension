// @ts-check

const msal = new Msal.UserAgentApplication({
    auth: {
        clientId: "36cb3b59-915a-424e-bc06-f8f557baa72f",
        redirectUri: "chrome-extension://cpjdlimhngdofgkacmlmgnkpkoadhldk/popup.html"
    },
    cache: {
        cacheLocation: "localStorage"
    }
});

const TASKS_SCOPES = [
    "Tasks.ReadWrite",
    "openid",
    "profile"
];

const TASKS_FOLDER_NAME = "Links";

async function getAccessToken(scopes) {
    try {
        const { accessToken } = await msal.acquireTokenSilent({ scopes });

        return accessToken;
    } catch (e) {
        if (e.name === "InteractionRequiredAuthError") {
            const { accessToken } = await msal.acquireTokenPopup({ scopes });

            return accessToken;
        }

        throw e;
    }
}

async function login() {
    return msal.loginPopup({
        scopes: TASKS_SCOPES
    });
}

async function logout() {
    msal.logout();
}

async function fetchJson(url, method, headers, body) {
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

    return response;
}

async function getGraphJson(url, scopes) {
    const accessToken = await getAccessToken(scopes);

    const headers = {
        Authorization: `Bearer ${accessToken}`
    };

    return fetchJson(url, "GET", headers);
}

async function postGraphJson(url, scopes, body) {
    const accessToken = await getAccessToken(scopes);

    const headers = {
        Authorization: `Bearer ${accessToken}`
    };

    return fetchJson(url, "POST", headers, body);
}

async function getCurrentTab() {
    return new Promise((resolve, reject) => {
        chrome.tabs.getSelected(null, tab => tab ? resolve(tab) : reject());
    });
}

async function getOrCreateTaskFolder(name) {
    const existingFolders = await getGraphJson(`https://graph.microsoft.com/beta/me/outlook/taskFolders?$filter=startswith(name, '${name}')`, TASKS_SCOPES);

    if (existingFolders.value.length > 0) {
        return existingFolders.value[0];
    }

    const newFolder = await postGraphJson('https://graph.microsoft.com/beta/me/outlook/taskFolders', TASKS_SCOPES, { name });

    return newFolder;
}

async function addTaskForCurrentTab() {
    const { id: folderId } = await getOrCreateTaskFolder(TASKS_FOLDER_NAME);

    const {
        title: tabTitle,
        url: tabUrl
    } = await getCurrentTab();

    const newTask = await postGraphJson(`https://graph.microsoft.com/beta/me/outlook/taskFolders/${folderId}/tasks`, TASKS_SCOPES, {
        subject: tabTitle,
        body: {
            contentType: "Text",
            content: tabUrl
        }
    });

    return newTask;
}

const loggedIn = document.querySelector("#logged-in");
const loggedOut = document.querySelector("#logged-out");

const logInButton = document.querySelector("#login");
const logOutButton = document.querySelector("#logout");
const saveButton = document.querySelector("#save");

function checkLogin() {
    if (msal.getAccount()) {
        loggedIn.classList.remove("hidden");
        loggedOut.classList.add("hidden");
    } else {
        loggedIn.classList.add("hidden");
        loggedOut.classList.remove("hidden");
    }
}

checkLogin();

saveButton.addEventListener("click", async () => {
    const newTask = await addTaskForCurrentTab();

    console.log("task", newTask);
});

logInButton.addEventListener("click", async () => {
    await login();
    checkLogin();
})

logOutButton.addEventListener("click", () => {
    logout();

    checkLogin();
})

