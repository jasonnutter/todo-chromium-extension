import { AccountInfo, AuthenticationResult, PublicClientApplication } from "@azure/msal-browser";

// Set the redirect URI to the chromiumapp.com provided by Chromium
export const redirectUri = typeof chrome !== "undefined" && chrome.identity ?
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

export const TASKS_SCOPES = [
    "Tasks.ReadWrite",
    "User.Read",
    "openid",
    "profile"
];

export async function getAccessToken(scopes: string[]): Promise<string> {
    try {
        const { accessToken } = await msal.acquireTokenSilent({ scopes, account: msal.getAllAccounts()[0] });

        return accessToken;
    } catch (e) {
        const acquireTokenUrl = await getLoginUrl();

        const result = await launchWebAuthFlow(acquireTokenUrl);

        return result?.accessToken || "";
    }
}

export async function getCachedAccessToken(scopes: string[]): Promise<string | null> {
    // acquireTokenSilent will throw an error in Chrome extensions if a network request is made.
    const response = await msal.acquireTokenSilent({ scopes, account: msal.getAllAccounts()[0] }).catch(() => null);

    return response && response.accessToken;
}

export async function getLoginUrl(loginHint?: string): Promise<string> {
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

export async function getLogoutUrl(): Promise<string> {
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
export async function launchWebAuthFlow(url: string): Promise<AuthenticationResult | null> {
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

export async function login(loginHint?: string): Promise<AuthenticationResult | null> {
    const loginUrl = await getLoginUrl(loginHint);

    return launchWebAuthFlow(loginUrl);
}

export async function logout(): Promise<void> {
    const logoutUrl = await getLogoutUrl();

    await launchWebAuthFlow(logoutUrl);
}

export function getActiveAccount(): AccountInfo | null {
    return msal.getAllAccounts()[0];
}
