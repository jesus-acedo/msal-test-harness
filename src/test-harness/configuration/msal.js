import { LogLevel } from "@azure/msal-browser";

const apiUrl = process.env.CLIENT_API_URL;
const clientId = process.env.CLIENT_CLIENT_ID;
const authority = process.env.CLIENT_AUTHORITY;
const redirectUri = process.env.CLIENT_REDIRECT_URI;
//const appInsightsConnectionString = process.env.APPLICATION_INSIGHTS_CONNECTION_STRING;
const scopes = (process.env.CLIENT_SCOPES || "").split(";").map(s => s.trim()).filter(Boolean);

export const appConfig = {
    apiUrl: apiUrl,
    dateFormat: "YYYY/MM/DD",
};

export const msalConfig = {
    auth: {
        clientId: clientId,
        authority: authority,
        redirectUri: redirectUri,
        // postLogoutRedirectUri: '/',
        validateIssuer: true,
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            logLevel: LogLevel.Verbose,
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    case LogLevel.Error:
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        console.warn(message);
                        return;
                    default:
                        return;
                }
            }
        }
    }
};


export const loginRequest = {    
    scopes: scopes
};
