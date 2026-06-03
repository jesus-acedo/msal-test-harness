import React, { useState } from "react";
import { useMsal, AuthenticatedTemplate, UnauthenticatedTemplate } from "@azure/msal-react";
import { loginRequest } from "../../configuration/msal";

const fallbackApiOptions = [
    "https://localhost:30122"
];

const apiOptions = (process.env.CLIENT_API_OPTIONS || "")
    .split(";")
    .map((url) => url.trim())
    .filter(Boolean);

const resolvedApiOptions = apiOptions.length > 0 ? apiOptions : fallbackApiOptions;

function ApiTester() {
    const { instance, accounts } = useMsal();
    const [duration, setDuration] = useState(null);
    const [apiBaseUrl, setApiBaseUrl] = useState(resolvedApiOptions[0]);
    const [method, setMethod] = useState("GET");
    const [path, setPath] = useState("/api/lbhealth");
    const [queryString, setQueryString] = useState("");
    const [payload, setPayload] = useState("{}");
    const [bearerToken, setBearerToken] = useState("");
    const [response, setResponse] = useState("");

    const isAuthenticated = Array.isArray(accounts) && accounts.length > 0;

    async function handleSubmit() {
        let url = `${apiBaseUrl.replace(/\/$/, "")}/${path.replace(/^\//, "")}`;

        if (queryString.trim() !== "") {
            url += `?${queryString.trim()}`;
        }

        console.log("uri:", url);
        console.log("payload:", payload);

        try {
            let authorization;

            if (bearerToken.trim() !== "") {
                const tokenValue = bearerToken.trim();
                authorization = tokenValue.match(/^Bearer\s+/i)
                    ? tokenValue
                    : `Bearer ${tokenValue}`;
            } else {
                if (!isAuthenticated) {
                    throw new Error("Please sign in or enter a bearer token before sending the request.");
                }

                const tokenResponse = await instance.acquireTokenSilent({
                    ...loginRequest,
                    account: accounts[0]
                });
                authorization = `Bearer ${tokenResponse.accessToken}`;
            }

            const options = {
                method,
                headers: {
                    "Content-Type": "application/json",
                    Authorization: authorization
                }
            };

            if (method !== "GET") {
                options.body = JSON.stringify(JSON.parse(payload));
            }

            const start = performance.now();
            const res = await fetch(url, options);
            const end = performance.now();

            setDuration(((end - start) / 1000).toFixed(3));

            const text = await res.text();

            try {
                setResponse(JSON.stringify(JSON.parse(text), null, 2));
            } catch {
                setResponse(text);
            }
        } catch (err) {
            setResponse(`Error: ${err.message}`);
        }
    }

    return (
        <>
            <div style={{ marginTop: "20px" }}>
                <div style={{ marginBottom: "12px", fontSize: "0.95em", color: "#333" }}>
                    {isAuthenticated ? (
                        <span>Signed in with MSAL. Provide a bearer token to override the MSAL access token.</span>
                    ) : (
                        <span>Not signed in. Enter a bearer token to send requests without logging in.</span>
                    )}
                </div>

                <div style={{ marginTop: "10px" }}>
                    <label>API Base URL:</label>
                    <select
                        value={apiBaseUrl}
                        onChange={(e) => setApiBaseUrl(e.target.value)}
                    >
                        {resolvedApiOptions.map((u) => (
                            <option key={u} value={u}>
                                {u}
                            </option>
                        ))}
                    </select>
                </div>

                <div style={{ marginTop: "10px" }}>
                    <label>Method:</label>
                    <select
                        value={method}
                        onChange={(e) => setMethod(e.target.value)}
                    >
                        <option>GET</option>
                        <option>POST</option>
                        <option>PUT</option>
                        <option>DELETE</option>
                    </select>

                    <input
                        type="text"
                        placeholder="/api/values"
                        value={path}
                        onChange={(e) => setPath(e.target.value)}
                        style={{ marginLeft: "10px", width: "300px" }}
                    />
                </div>

                <div style={{ marginTop: "10px" }}>
                    <label>Querystring (without ?):</label>
                    <input
                        type="text"
                        placeholder="key=value&foo=bar"
                        value={queryString}
                        onChange={(e) => setQueryString(e.target.value)}
                        style={{ marginLeft: "10px", width: "300px" }}
                    />
                </div>

                <div style={{ marginTop: "10px" }}>
                    <label>Bearer token (optional):</label>
                    <input
                        type="text"
                        placeholder="Enter custom bearer token"
                        value={bearerToken}
                        onChange={(e) => setBearerToken(e.target.value)}
                        style={{ marginLeft: "10px", width: "500px" }}
                    />
                    <div style={{ fontSize: "0.85em", color: "#555", marginTop: "4px" }}>
                        If provided, this token is used instead of the MSAL access token.
                    </div>
                </div>

                <div style={{ marginTop: "10px" }}>
                    <label>Payload (JSON):</label>
                    <br />
                    <textarea
                        rows={6}
                        cols={80}
                        value={payload}
                        onChange={(e) => setPayload(e.target.value)}
                    />
                </div>

                <div style={{ marginTop: "10px" }}>
                    <button onClick={handleSubmit}>Send Request</button>
                </div>

                <div style={{ marginTop: "20px" }}>
                    <label>
                        Response {duration !== null ? `(took: ${duration} s)` : ""}:
                    </label>
                    <pre
                        style={{
                            background: "#f5f5f5",
                            padding: "10px",
                            borderRadius: "6px",
                            whiteSpace: "pre-wrap"
                        }}
                    >
                        {response}
                    </pre>
                </div>
            </div>
        </>
    );
}

export function HomePage() {
    const { instance } = useMsal();

    async function handleLogin(e) {
        e.preventDefault();
        try {
            const loginResponse = await instance.loginPopup();
            console.log(loginResponse);
        } catch (error) {
            console.log(error);
        }
    }

    function handleLogout(e) {
        e.preventDefault();
        instance.logoutRedirect();
    }

    return (
        <>
            <UnauthenticatedTemplate>
                <div className="container-fluid" style={{ padding: "20px" }}>
                    <button onClick={handleLogin}>Sign in</button>
                    <ApiTester />
                </div>
            </UnauthenticatedTemplate>

            <AuthenticatedTemplate>
                <div className="container-fluid" style={{ padding: "20px" }}>
                    <button onClick={handleLogout}>Sign out</button>
                    <ApiTester />
                </div>
            </AuthenticatedTemplate>
        </>
    );
}
