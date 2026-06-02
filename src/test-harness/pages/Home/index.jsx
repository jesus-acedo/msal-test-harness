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

export function HomePage() {
    const { instance, accounts, inProgress } = useMsal();

    const [duration, setDuration] = useState(null);
    const [apiBaseUrl, setApiBaseUrl] = useState(resolvedApiOptions[0]);
    const [method, setMethod] = useState("GET");
    const [path, setPath] = useState("/lbhealth");
    const [queryString, setQueryString] = useState("");
    const [payload, setPayload] = useState("{}");
    const [response, setResponse] = useState("");

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

    async function handleSubmit() {
        let url = `${apiBaseUrl.replace(/\/$/, "")}/${path.replace(/^\//, "")}`;

        if (queryString.trim() !== "") {
            url += `?${queryString.trim()}`;
        }


        console.log("uri:", url);
        console.log("payload:", payload);

        try {
            const tokenResponse = await instance.acquireTokenSilent(loginRequest);

            let options = {
                method,
                headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${tokenResponse.accessToken}`
                }
            };

            if (method !== "GET") {
                options.body = JSON.stringify(JSON.parse(payload));
            }

            const start = performance.now();
            const res = await fetch(url, options);
            const end = performance.now();

            setDuration(((end - start) / 1000).toFixed(3)); // seconds

            const text = await res.text();

            // Try JSON parse
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
            <UnauthenticatedTemplate>
                <div className="container-fluid">
                    <button onClick={handleLogin}>Sign in</button>
                </div>
            </UnauthenticatedTemplate>

            <AuthenticatedTemplate>
                <div className="container-fluid" style={{ padding: "20px" }}>
                    <button onClick={handleLogout}>Sign out</button>

                    {/* API Base URL Selector */}
                    <div style={{ marginTop: "20px" }}>
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

                    {/* Method + Path */}
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
                    {/* Query String */}
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


                    {/* Payload */}
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

                    {/* Submit */}
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
            </AuthenticatedTemplate>
        </>
    );
}
