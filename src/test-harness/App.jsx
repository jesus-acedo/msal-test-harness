import { StrictMode } from 'react';
import { PublicClientApplication, EventType } from '@azure/msal-browser';
import { MsalProvider } from '@azure/msal-react';
import { msalConfig } from './configuration/msal';
import { HomePage } from './pages/Home';

// Do not upgrade MSAL past 5.9.0 / 5.3.0 without retesting CIAM login.
// MSAL 5.10+ introduced issuer validation that can break valid CIAM authorities
// with endpoints_resolution_error even when .well-known discovery succeeds.
// See AzureAD/microsoft-authentication-library-for-js#8592 / PR #8570.

const pca = new PublicClientApplication(msalConfig);

pca.initialize().then(() => {
  // Default to using the first account if no account is active on page load
  if (!pca.getActiveAccount() && pca.getAllAccounts().length > 0) {
    // Account selection logic is app dependent. Adjust as needed for different use cases.
    pca.setActiveAccount(pca.getAllAccounts()[0]);
  }

  pca.addEventCallback((event) => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload) {
      const account = event.payload;
      pca.setActiveAccount(account);
    }
  });
});

export function App() {
  return (
    <StrictMode>
      <MsalProvider instance={pca}>
        <HomePage />
      </MsalProvider>
    </StrictMode>
  );
}