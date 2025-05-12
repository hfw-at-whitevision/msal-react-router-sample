import React from 'react';
import ReactDOM from 'react-dom/client';
import { createBrowserRouter, RouterProvider } from "react-router-dom";
import { ThemeProvider } from "@mui/material/styles";
import { theme } from "./styles/theme";
import App from './App';

// MSAL imports
import { PublicClientApplication, EventType } from "@azure/msal-browser";
import { msalConfig } from "./authConfig";
// Import all the page components
import { Home, loader as homeLoader } from "./pages/Home";
import { Profile, loader as profileLoader, ErrorBoundary as ProfileErrorBoundary } from "./pages/Profile";
import { ProfileWithMsal } from "./pages/ProfileWithMsal";
import { ProfileRawContext } from "./pages/ProfileRawContext";
import { ProfileUseMsalAuthenticationHook } from "./pages/ProfileUseMsalAuthenticationHook";
import { Logout } from "./pages/Logout";

export const msalInstance = new PublicClientApplication(msalConfig);

msalInstance.initialize().then(() => {
  // Default to using the first account if no account is active on page load
  if (!msalInstance.getActiveAccount() && msalInstance.getAllAccounts().length > 0) {
    // Account selection logic is app dependent. Adjust as needed for different use cases.
    msalInstance.setActiveAccount(msalInstance.getAllAccounts()[0]);
  }

  // Optional - This will update account state if a user signs in from another tab or window
  msalInstance.enableAccountStorageEvents();

  msalInstance.addEventCallback((event) => {
    if (event.eventType === EventType.LOGIN_SUCCESS && event.payload.account) {
      const account = event.payload.account;
      msalInstance.setActiveAccount(account);
    }
  });

  // Create the router with routes configuration
  const router = createBrowserRouter([
    {
      path: "/",
      element: <App pca={msalInstance} />,
      children: [
        { 
          index: true, 
          element: <Home />,
          loader: homeLoader
        },
        { 
          path: "profile", 
          element: <Profile />,
          loader: profileLoader,
          errorElement: <ProfileErrorBoundary />
        },
        { 
          path: "profileWithMsal", 
          element: <ProfileWithMsal /> 
        },
        { 
          path: "profileRawContext", 
          element: <ProfileRawContext /> 
        },
        { 
          path: "profileUseMsalAuthenticationHook", 
          element: <ProfileUseMsalAuthenticationHook /> 
        },
        { 
          path: "logout", 
          element: <Logout /> 
        }
      ]
    }
  ]);

  const container = document.getElementById("root");
  const root = ReactDOM.createRoot(container);

  root.render(
    <ThemeProvider theme={theme}>
      <RouterProvider router={router} />
    </ThemeProvider>
  );
});
