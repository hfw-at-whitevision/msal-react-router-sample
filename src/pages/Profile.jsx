import { useLoaderData, useRouteError } from "react-router-dom";

// Msal imports
import { MsalAuthenticationTemplate, useMsal } from "@azure/msal-react";
import { InteractionType, InteractionRequiredAuthError } from "@azure/msal-browser";
import { loginRequest } from "../authConfig";

// Sample app imports
import { ProfileData } from "../ui-components/ProfileData";
import { Loading } from "../ui-components/Loading";
import { ErrorComponent } from "../ui-components/ErrorComponent";
import { callMsGraph } from "../utils/MsGraphApiCall";

// Material-ui imports
import Paper from "@mui/material/Paper";

// Loader function for Data Router
export async function loader() {
  try {
    const graphData = await callMsGraph();
    return { graphData };
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      throw new Response("Unauthorized", { status: 401 });
    }
    throw error;
  }
}

const ProfileContent = () => {
    const { graphData } = useLoaderData();
  
    return (
        <Paper>
            <ProfileData graphData={graphData} />
        </Paper>
    );
};

// Error boundary for the route
export function ErrorBoundary() {
  const error = useRouteError();
  const { instance } = useMsal();
  
  if (error.status === 401) {
    // Handle authentication error
    instance.acquireTokenRedirect({
      ...loginRequest,
      account: instance.getActiveAccount()
    });
    return <Loading />;
  }
  
  return <ErrorComponent error={error} />;
}

export function Profile() {
    const authRequest = {
        ...loginRequest
    };

    return (
        <MsalAuthenticationTemplate 
            interactionType={InteractionType.Popup} 
            authenticationRequest={authRequest} 
            errorComponent={ErrorComponent} 
            loadingComponent={Loading}
        >
            <ProfileContent />
        </MsalAuthenticationTemplate>
    );
};