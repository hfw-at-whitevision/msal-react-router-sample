import { Outlet, useNavigate } from "react-router-dom";
// Material-UI imports
import Grid from "@mui/material/Grid";

// MSAL imports
import { MsalProvider } from "@azure/msal-react";
import { CustomNavigationClient } from "./utils/NavigationClient";

// Sample app imports
import { PageLayout } from "./ui-components/PageLayout";
import { ThemeProvider } from "@mui/material/styles";
import { theme } from "./styles/theme";

function App({ pca }) {
    // The next 3 lines are optional. This is how you configure MSAL to take advantage of the router's navigate functions when MSAL redirects between pages in your app
    const navigate = useNavigate();
    const navigationClient = new CustomNavigationClient(navigate);
    pca.setNavigationClient(navigationClient);

    return (
        <MsalProvider instance={pca}>
            <ThemeProvider theme={theme}>
                <PageLayout>
                    <Grid container justifyContent="center">
                        <Outlet />
                    </Grid>
                </PageLayout>
            </ThemeProvider>
        </MsalProvider>
    );
}

export default App;
