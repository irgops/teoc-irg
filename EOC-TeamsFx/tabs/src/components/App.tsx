import { Loader, Provider, ThemeInput, teamsTheme } from "@fluentui/react-northstar";
import { PublicClientApplication, Configuration } from "@azure/msal-browser";
import { ApplicationInsights } from "@microsoft/applicationinsights-web";
import { SeverityLevel } from "@microsoft/applicationinsights-common";
import * as microsoftTeams from "@microsoft/teams-js";
import { useState, useEffect } from "react";
import { unstable_batchedUpdates } from "react-dom";
import { Redirect, Route, HashRouter as Router } from "react-router-dom";
import { TeamsFxContext } from "./Context";
import Tab from "./Tab";
import TabConfig from './TabConfig';

const clientId = process.env.REACT_APP_CLIENT_ID!;

export default function App() {
  const [loading, setLoading] = useState(true);
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | undefined>(undefined);
  const [theme, setTheme] = useState<any>(undefined);
  const [themeString, setThemeString] = useState("");

  useEffect(() => {
    async function init() {
      await microsoftTeams.app.initialize();
      const ctx = await microsoftTeams.app.getContext();
      setThemeString(ctx.app.theme ?? "default");

      const config: Configuration = {
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${ctx.user?.tenant?.id ?? "common"}`,
          supportsNestedAppAuth: true,
        },
      };

      const appInsights = new ApplicationInsights({
        config: {
          instrumentationKey: process.env.REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY || ""
        }
      });
      appInsights.loadAppInsights();

      const msal = new PublicClientApplication(config);
      await msal.initialize();

      // Set the active account before the instance reaches context so that
      // EOCHome's acquireTokenSilent never runs against an account-less instance.
      try {
        const result = await msal.ssoSilent({
          loginHint: ctx.user?.userPrincipalName,
          scopes: ["openid", "profile"],
        });
        msal.setActiveAccount(result.account);
      } catch (ssoError: any) {
        console.error("NAA ssoSilent failed", {
          errorCode: ssoError?.errorCode,
          errorMessage: ssoError?.errorMessage,
          subError: ssoError?.subError,
          name: ssoError?.name,
          loginHint: ctx.user?.userPrincipalName,
        }, ssoError);
        appInsights.trackException(
          { exception: ssoError, severityLevel: SeverityLevel.Error },
          {
            Component: "App",
            Method: "ssoSilent",
            User: ctx.user?.userPrincipalName,
            ErrorCode: ssoError?.errorCode,
            SubError: ssoError?.subError,
          }
        );
        const accounts = msal.getAllAccounts();
        if (accounts.length > 0) {
          msal.setActiveAccount(accounts[0]);
        }
      }

      // React 16 does not auto-batch state updates in async callbacks — each
      // setter would fire its own render cycle. Without batching, there is a
      // render where msalInstance is set in context but loading is still true,
      // which can race with consumers. unstable_batchedUpdates collapses both
      // updates into one render so the instance only reaches context the same
      // frame that loading becomes false and EOCHome mounts.
      unstable_batchedUpdates(() => {
        setMsalInstance(msal);
        setLoading(false);
      });
    }

    init().catch((err) => {
      console.error("NAA MSAL init failed", err);
      setLoading(false);
    });
  }, []);

  return (
    <TeamsFxContext.Provider value={{ theme, themeString, msalInstance }}>
      <Provider theme={(theme as ThemeInput) || teamsTheme} styles={{ backgroundColor: "#eeeeee" }}>
        <Router>
          <Route exact path="/">
            <Redirect to="/tab" />
          </Route>
          {loading ? (
            <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh" }}>
              <Loader size="largest" label="Signing you in..." />
            </div>
          ) : (
            <>
              <Route exact path="/tab" component={Tab} />
              <Route exact path="/config" component={TabConfig} />
            </>
          )}
        </Router>
      </Provider>
    </TeamsFxContext.Provider>
  );
}
