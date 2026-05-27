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
const authScopes = ["openid", "profile"];

export default function App() {
  const [loading, setLoading] = useState(true);
  const [msalInstance, setMsalInstance] = useState<PublicClientApplication | undefined>(undefined);
  const [loginHint, setLoginHint] = useState<string | undefined>(undefined);
  const [theme, setTheme] = useState<any>(undefined);
  const [themeString, setThemeString] = useState("");

  useEffect(() => {
    async function init() {
      await microsoftTeams.app.initialize();
      const ctx = await microsoftTeams.app.getContext();
      setThemeString(ctx.app.theme ?? "default");

      const appInsights = new ApplicationInsights({
        config: { instrumentationKey: process.env.REACT_APP_APPINSIGHTS_INSTRUMENTATIONKEY || "" }
      });
      appInsights.loadAppInsights();

      // Pull loginHint before the cascade so it is available synchronously in
      // Layer 2 without an async call inside a catch block.
      const teamsLoginHint = ctx.user?.loginHint ?? ctx.user?.userPrincipalName;

      const config: Configuration = {
        auth: {
          clientId,
          authority: `https://login.microsoftonline.com/${ctx.user?.tenant?.id ?? "common"}`,
          supportsNestedAppAuth: true,
        },
      };

      const msal = new PublicClientApplication(config);
      await msal.initialize();

      // Three-layer NAA auth cascade for iOS WKWebView compatibility.
      //
      // Layer 1 — ssoSilent: works on Teams desktop and web. Fails with
      //   monitor_window_timeout on iOS because WKWebView blocks the hidden
      //   iframe that ssoSilent uses to contact the authority.
      //
      // Layer 2 — acquireTokenSilent with loginHint: bypasses the iframe
      //   entirely. With supportsNestedAppAuth, MSAL routes this directly
      //   through the Teams NAA broker cache — no popup, no redirect.
      //
      // Layer 3 — acquireTokenPopup: last resort only; should not be reached
      //   on any NAA-capable Teams client.
      let authResult: any;
      try {
        authResult = await msal.ssoSilent({ scopes: authScopes, loginHint: teamsLoginHint });
      } catch (err: any) {
        if (err.errorCode === "monitor_window_timeout" || err.errorCode === "interaction_required") {
          console.warn("NAA ssoSilent Layer 1 failed (iOS WKWebView expected), falling to Layer 2", err.errorCode);
          try {
            authResult = await msal.acquireTokenSilent({
              scopes: authScopes,
              loginHint: teamsLoginHint,
            });
          } catch (silentErr: any) {
            console.error("NAA Layer 2 acquireTokenSilent failed, falling to Layer 3 popup", {
              errorCode: silentErr?.errorCode,
              errorMessage: silentErr?.errorMessage,
              subError: silentErr?.subError,
              name: silentErr?.name,
            }, silentErr);
            appInsights.trackException(
              { exception: silentErr, severityLevel: SeverityLevel.Error },
              {
                Component: "App",
                Method: "acquireTokenSilent_L2",
                User: teamsLoginHint,
                ErrorCode: silentErr?.errorCode,
                SubError: silentErr?.subError,
              }
            );
            authResult = await msal.acquireTokenPopup({ scopes: authScopes });
          }
        } else {
          // Unexpected ssoSilent error — not a timeout or interaction issue.
          console.error("NAA ssoSilent unexpected error", {
            errorCode: err?.errorCode,
            errorMessage: err?.errorMessage,
            subError: err?.subError,
            name: err?.name,
            loginHint: teamsLoginHint,
          }, err);
          appInsights.trackException(
            { exception: err, severityLevel: SeverityLevel.Error },
            {
              Component: "App",
              Method: "ssoSilent_unexpected",
              User: teamsLoginHint,
              ErrorCode: err?.errorCode,
              SubError: err?.subError,
            }
          );
          throw err;
        }
      }

      msal.setActiveAccount(authResult.account);

      // React 16 does not auto-batch state updates in async callbacks — each
      // setter would fire its own render cycle. unstable_batchedUpdates collapses
      // all three into one render so msalInstance and loginHint never appear in
      // context while loading is still true.
      unstable_batchedUpdates(() => {
        setMsalInstance(msal);
        setLoginHint(teamsLoginHint);
        setLoading(false);
      });
    }

    init().catch((err) => {
      console.error("NAA MSAL init failed", err);
      setLoading(false);
    });
  }, []);

  return (
    <TeamsFxContext.Provider value={{ theme, themeString, msalInstance, loginHint }}>
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
