import { Loader, Provider, ThemeInput, teamsTheme } from "@fluentui/react-northstar";
import { PublicClientApplication, Configuration } from "@azure/msal-browser";
import * as microsoftTeams from "@microsoft/teams-js";
import { useState, useEffect } from "react";
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

      const msal = new PublicClientApplication(config);
      await msal.initialize();

      // Establish the active account before EOCHome's acquireTokenSilent runs.
      // ssoSilent routes through the Teams NAA broker — no popup, no redirect.
      try {
        const result = await msal.ssoSilent({
          loginHint: ctx.user?.userPrincipalName,
          scopes: ["openid", "profile"],
        });
        msal.setActiveAccount(result.account);
      } catch (ssoError) {
        console.warn("NAA ssoSilent failed, falling back to getAllAccounts()", ssoError);
        const accounts = msal.getAllAccounts();
        if (accounts.length > 0) {
          msal.setActiveAccount(accounts[0]);
        }
      }

      setMsalInstance(msal);
      setLoading(false);
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
            <Loader style={{ margin: 100 }} />
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
