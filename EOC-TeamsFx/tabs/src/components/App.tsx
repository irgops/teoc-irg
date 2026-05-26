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
