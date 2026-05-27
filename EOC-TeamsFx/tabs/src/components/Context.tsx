import { IPublicClientApplication } from "@azure/msal-browser";
import { createContext } from "react";
import { ThemePrepared } from "@fluentui/react-northstar";

export const TeamsFxContext = createContext<{
  theme?: ThemePrepared;
  themeString: string;
  msalInstance?: IPublicClientApplication;
  loginHint?: string;
}>({
  theme: undefined,
  themeString: "",
  msalInstance: undefined,
  loginHint: undefined,
});
