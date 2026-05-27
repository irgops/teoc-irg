import { TeamsFxContext } from "./Context";
import loadable from "@loadable/component";

const EOCHome = loadable(() => import("./EOCHome"));

export default function Tab() {
  return (
    <div>
      <TeamsFxContext.Consumer>
        {(value) => <EOCHome msalInstance={value.msalInstance!} loginHint={value.loginHint} />}
      </TeamsFxContext.Consumer>
    </div>
  );
}
