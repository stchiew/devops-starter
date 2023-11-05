import * as React from "react";
import styles from "./DevOpsStarter.module.scss";
import { IDevOpsStarterProps } from "./IDevOpsStarterProps";
import { escape } from "@microsoft/sp-lodash-subset";

export default class DevOpsStarter extends React.Component<
  IDevOpsStarterProps,
  {}
> {
  public render(): React.ReactElement<IDevOpsStarterProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;

    return (
      <section
        className={`${styles.devOpsStarter} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className={styles.welcome}>
          <img
            alt=""
            src={
              isDarkTheme
                ? require("../assets/welcome-dark.png")
                : require("../assets/welcome-light.png")
            }
            className={styles.welcomeImage}
          />
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>
            Web part property value: <strong>{escape(description)}</strong>
          </div>
        </div>
        <div>
          <h3>Welcome to SharePoint Framework!</h3>
          <p>Second change to test commits to main branch</p>
        </div>
      </section>
    );
  }
}
