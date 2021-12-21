import * as React from "react";
import styles from "./GraphToolkit.module.scss";
import { IGraphToolkitProps } from "./IGraphToolkitProps";
import { escape } from "@microsoft/sp-lodash-subset";

export default class GraphToolkit extends React.Component<
  IGraphToolkitProps,
  {}
> {
  public render(): React.ReactElement<IGraphToolkitProps> {
    return (
      <div className={styles.graphToolkit}>
        <div className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <span className={styles.title}>Welcome to SharePoint!</span>
              <p className={styles.subTitle}>
                Customize SharePoint experiences using Web Parts.
              </p>
              <p className={styles.description}>
                {escape(this.props.description)}
              </p>
              <div
                dangerouslySetInnerHTML={{
                  __html: `<mgt-person person-query="me" view="twolines"></mgt-person>`,
                }}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
