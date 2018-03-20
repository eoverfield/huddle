import * as React from 'react';
import { Fabric } from 'office-ui-fabric-react';
import { IReactFooterProps } from './IReactFooterProps';

import styles from './ReactFooter.module.scss';

export class ReactFooter extends React.Component<IReactFooterProps, {}> {
  public render(): JSX.Element {
    return (
      <Fabric>
        <div className={styles.app}>
          <div className={`ms-bgColor-themeDark ms-fontColor-white ${styles.footer}`}>
            <b>{this.props.description || 'A default footer message'}</b>
          </div>
        </div>
      </Fabric>
    );
  }
}