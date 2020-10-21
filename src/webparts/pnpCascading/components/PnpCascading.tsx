import * as React from 'react';
import styles from './PnpCascading.module.scss';
import { IPnpCascadingProps } from './IPnpCascadingProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class PnpCascading extends React.Component<IPnpCascadingProps, {}> {
  public render(): React.ReactElement<IPnpCascadingProps> {
    return (
      <div className={styles.pnpCascading}>
        <p>Selected List: {this.props.list}</p>
        <p>Selected Fields: {this.props.fields.toString()}</p>
      </div>
    );
  }
}
