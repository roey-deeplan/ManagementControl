import * as React from 'react';
import styles from './ManagementControl.module.scss';
import type { IManagementControlProps } from './IManagementControlProps';

export default class ManagementControl extends React.Component<IManagementControlProps, {}> {
  public render(): React.ReactElement<IManagementControlProps> {
    const {
      Title,
      ProductsListId
    } = this.props;

    return (
      <section >
        <h1>{Title}</h1>
      </section>
    );
  }
}
