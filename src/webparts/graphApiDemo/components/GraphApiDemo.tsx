import * as React from 'react';
import styles from './GraphApiDemo.module.scss';
import { IGraphApiDemoProps } from './IGraphApiDemoProps';
import { IGraphApiDemoState } from './IGraphApiDemoState';
import { escape } from '@microsoft/sp-lodash-subset';

export default class GraphApiDemo extends React.Component<IGraphApiDemoProps, IGraphApiDemoState> {

  constructor(props: IGraphApiDemoProps) {
    super(props);
    this.state = {
      messages: [{
        subject: ''
      }]
    }
  }

  public getDriveItems() {
    let getMessages: string = "me/messages";

    if (!this.props.graphClient) {
      return;
    }

    this.props.graphClient
      .api(getMessages)
      .version("v1.0")
      .select("subject,sentDateTime,webLink")
      .top(5)
      .get((err: any, res: any): void => {
        if (err) {
          console.log("Getting error in retrieving mesages =>", err)
        }
        if (res) {
          console.log("Success");
          if (res && res.value.length) {
            console.log(res.value);
            this.setState({
              messages: res.value
            })
          }
        }
      });
  }

  public componentDidMount() {
    this.getDriveItems();
  }

  public render(): React.ReactElement<IGraphApiDemoProps> {
    return (
      <div className={styles.graphApiDemo}>
        { this.state.messages.map(m => <><span>{m.subject}</span><br /></>)}
      </div>
    );
  }
}
