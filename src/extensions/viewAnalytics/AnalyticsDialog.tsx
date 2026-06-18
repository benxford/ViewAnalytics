import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  PrimaryButton,
  DialogFooter,
  DialogContent,
} from 'office-ui-fabric-react';

interface IAnalyticsDialogContentProps {
  close: () => void;
  viewsAllTime?: string;
  viewsLastSevenDays?: string;
}
interface IAnalyticsDialogContentState {
}

class AnalyticsDialogContent extends React.Component<IAnalyticsDialogContentProps, IAnalyticsDialogContentState> {
  constructor(props: IAnalyticsDialogContentProps) {
    super(props);
  }

  public render(): JSX.Element {
    return <DialogContent
      title='Views'
      onDismiss={this.props.close}
      showCloseButton={true}
    >
      <p>
        <b>All Time:</b> {this.props.viewsAllTime}<br/>
        <b>Last Seven Days:</b> {this.props.viewsLastSevenDays}
      </p>

      <DialogFooter>
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.close(); }} />
      </DialogFooter>
    </DialogContent>;
  }
}

export default class AnalyticsDialog extends BaseDialog {
  viewsAllTime?: string;
  viewsLastSevenDays?: string;

  constructor(viewsAllTime: string, viewsLastSevenDays: string) {
    super();
    this.viewsAllTime = viewsAllTime;
    this.viewsLastSevenDays = viewsLastSevenDays;
  }

  public render(): void {
    ReactDOM.render(<AnalyticsDialogContent
      close={ this.close }
      viewsAllTime={ this.viewsAllTime }
      viewsLastSevenDays={ this.viewsLastSevenDays }
    />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }
}