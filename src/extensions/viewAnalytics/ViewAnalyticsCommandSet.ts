import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { SPHttpClient } from '@microsoft/sp-http';
import AnalyticsDialog from './AnalyticsDialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IViewAnalyticsCommandSetProperties {
}

const LOG_SOURCE: string = 'ViewAnalyticsCommandSet';

export default class ViewAnalyticsCommandSet extends BaseListViewCommandSet<IViewAnalyticsCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized ViewAnalyticsCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    compareOneCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_1':
        const webUrl: string = this.context.pageContext.web.absoluteUrl;
        const siteId: string = this.context.pageContext.site.id.toString();
        const webId: string = this.context.pageContext.web.id.toString();
        const listId: string = this.context.pageContext.list?.id.toString() || '';
        const itemId: string = this.context.listView.selectedRows?.[0]?.getValueByName("UniqueId") || '';

        const response = await this.context.spHttpClient.get(`${webUrl}/_api/v2.1/sites/${siteId},${webId}/lists/${listId}/items/${itemId}//driveItem?$select=id,analytics&$expand=analytics($expand=allTime,lastSevenDays)`, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json;odata.metadata=none'
          }
        }); 
        const analytics = await response.json();

        const allTime = analytics.analytics.allTime.access.actionCount;
        const lastSevenDays = analytics.analytics.lastSevenDays.access.actionCount;

        const dialog: AnalyticsDialog = new AnalyticsDialog(allTime, lastSevenDays);
        await dialog.show();

        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');


    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
