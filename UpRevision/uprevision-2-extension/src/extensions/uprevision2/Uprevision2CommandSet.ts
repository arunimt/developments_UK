import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'Uprevision2CommandSetStrings';
import { SPPermission } from '@microsoft/sp-page-context';
import UpRevisionPanel from './UpRevision/UpRevisionPanel';
import RelatedDocsUploadPanel from './RelatedDocsUpload/RelatedDocsUploadPanel';


/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IUprevision2CommandSetProperties {
  // This is an example; replace with your own properties
  // sampleTextOne: string;
  // sampleTextTwo: string;
}

const LOG_SOURCE: string = 'Uprevision2CommandSet';

export default class Uprevision2CommandSet extends BaseListViewCommandSet<IUprevision2CommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized Uprevision2CommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    Log.info(LOG_SOURCE, this.context.pageContext.list.title);
    const UpRevisionCommand: Command = this.tryGetCommand('UPREVISION');
    if (UpRevisionCommand) {
      UpRevisionCommand.visible = this.context.pageContext.list.title === 'Draft' 
      && this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems);
    }
    const UploadCommand: Command = this.tryGetCommand('UPLOADDOCS');
    if (UploadCommand) {
      UploadCommand.visible = this.context.pageContext.list.title === 'Contract Related Documents' 
      && this.context.pageContext.list.permissions.hasPermission(SPPermission.editListItems);
    }
    
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'UPREVISION':
        var uprevisonDialog: UpRevisionPanel = new UpRevisionPanel();
        uprevisonDialog.show();
        break;
      case 'UPLOADDOCS':
        var uploadRelatedDocDialog: RelatedDocsUploadPanel = new RelatedDocsUploadPanel();
        uploadRelatedDocDialog.show();
         break;
      default:
        throw new Error('Unknown command');
    }
  }
}
