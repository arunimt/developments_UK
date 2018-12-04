import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  PrimaryButton,
  Button,
  DialogFooter,
  DialogContent,
  Label,
  TextField
} from 'office-ui-fabric-react';
import 'jquery';

interface IUpRevisionPanelDialogContentProps {
  submit: () => void;
}

export default class UpRevisionPanel extends BaseDialog {
  public render(): void {

    const owuScript = require('./OverWriteUpload-Script.js');
    require('sp-init');
    require('microsoft-ajax');
    require('sp-runtime');
    require('sharepoint');

    ReactDOM.render(
      <DialogContent
        title='Upload a New Version of a Contract'
        showCloseButton={true}
        onDismiss={() => this.close()}>
        <div id = 'search'>
          <table>
            <tr>
              <td><Label required={true}>Please Enter the Contract ID </Label></td>
              <td><input id="ContactId" type="text" /></td>
              <td><PrimaryButton text='Confirm' title='Confirm' onClick={owuScript.OverWriteUpload} /></td>
            </tr>
          </table>
        </div>
        <div id='metadata' style={{display: 'none'}}>
          <table>
            <tr>
              <td><Label>Contract ID </Label></td>
              <td><input id="ContactIdEdit" type="text" disabled/></td>
            </tr>
            <tr>
              <td></td>
              <td><Label id = "lblContactIdErrMsg"></Label></td>
            </tr>
            <tr>
              <td><Label>Counter Party </Label></td>
              <td><input id="CounterPartyEdit" type="text" disabled/></td>
            </tr>
            <tr>
              <td></td>
              <td><Label id = "lblCounterPartyErrMsg"></Label></td>
            </tr>
            <tr>
              <td><Label>Document Type </Label></td>
              <td><input id="DocumentTypeEdit" type="text" disabled/></td>
            </tr>
            <tr>
              <td></td>
              <td><Label id = "lblDocumentTypeErrMsg"></Label></td>
            </tr>
            <tr>
              <td><Label required={true}>Document Title </Label></td>
              <td><input id="DocumentTitleEdit" type="text" disabled/></td>
            </tr>
            <tr>
              <td></td>
              <td><Label id = "lblDocumentTitleErrMsg"></Label></td>
            </tr>
            <tr>
              <td><Label required={true}>Effective Date </Label></td>
              <td><input id="EffectiveDateEdit" type="date" disabled/></td>
            </tr>
            <tr>
              <td></td>
              <td><Label id = "lblEffectiveDateErrMsg"></Label></td>
            </tr>
            <tr>
              <td><Label required={true}>Master Agreement Number </Label></td>
              <td><input id="MasterAgreementNumEdit" type="text" disabled/></td>
            </tr>
            <tr>
              <td></td>
              <td><Label id = "lblMasterAgreementNumErrMsg"></Label></td>
            </tr>
            <tr>
              <td><Label required={true}>Version </Label></td>
              <td><input id="VersionEdit" type="text" disabled/></td>
            </tr>
          </table>
          
        </div>
        <div id='buttonmenu' style={{display: 'none'}}>
            <div id='returnToMainMenu' style={{display: 'inline'}}><Button text='Back' title='Back' onClick={owuScript.ReturnToMainMenu}/></div>
            <div id='blank' style={{display: 'inline'}}><span>                                    </span></div>
            <div id='editMetadata' style={{display: 'none'}}><Button text='Edit' title='Edit' onClick={owuScript.EnableMetadata}/></div>
            <div id='saveMetadata'  style={{display: 'none'}} ><Button text='Save'title='Save' onClick={owuScript.SaveMetaData}/></div>
          </div>
          <div id='uploadfile' style={{display: 'none'}}>
            <table>
              <tr>
                <td><Label required={true}>Please select the file to upload </Label></td>
                <td><input id="file" type="file" onChange={owuScript.CheckFileType}/></td>
              </tr>  
            </table>
          </div>
        <div id='confirmupload' style={{display: 'none'}}>
          <table>
          <tr>
            <td><Label required={true}>Please enter Check-in comments </Label></td>
            <td><TextField id='comment' multiline={true}/></td>
          </tr>
              <tr>
                <td></td>
                <td><Label id = "lblErrMsg"></Label></td>
              </tr>
              <tr>
                <td><Label id ="confirmUploadMsg" required={true}> </Label></td>
              </tr>
              <tr>
                <td><PrimaryButton text='Yes' title='Yes' onClick={owuScript.ConfirmUpload} /></td>
                <td><Button text='No' title='No' onClick={owuScript.ReturnToMainMenu} /></td>
              </tr>  
            </table>
          </div>  
        <div id='message' color='Red' style={{display: 'none'}}> 
          <table>
              <tr>
                <td><label id = 'msg'></label></td>
              </tr>
              <tr>
                <td><Button text='Cancel' title='Cancel' onClick={owuScript.ReturnToMainMenu} /></td>
              </tr>
            </table>
        </div>
        <div id='success' style={{display: 'none'}}> 
          <table>
              <tr>
                <td><label id = 'successMsg'></label></td>
              </tr>
              <tr>
                <td><Button text='OK' title='OK' onClick={this.close} /></td>
              </tr>
            </table>
        </div>
        <div id='progress' style={{display: 'none'}}> 
          <table>
              <tr>
                <td><label id = 'progressMsg'></label></td>
              </tr>
            </table>
        </div>
      </DialogContent>, this.domElement);
  }
}