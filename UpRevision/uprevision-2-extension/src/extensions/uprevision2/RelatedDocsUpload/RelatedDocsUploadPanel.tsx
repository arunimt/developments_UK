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

interface IRelatedDocsUploadPanelDialogContentProps {
  submit: () => void;
}

export default class RelatedDocsUploadPanel extends BaseDialog {
  public render(): void {

    const owuScript = require('./RelatedDocsUpload-Script.js');
    require('sp-init');
    require('microsoft-ajax');
    require('sp-runtime');
    require('sharepoint');

    ReactDOM.render(
      <DialogContent
        title='Upload Related Document'
        showCloseButton={true}
        onDismiss={() => this.close()}>
        <div id = 'search'>
          <table>
            <tr>
              <td><Label required={true}>contract ID</Label></td>
              <td><input id="ContactId" type="text" /></td>
              <td><PrimaryButton text='Confirm' title='Confirm' onClick={owuScript.RelatedDocsUpload} /></td>
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
              <td><Label>Document Type </Label></td>
              <td><input id="DocumentTypeEdit" type="text" disabled/></td>
            </tr>            
            <tr>
              <td><Label>Counter Party </Label></td>
              <td><input id="CounterPartyEdit" type="text" disabled/></td>
            </tr>
            <tr>
              <td><Label>Master Agreement Number </Label></td>
              <td><input id="MasterAgreementNumEdit" type="text" disabled/></td>
            </tr>
            <tr>
                <td><Label required={true}>Please select the file to upload </Label></td>
                <td><input id="file" type="file"/></td>
            </tr>
            <tr>
                <td></td>
                <td><Label id = "lblFileMsg"></Label></td>
            </tr>
            <tr>
              <td><Label required={true}>Please enter Check-in comments </Label></td>
              <td><TextField id='comment' multiline={true}/></td>
            </tr>
            <tr>
                <td></td>
                <td><Label id = "lblErrMsg"></Label></td>
            </tr>
            <tr>
                <td><PrimaryButton text='Submit' title='Submit' onClick={owuScript.ConfirmUpload} /></td>
                <td><Button text='Cancel' title='Cancel' onClick={owuScript.ReturnToMainMenu} /></td>
            </tr>
          </table>
        </div>
        <div id='message' color='Red' style={{display: 'none'}}> 
          <table>
              <tr>
                <td><label id = 'msg'></label></td>
              </tr>
              <tr>
                <td><Button text='Cancel' title='Cancel' onClick={this.close} /></td>
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