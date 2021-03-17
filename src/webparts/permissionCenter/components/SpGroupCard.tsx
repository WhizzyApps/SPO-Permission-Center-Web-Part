import * as React from 'react';
import { useBoolean } from '@uifabric/react-hooks';
import { useState } from 'react';
import { IButtonStyles } from 'office-ui-fabric-react';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

import { SPHttpClient, ISPHttpClientOptions} from '@microsoft/sp-http';

import cssStyles from './PermissionCenter.module.scss';

interface Props {
  spGroup: any;
  state;
  props;
}

const GroupCard: React.FC<Props> = ( ({spGroup, state, props}) => {

  try {

    const actionButtonStyles: IButtonStyles = {
      root: {
        width: "-webkit-fill-available",
        height: "-webkit-fill-available",
        margin: "0.1875rem 0",
        padding: "0.4375rem 1rem",
        maxWidth: "13.125rem"
      },
    };
    
    // dialog variables
    const [hideDialog, { toggle: toggleHideDialog }] = useBoolean(true);
    const [hideMessage, { toggle: toggleHideMessage }] = useBoolean(true);
    const [messageTitle, setMessageTitle] = useState('');
    const [message, setMessage] = useState('');

    // event handlers
    const _deleteGroup = async () => {
      toggleHideDialog();
      const url = props.siteCollectionURL + `/_api/web/sitegroups/removebyid(${spGroup.id})`;
      const clientOptions: ISPHttpClientOptions = {
        headers: new Headers(),
        method: 'post',
        mode: 'cors'
      };
      try {
        const response = await props.spHttpClient.post(url, SPHttpClient.configurations.v1, clientOptions);
        if (props.config.logComponentVars) {console.log("delete group response", response);}
        // if success
        if (response.ok) {
          // show message
          setMessageTitle('Success');
          setMessage(`${spGroup.groupName} deleted.`);
          toggleHideMessage();
        }
        // if error
        else {
          const responseJson = await response.json();
          if (props.config.logComponentVars) {console.log('response', responseJson);}
          // show message
          setMessageTitle('SharePoint Error');
          setMessage(responseJson["odata.error"].message.value);
          toggleHideMessage();
        }
          
      } catch (error) {
        setMessageTitle('Error');
        setMessage(`${spGroup.groupName} could not be deleted.`);
        toggleHideMessage();
        if (props.config.logComponentVars) {console.log(error);}
        if (props.config.throwErrors) {throw error;}
      }
    };

    const _deleteSharingGroup = async () => {
      // if item (in SharePoint) of group is deleted, group did not get item properties in _getPropsOfSharingGroups in PermissionCenter.tsx
      // so there is no group property item, delete group via _deleteGroup
      if (!spGroup.item) {
        _deleteGroup();
      }
      // if spGroup.item exists, unshare link: group and link will be deleted.
      else {
        toggleHideDialog();
        const listId = spGroup.item.listId;
        const itemId = spGroup.item.guid;
        const shareId = spGroup.item.shareId;
        const url = `${props.siteCollectionURL}/_api/web/Lists('${listId}')/GetItemByUniqueId('${itemId}')/UnshareLink`;

        const requestHeaders: Headers = new Headers();  
        requestHeaders.append('Content-type', 'application/json'); 
        requestHeaders.append('Accept', 'application/json'); 
        requestHeaders.append('Authorization', 'Bearer'); 

        const body: string = JSON.stringify({
          shareId: shareId
        });
        
        const clientOptions: ISPHttpClientOptions = {
          headers: requestHeaders,
          method: 'POST',
          mode: 'cors',
          body: body
        };

        try {
          const response = await props.spHttpClient.post(url, SPHttpClient.configurations.v1, clientOptions);
                    
          // if success
          if (response.ok) {
            // show message
            setMessageTitle('Success');
            setMessage(`${spGroup.groupName} deleted.`);
            toggleHideMessage();
          }
          // if error
          else {
            const responseJson = await response.json();
            if (props.config.logComponentVars) {console.log('response', responseJson);}
            // show message
            setMessageTitle('SharePoint Error');
            setMessage(responseJson["odata.error"].message.value);
            toggleHideMessage();
          }
            
        } catch (error) {
          setMessageTitle('Error');
          setMessage(`${spGroup.groupName} could not be deleted.`);
          toggleHideMessage();
          if (props.config.logComponentVars) {console.log(error);}
          if (props.config.throwErrors) {throw error;}
        }
      }
    };
    
    const _editGroup = (event) => {
      if (event.currentTarget.className === "myClick") { // fix the bug: outer click
        window.open(props.siteCollectionURL + `/_layouts/15/editgrp.aspx?Group=${spGroup.groupName}`);
      }
    };
    const _handlerDeleteGroup = (event) => {
      if (event.currentTarget.className === "myClick") { // fix the bug: outer click
        // if user is no owner or admin
        if ((state.mode == "member") || (state.mode == "visitor")) {
          // show message
          setMessageTitle('Note');
          setMessage(`You need to be owner to delete a group.`);
          toggleHideMessage();
        }
        // if user is owner or admin
        else {
          toggleHideDialog();
        }
      }
    };
    const _showGroupinSp = (event) => {
      if (event.currentTarget.className === "myClick") { // fix the bug: outer click
        window.open(props.siteCollectionURL + `/_layouts/15/people.aspx?MembershipGroupId=${spGroup.id}`);
      }
    };
    const _editGroupPermissions = (event) => {
      if (event.currentTarget.className === "myClick") { // fix the bug: outer click
        window.open(props.siteCollectionURL + `/_layouts/15/editprms.aspx?sel=${spGroup.id}`);
      }
    };

    //dialog variables
    const dialogContentProps = {
      type: DialogType.normal,
      title: 'Delete group',
      closeButtonAriaLabel: 'Close'
    };
    const messageContentProps = {
      type: DialogType.normal,
      title: messageTitle,
      closeButtonAriaLabel: 'Close',
    };

    // display "set as" if default group
    let isDefaultGroup = false;
    if (spGroup.defaultGroup) {
      isDefaultGroup = true;
    }

    //-------------------------------
    return (
      <div className={ cssStyles.groupCardContainer }>

        {/* info content */}
        <div className={ cssStyles.dataContainer }>
          <div className={ cssStyles.dataRow }>
            <span className={ cssStyles.dataKeyShort }>Name </span>
            <span className={ cssStyles.dataValue }>{spGroup.groupName}</span>
          </div>
          <div className={ cssStyles.dataRow }>
            <span className={ cssStyles.dataKeyShort }>Type </span>
            <span className={ cssStyles.dataValue }>{spGroup.type}</span>
          </div>
          {isDefaultGroup && (
            <div className={ cssStyles.dataRow }>
            <span className={ cssStyles.dataKeyShort }>Set as </span>
            <span className={ cssStyles.dataValue }>{spGroup.defaultGroup}</span>
          </div>
          )}
          <div className={ cssStyles.dataRow }>
            <span className={ cssStyles.dataKeyLong }>Group owner </span>
            <span className={ cssStyles.dataValue }>{spGroup.owner}</span>
          </div>
          {spGroup.permissionLevel && spGroup.permissionLevel[0]
          && <div className={ cssStyles.dataRow }>
            <span className={ cssStyles.dataKeyLong }>Permission level </span>
            <span className={ cssStyles.dataValue }>{spGroup.permissionLevel.join(', ')}</span>
          </div>
          }
          
        </div>
  
        {/* buttons */}
        <div className={ cssStyles.groupCardButtonContainer }>

          {/* Edit group */}
          {props.config.showCardButtons && !spGroup.displayName.startsWith('Sharing') && 
            <div className={"myClick"} onClick={(e)=>{_editGroup(e);}}>
              <DefaultButton text="Edit group" styles={actionButtonStyles} href={'#'}/>
            </div>
          }
          {/* Delete group */}
          {props.config.showCardButtons &&
            <div className={"myClick"} onClick={(e)=>{_handlerDeleteGroup(e);}}>
              <DefaultButton text="Delete group" styles={actionButtonStyles} href={'#'} />
            </div>
          }
          {/* Show group in SharePoint */}
          <div className={"myClick"} onClick={(e)=>{_showGroupinSp(e);}}>
            <DefaultButton text="Show group in SharePoint" styles={actionButtonStyles} href={'#'} />
          </div>

          {/* Edit group permissions */}
          {props.config.showCardButtons && !spGroup.displayName.startsWith('Sharing') && 
            <div className={"myClick"} onClick={(e)=>{_editGroupPermissions(e);}}>
              <DefaultButton text="Edit group permissions" styles={actionButtonStyles} href={'#'}/>
            </div>
          }

        </div>
        
        
        <Dialog
          hidden={hideDialog}
          onDismiss={toggleHideDialog}
          dialogContentProps={dialogContentProps}
        >
          <div>Do you want to delete "{spGroup.groupName}"? </div>
          <DialogFooter>
            <PrimaryButton onClick={((spGroup.loginName) && (spGroup.loginName.startsWith("SharingLinks."))) ? _deleteSharingGroup : _deleteGroup} text="Delete" />
            <DefaultButton onClick={toggleHideDialog} text="Cancel" />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={hideMessage}
          onDismiss={toggleHideMessage}
          dialogContentProps={messageContentProps}
          >
          <div> {message} </div>
          <PrimaryButton onClick={toggleHideMessage} text="Close" style={{float:"right", margin:"1.25rem 0"}}/>
        </Dialog>

      </div>
    );

  } catch (error) {
    if (props.config.logComponentVars) {console.log(error);}
    if (props.config.throwErrors) {throw error;}
  }
});
export default GroupCard;