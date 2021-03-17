// import react 
import * as React from 'react';
import { useState, useRef, memo  } from 'react';
import { useBoolean } from '@uifabric/react-hooks';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { Persona, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import {
  getTheme,
  mergeStyleSets,
  FontWeights,
  ContextualMenu,
  DefaultButton,
  PrimaryButton,
  Modal,
  IDragOptions,
  IconButton,
  IIconProps,
  IButtonStyles
} from 'office-ui-fabric-react';

// import APIs
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { MSGraphClient } from '@microsoft/sp-http';

// import components
import { _reload } from './PermissionCenter';
import treeStyles from './tree.module.scss';
import cssStyles from './PermissionCenter.module.scss';

// variables
interface Props {
  state: any;
  props;
  userEntry: string;
}
let removeUserGroupsArray = [];
let addUserGroupsArray = [];
const responseErrorStatusArray = [401, 403, 407, 507];

// main function
const UserCard: React.FC<Props> = (({state, props, userEntry}) => {
  
  try {

    // ------------ api calls --------------

    const _spApiGet = async (url: string): Promise<object> => {
  
      const clientOptions: ISPHttpClientOptions = {
        headers: new Headers(),
        method: 'GET',
        mode: 'cors'
      };
      try {
        const response = await props.spHttpClient.get(url, SPHttpClient.configurations.v1, clientOptions);
        const responseJson = await response.json();
        responseJson['status'] = response.status;
        if (!responseJson.value) {
          responseJson['value'] = [];
        }
        return responseJson;
      } 
      catch (error) {
        if (props.config.logErrors) {console.log(error);}
        if (props.config.throwErrors) {throw error;}
        error['value'] = [];
        error['status'] = "error";
        return error;
      }
    };
    
    const _graphApiGet = async (url:string): Promise<any> => {
      return new Promise<any> (
        (resolve) => {
          props.context.msGraphClientFactory.getClient().then(
            (client: MSGraphClient): any => {
              client.api(url).get(
                (error, response: any) => {
                  if (response) {
                    resolve(response);
                  } else
                  if (error) {
                    resolve(error);
                  }
                }
              );
            }
          )
          // catch error for getClient
          .catch(
            (error) => {
              if (props.config.logErrors) {console.log(error);}
              resolve(error);
            }
          );
        }
      );
    } ;
    
    // ------------ user foto ------------

    const currentUserFotoUrl = useRef('');
    const [userFotoUrl, setUserFotoUrl] = useState('');

    const _getFotoUrl = async () => {
      const url = props.siteCollectionURL + `/_api/sp.userprofiles.peoplemanager/getpropertiesfor(AccountName=@v)?@v='i%3A0%23.f%7Cmembership%7C${state.users[userEntry].principalName}'&$select=PictureUrl`;
      const clientOptions: ISPHttpClientOptions = {
        headers: {'Access-Control-Allow-Origin': "*" },
        method: 'GET',
        mode: 'cors'
      };
      try {
        const response = await props.spHttpClient.get(url, SPHttpClient.configurations.v1, clientOptions);
        const responseJson = await response.json();
        if (props.config.logComponentVars) {console.log(response, responseJson);}
        // if user is external, he has no property PictureUrl. to tell the code that we executed _getFotoUrl for this user, set PictureUrl = null
        if (responseJson.PictureUrl == undefined) {responseJson.PictureUrl = null;}
        return responseJson.PictureUrl;
      } catch (error) {
        if (props.config.logErrors) {console.log(error);}
        if (props.config.throwErrors) {throw error;}
      }
    };
    // if userFotoUrl contains no url
    // check if since load of webpart, we already got the url once, so it would be saved in props.userAndFoto
    if (!userFotoUrl) {
      // if there is an url in props.userAndFoto, then setUserFotoUrl
      if (props.userAndFoto[userEntry]) {
        if (props.config.logComponentVars) {console.log("props.userAndFoto[userEntry] exists", props.userAndFoto[userEntry]);}
        setUserFotoUrl (props.userAndFoto[userEntry]);
      // if there is no url in props.userAndFoto
      } else {
        // if props.userAndFoto would be null, we would have tried to get the url once since loading of web part, because if there is no picture for this user, the response is null
        if (props.userAndFoto[userEntry]!==null) {
          if (props.config.logComponentVars) {console.log("props.userAndFoto[userEntry] empty. ");}
          _getFotoUrl().then(
            (url) => {
              props.userAndFoto[userEntry] = url;
              currentUserFotoUrl.current = url;
              setUserFotoUrl(currentUserFotoUrl.current);
              if (props.config.logComponentVars) {console.log("props.userAndFoto[userEntry] updated", url);}
            }
          );
        }
      }
    }
    
    // ----------- Group nesting events ----------

    // open Sp group
    const _openSpGroup = (groupEntry, event) => {
      if (props.config.logComponentVars) {console.log("event.currentTarget.className: ", event.currentTarget.className);}
      if (event.currentTarget.className.includes('nestedGroup') // fix the bug: outer click
        && (groupEntry !== "spGroup1") // leave admins out
        ) { 
        // if Access given directly group
        if (state.spGroups[groupEntry].groupName === 'Access given directly') {
          window.open(props.siteCollectionURL + `/_layouts/15/user.aspx`);
        } else {
          window.open(props.siteCollectionURL + `/_layouts/15/people.aspx?MembershipGroupId=${state.spGroups[groupEntry].id}`);
        }
      }
    };
    // open Azure group
    const _openAzureGroup = (groupEntry, event) => {
      if (props.config.logComponentVars) {console.log("event.currentTarget.className: ", event.currentTarget.className);}
      if (event.currentTarget.className.includes('nestedGroup')) { // fix the bug: outer click
        let role = "Members";
        let groupID = state.azureGroups[groupEntry].id;
        if (state.azureGroups[groupEntry].id.length > 36) {
          groupID = state.azureGroups[groupEntry].id.substring(0,36);
          role = "Owners";
        }
        window.open(`https://portal.azure.com/#blade/Microsoft_AAD_IAM/GroupDetailsMenuBlade/${role}/groupId/${groupID}`);
      }
    };
    
    // open user in Azure portal
    const _openUserInAzure = async (event) => {
      if (props.config.logComponentVars) {console.log("event.currentTarget.className: ", event.currentTarget.className);}
      if (event.currentTarget.className.includes("nestedGroup")) { // fix the bug: outer click
        if (state.users[userEntry].principalName.includes("@")) {
          // if state has azureId of user, open user in Azure
          window.open(`https://portal.azure.com/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/${state.users[userEntry].principalName}`);
        }
      }
    };
    
    // ------------- elements for render -----------------
    
    // some vars
    let userMembershipArray = [];
    const [userMembershipArrayState, setUserMembershipArrayState] = useState([]);
    const [membershipCheckboxChanged, setMembershipCheckboxChanged] = useState(false);
    const [membershipChanged, setMembershipChanged] = useState(false);
    const [changeMembershipExecuted, setChangeMembershipExecuted] = useState(false);
    
    const actionButtonStyles: IButtonStyles = {
      root: {
        width: "-webkit-fill-available",
        height: "-webkit-fill-available",
        margin: "0.1875rem 0",
        padding: "0.4375rem 1rem",
        maxWidth: "13.125rem"
      },
    };

    const GroupElement = (groupEntry) => {
      // return JSX element for group
      return (
        // for sp group
        groupEntry.startsWith("sp") ? 
          <div style={{display: "flex"}}>
            {/* group bubble */}
            <div 
              className={ `${cssStyles.nestedGroup} ${props.config.showCardUserLinks ? cssStyles.spGroup : cssStyles.spGroupInvalid}` } 
              onClick = { event =>
                props.config.showCardUserLinks
                && state.spGroups[groupEntry].displayName === 'Site Admins'
                  ? window.open(props.siteCollectionURL + '/_layouts/15/user.aspx')
                  : _openSpGroup(groupEntry, event)
              }
              title = {
                props.config.showCardUserLinks && (
                  state.spGroups[groupEntry].groupName === 'Access given directly' 
                  ? 'Open classic permissions page' 
                  : state.spGroups[groupEntry].displayName === 'Site Admins'
                    ? 'Open classic permissions page'
                    : `${state.spGroups[groupEntry].type}: Show in SharePoint`
                )
              }
              >
              {(state.spGroups[groupEntry].displayName!=='Access given directly' && state.spGroups[groupEntry].displayName!=='Site Admins') && (
                <div className={ cssStyles.type }> {state.spGroups[groupEntry].typeShort}:  </div>
              )}
              <div className={ cssStyles.name }> {state.spGroups[groupEntry].groupName} </div>
            </div>

            {/* permission level */}
            <div className={cssStyles.permissions}>
              {
                (state.spGroups[groupEntry].groupName === "Access given directly") 
                ? '(' + state.users[userEntry].permissionLevelDirectAccess.join(', ') + ')'
                : state.spGroups[groupEntry].permissionLevel && state.spGroups[groupEntry].permissionLevel[0]
                  ? '(' + state.spGroups[groupEntry].permissionLevel.join(', ') + ')' 
                  : null
              }
            </div>
          </div>
        // for azure group
        : groupEntry.startsWith("azure") ? 
          <div 
            className={`
              ${cssStyles.nestedGroup} 
              ${state.azureGroups[groupEntry].type.short==="M365" 
                ? props.config.showCardUserLinks
                  ? cssStyles.m365Group 
                  : cssStyles.m365GroupInvalid
                : props.config.showCardUserLinks
                  ? cssStyles.otherAzureGroup
                  : cssStyles.otherAzureGroupInvalid
              }
            `} 
            onClick={(event)=>
              props.config.showCardUserLinks
              && _openAzureGroup(groupEntry, event)
            } 
            title = {props.config.showCardUserLinks && `${state.azureGroups[groupEntry].type.long}: Show in Azure Portal`}
            >
            <div className={ cssStyles.type }> {state.azureGroups[groupEntry].type.short}:  </div>
            <div className={ cssStyles.name }> {state.azureGroups[groupEntry].name} </div>
          </div>
        : // for user 
          <div 
            className={ `${cssStyles.nestedGroup} ${props.config.showCardUserLinks ? cssStyles.user : cssStyles.userInvalid }` } 
            onClick={(event)=> props.config.showCardUserLinks && _openUserInAzure(event)}
            title = {props.config.showCardUserLinks && 'Show user in Azure Portal'}
            >
            {state.users[userEntry].name}
        </div>
      );
    }; 

    const _groupItem = (groupItem) => {
      return (<>
        {/*parent group*/
          GroupElement(groupItem.name)
        }
        {/*if has children => group, else => nothing*/
          groupItem.children[0] && (<ul>{groupItem.children.map(ChildGroupItem => <li>{_groupItem(ChildGroupItem)}</li>)}</ul>)
        }
      </>);
    };

    // prepare group nesting Jsx
    let groupNestingElement; 
    if (state.isGroupsLoading) {
      groupNestingElement = <div className={cssStyles.placeholderItem}> loading </div>;
    } else {
      groupNestingElement = state.users[userEntry].groupNesting.children.map(
        groupNestingBranchItem => {
          return <div className={treeStyles.tree}>
            {_groupItem(groupNestingBranchItem)}
          </div>;
        }
      );
    }
    
    // ---------- Button events ---------------

    // ----- open dialog: change group membership

    // modal properties from fluent ui 
    const theme = getTheme();
    const contentStyles = mergeStyleSets({
      container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
        padding: "1.25rem"
      },
      header: [
        // eslint-disable-next-line deprecation/deprecation
        theme.fonts.xLargePlus,
        {
          flex: '0.0625rem 0.0625rem auto',
          borderTop: `0.25rem solid ${theme.palette.themePrimary}`,
          display: 'flex',
          alignItems: 'center',
          fontWeight: FontWeights.semibold,
          padding: "0.75rem 0",
        },
      ],
      body: {
        flex: '.25rem .25rem auto',
        padding: '0 1.5rem 1.5rem 1.5rem',
        overflow: 'hidden',
        selectors: {
          p: { margin: '0.875rem 0' },
          'p:first-child': { marginTop: 0 },
          'p:last-child': { marginBottom: 0 },
        },
      },
    });
    const dragOptions: IDragOptions = {
      moveMenuItemText: 'Move',
      closeMenuItemText: 'Close',
      menu: ContextualMenu,
    };
    const iconButtonStyles = {
      root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '.25rem',
        marginRight: '.125rem',
      },
      rootHovered: {
        color: theme.palette.neutralDark,
      },
    };
    const cancelIcon: IIconProps = { iconName: 'Cancel' };

    const [isModalOpen, { setTrue: showModal, setFalse: hideModal }] = useBoolean(false);
    const [hideMessage_changeGroupMembership, { toggle: toggleHideMessage_changeGroupMembership }] = useBoolean(true);

    // logic for change group mebership

    const _getGroupChildren = (parentGroup) => {
      
      parentGroup.children.forEach( childItem => {
        if (!childItem.name.startsWith('user')) {
          _getGroupChildren(childItem);
        } else {
          userMembershipArray.push(parentGroup.name);
        }
      });
    };
      
    const _getUserMembership = () => {
      state.users[userEntry].groupNesting.children.forEach(
        spGroupItem => _getGroupChildren(spGroupItem) // get last group of group nesting branch in user to delete user from it
      );
      // delete group doubles from array
      // userMembershipArray = [...new Set(userMembershipArray)];
      // instead for IE11
      userMembershipArray = userMembershipArray.filter((v, i, a) => a.indexOf(v) === i);


      setUserMembershipArrayState(userMembershipArray);
    };
    
    const _openDialog_changeGroupMembership = (event) => {
      // reset arrays
      removeUserGroupsArray = [];
      addUserGroupsArray = [];
      if (event.currentTarget.className.includes("myClick")) { // fix the bug: outer click
        // get actual user membership
        _getUserMembership();
        // show dialog (modal)
        showModal();
      }
    };
    
    const _addORRemoveAdmin = async (userLoginName, userId, isRemove) => {
      
      // if no userId, get sp Id
      if (!userId) {
        // first, try to get sp user id from sp api. if user has a profile in sp, reponse will be successful
        const response = await _spApiGet (`${props.siteCollectionURL}/_api/web/siteusers(@v)?@v='${encodeURIComponent(userLoginName)}'&$select=Id`);
        // if response has id, add Id to state.users
        if (response["Id"]) {
          userId = response["Id"];
          state.users[userEntry].spId = userId;
        }
        // if user has no user profile in sp, response will be an error. so call "ensureuser" to create his sp user profile.
        else {
          // for add admin: if not in site collection, add user to site collection
          if (!isRemove) {

            // parameter
            let ensureAdminUrl = props.siteCollectionURL + `/_api/web/ensureuser`;
            const ensureAdminOpts = {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': '',
              },
              body: JSON.stringify({
                'logonName': userLoginName
              }),
              mode: 'cors'
            };

            // http request
            try {
              // call ensureuser
              await props.spHttpClient.post(ensureAdminUrl, SPHttpClient.configurations.v1, ensureAdminOpts);
              // get his user id
              const responseGetUserId = await _spApiGet (`${props.siteCollectionURL}/_api/web/siteusers(@v)?@v='${encodeURIComponent(userLoginName)}'&$select=Id`);
              // if response has id, add Id to state.users
              userId = responseGetUserId["Id"];
              if (userId) {
                state.users[userEntry].spId = userId;
              }
            } 
            catch (error) {
              if (props.config.logErrors) {console.log(error);}
              if (props.config.throwErrors) {throw error;}
              error['value'] = [];
              error['status'] = "error";
            }
          }
        }
      }
      
      // add or remove admin
      if (userId) {
        // parameter
        let requestUrl = props.siteCollectionURL + `/_api/web/GetUserById(${userId})`;
        let dataToPost = JSON.stringify({
          '__metadata': { 'type': 'SP.User' },
          'IsSiteAdmin': isRemove ? 'false' : 'true',
        });
        const spOpts = {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=verbose',
            'odata-version': '',
            "X-HTTP-Method": "MERGE"
          },
          body: dataToPost,
          mode: 'cors'
        };

        // http request
        try {
          const response = await props.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, spOpts);
          return response.status;
        } 
        catch (error) {
          if (props.config.logErrors) {console.log(error);}
          if (props.config.throwErrors) {throw error;}
          return error;
        }
      } 
      else {
        return "error";
      }
    };

    const _addAndRemoveUserFromSpGroup = async (userLoginName, groupId, isRemove) => {
      // parameter
      // for add user
      let requestUrl = props.siteCollectionURL + `/_api/web/sitegroups/getbyid(${groupId})/users`;
      let dataToPost = JSON.stringify({
        '__metadata': { 'type': 'SP.User' },
        'LoginName': userLoginName
      });
      // for remove user
      if (isRemove) {requestUrl += "/removeByLoginName";}
      if (isRemove) {
        dataToPost = JSON.stringify({
          'loginName': userLoginName
        });
      }
      // general
      const spOpts = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': ''
        },
        body: dataToPost,
        mode: 'cors'
      };
    
      // http request
      try {
        const response = await props.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, spOpts);
        return response.status;
      } catch (error) {
        if (props.config.logErrors) {console.log(error);}
        if (props.config.throwErrors) {throw error;}
      }
    };
    
    const _addOrRemoveUserFromAzureGroup = (userAzureId, azureGroupEntry, isRemove): Promise<number> => {

      const azureGroupId = state.azureGroups[azureGroupEntry].id;
    
      return props.context.msGraphClientFactory.getClient()
      .then(
        async (client: MSGraphClient) => {
          // remove member
          if (isRemove) {
            let url = `/groups/${azureGroupId}/members/${userAzureId}/$ref`;
            // for M365 Owners
            if (state.azureGroups[azureGroupEntry].id.includes("_o")) {
              url = `/groups/${azureGroupId.substring(0,36)}/owners/${userAzureId}/$ref`;
            }
            await client.api(url).delete();
            return 200;
          
          // add member
          } else {
            let url = `/groups/${azureGroupId}/members/$ref`;
            // for M365 Owners
            if (state.azureGroups[azureGroupEntry].id.includes("_o")) {
              url = `/groups/${azureGroupId.substring(0,36)}/owners/$ref`;
            }
            const directoryObject = {
              '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${userAzureId}`
            };
            await client.api(url).post(directoryObject);
            return 200;
          }
        }
      )
      .catch(
        error => {
          if (props.config.logErrors) {console.log(error);}
          return error.statusCode;
        }
      );
    };

    // change group membership when click on button
    let changeGroupMembershipFeedback = {};
    const [changeGroupMembershipFeedbackState, setChangeGroupMembershipFeedbackState] = useState(changeGroupMembershipFeedback);

    const _changeGroupMembership = async () => {
      setChangeMembershipExecuted(true);
      // take membership from _getUserMembership in _openDialog_changeGroupMembership
      if (props.config.logComponentVars) {console.log("userMembershipArrayState on change membership button", userMembershipArrayState);}
      if (props.config.logComponentVars) {console.log("add user to groups:", addUserGroupsArray);}
      if (props.config.logComponentVars) {console.log("remove user from groups:", removeUserGroupsArray);}

      // some variables
      const userLoginName = `i:0#.f|membership|${state.users[userEntry].principalName}`;
      let userAzureId = state.users[userEntry].azureId;
      if (
        (
          addUserGroupsArray.some(item=>item.startsWith('azure')) 
          || removeUserGroupsArray.some(item=>item.startsWith('azure'))
        )
        && !userAzureId
      ) {
        const response = await _graphApiGet(`/users/${state.users[userEntry].principalName}`);
        // if error
        if (response.statusCode) {
          if (responseErrorStatusArray.includes(response.statusCode)) {
            userAzureId = "no access";
          } else {
            userAzureId = "error";
          }
        // if response ok
        } else {
          userAzureId = response['id'];
          state.users[userEntry].azureId = userAzureId;
        }
      }
      

      // add/remove user
      Promise.all([

        // remove user
        Promise.all(
          removeUserGroupsArray.map(
            async removeGroupItem=>{
              // prepare feedback
              let status = "";
              let removeResponse;

              // from sp groups
              if (removeGroupItem.startsWith('sp')) {
                // admins
                if(removeGroupItem==="spGroup1") {
                  removeResponse = await _addORRemoveAdmin (userLoginName, state.users[userEntry].spId, true);
                  // if success / error: if user has no access, error status code will be 500, so it is not possible to display the correct reaosn
                  if (removeResponse.toString().startsWith('2')) {
                    status="success";
                  } else if (responseErrorStatusArray.includes(removeResponse)) {
                    status = "accessDenied";
                  }
                  else status="error";
                }
                // normal sp groups
                else {
                  const groupId = state.spGroups[removeGroupItem].id;
                  removeResponse = await _addAndRemoveUserFromSpGroup(userLoginName, groupId, true);
                  // if success / error
                  if ((removeResponse==409) || removeResponse.toString().startsWith('2')) {
                    status = "success";
                  } else if (responseErrorStatusArray.includes(removeResponse)) {
                    status = "accessDenied";
                  }
                  else status = "error";
                }

              // from azure groups
              } else {
                if (userAzureId == "no access") {
                  status = "accessDenied";
                }
                else if (userAzureId == "error") {
                  status = "error";
                }
                else {
                  removeResponse = await _addOrRemoveUserFromAzureGroup(userAzureId, removeGroupItem, true);
                  // if success
                  if (removeResponse == 200) {
                    status = "success";
                  } 
                  // if access denied error
                  else if (responseErrorStatusArray.includes(removeResponse)) {
                    status = "accessDenied";
                  }
                  // if other error
                  else status = "error";
                }
              }
              changeGroupMembershipFeedback[removeGroupItem] = {status: status};
              if (props.config.logComponentVars) {console.log(removeGroupItem, " removeUserResponse: ", removeResponse);}
              return removeResponse;
            }
          )
        ),
        
        // add user
        Promise.all(
          addUserGroupsArray.map(
            async addGroupItem=>{
              // prepare feedback
              let status = '';
              let addResponse;

              // to sp groups
              if (addGroupItem.startsWith('sp')) {
                // admins
                if (addGroupItem==="spGroup1") {
                  addResponse = await _addORRemoveAdmin (userLoginName, state.users[userEntry].spId, false);
                  // if success / error: if user has no access, error status code will be 500, so it is not possible to display the correct reaosn
                  if (addResponse.toString().startsWith('2')) {
                    status="success";
                  } else if (responseErrorStatusArray.includes(addResponse)) {
                    status = "accessDenied";
                  }
                  else status="error";
                // normal sp groups
                } else {
                  const groupId = state.spGroups[addGroupItem].id;
                  addResponse = await _addAndRemoveUserFromSpGroup(userLoginName, groupId, false);
                  if ((addResponse==409) || addResponse.toString().startsWith('2')) {
                    status = "success";
                  } else if ( responseErrorStatusArray.includes(addResponse)) {
                    status = "accessDenied";
                  }
                  else status = "error";
                }

              // to azure groups
              } else {
                if (userAzureId == "no access") {
                  status = "accessDenied";
                }
                else if (userAzureId == "error") {
                  status = "error";
                }
                else {
                  addResponse = await _addOrRemoveUserFromAzureGroup(userAzureId, addGroupItem, false);
                  if (addResponse == 200) {
                    status = "success";
                  } else if (responseErrorStatusArray.includes(addResponse)) {
                    status = "accessDenied";
                  }
                  else status = "error" ;
                }
              }
              if (props.config.logComponentVars) {console.log(addGroupItem, " addUserResponse:", addResponse);}

              changeGroupMembershipFeedback[addGroupItem] = {status: status};
              return addResponse;
            }
          )
        )
      ]).then(()=>{
          // when done, show message
          if (props.config.logComponentVars) {console.log("changeGroupMembershipFeedback", changeGroupMembershipFeedback);}
          setChangeGroupMembershipFeedbackState(changeGroupMembershipFeedback);
          setMembershipChanged(true);
      });
    };

    // get checkbox input
    const _onChangeCheckbox = (event) => {
      const groupEntry = event.currentTarget.getAttribute("data-groupentry");
      // the checked attribute of the check box input behaves contrary: it is true if box is unchecked

      // if unchecked
      if (event.currentTarget.checked) {
        // add groupEntry to userMembershipArrayState
        setUserMembershipArrayState(userMembershipArrayState.concat(groupEntry));
        // if exists in removeUserGroupsArray
        if (removeUserGroupsArray.includes(groupEntry)) {
          // remove groupEntry from removeUserGroupsArray
          removeUserGroupsArray = removeUserGroupsArray.filter(group=>group!==groupEntry);
        // if doesn't exist in removeUserGroupsArray
        } else {
          // add groupEntry to addUserGroupsArray
          addUserGroupsArray.push(groupEntry);
        }

      // if checked
      } else {
        // remove groupEntry from userMembershipArrayState
        setUserMembershipArrayState(userMembershipArrayState.filter(group=>group!==groupEntry));
        // if exists in addUserGroupsArray
        if (addUserGroupsArray.includes(groupEntry)) {
          // remove groupEntry from addUserGroupsArray
          addUserGroupsArray = addUserGroupsArray.filter(group=>group!==groupEntry);
        // if doesn't exist in addUserGroupsArray
        } else {
          // add groupEntry to removeUserGroupsArray
          removeUserGroupsArray.push(groupEntry);
        }
      }

      // set membershipCheckboxChanged for disabeling button "Change membership"
      if (addUserGroupsArray[0] || removeUserGroupsArray[0]) {
        setMembershipCheckboxChanged(true);
      } else {
        setMembershipCheckboxChanged(false);
      }
    };

    // ----- Dialog: Delete user from site

    const [hideDialog_deleteUserFromSite, { toggle: toggleHideDialog_deleteUserFromSite }] = useBoolean(true);
    const [hideMessage_deleteUserFromSite, { toggle: toggleHideMessage_deleteUserFromSite }] = useBoolean(true);
    const [deleteUserFromSiteMessageTitle, setDeleteUserFromSiteMessageTitle] = useState('');
    const [deleteUserFromSiteMessage, setDeleteUserFromSiteMessage] = useState('');
    const _openDialog_deleteUserFromSite = (event) => {
      if (event.currentTarget.className.includes("myClick")) { // fix the bug: outer click
        // if user has spId (has membership in SP, else is just member of azure groups)
        if (state.users[userEntry].spId) {
          // show dialog
          toggleHideDialog_deleteUserFromSite();
        } else {
          // show message
          setDeleteUserFromSiteMessageTitle('Note');
          setDeleteUserFromSiteMessage('User has no membership in SharePoint, just in Azure groups.');
          toggleHideMessage_deleteUserFromSite();
        }
      }
    };

    // remove user from sharepoint
    const _deleteUserFromSite = async () => {
      // hide dialog
      toggleHideDialog_deleteUserFromSite();
      // parameter
      const spOpts = {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=verbose',
          'odata-version': '',
          'X-HTTP-Method': 'DELETE'
        },
        mode: 'cors'
      };
      const userSpId = state.users[userEntry].spId;
      const requestUrl = props.siteCollectionURL + `/_api/web/GetUserById(${userSpId})`;
      // http request
      try {
        const response = await props.spHttpClient.post(requestUrl, SPHttpClient.configurations.v1, spOpts);
        // if success
        if (response.ok) {
          if (props.config.logComponentVars) {console.log('_deleteUserFromSite response', response);}
          // show message
          setDeleteUserFromSiteMessageTitle('Success');
          setDeleteUserFromSiteMessage(`${state.users[userEntry].name} removed from all SharePoint groups.`);
          toggleHideMessage_deleteUserFromSite();
        } 
        // if error
        else {
          const responseJson = await response.json();
          if (props.config.logComponentVars) {console.log('response', responseJson);}
          // show message
          setDeleteUserFromSiteMessageTitle('SharePoint Error');
          setDeleteUserFromSiteMessage(responseJson["odata.error"].message.value);
          toggleHideMessage_deleteUserFromSite();
        }
      // if other error
      } catch (error) {
        if (props.config.logErrors) {console.log('_deleteUserFromSite error: ', error);}
        if (props.config.throwErrors) {throw error;}
        // show message
        setDeleteUserFromSiteMessageTitle('Error');
        setDeleteUserFromSiteMessage('User could not be removed.');
        toggleHideMessage_deleteUserFromSite();
      } 
    };
    
    // ----- open classic property page
      
    // Dialog: no sp user, open user in azure portal insted
    const [hideDialog_openUserInAzure, { toggle: toggleHideDialog_openUserInAzure }] = useBoolean(true);
    const _openDialog_openUserInAzure = () => {
      toggleHideDialog_openUserInAzure();
    };

    const _openClassicPropertyPage = async (event) => {
      if (props.config.logComponentVars) {console.log("event.currentTarget.className", event.currentTarget.className);}
      if (event.currentTarget.className.includes("myClick")) { // fix the bug: outer click
        if (state.users[userEntry].spId) {
          // if sp id exists, open classic property page
          window.open(props.siteCollectionURL + `/_layouts/15/userdisp.aspx?ID=${state.users[userEntry].spId}&force=1`);
        } else {
          // if not, get sp Id
          const response = await _spApiGet (`${props.siteCollectionURL}/_api/web/siteusers(@v)?@v='i%3A0%23.f%7Cmembership%7C${encodeURIComponent(state.users[userEntry].principalName)}'&$select=Id`);
          // if response has id, open classic property page
          if (response["Id"]) {
            state.users[userEntry].spId = response["Id"];
            window.open(props.siteCollectionURL + `/_layouts/15/userdisp.aspx?ID=${response["Id"]}&force=1`);
            if (props.config.logComponentVars) {console.log("_openClassicPropertyPage user id response",response["Id"]);}
          } 
          // if not, user does not exist in Sp, so open dialog
          else {_openDialog_openUserInAzure();}
        }
      }
    };

    // ----- open user in Azure portal
    
    const _openUserInAzureClassic = async () => {
      toggleHideDialog_openUserInAzure();
      if (state.users[userEntry].principalName) {
        // if state has azureId of user, open user in Azure
        window.open(`https://portal.azure.com/#blade/Microsoft_AAD_IAM/UserDetailsMenuBlade/Profile/userId/${state.users[userEntry].principalName}`);
      } else {
        alert('Something went wrong. Try it manually.');
      }
    };
    
    //------------------------------- return ---------------------------------
    return (
      <div className={ cssStyles.userCard }>
        
        {/* userCardContainer */}
        <div className={ cssStyles.userCardContainer }>

          {/* info content */}
          <div>

            {props.config.logComponentVars && console.log("UserCard return - currentUserFotoUrl =", currentUserFotoUrl, "userFotoUrl =", userFotoUrl )}
            <div style={{ display:"flex", padding: '0 0'}}>
              <div className = {cssStyles.persona}>
                <Persona
                  showUnknownPersonaCoin={userFotoUrl ? false : true}
                  imageUrl={userFotoUrl}
                  size={PersonaSize.size40}
                />
              </div>
              <div style={{ padding: '0 0'}}>
                <div style={{wordBreak: 'break-word'}}>
                  {state.users[userEntry].email ? state.users[userEntry].email : "No email adress"}
                </div>
                <div style={{wordBreak: 'break-word'}}>
                  {state.users[userEntry].permissionLevel[0] && (`Permission: ${state.users[userEntry].permissionLevel.join(', ')}`)}
                </div>
              </div>
            </div>

            <div className={ cssStyles.dataRow }>
              <span className={ cssStyles.dataValue }>This user is member of the following groups: </span>
            </div>

            {groupNestingElement}

          </div>

          {/* buttons */}
          
          <div className={ cssStyles.groupCardButtonContainer }>
            {props.config.showCardButtons &&
              <DefaultButton 
                className={"myClick"} 
                text="Change membership" 
                styles={actionButtonStyles} 
                href={'#'} 
                onClick={(e)=>{ if (!state.isGroupsLoading) {_openDialog_changeGroupMembership(e);}}}
              />
            }
            {props.config.showCardButtons &&
              <DefaultButton 
                className={"myClick"} 
                text="Delete user from site" 
                styles={actionButtonStyles} 
                href={'#'} 
                onClick={(e)=>{_openDialog_deleteUserFromSite(e);}}
              />
            }
            <DefaultButton 
              className={"myClick"} 
              text="Classic property page" 
              styles={actionButtonStyles} 
              href={'#'} 
              onClick={ (e) => _openClassicPropertyPage(e)}
            />
          </div>
          
          
          {/* Change group membership */}

          <Modal
            isOpen={isModalOpen}
            onDismiss={hideModal}
            isBlocking={false}
            containerClassName={contentStyles.container}
            dragOptions={dragOptions}
            >
            <div style={{margin: '0.625rem'}}>

              <div className={contentStyles.header}>
                <span>Change group membership</span>
                <IconButton
                  styles={iconButtonStyles}
                  iconProps={cancelIcon}
                  ariaLabel="Close popup modal"
                  onClick={hideModal}
                />
              </div>
              <div>User: {state.users[userEntry].name}</div>
              <br/>
              <div style={{display:"flex"}}>
                {/* SharePoint groups */}
                <div className={cssStyles.modalColumn} style={{marginRight:'0.625rem'}}> 
                  <div className={cssStyles.headline} > SharePoint groups:</div>
                  <br/>
                  <div >
                    {Object.keys(state.spGroups).map(
                      (spGroupEntryItem)=>
                      <div style={{display:"flex", marginBottom:'0.625rem'}}>
                        {/* checkbox */}
                        <input 
                          type="checkbox" 
                          data-groupEntry={spGroupEntryItem} 
                          checked={userMembershipArrayState.filter(group=>group===spGroupEntryItem)[0] ? true : false}
                          onClick={(e)=>_onChangeCheckbox(e)}
                          className={cssStyles.modalCheckbox}
                          disabled = {(
                            state.spGroups[spGroupEntryItem].displayName==="Access given directly"
                            || changeMembershipExecuted
                            ) ? true : false}
                        />
                        {/* group name */}
                        <div 
                          style={{ paddingLeft:"0.3125rem"}}
                          className={state.spGroups[spGroupEntryItem].displayName==="Access given directly" && cssStyles.openLink}
                          onClick={state.spGroups[spGroupEntryItem].displayName==="Access given directly" && (()=>window.open(props.siteCollectionURL + `/_layouts/15/user.aspx`))}
                          title = {state.spGroups[spGroupEntryItem].displayName==="Access given directly" && "To manage members open 'classic permissions page'"}
                          >
                          {state.spGroups[spGroupEntryItem].displayName}
                        </div>

                        {/* Feedback for user if change success */}
                        <div className={cssStyles.iconContainer}>
                          {
                          changeGroupMembershipFeedbackState[spGroupEntryItem] 
                          ? changeGroupMembershipFeedbackState[spGroupEntryItem].status == "success" 
                            ? <Icon iconName={"CheckMark"} className={`${cssStyles.successIcon}`} title='Change successful' /> 
                            : changeGroupMembershipFeedbackState[spGroupEntryItem].status == "accessDenied"
                              ? <Icon iconName={"Cancel"} className={`${cssStyles.errorIcon}`} title="No permission to change membership" /> 
                              : <Icon iconName={"Help"} className={`${cssStyles.errorIcon}`} title="An error occurred while changing membership."/> 
                          : null
                          }
                        </div>

                      </div>
                    )}
                  </div>
                </div>
                
                {/* Azure groups */}
                <div className={cssStyles.modalColumn}> 
                  <div className={cssStyles.headline} > 
                    {state.azureGroupArraySorted[0]
                      && 'Azure groups: '
                    }
                  </div>
                  <br/>
                  <div className={cssStyles.modalColumn} >
                    {state.azureGroupArraySorted.map(
                      (azureGroupObjectItem)=>
                      <div style={{display:"flex", marginBottom:'0.625rem'}}>
                        {/* checkbox */}
                        <input 
                          type="checkbox" 
                          data-groupEntry={azureGroupObjectItem.key} 
                          checked={userMembershipArrayState.filter(groupEntry=>groupEntry===azureGroupObjectItem.key)[0] ? true : false}
                          onClick={(e)=>_onChangeCheckbox(e)}
                          className={cssStyles.modalCheckbox}
                          disabled={(
                            azureGroupObjectItem.type.long.startsWith("Dist") 
                            || azureGroupObjectItem.type.long.startsWith("Mail")
                            || changeMembershipExecuted
                            ) ? true : false}
                        />
                        {/* group name */}
                        <div 
                          style={{ paddingLeft:"0.3125rem"}}
                          onClick={(
                            azureGroupObjectItem.type.long.startsWith("Dist") || 
                            azureGroupObjectItem.type.long.startsWith("Mail")
                            ) && (()=>window.open(`https://admin.microsoft.com/AdminPortal/Home#/groups/:/GroupDetails/${azureGroupObjectItem.id}/2`))
                          }
                          title={ (
                            azureGroupObjectItem.type.long.startsWith("Dist") || 
                            azureGroupObjectItem.type.long.startsWith("Mail")
                            ) && "To manage members, open Microsoft 365 admin center." 
                          }
                          className={(
                            azureGroupObjectItem.type.long.startsWith("Dist") || 
                            azureGroupObjectItem.type.long.startsWith("Mail")
                            )
                            && cssStyles.openLink
                          }
                          >
                          {azureGroupObjectItem.name}
                        </div>

                        {/* Feedback for user if change success or error */}
                        <div className={cssStyles.iconContainer}>
                          {
                          changeGroupMembershipFeedbackState[azureGroupObjectItem.key] 
                          ? changeGroupMembershipFeedbackState[azureGroupObjectItem.key].status == "success" 
                            ? <Icon iconName={"CheckMark"} className={`${cssStyles.successIcon}`} title='Change successful' /> 
                            : changeGroupMembershipFeedbackState[azureGroupObjectItem.key].status == "accessDenied"
                              ? <Icon iconName={"Cancel"} className={`${cssStyles.errorIcon}`} title="No permission to change membership" /> 
                              : <Icon iconName={"Help"} className={`${cssStyles.errorIcon}`} title="An error occurred while changing membership."/> 
                          : null
                          }
                        </div>
                      </div>
                    )}
                  </div>
                </div>

              </div>

              <DialogFooter>
                {
                  membershipChanged
                  ? <PrimaryButton onClick={_reload} text="Reload web part" />
                  : <PrimaryButton onClick={_changeGroupMembership} text={changeMembershipExecuted ? "Changing membership" : "Change membership"} disabled={!membershipCheckboxChanged || changeMembershipExecuted}/>
                }
                <DefaultButton onClick={hideModal} text="Close" />
              </DialogFooter>
              
            </div>

          </Modal>

          {/* Delete user from site */}
          
          <Dialog
            hidden={hideDialog_deleteUserFromSite}
            onDismiss={toggleHideDialog_deleteUserFromSite}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Warning',
              closeButtonAriaLabel: 'Close'
            }}
          >
            <div>
              Do you want to delete {state.users[userEntry].name} from all SharePoint groups?
            </div>
            <DialogFooter>
              <PrimaryButton onClick={_deleteUserFromSite} text="Yes" />
              <DefaultButton onClick={toggleHideDialog_deleteUserFromSite} text="No" />
            </DialogFooter>
          </Dialog>

          <Dialog
            hidden={hideMessage_deleteUserFromSite}
            onDismiss={toggleHideMessage_deleteUserFromSite}
            dialogContentProps={{
              type: DialogType.normal,
              title: deleteUserFromSiteMessageTitle,
              closeButtonAriaLabel: 'Close',
            }}
            >
            <div> {deleteUserFromSiteMessage} </div>
            <PrimaryButton onClick={toggleHideMessage_deleteUserFromSite} text="Close" style={{float:"right", margin:"1.25rem 0"}}/>
          </Dialog>
          
          {/* Open user in Azure */}
          <Dialog
            hidden={hideDialog_openUserInAzure}
            onDismiss={toggleHideDialog_openUserInAzure}
            dialogContentProps={{
              type: DialogType.normal,
              title: 'Note',
              closeButtonAriaLabel: 'Close',
            }}
            >
            <div>{state.users[userEntry].name} has no property page in SharePoint.</div>
            <br/>
            <PrimaryButton onClick={_openUserInAzureClassic} text="Show user in Azure Portal" style={{width:'18.125rem'}}/>
            <DefaultButton onClick={toggleHideDialog_openUserInAzure} text="Close" style={{width:'18.125rem'}} />
          </Dialog>

        </div>
      </div>
    );
  }
  catch (error) {
    if (props.config.logErrors) {console.log(error);}
    if (props.config.throwErrors) {throw error;}
  }
});
export default memo(UserCard);
