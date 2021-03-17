// import react
import * as React from 'react';
import { initializeIcons } from '@uifabric/icons';
import { IconButton } from '@fluentui/react/lib/Button';

// import APIs
import { SPHttpClient, ISPHttpClientOptions} from '@microsoft/sp-http';
import { MSGraphClient } from '@microsoft/sp-http';

// import components
import { IPermissionCenterProps } from './IPermissionCenterProps';
import cssStyles from './PermissionCenter.module.scss';
import SpMenu from './SpMenu';
import SpGroupContainer from './SpGroupContainer';
import AllUsers from './AllUsers';

// set variables
let userCount = 1;
let azureGroupCount = 1;
let spGroupCount = 5;
export let _reload;
const logLastState = true;
const showLogs = false;
const logErrors = false;
const showLogsOfUserDouble = false;
const showLogsOfAzureDouble = false;
const filterM365OwnersGroupFromSpDefaultOwnersGroup = false;
const regExMail = /^(([^<>()\[\]\.,;:\s@\"]+(\.[^<>()\[\]\.,;:\s@\"]+)*)|(\".+\"))@(([^<>()[\]\.,;:\s@\"]+\.)+[^<>()[\]\.,;:\s@\"]{2,})$/i;
const errorStatusAccessDeniedArray = [401, 403, 407, 507];

// main class
export default class PermissionCenter extends React.Component<IPermissionCenterProps, {}> {
  
  // ------------- set properties ---------------

  // set initial state object
  private initialState = {
    users: {},
    spGroups: {
      spGroup1: {
        type: "SharePoint",
        displayName: "Site Admins",
        groupName: "Site Admins",
        permissionLevel: ["Administrator"],
        users: [null]
      },
      spGroup2: {
        type: "SharePoint",
        displayName: "Site Owners",
        permissionLevel: ["Full control"],
        defaultGroup: "Default owners group of site",
        users: [null]
      },
      spGroup3: {
        type: "SharePoint",
        displayName: "Site Members",
        permissionLevel: ["Edit"],
        defaultGroup: "Default members group of site",
        users: [null]
      },
      spGroup4: {
        type: "SharePoint",
        displayName: "Site Visitors",
        permissionLevel: ["Read"],
        defaultGroup: "Default members group of site",
        users: [null]
      },
    },
    azureGroups: {},
    azureGroupArraySorted: [],
    selectedTab: 'Groups',
    isGroupsLoading: true,
    hiddenGroupsExist: false,
    mode: this.props.mode,
  };

  // set state of class
  public state = this.initialState;
  
  // define properties of class
  private users = this.state.users;
  private spGroups = this.state.spGroups;
  private azureGroups= this.state.azureGroups;
  private azureGroupArraySorted = this.state.azureGroupArraySorted;
  private sitePermissionLevels;
  private allPermResponse;
  private hiddenGroupsExist = false;
  
  // -------------- basic api calls -------------

  // get data from SharePoint REST Api
  private async _spApiGet (url: string): Promise<object> {
    const clientOptions: ISPHttpClientOptions = {
      headers: new Headers(),
      method: 'GET',
      mode: 'cors'
    };
    try {
      const response = await this.props.spHttpClient.get(url, SPHttpClient.configurations.v1, clientOptions);
      const responseJson = await response.json();
      responseJson['status'] = response.status;
      if (!responseJson.value) {
        responseJson['value'] = [];
      }
      return responseJson;
    } 
    catch (error) {
      if (logErrors) {console.log(error);}
      if (this.props.throwErrors) {throw error;}
      error['value'] = [];
      error['status'] = "error";
      return error;
    }
  } 

  // post data to SharePoint REST Api
  private async _spApiPost (url: string): Promise<object> {
    const requestHeaders: Headers = new Headers();  
    requestHeaders.append('Content-type', 'application/json'); 
    requestHeaders.append('Accept', 'application/json'); 
    requestHeaders.append('Authorization', 'Bearer'); 
    
    const clientOptions: ISPHttpClientOptions = {
      headers: requestHeaders,
      method: 'POST',
      mode: 'cors',
    };
    try {
      const response = await this.props.spHttpClient.post(url, SPHttpClient.configurations.v1, clientOptions);
      const responseJson = await response.json();
      responseJson['status'] = response.status;
      if (!responseJson.value) {
        responseJson['value'] = [];
      }
      return responseJson;
    } 
    catch (error) {
      if (logErrors) {console.log(error);}
      if (this.props.throwErrors) {throw error;}
      error['value'] = [];
      error['status'] = "error";
      return error;
    }
  } 
  
  // get data from Microsoft Graph Api
  private _graphApiGet (url: string): Promise<any> {
    return new Promise<any> (
      (resolve, reject) => {
        this.props.context.msGraphClientFactory.getClient()
        .then(
          (client: MSGraphClient): any => {
            client.api(url).get(
              async (error, response: any) => {
                if (response) {
                  resolve(response);
                } else if (error) {
                  resolve(error);
                }
              }
            );
          }
        )
        // catch error for getClient
        .catch(
          (error) => {
            if (logErrors) {console.log(url, error);}
            resolve(error);
          }
        );
      }
    );
  }  

  // -------------- secondary methods ---------------


  // create object for azure group, get its members recursively and write them into memory
  private async _buildAzureGroupAndGetChildItemsRecursive (azureGroupName: string, azureGroupId: string, spGroupEntry: string, groupNestingBranch: string[], parentAzureGroupEntry: string, parentAzureGroupType, firstLevelAzureGroupPermissionLevel:string[]): Promise<any> {

    //create new azureGroup in this.azureGroups
    const azureGroupEntry = `azureGroup${azureGroupCount}`; azureGroupCount += 1;
    this.azureGroups[azureGroupEntry] = new Object;
    const newGroupNestingBranch = groupNestingBranch.concat(azureGroupEntry);
    //write properties
    this.azureGroups[azureGroupEntry]["name"] = azureGroupName;
    this.azureGroups[azureGroupEntry]["id"] = azureGroupId;
    this.azureGroups[azureGroupEntry]["type"] = parentAzureGroupType;

    // prepare get members
    let url = "/groups/" + azureGroupId + "/members?$top=999" + "&$c=" + Date.now().toString(); // "&$c=" + Date.now().toString() for IE11 because otherwise it takes the cached request
    // if M365OwnersGroup
    if (azureGroupId.length > 36) {
      const azureGroupIdCut = azureGroupId.substring(0,36); // cut "_o" from end of id of M365OwnersGroup
      url = "/groups/" + azureGroupIdCut + "/owners" + "?$c=" + Date.now().toString(); // "&$c=" + Date.now().toString() for IE11 because otherwise it takes the cached request
    }
    // get members
    let response = await this._graphApiGet(url);
    let newResponse = response;
    let loopCount = 0;
    // server side paging: if more pages
    while (newResponse["@odata.nextLink"] && loopCount < 100) {
      loopCount +=1;
      url = newResponse["@odata.nextLink"];
      newResponse = await this._graphApiGet(url);
      response.value = response.value.concat(newResponse.value);
    }
    if (showLogs) { console.log(`${azureGroupEntry} with name ${azureGroupName} members response:`, response );}
    // if any error just not displaying members
    // if no "response.statusCode", no error occured
    if (!response.statusCode) {
      return await Promise.all (
        response.value.map(
          // for each member             
          async (memberItem) => {
            switch (memberItem['@odata.type']) {
              
              // if azureGroupMember = user
              case "#microsoft.graph.user": {

                //create user in this.users
                const userEntry = `user${userCount}`; userCount += 1;
                this.users[userEntry] = new Object;
                //write properties of user
                this.users[userEntry]["name"] = memberItem.displayName;
                this.users[userEntry]["azureId"] = memberItem.id;
                this.users[userEntry]["principalName"] = memberItem.userPrincipalName;
                if (this.spGroups[spGroupEntry].groupName !== "Access given directly") {
                  this.users[userEntry]["permissionLevel"] = this.spGroups[spGroupEntry].permissionLevel;
                } else {
                  this.users[userEntry]["permissionLevel"] = firstLevelAzureGroupPermissionLevel;
                  this.users[userEntry]["permissionLevelDirectAccess"] = firstLevelAzureGroupPermissionLevel;
                }
                this.users[userEntry]["groupNesting"] = [newGroupNestingBranch];
                this.users[userEntry]["spGroup"] = spGroupEntry;
                this.users[userEntry]["azureGroup"] = azureGroupEntry + ': ' + azureGroupName;
                // if user has email
                if (memberItem.mail) {
                  this.users[userEntry]["email"] = memberItem.mail;
                  // if not, check if principalName is email
                } else if (regExMail.test(memberItem.userPrincipalName)) {
                  this.users[userEntry]["email"] = memberItem.userPrincipalName;
                  // if neighter, no email
                } else {this.users[userEntry]["email"] = '';}

                //write user into this.spGroups
                this.spGroups[spGroupEntry].users.push(userEntry);
                return userEntry;
              }

              // if azureGroupMember = azure group
              case "#microsoft.graph.group": {
                parentAzureGroupEntry = azureGroupEntry;
                const azureGroupType = this._evalAzureGroupType (memberItem["groupTypes"][0], memberItem["mailEnabled"], memberItem["securityEnabled"]);
                // handle child azure groups
                return await this._buildAzureGroupAndGetChildItemsRecursive(memberItem.displayName, memberItem.id, spGroupEntry, newGroupNestingBranch, parentAzureGroupEntry, azureGroupType, firstLevelAzureGroupPermissionLevel);
              }
              default: return "nix";
            }
          }
        )
      );
    }
  }

  // evaluate type of azure group
  private _evalAzureGroupType (groupTypes, mailEnabled, securityEnabled) {

    let azureGroupType;
    if (groupTypes == "Unified") {azureGroupType = {long: 'Microsoft 365 group', short: 'M365'};}
    else if (!mailEnabled) {azureGroupType = {long: 'Security group', short: 'SEC'};}
    else if (securityEnabled) {azureGroupType = {long: 'Mail-enabled security group', short: 'MSEC'};}
    else {azureGroupType = {long: 'Distribution list', short: 'DL'};}
    return azureGroupType;
  }

  // write members of SharePoint group into memory
  private async _buildSpGroupMembers (spGroupEntry: string, spGroupMembersResponse: any): Promise<any> {
    const groupNestingBranch = [spGroupEntry];

    if (spGroupMembersResponse == "no access") {
      this.spGroups[spGroupEntry].users = ['no access'];
    } 
    else {
      return await Promise.all (
        spGroupMembersResponse.map(async (memberItem) => {
          const loginNameSplit = memberItem.LoginName.split("|");
          let firstLevelAzureGroupPermissionLevel = [];

          // if spGroupMember = user
          if ((loginNameSplit[0] === 'i:0#.f') || (memberItem.Title === 'Company Administrator')) {
            
            //create user in this.users
            const userEntry = `user${userCount}`; userCount += 1;
            this.users[userEntry] = new Object;
            // write properties
            this.users[userEntry]["name"] = memberItem.Title;
            this.users[userEntry]["spId"] = memberItem.Id;
            this.users[userEntry]["principalName"] = loginNameSplit[2];
            if (this.spGroups[spGroupEntry].groupName !== "Access given directly") {
              // permission level for groups except Access given directly
              this.users[userEntry]["permissionLevel"] = this.spGroups[spGroupEntry].permissionLevel;
            } else {
              // permission level for Access given directly group
              this.users[userEntry]["permissionLevel"] = memberItem.permissionLevel;
              this.users[userEntry]["permissionLevelDirectAccess"] = memberItem.permissionLevel;
            }
            this.users[userEntry]["groupNesting"] = [groupNestingBranch];
            this.users[userEntry]["spGroup"] = spGroupEntry;
            // if user has email
            if (memberItem.Email) {
              this.users[userEntry]["email"] = memberItem.Email;
              // if not, check if principalName is email
            } else if (regExMail.test(loginNameSplit[2])) {
              this.users[userEntry]["email"] = loginNameSplit[2];
              // if neighter, no email
            } else {this.users[userEntry]["email"] = '';}
            

            //write user into this.spGroups.users
            this.spGroups[spGroupEntry].users.push(userEntry);
            return userEntry;
          } 

          // if spGroupMember = azure group
          else if ((loginNameSplit[0] === 'c:0t.c') || (loginNameSplit[0] === 'c:0o.c')) { 

            const azureGroupId = loginNameSplit[2];
            const parentAzureGroupEntry = "spGroup";
            // set firstLevelAzureGroupPermissionLevel for Access given directly group having members with individual permission levels
            if (this.spGroups[spGroupEntry].groupName === "Access given directly") {
              firstLevelAzureGroupPermissionLevel = memberItem.permissionLevel;
            }
            // get type properties
            let url = "/groups/" + azureGroupId + "?$select=groupTypes,mailEnabled,securityEnabled" + "&$c=" + Date.now().toString(); // "&$c=" + Date.now().toString() for IE11 because otherwise it takes the cached request
            // if M365OwnersGroup
            if (azureGroupId.length > 36) {
              const azureGroupIdCut = azureGroupId.substring(0,36); // cut "_o" from end of id of M365OwnersGroup
              url = "/groups/" + azureGroupIdCut + "?$select=groupTypes,mailEnabled,securityEnabled" + "&$c=" + Date.now().toString(); // "&$c=" + Date.now().toString() for IE11 because otherwise it takes the cached request
            }
            const typePropertiesResponse = await this._graphApiGet(url);
            let azureGroupType;
            // if any error, just display "Azure group" for type
            if (typePropertiesResponse.statusCode) {
              azureGroupType = {long: "Azure group", short: "AG"};
            }
            else {
              azureGroupType = this._evalAzureGroupType (typePropertiesResponse["groupTypes"][0], typePropertiesResponse["mailEnabled"], typePropertiesResponse["securityEnabled"]);
            }
            return await this._buildAzureGroupAndGetChildItemsRecursive(memberItem.Title, azureGroupId, spGroupEntry, groupNestingBranch, parentAzureGroupEntry, azureGroupType, firstLevelAzureGroupPermissionLevel);
          } 
          else return `${memberItem.LoginName} is no user nor azure group`;
        })
      );
    }
  }

  // get members of SharePoint group
  private async _getSpGroupMembers (spGroupId: number): Promise<object> {
    const url = `${this.props.siteCollectionURL}/_api/web/SiteGroups/GetById(${spGroupId})/users?$select=Title,LoginName,Email,Id&$top=5000`; 
    const response = await this._spApiGet(url);
    // console.log("_getSpGroupMembers of ", spGroupId, response);
    return response;

  }

  // get permissions of SharePoint group and write them into memory
  private async _getAndBuildSpGroupPermissions (spGroupId: string, spGroupEntry:string): Promise<any> {
    // check if sp group has permission level, to avoid throwing an error when trying to get perm level for group from api and it hasn't one.
    // this.allPermResponse gives an Array with all sp groups, that have permission levels. those who don't have, will be hidden.
    // For default groups assume that they have one, because this.allPermResponse is executed after getting default groups to improve performance on load
    // if this.allPermResponse has error, the sp group won't have permission and group may get hidden.
    let hasPermissionLevel = true;
    if (this.allPermResponse) {hasPermissionLevel = this.allPermResponse.value.some(item=>item.PrincipalId==spGroupId);}
    // if sp group has permission level
    if (hasPermissionLevel) {
      const url = `${this.props.siteCollectionURL}/_api/Web/RoleAssignments/GetByPrincipalId(${spGroupId})/RoleDefinitionBindings?$select=Name`;
      const response = await this._spApiGet(url);
      this.spGroups[spGroupEntry]['permissionLevel'] = new Array;
      this.spGroups[spGroupEntry]['isHidden'] = false;
      response["value"].map((permLevelEntry)=>{
        this.spGroups[spGroupEntry].permissionLevel.push(permLevelEntry.Name);
      });
      // add flag isHidden if just permission level Limited Access
      if ((response["value"].length==1) && (this.spGroups[spGroupEntry].permissionLevel.includes('Limited Access'))) {
        this.spGroups[spGroupEntry].isHidden = true;
      }
    }  else {
      // add flag isHidden if has't permissions
      this.spGroups[spGroupEntry].isHidden = true;
    }
    return "permission level added";
  }

  // create object of SharePoint group and get its members
  private async _buildSpGroupAndGetMembers (spGroupResponse, spGroupEntry:string): Promise<any> {
    
    if (!this.spGroups[spGroupEntry]) {
      this.spGroups[spGroupEntry] = new Object;
    }
    this.spGroups[spGroupEntry]['id'] = spGroupResponse.Id;
    this.spGroups[spGroupEntry]['loginName'] = spGroupResponse.LoginName;
    this.spGroups[spGroupEntry]['groupName'] = spGroupResponse.Title;
    this.spGroups[spGroupEntry]['owner'] = spGroupResponse.OwnerTitle;
    this.spGroups[spGroupEntry]['description'] = '';
    this.spGroups[spGroupEntry]['type'] = "SharePoint group";
    this.spGroups[spGroupEntry]['typeShort'] = "SP";
    this.spGroups[spGroupEntry]['users'] = [];
    //keep displayname of default groups
    if (!this.spGroups[spGroupEntry].displayName) {
      this.spGroups[spGroupEntry].displayName = spGroupResponse.Title;
    }

    // get sp group permissions
    await this._getAndBuildSpGroupPermissions(spGroupResponse.Id, spGroupEntry);
    if (showLogs) { console.log(`${spGroupEntry} permission levels after adding new ones:`, this.spGroups[spGroupEntry]['permissionLevel']);}

    // get members
    const response = await this._getSpGroupMembers(spGroupResponse.Id);
    let spGroupMembersResponse = response["value"];

    // if response status ok
    if (response["status"].toString().startsWith('2')) {
      // filter M365OwnersGroup from spDefaultOwnersGroup
      if (filterM365OwnersGroupFromSpDefaultOwnersGroup) {
        if (spGroupEntry === "spGroup2") {
          spGroupMembersResponse.map(
            (memberItem, index) => {
              const loginNameSplit = memberItem.LoginName.split("_");
              if (loginNameSplit[1]) {
                spGroupMembersResponse.splice(index, 1);
              }
            }
          );
        }
      }
    } 
    // if no access
    else if (errorStatusAccessDeniedArray.includes(response["status"])) {
      spGroupMembersResponse = "no access";
    }
    // if any other error, spGroupMembersResponse = empty array or undefined
    if (showLogs) { console.log( `${spGroupEntry} Members response: `, spGroupMembersResponse); }
    // write members
    return await this._buildSpGroupMembers(spGroupEntry, spGroupMembersResponse);
  }
  
  // ------------ primary methods  ---------------

  // get site admins
  private async _getAdmins (): Promise<any> { 
    const url = `${this.props.siteCollectionURL}/_api/web/siteusers?$filter=IsSiteAdmin eq true`;
    if (showLogs) { console.log("----------------_getAdmins executed");}
    const response = await this._spApiGet(url);
    const spGroupEntry = "spGroup1";
    this.spGroups.spGroup1.users = [];
    let getAdminsResponse;
    // to display "no access": if response.value = empty: since there is always at least one site admin, we can assume that user has no permission to read admins
    if (!response["value"][0]) {
      getAdminsResponse = "no access";
    } else getAdminsResponse = response["value"];
    return await this._buildSpGroupMembers(spGroupEntry, getAdminsResponse);
  }

  // get members with access given directly
  private async _getDirectAccess (): Promise<any> { 
    const url = `${this.props.siteCollectionURL}/_api/web/RoleAssignments?$expand=Member,RoleDefinitionBindings&$filter=Member/PrincipalType ne 8`;
    const response = await this._spApiGet(url);

    const spGroupEntry = `spGroup${spGroupCount}`; spGroupCount += 1;
    // write sp group properties
    this.spGroups[spGroupEntry] = new Object;
    this.spGroups[spGroupEntry].id = 'no sp group';
    this.spGroups[spGroupEntry].groupName = 'Access given directly';
    this.spGroups[spGroupEntry].displayName = 'Access given directly';
    this.spGroups[spGroupEntry].owner = 'No owner';
    this.spGroups[spGroupEntry].description = 'Users or groups with access given directly to the site.';
    this.spGroups[spGroupEntry].type = "SharePoint";
    this.spGroups[spGroupEntry].users = [];
    this.spGroups[spGroupEntry].permissionLevel = [''];
    this.spGroups[spGroupEntry].isHidden = false;
    // prepare members from response
    // filter hidden members
    const visibleMembers = response["value"].filter(member=>{
      // loop through permission levels of member, if it has at least one permission level with hidden=false, return true
      let hasVisiblePermissionLevel = false;
      member.RoleDefinitionBindings.forEach(
        permissionLevelItem=>{
          if (permissionLevelItem.Hidden === false) {
            return hasVisiblePermissionLevel = true;
          }
        }
      );
      return hasVisiblePermissionLevel;
    });

    let getDirectAccessResponse;
    // error handling: if any error, response.value = empty array, so "no users" is displayed. 
    // in case of no access, "no access" will be displayed
    if (errorStatusAccessDeniedArray.includes(response["status"])) {
      getDirectAccessResponse = "no access";
    } else {
      // props: loginName, Id, Title, Email, permissionLevel
      getDirectAccessResponse = visibleMembers.map(
        memberItem=>{
          let prettyMember = {};
          prettyMember['LoginName'] = memberItem.Member.LoginName;
          prettyMember['Id'] = memberItem.Member.Id;
          prettyMember['Title'] = memberItem.Member.Title;
          prettyMember['Email'] = memberItem.Member.Email;
          prettyMember['permissionLevel'] = memberItem.RoleDefinitionBindings.map(permissionLevelItem=>permissionLevelItem.Name);
          return prettyMember;
        }
      );
    }
    // write members
    
    if (showLogs) { console.log( `${spGroupEntry} Members response: `, getDirectAccessResponse);}
    return await this._buildSpGroupMembers(spGroupEntry, getDirectAccessResponse);
  }

  // get all SharePoint groups
  private async _getOtherSpGroups (): Promise<any> {

    const url = `${this.props.siteCollectionURL}/_api/web/sitegroups?$select=Id,Title,OwnerTitle,LoginName`;
    if (showLogs) { console.log("----------------_getOtherSpGroups executed");}
    const allSiteGroupsResponse = await this._spApiGet(url);
    
    // error handling: not needed, because if sp api gives error back, response.value will be an empty array. in this case web part is just not displaying any groups
    if (showLogs) { console.log('all sitegroups response: ', allSiteGroupsResponse["value"]);}
    return await Promise.all (
      allSiteGroupsResponse["value"].map(
        async (spGroupResponseItem) => {
          //filter default groups
          switch (spGroupResponseItem.Id) {
            case this.spGroups.spGroup2['id'] : return "spGroup already exists";
            case this.spGroups.spGroup3['id'] : return "spGroup already exists";
            case this.spGroups.spGroup4['id'] : return "spGroup already exists";
            default : {
              const spGroupEntry = `spGroup${spGroupCount}`; spGroupCount += 1;
              if (showLogs) { console.log(`response entry for ${spGroupEntry}: `, spGroupResponseItem);}
              return await this._buildSpGroupAndGetMembers (spGroupResponseItem, spGroupEntry);
            }
          }
        }
      )
      .concat(
        await this._getDirectAccess()
      )
    );
  }
  
  // get default SharePoint groups
  private async _getDefaultSpGroups (defaultGroupArray: string[]): Promise<any> {
    if (showLogs) { console.log("----------------_getDefaultSpGroups executed");}
    return await Promise.all(
      defaultGroupArray.map(
        async (defaultGroup) => {
          const url = `${this.props.siteCollectionURL}/_api/web/Associated${defaultGroup}Group?$select=Id,Title,OwnerTitle`;
          const defaultGroupResponse = await this._spApiGet(url);
          if (showLogs) { console.log(`default ${defaultGroup} response: `, defaultGroupResponse);}
          let groupNr: string;
          switch (defaultGroup) {
            case 'Owner': groupNr = "2"; break;
            case 'Member': groupNr = "3"; break;
            case 'Visitor': groupNr = "4"; break;
          }
          const spGroupEntry = `spGroup${groupNr}`;
          return await this._buildSpGroupAndGetMembers(defaultGroupResponse, spGroupEntry);
        }
      )
      .concat(
        await this._getAdmins()
      )
    );
  }

  // ---------- prettify results --------------

  // remove duplicates of users: a user can be a member of more than one parent group
  private _removeDuplicateUsers () {
    let userEntryDoubleForDeleteArray = [];
    const usersArray = Object.keys(this.users);
    if (usersArray.length > 1) {
      // first loop through this.users
      usersArray.forEach(
        (userEntry1Item, userEntry1Index) => {
          let userEntryItem2Array = [];
          // second loop through this.users to compare principalName
          for (let userEntry2Index = userEntry1Index +1; userEntry2Index < usersArray.length; userEntry2Index++ ) {
            const userEntry2Item = usersArray[userEntry2Index];

            // if user duplicate (principalName)
            if (this.users[userEntry1Item].principalName === this.users[userEntry2Item].principalName) {
              if (showLogsOfUserDouble) { console.log("duplicate user in this.users: userEntry1Item = " , userEntry1Item, ", userEntry2Item = ", userEntry2Item);}
              // put all user doubles (userEntry2Item) for actual userEntry1Item in array to handle them after second loop through this.users
              userEntryItem2Array.push(userEntry2Item);
            }
          }
          // put userEntryItem2Array into const for safety
          const userEntryItem2ArrayConst = userEntryItem2Array;

          // handle user doubles
          userEntryItem2ArrayConst.forEach(
            userEntry2Item => {
              // add permission levels to user, even if it already exists, doubles will be deleted after _removeDuplicateUsers
              if (this.users[userEntry2Item].permissionLevel) {
                if (!this.users[userEntry1Item].permissionLevel) {this.users[userEntry1Item]['permissionLevel'] = [];}
                this.users[userEntry1Item].permissionLevel = this.users[userEntry1Item].permissionLevel.concat(this.users[userEntry2Item].permissionLevel);
              }

              // add permission level of dirct access
              // if user 2 has permissionLevelDirectAccess
              if (this.users[userEntry2Item].permissionLevelDirectAccess) {
                // if user 1 doesn't have permissionLevelDirectAccess, then create array
                if (!this.users[userEntry1Item].permissionLevelDirectAccess) {this.users[userEntry1Item]['permissionLevelDirectAccess'] = [];}
                // add permissionLevelDirectAccess of user 2 to user 1
                this.users[userEntry1Item].permissionLevelDirectAccess = this.users[userEntry1Item].permissionLevelDirectAccess.concat(this.users[userEntry2Item].permissionLevelDirectAccess);
              }// handle doubles in permissionLevelDirectAccess later

              // push groupNestingBranch of second user into groupNesting of first user
              this.users[userEntry1Item].groupNesting.push(this.users[userEntry2Item].groupNesting[0]);

              // put azureId to userEntry1Item, if userEntry1Item has no azureId and userEntry2Item has one
              if (!this.users[userEntry1Item].azureId && this.users[userEntry2Item].azureId)  {
                this.users[userEntry1Item].azureId = this.users[userEntry2Item].azureId;
              }
              // put spId to userEntry1Item, if userEntry1Item has no spId and userEntry2Item has one
              if (!this.users[userEntry1Item].spId && this.users[userEntry2Item].spId)  {
                this.users[userEntry1Item].spId = this.users[userEntry2Item].spId;
              }

              // look in all spGroups.users for userEntry2Item, if found, check for userEntry1Item, if exists delete second, if not replace the second with first
              // loop through spGroups
              Object.keys(this.spGroups).forEach(
                spGroupEntryItem => {
                  let userEntryItem1Exists = false;
                  let userEntryItem2Exists = false;
                  if (showLogsOfUserDouble) { console.log(spGroupEntryItem,": checking users");}
                  if (showLogsOfUserDouble) { console.log(`this.spGroups.${spGroupEntryItem}.permissionLevel: `, this.spGroups[spGroupEntryItem].permissionLevel);}
                  // for each spGroup, loop through their users (userEntryItems are all unique, even if it is the same user)
                  let spUserIndexOfUserEntryItem2;
                  this.spGroups[spGroupEntryItem].users.forEach(
                    (groupEntryItem2, spUserIndex2) => {
                      
                      // if userEntry2Item exists in spGroup.users
                      if (groupEntryItem2 === userEntry2Item) {
                        userEntryItem2Exists = true;
                        spUserIndexOfUserEntryItem2 = spUserIndex2;
                        if (showLogsOfUserDouble) { console.log("userEntry2Item ", userEntry2Item, "exists in ", spGroupEntryItem);}
                        // check if also userEntry1Item exists in spGroup.users
                        
                        this.spGroups[spGroupEntryItem].users.forEach(
                          groupEntryItem1 => {
                            // if userEntry1Item exist in spGroup.users, set userEntry1Exists = true
                            if (groupEntryItem1 === userEntry1Item) {
                              userEntryItem1Exists = true;
                              if (showLogsOfUserDouble) { console.log("userEntry1Item ", userEntry1Item, "exists in ", spGroupEntryItem);}
                            }
                          }
                        );
                      }
                    }
                  );
                  
                  // if userEntry1Item and userEntry2Item exist in spGroup.users, delete userEntry2Item (after finishing loop of spGroup.users)
                  if (userEntryItem1Exists && userEntryItem2Exists) {
                    this.spGroups[spGroupEntryItem].users.splice(spUserIndexOfUserEntryItem2, 1); 
                    if (showLogsOfUserDouble) { console.log("userEntry1Item and userEntry2Item exist in spGroup.users, delete userEntry2Item: ", userEntry2Item);}
                  }
                  // if just userEntry2Item exist in spGroup.users, replace userEntry2Item with userEntry1Item
                  if (!userEntryItem1Exists && userEntryItem2Exists) {
                    this.spGroups[spGroupEntryItem].users[spUserIndexOfUserEntryItem2] = userEntry1Item;
                    if (showLogsOfUserDouble) { console.log(`just userEntry2Item exist in spGroup.users, replace userEntry2Item ${userEntry2Item} with userEntry1Item ${userEntry1Item} `);}
                  }
                }
              );
              userEntryDoubleForDeleteArray.push(userEntry2Item);
              if (showLogsOfUserDouble) { console.log("userEntry2Item ", userEntry2Item, "will be deleted from this.users");}
            }
          );
        }
      );
      // delete (duplicate) userEntryItem2s from userEntryDoubleArray in this.users
      if (showLogsOfUserDouble) { console.log("userEntryDoubleForDeleteArray:", userEntryDoubleForDeleteArray);}
      userEntryDoubleForDeleteArray.forEach(
        (userEntryToDeleteItem) => {
          delete this.users[userEntryToDeleteItem];
          if (showLogsOfUserDouble) { console.log("userEntry2Item deleted: ", userEntryToDeleteItem);}
        }
      );
    }
  }

  // get all permission levels of site
  private async _getSitePermissionLevels(): Promise<Array<string>> {

    //get all permission levels of site
    const permissionLevelsResult = await this._spApiGet(`${this.props.siteCollectionURL}/_api/web/roleDefinitions`);
    // put their names in  array sitePermissionLevels
    let sitePermissionLevels = permissionLevelsResult["value"].map(
      (permissionLevelObjectItem) => {
        return permissionLevelObjectItem.Name;
      }
    );
    // add Administator to sitePermissionLevels
    sitePermissionLevels.unshift('Administrator');
    return sitePermissionLevels;
  }

  // prettify user permission levels: remove duplicates, sort
  private _prettifyUserPermissionLevels (sitePermissionLevels) {

    //delete doubles of permission levels in user, sort them
    Object.keys(this.users).forEach(
      (userEntryItem: string) => {

        const userPermLevelsArrayOrig = this.users[userEntryItem].permissionLevel;
        // filter duplicates
        // let userPermLevelsArrayUniq = [...new Set(userPermLevelsArrayOrig)];
        // instead for IE11
        let userPermLevelsArrayUniq = userPermLevelsArrayOrig.filter((v, i, a) => a.indexOf(v) === i);
        let userPermLevelsArrayUniqSorted = userPermLevelsArrayUniq;
        // if more than one permissionLevel
        if (userPermLevelsArrayUniq.length>1) {
          //sort permission levels by filtering unique permission levels out of sitePermissionLevels to keep the original order
          userPermLevelsArrayUniqSorted = sitePermissionLevels.filter( i => userPermLevelsArrayUniq.includes(i) );
        }
        // write unique sorted permission levels back to user
        this.users[userEntryItem].permissionLevel = userPermLevelsArrayUniqSorted;

        // if user has permissionLevelDirectAccess, then delete duplicate permission levels
        if (this.users[userEntryItem].permissionLevelDirectAccess) {
          // this.users[userEntryItem].permissionLevelDirectAccess = [...new Set(this.users[userEntryItem].permissionLevelDirectAccess)];
          // instead for IE11
          this.users[userEntryItem].permissionLevelDirectAccess = [ this.users[userEntryItem].permissionLevelDirectAccess.filter((v, i, a) => a.indexOf(v) === i)];
      }
      }
    );

  }

  // remove duplicates of azure gorups: an azure group can be a member of more than one parent group
  private _removeDuplicateAzureGroups () {
    let azureGroupEntryDoubleForDeleteArray = [];
    const azureGroupsArray = Object.keys(this.azureGroups);
    if (azureGroupsArray.length > 1) {
      // first loop through this.azureGroups
      azureGroupsArray.forEach(
        (azureGroupEntryItem1, azureGroupEntryIndex1) => {
          let azureGroupEntryItem2Array = [];
          // second loop through this.azureGroups
          for (let azureGroupEntryIndex2 = azureGroupEntryIndex1 +1; azureGroupEntryIndex2 < azureGroupsArray.length; azureGroupEntryIndex2++ ) {
            const azureGroupEntryItem2 = azureGroupsArray[azureGroupEntryIndex2];

            // if azureGroup duplicate
            if (this.azureGroups[azureGroupEntryItem1].id === this.azureGroups[azureGroupEntryItem2].id) {
              if (showLogsOfAzureDouble ) { console.log("duplicate azureGroup in this.azureGroups: azureGroupEntryItem1 = " , azureGroupEntryItem1, ", azureGroupEntryItem2 = ", azureGroupEntryItem2);}
              // put all azureGroup doubles (azureGroupEntryItem2) for actual azureGroupEntryItem1 in array to handle them after second loop through this.azureGroups
              azureGroupEntryItem2Array.push(azureGroupEntryItem2);
            }
          }
          // handle azureGroup doubles
          azureGroupEntryItem2Array.forEach(
            (azureGroupEntryItem2) => {

              // look in all user.groupNesting for azureGroupEntryItem2, if found replace the second with first
              // loop through users
              Object.keys(this.users).forEach(
                (userEntry) => {
                  // for each user, loop through their groupNesting (azureGroupEntryItems are all unique, even if it is the same azureGroup)
                  
                  // first loop through groupNesting (groupnesting is an array of arrays)
                  this.users[userEntry].groupNesting.forEach(
                    (nestingBranchItem, nestingBranchIndex) => {
                      // second loop through groupNestingBranch
                      nestingBranchItem.forEach(
                        (groupEntryItem2, groupEntryIndex2) => {
                          
                          // if azureGroupEntryItem2 exists in groupNestingBranch of userEntry
                          if (groupEntryItem2 === azureGroupEntryItem2) {
                            // replace the second with first
                            this.users[userEntry].groupNesting[nestingBranchIndex][groupEntryIndex2] = azureGroupEntryItem1;
                            if (showLogsOfAzureDouble ) { console.log(azureGroupEntryItem2, " replaced with ", azureGroupEntryItem1) ;}
                          }
                        }
                      );
                    }
                  );
                }
              );
              azureGroupEntryDoubleForDeleteArray.push(azureGroupEntryItem2);
              if (showLogsOfAzureDouble ) { console.log("azureGroupEntryItem2 ", azureGroupEntryItem2, "will be deleted from this.azureGroups") ;}
            }
          );
        }
      );
      // delete (duplicate) azureGroupEntryItem2s from azureGroupEntryDoubleArray in this.azureGroups
      if (showLogsOfAzureDouble ) { console.log("azureGroupEntryDoubleForDeleteArray:", azureGroupEntryDoubleForDeleteArray) ;}
      azureGroupEntryDoubleForDeleteArray.forEach(
        (azureGroupEntryToDeleteItem) => {
          delete this.azureGroups[azureGroupEntryToDeleteItem];
          if (showLogsOfAzureDouble ) { console.log("azureGroupEntryItem2 deleted: ", azureGroupEntryToDeleteItem) ;}
        }
      );
    }
  }

  // sort azure groups by type
  private _sortAzureGroups () {
    // prepare for user card:
    // sort azure groups first by group type, second by group name
    let azureGroupArraySorted = [];
    let azureGroupArray = [];
    // put all azure group objects in azureGroupArray to sort them
    Object.keys(this.azureGroups).forEach(
      azureGroupEntryItem => {
        this.azureGroups[azureGroupEntryItem].key = azureGroupEntryItem;
        azureGroupArray.push(this.azureGroups[azureGroupEntryItem]);
      }
    );
    // filter azure groups by type, and put them in array, put those arrays in array of arrays
    let azureGroupArrayArray = [];
    azureGroupArrayArray.push(azureGroupArray.filter(azureGroupItem => azureGroupItem.type.long.startsWith("Micro")));
    azureGroupArrayArray.push(azureGroupArray.filter(azureGroupItem => azureGroupItem.type.long.startsWith("Security")));
    azureGroupArrayArray.push(azureGroupArray.filter(azureGroupItem => azureGroupItem.type.long.startsWith("Dist")));
    azureGroupArrayArray.push(azureGroupArray.filter(azureGroupItem => azureGroupItem.type.long.startsWith("Mail")));
    // sort the arrays of azure group objects by name
    const azureGroupArrayArraySorted = azureGroupArrayArray.map(
      azureGroupTypeArrayItem => azureGroupTypeArrayItem.sort((a, b) => {
        return a.name.localeCompare(b.name);
      })
    );
    // put the azure gorup objects all in one array, that is now sorted first by group type, second by group name
    azureGroupArrayArraySorted.forEach(
      azureGroupTypeArrayItem => azureGroupTypeArrayItem.forEach(azureGroupItem => 
        azureGroupArraySorted.push(azureGroupItem)
      )
    );
    // supply azureGroupArraySorted to state to be available in group card
    this.azureGroupArraySorted = azureGroupArraySorted;
  }

  // unify parallel parts of group nesting paths for each user and create tree
  private _unifyGroupNestingOfUser () {

    // unify group nesting branches and create tree
    Object.keys(this.users).forEach(
      (userEntryItem) => {

        let groupNesting = this.users[userEntryItem].groupNesting;
        // add user to every group nesting branch at the end for treeview in group card
        groupNesting = groupNesting.map(branchItem=>branchItem.concat(userEntryItem));

        // write new groupNesting tree back into user
        this.users[userEntryItem].groupNesting = createTree(groupNesting);

        // create tree function
        function createTree(structure) {
          const node = (name) => ({name, children: []});
          const addNode = (parent, child) => (parent.children.push(child), child);
        
          const findNamed = (name, parent) => {
              for (const child of parent.children) {
                  if (child.name === name) { return child; }
                  const found = findNamed(name, child);
                  if (found) { return found; }            
              }
          };
          const TOP_NAME = "groupNesting", top = node(TOP_NAME);
          for (const children of structure) {
              let parent = top;
              for (const name of children) {
                  const found = findNamed(name, parent);
                  parent = found ? found : addNode(parent, node(name));
              }
          }
          return top;
        }

      }
    );
    
    // sort array groupNestingBranch
    Object.keys(this.users).forEach(
      (userEntryItem) => {
        // it more than one group nesting branch
        if (this.users[userEntryItem].groupNesting.children.length > 1) {
          const unsorted = this.users[userEntryItem].groupNesting.children;
          // sort array groupNestingBranch by spGroup entry , which is the first element in array
          this.users[userEntryItem].groupNesting.children = unsorted.sort((a, b) => {
            const item1 = Number(a['name'].replace( /^\D+/g, ''));
            const item2 = Number(b['name'].replace( /^\D+/g, ''));
            let compare = 1;
            if (item1 < item2) {(compare = -1);}
            return compare;
          });
        }
      }
    );
  }
  
  // get details about Sharing groups
  private async _getPropsOfSharingGroups () {
    // get name of organization
    const getOrgaNameResponse = await this._graphApiGet('/organization?$select=displayName' + "&$c=" + Date.now().toString()); // "&$c=" + Date.now().toString() for IE11 because otherwise it takes the cached request
    let orgaName;
    // if no error
    if (!getOrgaNameResponse.statusCode) {
      orgaName = getOrgaNameResponse['value'][0].displayName;
    } 
    // if error
    else {
      orgaName = "your organization";
    }

    let groupObjectArray = [];
    let guidArray = [];
    // look in spGroups for "SharingLinks", get guid from loginName, create object for groups and put it in groupObjectArray
    Object.keys(this.spGroups).forEach(spGroupEntryItem=>{
      const loginName = this.spGroups[spGroupEntryItem].loginName;
      if (loginName) {
        if (loginName.startsWith('SharingLinks.')) {
          const loginNameSplit = loginName.split('.');
          const guid = loginNameSplit[1].toUpperCase();
          guidArray.push(guid);
          const shareId = loginNameSplit[3];
          groupObjectArray.push({spGroupEntry: spGroupEntryItem, item: {guid: guid, shareId: shareId}});
        }
      }
    });
    // if there is at least one sharing group
    if (guidArray[0]) {
      // get items from sp api search
      const queryText = guidArray.join(' OR ');
      const url = `${this.props.siteCollectionURL}/_api/search/query?querytext='${queryText}'&selectproperties='Title,Path,ListId,ParentLink'`;
      const getItemResponse = await this._spApiGet(url);

      // if some result (in case of no sharing groups for not deleted items)
      // error handling: if getItemResponse contains an error, just dispaying regular name of sharing group
      try {
        const getItemResult = getItemResponse['PrimaryQueryResult'].RelevantResults.Table.Rows;
        
        // for each element of getItemResult extract title, path and guid
        const siteCollectionUrlLenght = this.props.siteCollectionURL.length;
        const result = await Promise.all (
          getItemResult.map(async resultItem => {
            const title = resultItem.Cells.filter(cell=>cell.Key==='Title')[0].Value;
            const fullPath = resultItem.Cells.filter(cell=>cell.Key==='Path')[0].Value;
            const path = fullPath.substring(siteCollectionUrlLenght+1);
            const fullGuid = resultItem.Cells.filter(cell=>cell.Key==='UniqueId')[0].Value;
            const guid = fullGuid.substring(1,37).toUpperCase();
            const listId = resultItem.Cells.filter(cell=>cell.Key==='ListId')[0].Value;
            const parentLink = resultItem.Cells.filter(cell=>cell.Key==='ParentLink')[0].Value;
            // for each group object, if same guid, add title and path. add item to this.spGroups
            return await Promise.all (groupObjectArray.map(async groupObjectItem=>{
              let getPermResult;
              if (groupObjectItem.item.guid===guid) {
                groupObjectItem.item.title = title;
                groupObjectItem.item.path = path;
                groupObjectItem.item.listId = listId;
                groupObjectItem.item.parentLink = parentLink;
                // take spGroupEntry from groupObjectItem and add groupObjectItem.item to spGroup
                this.spGroups[groupObjectItem.spGroupEntry]['item'] = groupObjectItem.item;
                // make group name from item.title 
                this.spGroups[groupObjectItem.spGroupEntry]['subTitle'] = this.spGroups[groupObjectItem.spGroupEntry].displayName;
                this.spGroups[groupObjectItem.spGroupEntry]['displayName'] = 'Sharing: ' + title;
                this.spGroups[groupObjectItem.spGroupEntry]['groupName'] = 'Sharing: ' + title;

                // make description from path and link type

                // link type and permission from loginname
                const linkTypeRaw = this.spGroups[groupObjectItem.spGroupEntry].loginName.split('.')[2];
                let linkType = '';
                let permissionLevel = '';
                let permissionText = '';

                // case people with the link
                if (linkTypeRaw.startsWith('Orga')) {
                  linkType = `People in ${orgaName} with the link`;
                  if (linkTypeRaw.includes('Edit')) {permissionText = 'can edit'; permissionLevel = 'Contribute';}
                  if (linkTypeRaw.includes('View')) {permissionText = 'can view'; permissionLevel = 'Read';}
                }

                // case specific people
                else if (linkTypeRaw.startsWith('Flex')) {
                  linkType = `Specific people`;
                  // get permission level
                  const requestUrl = `${this.props.siteCollectionURL}/_api/web/Lists('${listId}')/GetItemByUniqueId('${guid}')/GetSharingInformation?$Expand=permissionsInformation&$Select=links,Id`;
                  getPermResult = await this._spApiPost(requestUrl);
                  // error handling: if error, just not displaying permission
                  // if response ok, display permission level
                  if (getPermResult.status.toString().startsWith('2')) {
                    getPermResult['permissionsInformation'].links.forEach(permResultItem=>{
                      if (permResultItem.linkDetails.ShareId===groupObjectItem.item.shareId) {
                        if (permResultItem.linkDetails.isEditLink) {
                          permissionText = 'can edit'; permissionLevel = 'Contribute';
                        } else {
                          permissionText = 'can view'; permissionLevel = 'Read';
                        }
                      }
                    });
                  }
                }

                // case anyone with the link
                else if (linkTypeRaw.startsWith('Anon')) {
                  linkType = `Anyone with the link`;
                  if (linkTypeRaw.includes('Edit')) {permissionText = 'can edit'; permissionLevel = 'Contribute';}
                  if (linkTypeRaw.includes('View')) {permissionText = 'can view'; permissionLevel = 'Read';}
                }

                // write them in spGroup
                this.spGroups[groupObjectItem.spGroupEntry]['description'] = `${linkType} ${permissionText} "${path}"`;
                this.spGroups[groupObjectItem.spGroupEntry]['permissionLevel'] = [permissionLevel];
              }
              return;
            }));
          })
        );
      }
      catch (error) {
        if (logErrors) {console.log(error);}
        if (this.props.throwErrors) {throw error;}
      }
    }
    return;
  }

  // ---------- call batches -----------------
  
  // get and display default Sharepoint groups and their members
    private async _firstCallBatch () {
    const resultDefaultGroups = await this._getDefaultSpGroups(["Owner", "Member", "Visitor"]);
    if (showLogs) { console.log("----------------_getDefaultSpGroups result: ", resultDefaultGroups);}
    this._removeDuplicateAzureGroups ();
    this._removeDuplicateUsers();
    this.sitePermissionLevels = await this._getSitePermissionLevels();
    this.allPermResponse = await this._spApiGet(`${this.props.siteCollectionURL}/_api/Web/RoleAssignments`);
    this._prettifyUserPermissionLevels(this.sitePermissionLevels);
    return ;
  }

  // get and display all other Sharepoint groups and their members
  private async _secondCallBatch () {
    if (showLogs) {console.log("_secondCallBatch executed");}
    const resultOtherGroups = await this._getOtherSpGroups();
    if (showLogs) { console.log("----------------_getOtherSpGroups result: ", resultOtherGroups);}
    this._removeDuplicateAzureGroups ();
    this._sortAzureGroups();
    this._removeDuplicateUsers();
    this._prettifyUserPermissionLevels(this.sitePermissionLevels);
    this._unifyGroupNestingOfUser();
    
    // if hidden groups
    if (Object.keys(this.state.spGroups).filter(spGroupEntryItem=>this.spGroups[spGroupEntryItem].isHidden===true)[0]) {
      this.hiddenGroupsExist = true;
    }

    await this._getPropsOfSharingGroups();
    if (showLogs) {console.log("_secondCallBatch finished");}
    return ;
  }

  // -------- calling functions -----------

  // start for features of web part
  public async componentDidMount() {

    try {
      await this._firstCallBatch();
      this.setState(
        {spGroups: this.spGroups, users: this.users, azureGroups: this.azureGroups},
        /* react sometimes executes setState asyncron, then it causes false display in UI, to avoid execute _secondCallBatch after state is actually set (still not shure if it works always)*/
        async ()=>{
          // the callback after setState seems not to bubble an error, so it must be wrapped in try catch
          try {
            if (showLogs) {console.log('Default groups ready');}
            await this._secondCallBatch();
          } catch (error) {
            if (logErrors) {console.log(error);}
            if (this.props.throwErrors) {throw error;}
          }
          finally {
            this.setState({spGroups: this.spGroups, users: this.users, azureGroups: this.azureGroups, azureGroupArraySorted: this.azureGroupArraySorted, isGroupsLoading: false, hiddenGroupsExist: this.hiddenGroupsExist});
            if (logLastState) {console.log(this.state);}
          }
        }
      );
    } 
    catch (error) {
      if (logErrors) {console.log(error);}
      if (this.props.throwErrors) {throw error;}
    }
  }

  public componentWillMount(){
    initializeIcons();
  }

  public render(): React.ReactElement<IPermissionCenterProps> {
    
    try {
      _reload = () => {
        userCount = 1;
        azureGroupCount = 1;
        spGroupCount = 5;
        this.props.reload();
      };
      
      const _getTabTitle = (event) => {
        this.setState({selectedTab: event.currentTarget.dataset.id});
      };
      
      return (
        <div className={ cssStyles.permissionCenter }>
          <div className={ cssStyles.container } >
            <div className={ cssStyles.row }>
              <div className={ cssStyles.column }>

                {/* Webpart header */}
                <div className={ cssStyles.titleContainer }>
                  <span className={ cssStyles.title }>Permission Center</span>
                  {this.props.config.showMenu &&
                    <SpMenu siteCollectionURL={this.props.siteCollectionURL}/>
                  }
                </div>
                
                {/* Tab headers */}
                <div className={cssStyles.tabContainer}>
                  <div className={(this.state.selectedTab === "Groups") ? cssStyles.tabActiveItem : cssStyles.tabItem} data-id = "Groups" onClick={(event)=>_getTabTitle(event)} >Groups</div>
                  {this.props.config.showTabUsers &&
                    <div className={(this.state.selectedTab === "Users") ? cssStyles.tabActiveItem : cssStyles.tabItem} data-id = "Users" onClick={(event)=>_getTabTitle(event)} >Users</div>
                  }
                  {this.props.config.showTabHidden &&
                    <div className={(this.state.selectedTab === "Hidden groups") ? cssStyles.tabActiveItem : cssStyles.tabItem} data-id = "Hidden groups" onClick={(event)=>_getTabTitle(event)} >Hidden groups</div>
                  }
                  <div style={{flexShrink: 20, flexGrow: 1}}/>
                  <IconButton iconProps={{ iconName: 'Refresh' }} title="Reload Webpart with new data" ariaLabel="Refresh" className={`${cssStyles.refresh}`} onClick={_reload}  />
                </div>

                {/* Tab contents */}
                <div className={cssStyles.tabContent}> 

                  {/* Groups tab content */}
                  {(this.state.selectedTab === 'Groups') && (
                    Object.keys(this.state.spGroups).map(
                      (spGroupEntryItem) => 
                      <SpGroupContainer 
                        spGroupEntry={spGroupEntryItem} 
                        state={this.state} 
                        props={this.props} 
                        hideGroup = {this.spGroups[spGroupEntryItem].isHidden}
                        isGroupsLoading={this.state.isGroupsLoading}  
                      />
                    ).concat(
                    this.state.isGroupsLoading && <div className={`${cssStyles.placeholderItem} ${cssStyles.userContainer}`}>Loading groups</div>)
                  )}
                  {/* Users tab content */}
                  {(this.state.selectedTab === 'Users') && (
                    <AllUsers 
                      state={this.state} 
                      props={this.props} 
                      isGroupsLoading={this.state.isGroupsLoading} 
                    />
                  )}
                  {/* Hidden groups tab content */}
                  { this.state.selectedTab === 'Hidden groups' && (
                      this.state.isGroupsLoading 
                      ? <div className={`${cssStyles.placeholderItem} ${cssStyles.userContainer}`}>Loading groups</div>
                      : ( 
                        this.state.hiddenGroupsExist
                        ? Object.keys(this.state.spGroups).map(
                            (spGroupEntryItem) => 
                            <SpGroupContainer 
                              spGroupEntry={spGroupEntryItem} 
                              state={this.state} 
                              props={this.props} 
                              hideGroup = {!this.spGroups[spGroupEntryItem].isHidden}
                              isGroupsLoading={this.state.isGroupsLoading}  
                            />
                          ) 
                        : <div>No hidden groups</div>
                      )
                    )
                  }

                </div>
              </div>
            </div>
          </div>
        </div>
      );
    }
    catch (error) {
      if (logErrors) {console.log(error);}
      if (this.props.throwErrors) {throw error;}
    }
  }
}