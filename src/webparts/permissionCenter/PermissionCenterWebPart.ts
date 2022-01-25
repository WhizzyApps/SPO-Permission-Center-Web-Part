import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneToggle, PropertyPaneLabel, PropertyPaneLink, PropertyPaneButton, PropertyPaneButtonType, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneField, IPropertyPaneCustomFieldProps, PropertyPaneFieldType } from '@microsoft/sp-property-pane';

import PermissionCenter from './components/PermissionCenter';
import { IPermissionCenterProps } from './components/IPermissionCenterProps';

const packageSolution = require("../../../config/package-solution.json");
const timeStampFile = require("../../../config/timeStamp.json");
const buildTimeStamp = timeStampFile["buildTimeStamp"];

export interface IPermissionCenterWebPartProps {
  configBasedOnPermissions: boolean;
  selectedRoleForConfig: string;

  manualShowMenu: boolean;
  manualShowTabUsers: boolean;
  manualShowTabHidden: boolean;
  manualShowCardGroup: boolean;
  manualShowCardUser: boolean;
  manualShowCardButtons: boolean;
  manualShowCardUserLinks: boolean;
  manualShowAdmins: boolean;
  manualShowOwners: boolean;
  manualShowMembers: boolean;
  manualShowDirectAccess: boolean;
  manualShowButtons: boolean;
  
  ownerShowMenu: boolean;
  ownerShowTabUsers: boolean;
  ownerShowTabHidden: boolean;
  ownerShowCardGroup: boolean;
  ownerShowCardUser: boolean;
  ownerShowCardButtons: boolean;
  ownerShowCardUserLinks: boolean;
  ownerShowAdmins: boolean;
  ownerShowOwners: boolean;
  ownerShowMembers: boolean;
  ownerShowDirectAccess: boolean;
  ownerShowButtons: boolean;

  memberShowMenu: boolean;
  memberShowTabUsers: boolean;
  memberShowTabHidden: boolean;
  memberShowCardGroup: boolean;
  memberShowCardUser: boolean;
  memberShowCardButtons: boolean;
  memberShowCardUserLinks: boolean;
  memberShowAdmins: boolean;
  memberShowOwners: boolean;
  memberShowMembers: boolean;
  memberShowDirectAccess: boolean;
  memberShowButtons: boolean;
  
  visitorShowMenu: boolean;
  visitorShowTabUsers: boolean;
  visitorShowTabHidden: boolean;
  visitorShowCardGroup: boolean;
  visitorShowCardUser: boolean;
  visitorShowCardButtons: boolean;
  visitorShowCardUserLinks: boolean;
  visitorShowAdmins: boolean;
  visitorShowOwners: boolean;
  visitorShowMembers: boolean;
  visitorShowDirectAccess: boolean;
  visitorShowButtons: boolean;

  debugMode: boolean;
  logState: boolean;
  logErrors: boolean;
  throwErrors: boolean;
  logPermCenterVars: boolean;
  logComponentVars: boolean;
  disableAnimateHeightUserCard: boolean;
  preloadAzureGroups: boolean;
  preloadAzureGroupsAmount: boolean;
  exportOrImportApiResponse: boolean;
  exportOrImportDropdown: string;
  importApiResponse: boolean;
  importApiResponseData;

}

export default class PermissionCenterWebPart extends BaseClientSideWebPart <IPermissionCenterWebPartProps> {
  
  private allowEditProps = false;
  
  // get data from SharePoint REST API
  private async _spApiGet (url: string): Promise<object> {

    const clientOptions: ISPHttpClientOptions = {
      headers: new Headers(),
      method: 'GET',
      mode: 'cors'
    };
    try {
      const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, clientOptions);
      //Since the fetch API only throws errors on network failure, We have to explicitly check for 404's etc.
      const responseJson = await response.json();

      responseJson['status'] = response.status;

      if (!responseJson.value) {
        responseJson['value'] = [];
      }
      return responseJson;
    } 
    catch (error) {
      return error;
    }
  } 

  // get permissions of current user
  private async _getUserPermissions () {
    let currentUserRole;
    // get permission of current users
    const urlAdmin = this.context.pageContext.web.absoluteUrl + '/_api/web/currentuser/isSiteAdmin';
    const isSiteAdminResponse = await this._spApiGet(urlAdmin);
    const isSiteAdmin = isSiteAdminResponse['value'];
    if (isSiteAdmin == true) {
      currentUserRole = "admin";
    } 
    else {
      const urlPerm = this.context.pageContext.web.absoluteUrl + `/_api/web/effectiveBasePermissions`;
      const permResponse = await this._spApiGet(urlPerm);
      let permArray = [];
      if (permResponse["Low"]) {
        permArray = this._convertUserPermissions (permResponse['Low'], permResponse['High']);
        if (permArray.includes("ManagePermissions")) {
          currentUserRole = "owner";
        } else if (permArray.includes("EditListItems")) {
          currentUserRole = "member";
        } else {
          currentUserRole = "visitor";
        }
      }
      else {
        currentUserRole = "visitor";
      }

    }
    return currentUserRole;
  }

  // convert permissions of current user
  private _convertUserPermissions (lowPermDec, highPermDec) {
    let Flags = {
      Low: [
        // Lists and Documents
        { EmptyMask: 0 },
        { ViewListItems: 1 << 0 },
        { AddListItems: 1 << 1 },
        { EditListItems: 1 << 2 },
        { DeleteListItems: 1 << 3 },
        { ApproveItems: 1 << 4 },
        { OpenItems: 1 << 5 },
        { ViewVersions: 1 << 6 },
        { DeleteVersions: 1 << 7 },
        { OverrideListBehaviors: 1 << 8 },
        { ManagePersonalViews: 1 << 9 },
        { ManageLists: 1 << 11 },
        { ViewApplicationPages: 1 << 12 },

        // Web Level
        { Open: 1 << 16 },
        { ViewPages: 1 << 17 },
        { AddAndCustomizePages: 1 << 18 },
        { ApplyThemAndBorder: 1 << 19 },
        { ApplyStyleSheets: 1 << 20 },
        { ViewAnalyticsData: 1 << 21 },
        { UseSSCSiteCreation: 1 << 22 },
        { CreateSubsite: 1 << 23 },
        { CreateGroups: 1 << 24 },
        { ManagePermissions: 1 << 25 },
        { BrowseDirectories: 1 << 26 },
        { BrowseUserInfo: 1 << 27 },
        { AddDelPrivateWebParts: 1 << 28 },
        { UpdatePersonalWebParts: 1 << 29 },
        { ManageWeb: 1 << 30 }
      ],
      High: [
        // High Bits
        { UseClientIntegration: 1 << 4 },
        { UseRemoteInterfaces: 1 << 5 },
        { ManageAlerts: 1 << 6 },
        { CreateAlerts: 1 << 7 },
        { EditPersonalUserInformation: 1 << 8 },

        // Special Permissions
        { EnumeratePermissions: 1 << 30 }
        //FullMask          :   2147483647 // Invisible in WebUI, not useful since it's always true when &'ed
      ]
    };
    // low permissions
    // Permissions: make array of objects
    const flagsLow = Flags.Low.map((objectItem) => {
      const code = (Object.values(objectItem)[0] >>> 0).toString(2);
      return { name: Object.keys(objectItem)[0], code: code };
    });
    //console.log('flagsLow', flagsLow);
    let zeros = "";
    const permLowArr = (lowPermDec >>> 0)
      .toString(2)
      .split("")
      .reverse()
      .map((item, index) => {
        const result = item + zeros;
        zeros += "0";
        return result;
      });
    //console.log('permLowArr', permLowArr);
    const flagsLowFiltered = flagsLow.filter((objectItem) =>
      permLowArr.includes(objectItem.code)
    );
    const lowPermArray = flagsLowFiltered.map((item) => item.name);
    let perm = lowPermArray;

    // high permissions
    // Permissions: make array of objects
    const flagsHigh = Flags.High.map((objectItem) => {
      const code = (Object.values(objectItem)[0] >>> 0).toString(2);
      return { name: Object.keys(objectItem)[0], code: code };
    });
    zeros = "";
    const permHighArr = (highPermDec >>> 0)
      .toString(2)
      .split("")
      .reverse()
      .map((item, index) => {
        const result = item + zeros;
        zeros += "0";
        return result;
      });
    const flagsHighFiltered = flagsHigh.filter((objectItem) =>
      permHighArr.includes(objectItem.code)
    );
    const highPermArray = flagsHighFiltered.map((item) => item.name);
    perm.concat(highPermArray);
    return perm;
  }
  
  private exportApiResponse = false;
  private importApiResponse = false;
  private currentUserRole;
  private _reRenderAndRecordAndDownloadApiResponse = () => {};
  private _rerenderWithImportedApiResponse = () => {};
  

  public async render() {

    try {

      if (!this.currentUserRole) {
        this.currentUserRole = await this._getUserPermissions();
      }
      
      if ((this.currentUserRole == "admin") || (this.currentUserRole == "owner") ) {
        this.allowEditProps = true;
      }

      const featureConfig = () => {
        if (this.properties.configBasedOnPermissions) {
          if (this.currentUserRole == 'visitor') {
            return {
              configBasedOnPermissions: this.properties.configBasedOnPermissions,
              showMenu: this.properties.visitorShowMenu,
              showTabUsers: this.properties.visitorShowTabUsers,
              showTabHidden: this.properties.visitorShowTabHidden,
              showCardGroup: this.properties.visitorShowCardGroup,
              showCardUser: this.properties.visitorShowCardUser,
              showCardButtons: this.properties.visitorShowCardButtons,
              showCardUserLinks: this.properties.visitorShowCardUserLinks,
              showAdmins: this.properties.visitorShowAdmins,
              showOwners: this.properties.visitorShowOwners,
              showMembers: this.properties.visitorShowMembers,
              showDirectAccess: this.properties.visitorShowDirectAccess
            };
          } else if (this.currentUserRole == 'member') {
            return {
              configBasedOnPermissions: this.properties.configBasedOnPermissions,
              showMenu: this.properties.memberShowMenu,
              showTabUsers: this.properties.memberShowTabUsers,
              showTabHidden: this.properties.memberShowTabHidden,
              showCardGroup: this.properties.memberShowCardGroup,
              showCardUser: this.properties.memberShowCardUser,
              showCardButtons: this.properties.memberShowCardButtons,
              showCardUserLinks: this.properties.memberShowCardUserLinks,
              showAdmins: this.properties.memberShowAdmins,
              showOwners: this.properties.memberShowOwners,
              showMembers: this.properties.memberShowMembers,
              showDirectAccess: this.properties.memberShowDirectAccess
            };
          } else if (this.currentUserRole == 'owner') {
            return {
              configBasedOnPermissions: this.properties.configBasedOnPermissions,
              showMenu: this.properties.ownerShowMenu,
              showTabUsers: this.properties.ownerShowTabUsers,
              showTabHidden: this.properties.ownerShowTabHidden,
              showCardGroup: this.properties.ownerShowCardGroup,
              showCardUser: this.properties.ownerShowCardUser,
              showCardButtons: this.properties.ownerShowCardButtons,
              showCardUserLinks: this.properties.ownerShowCardUserLinks,
              showAdmins: this.properties.ownerShowAdmins,
              showOwners: this.properties.ownerShowOwners,
              showMembers: this.properties.ownerShowMembers,
              showDirectAccess: this.properties.ownerShowDirectAccess
            };
          } 
          // currentUserRole = admin 
          else {
            return {
              configBasedOnPermissions: this.properties.configBasedOnPermissions,
              showMenu: true,
              showTabUsers: true,
              showTabHidden: true,
              showCardGroup: true,
              showCardUser: true,
              showCardButtons: true,
              showCardUserLinks: true,
              showAdmins: true,
              showOwners: true,
              showMembers: true,
              showDirectAccess: true,
              showButtons: true
            };
          }
        } 
        // configBasedOnPermissions off
        else {
          return {
            configBasedOnPermissions: this.properties.configBasedOnPermissions,
            showMenu: this.properties.manualShowMenu,
            showTabUsers: this.properties.manualShowTabUsers,
            showTabHidden: this.properties.manualShowTabHidden,
            showCardGroup: this.properties.manualShowCardGroup,
            showCardUser: this.properties.manualShowCardUser,
            showCardButtons: this.properties.manualShowCardButtons,
            showCardUserLinks: this.properties.manualShowCardUserLinks,
            showAdmins: this.properties.manualShowAdmins,
            showOwners: this.properties.manualShowOwners,
            showMembers: this.properties.manualShowMembers,
            showDirectAccess: this.properties.manualShowDirectAccess
          };
        }
      };
      let debugConfig = {
        "logErrors": false,
        "logState": false,
        "throwErrors": false,
        "logComponentVars": false,
        "logPermCenterVars": false,
        "disableAnimateHeightUserCard": this.properties.disableAnimateHeightUserCard,
        "preloadAzureGroups": true,
        "preloadAzureGroupsAmount": false,
        "exportOrImportApiResponse": false,
        "exportApiResponse": false,
        "importApiResponse": false,
        "importApiResponseData": null
        
      };
      if (this.properties.debugMode) {
        debugConfig = {
          "logErrors": this.properties.logErrors,
          "logState": this.properties.logState,
          "throwErrors": this.properties.throwErrors,
          "logComponentVars": this.properties.logComponentVars,
          "logPermCenterVars": this.properties.logPermCenterVars,
          "disableAnimateHeightUserCard": this.properties.disableAnimateHeightUserCard,
          "preloadAzureGroups": this.properties.preloadAzureGroups,
          "preloadAzureGroupsAmount": this.properties.preloadAzureGroupsAmount,
          "exportOrImportApiResponse": this.properties.exportOrImportApiResponse,
          "exportApiResponse": this.exportApiResponse,
          "importApiResponse": this.importApiResponse,
          "importApiResponseData": this.properties.importApiResponseData
        };
      }

      const config =  {...featureConfig(), ...debugConfig };
      
      const _reRender = () => {
        element.props.userAndFoto = {};
        ReactDom.unmountComponentAtNode(this.domElement);
        this.exportApiResponse = false;
        element.props.config.exportApiResponse = false;
        this.importApiResponse = false;
        element.props.config.importApiResponse = false;
        ReactDom.render(element, this.domElement);
      };

      this._reRenderAndRecordAndDownloadApiResponse = () => {
        element.props.userAndFoto = {};
        this.exportApiResponse = true;
        element.props.config.exportApiResponse = true;
        this.importApiResponse = false;
        element.props.config.importApiResponse = false;
        ReactDom.unmountComponentAtNode(this.domElement);
        ReactDom.render(element, this.domElement);
      };
      
      this._rerenderWithImportedApiResponse = () => {
        element.props.userAndFoto = {};
        this.exportApiResponse = false;
        element.props.config.exportApiResponse = false;
        this.importApiResponse = true;
        element.props.config.importApiResponse = true;
        ReactDom.unmountComponentAtNode(this.domElement);
        ReactDom.render(element, this.domElement);
      };

      const element: React.ReactElement<IPermissionCenterProps> = React.createElement(
        PermissionCenter,
        {
          config: config,
          siteCollectionURL: this.context.pageContext.web.absoluteUrl,
          spHttpClient: this.context.spHttpClient,
          context: this.context,
          reload: _reRender.bind(this),
          currentUserRole: this.currentUserRole,
          userAndFoto: {}
        }
      );

      if (element.props.config.logPermCenterVars) {console.log("config: ", element.props.config);}

      ReactDom.render(element, this.domElement);
    }
    catch (error) {console.log(error);}
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    // initial variables
    let autoConfig = [];
    let roleConfig = [];
    let config = [];
    let autoConfigName = 'You need to be owner to configure this web part';
    let roleName = null;
    let configName = null;
    let debugMode = [];
    let debugFlags = [];
    let recordFlags = [];
    let optimizationFlags = [];
    let recordOrImportFlag = [];
    
    // define description for web part in property pane with version, buildTimeStamp and link for doc website
    const descriptionPage1 = [
      PropertyPaneLabel('version', {  
        text: "Version "  + packageSolution['solution'].version,
      }),
      PropertyPaneLabel('buildTimeStamp', {  
        text: buildTimeStamp,
      }),
      PropertyPaneLink('link', {  
        href: 'https://sharepoint-permission-center.com',
        text: 'Documentation',
        target: '_blank',
      })
    ];

    //export api response
    const exportApiResponse = () => {
      this._reRenderAndRecordAndDownloadApiResponse();
      return null;
    };

    //import api response
    const importApiResponse = () => {
      this._rerenderWithImportedApiResponse();
      return null;
    };

    // if current user is allowed to edit web part
    if (this.allowEditProps) {
      
      // Property pane page 1
      // --------------------
    
      autoConfigName = null;
      autoConfig = [ 
        PropertyPaneToggle('configBasedOnPermissions', {
          label: 'Configuration based on permissions',
          onText: "On",
          offText: "Off"
        }),
        PropertyPaneLabel('currentUserRoleLable', {  
          text:'You are Site ' + this.currentUserRole + '.'
        })
      ];

      // set default role for dropDown menu
      if (!this.properties.selectedRoleForConfig) {
        this.properties.selectedRoleForConfig = "member";
      }
      
      // configBasedOnPermissions
      if (this.properties.configBasedOnPermissions) {
        
        roleConfig = [
          PropertyPaneDropdown('selectedRoleForConfig', {
            label: 'Configure web part for',
            options: [
              { key: 'owner', text: 'Site owners' },
              { key: 'member', text: 'Site members' },
              { key: 'visitor', text: 'Site visitors' },
            ]
          })
        ];

        // configBasedOnPermissions owner
        if (this.properties.selectedRoleForConfig == "owner") {
          
          config = [
            PropertyPaneToggle('ownerShowMenu', {
              label: 'Show SharePoint menu',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowTabUsers', {
              label: 'Show Users tab',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowTabHidden', {
              label: 'Show Hidden groups tab',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowCardGroup', {
              label: 'Show group cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowCardUser', {
              label: 'Show user cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowCardButtons', {
              label: 'Show card edit buttons',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowCardUserLinks', {
              label: 'Show links in user cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowAdmins', {
              label: 'Show Site Admins',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowOwners', {
              label: 'Show Site Owners',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowMembers', {
              label: 'Show Site Members',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('ownerShowDirectAccess', {
              label: 'Show Access given directly',
              onText: "On",
              offText: "Off",
            })
          ];
        } 
        // configBasedOnPermissions member
        else if (this.properties.selectedRoleForConfig == "member") {

          config = [
            PropertyPaneToggle('memberShowMenu', {
              label: 'Show SharePoint menu',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowTabUsers', {
              label: 'Show Users tab',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowTabHidden', {
              label: 'Show Hidden groups tab',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowCardGroup', {
              label: 'Show group cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowCardUser', {
              label: 'Show user cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowCardButtons', {
              label: 'Show card edit buttons',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowCardUserLinks', {
              label: 'Show links in user cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowAdmins', {
              label: 'Show Site Admins',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowOwners', {
              label: 'Show Site Owners',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowMembers', {
              label: 'Show Site Members',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('memberShowDirectAccess', {
              label: 'Show Access given directly',
              onText: "On",
              offText: "Off",
            })
          ];
        } 
        // configBasedOnPermissions visitor
        else if (this.properties.selectedRoleForConfig == "visitor") {

          config = [
            PropertyPaneToggle('visitorShowMenu', {
              label: 'Show SharePoint menu',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowTabUsers', {
              label: 'Show Users tab',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowTabHidden', {
              label: 'Show Hidden groups tab',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowCardGroup', {
              label: 'Show group cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowCardUser', {
              label: 'Show user cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowCardButtons', {
              label: 'Show card edit buttons',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowCardUserLinks', {
              label: 'Show links in user cards',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowAdmins', {
              label: 'Show Site Admins',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowOwners', {
              label: 'Show Site Owners',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowMembers', {
              label: 'Show Site Members',
              onText: "On",
              offText: "Off",
            }),
            PropertyPaneToggle('visitorShowDirectAccess', {
              label: 'Show Access given directly',
              onText: "On",
              offText: "Off",
            })
          ];
        }
      }

      // configBasedOnPermissions off
      else {
        
        config = [
          PropertyPaneToggle('manualShowMenu', {
            label: 'Show SharePoint menu',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowTabUsers', {
            label: 'Show Users tab',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowTabHidden', {
            label: 'Show Hidden groups tab',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowCardGroup', {
            label: 'Show group cards',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowCardUser', {
            label: 'Show user cards',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowCardButtons', {
            label: 'Show card edit buttons',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowCardUserLinks', {
            label: 'Show links in user cards',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowAdmins', {
            label: 'Show Site Admins',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowOwners', {
            label: 'Show Site Owners',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowMembers', {
            label: 'Show Site Members',
            onText: "On",
            offText: "Off",
          }),
          PropertyPaneToggle('manualShowDirectAccess', {
            label: 'Show Access given directly',
            onText: "On",
            offText: "Off",
          })
        ];
      }
      
      // Property pane page 2
      // --------------------
      debugMode = [ 
        PropertyPaneToggle('debugMode', {
          label: 'Debug mode',
          onText: "On",
          offText: "Off"
        })
      ];

      // set default for export/import dropDown menu
      if (!this.properties.exportOrImportDropdown) {
        this.properties.exportOrImportDropdown = "export";
      }


      // debugmode
      if (this.properties.debugMode) {
          
        // performanceOptimization preparation
        let preloadAzureGroupsAmount: any;
        if (this.properties.preloadAzureGroups) {
          preloadAzureGroupsAmount = PropertyPaneToggle('preloadAzureGroupsAmount', {
            label: 'Preload Azure groups amount 100/1000',
            onText: "1000",
            offText: "100"
          });
        }
        else {
          preloadAzureGroupsAmount = PropertyPaneLabel('emptyLabel', {
            text: ""
          });
        }
        // --------
        debugFlags = [
          PropertyPaneToggle('logState', {
            label: 'Log state to console',
            onText: "On",
            offText: "Off"
          }),
          PropertyPaneToggle('throwErrors', {
            label: 'Throw errors',
            onText: "On",
            offText: "Off"
          }),
          PropertyPaneToggle('logErrors', {
            label: 'Log errors to console',
            onText: "On",
            offText: "Off"
          }),
          PropertyPaneToggle('logPermCenterVars', {
            label: 'Log variables of main component',
            onText: "On",
            offText: "Off"
          }),
          PropertyPaneToggle('logComponentVars', {
            label: 'Log variables of other components',
            onText: "On",
            offText: "Off"
          }),
          PropertyPaneToggle('disableAnimateHeightUserCard', {
            label: 'Disable animate height for user card',
            onText: "On",
            offText: "Off"
          }),
          PropertyPaneToggle('preloadAzureGroups', {
            label: 'Preload Azure groups',
            onText: "On",
            offText: "Off"
          }),
          preloadAzureGroupsAmount,
          PropertyPaneToggle('exportOrImportApiResponse', {
            label: 'Export or import API response',
            onText: "On",
            offText: "Off"
          }),
        ];

        if (this.properties.exportOrImportApiResponse) {
          recordFlags = [
            PropertyPaneDropdown('exportOrImportDropdown', {
              label:'',
              options: [
                { key: 'export', text: 'Export' },
                { key: 'import', text: 'Import' }
              ],
            })
          ];
          if (this.properties.exportOrImportDropdown == "export") {
            recordOrImportFlag = [
              PropertyPaneButton('exportApiResponseButton', {
                text: "Record and Download",
                buttonType: PropertyPaneButtonType.Normal,
                onClick: exportApiResponse
              })
            ];
          }
          else if (this.properties.exportOrImportDropdown == "import") {
            recordOrImportFlag = [
              PropertyPaneTextField('importApiResponseData', {
                label: 'Paste content of apiResponse.json',
                multiline: true,
                resizable: true,
              }),
              PropertyPaneButton('importApiResponseButton', {
                text: "Reload with imported API response",
                buttonType: PropertyPaneButtonType.Normal,
                onClick: importApiResponse
              })
              
            ];
          }
        }
      }
    }
    

    return {
      pages: [
        {
          header: {
            description: ""
          },
          groups: [
            {
              groupName: '',
              groupFields: descriptionPage1
            },
            {
              groupName: autoConfigName,
              groupFields: autoConfig
            },
            {
              groupName: roleName,
              groupFields: roleConfig
            },
            {
              groupName: configName,
              groupFields: config
            }
          ]
        },
        {
          header: {
            description: "Debug page"
          },
          groups: [
            {
              groupName: '',
              groupFields: debugMode
            },
            {
              groupName: '',
              groupFields: debugFlags
            },
            {
              groupName: '',
              groupFields: recordFlags
            },
            {
              groupName: '',
              groupFields: recordOrImportFlag
            }
          ]
        }
      ]
    };
  }
}