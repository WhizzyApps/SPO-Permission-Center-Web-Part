import * as React from 'react';
import * as ReactDom from 'react-dom';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IPropertyPaneConfiguration, PropertyPaneDropdown, PropertyPaneToggle, PropertyPaneLabel, PropertyPaneLink } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import PermissionCenter from './components/PermissionCenter';
import { IPermissionCenterProps } from './components/IPermissionCenterProps';

const throwErrors = false;
const showLogs = false;
const buildTimeStamp = "Build: 2021-03-05 20:23";
const packageSolution = require("../../../config/package-solution.json");

export interface IPermissionCenterWebPartProps {
  auto: boolean;
  role: string;

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
}

export default class PermissionCenterWebPart extends BaseClientSideWebPart <IPermissionCenterWebPartProps> {
  
  private allowEditProps = false;
  
  // get data from SharePoint REST Api
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
    let role;
    // get permission of current users
    const urlAdmin = this.context.pageContext.web.absoluteUrl + '/_api/web/currentuser/isSiteAdmin';
    const isSiteAdminResponse = await this._spApiGet(urlAdmin);
    const isSiteAdmin = isSiteAdminResponse['value'];
    if (isSiteAdmin == true) {
      role = "admin";
    } 
    else {
      const urlPerm = this.context.pageContext.web.absoluteUrl + `/_api/web/effectiveBasePermissions`;
      const permResponse = await this._spApiGet(urlPerm);
      let permArray = [];
      if (permResponse["Low"]) {
        permArray = this._convertUserPermissions (permResponse['Low'], permResponse['High']);
        if (permArray.includes("ManagePermissions")) {
          role = "owner";
        } else if (permArray.includes("EditListItems")) {
          role = "member";
        } else {
          role = "visitor";
        }
      }
      else {
        role = "visitor";
      }

    }
    return role;
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
  
  private mode;

  public async render() {

    try {

      if (!this.mode) {
        this.mode = await this._getUserPermissions();
      }
      
      if ((this.mode == "admin") || (this.mode == "owner") ) {
        this.allowEditProps = true;
      }
        
      const config = () => {
        if (this.properties.auto) {
          if (this.mode == 'visitor') {
            return {
              auto: this.properties.auto,
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
          } else if (this.mode == 'member') {
            return {
              auto: this.properties.auto,
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
          } else if (this.mode == 'owner') {
            return {
              auto: this.properties.auto,
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
          // admin mode
          else {
            return {
              auto: this.properties.auto,
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
        // manual mode
        else {
          return {
            auto: this.properties.auto,
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

      const _reRender = () => {
        element.props.userAndFoto = {};
        ReactDom.unmountComponentAtNode(this.domElement);
        ReactDom.render(element, this.domElement);
      };

      const element: React.ReactElement<IPermissionCenterProps> = React.createElement(
        PermissionCenter,
        {
          config: config(),
          throwErrors: throwErrors,
          siteCollectionURL: this.context.pageContext.web.absoluteUrl,
          spHttpClient: this.context.spHttpClient,
          context: this.context,
          reload: _reRender.bind(this),
          mode: this.mode,
          userAndFoto: {}
        }
      );


      if (showLogs) {console.log("config: ", element.props.config);}
      ReactDom.render(element, this.domElement);
    }
    catch (error) {console.log(error);}
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    
    let autoConfig = [];
    let roleConfig = [];
    let config = [];
    let autoConfigName = 'You need to be owner to configure this web part';
    let roleName = null;
    let configName = null;
    // set default role for dropDown menu
    if (!this.properties.role) {
      this.properties.role = "member";
    }

    // define description for web part in property pane with version, buildTimeStamp and link for doc website
    const descriptionName = "";
    const description = [
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

    // if current user is allowed to edit web part
    if (this.allowEditProps) {
      autoConfigName = null;
      autoConfig = [ 
        PropertyPaneToggle('auto', {
          label: 'Configuration based on permissions',
          onText: "On",
          offText: "Off"
        }),
        PropertyPaneLabel('currentUserRole', {  
          text:'You are Site ' + this.mode + '.'
        })
      ];

      // auto mode
      if (this.properties.auto) {
        
        roleConfig = [
          PropertyPaneDropdown('role', {
            label: 'Configure webpart for',
            options: [
              { key: 'owner', text: 'Site owners' },
              { key: 'member', text: 'Site members' },
              { key: 'visitor', text: 'Site visitors' },
            ]
          })
        ];

        // auto owner
        if (this.properties.role == "owner") {
          
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
        // auto member
        else if (this.properties.role == "member") {

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
        // auto visitor
        else if (this.properties.role == "visitor") {

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

      // manual mode
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
    }

    return {
      pages: [
        {
          header: {
            description: ""
          },
          groups: [
            {
              groupName: descriptionName,
              groupFields: description
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
        }
      ]
    };
  }
}