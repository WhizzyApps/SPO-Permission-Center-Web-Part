import { SPHttpClient } from '@microsoft/sp-http';

import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPermissionCenterProps {
  
  siteCollectionURL?: string;
  config?: {
    configBasedOnPermissions: boolean,
    showMenu: boolean,
    showTabUsers: boolean,
    showTabHidden: boolean,
    showCardGroup: boolean,
    showCardUser: boolean,
    showCardButtons: boolean,
    showCardUserLinks: boolean,
    showAdmins: boolean,
    showOwners: boolean,
    showMembers: boolean,
    showDirectAccess: boolean,
    logState: boolean,
    logErrors: boolean,
    throwErrors: boolean,
    logPermCenterVars: boolean,
    logComponentVars: boolean,
    disableAnimateHeightUserCard: boolean,
    preloadAzureGroups: boolean,
    preloadAzureGroupsAmount: boolean,
    exportOrImportApiResponse: boolean,
    exportApiResponse: boolean,
    importApiResponse: boolean,
    importApiResponseData
    
  };
  spHttpClient?: SPHttpClient;
  context?: WebPartContext;
  reload?;
  currentUserRole?;
  userAndFoto?;
}
