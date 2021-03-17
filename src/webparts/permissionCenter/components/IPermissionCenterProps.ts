import { SPHttpClient } from '@microsoft/sp-http';

import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPermissionCenterProps {
  
  siteCollectionURL?: string;
  config?: {
    auto: boolean,
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
  };
  throwErrors?: boolean;
  spHttpClient?: SPHttpClient;
  context?: WebPartContext;
  reload?;
  mode?;
  userAndFoto?;
}
