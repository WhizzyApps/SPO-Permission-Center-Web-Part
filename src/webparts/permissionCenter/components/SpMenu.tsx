import * as React from 'react';
import { memo } from 'react';
import { IContextualMenuProps } from 'office-ui-fabric-react/lib/ContextualMenu';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { useConst } from '@uifabric/react-hooks';

import { IPermissionCenterProps } from './IPermissionCenterProps';

export const SpMenu: React.FunctionComponent<IPermissionCenterProps> = (props) => {
  
  try {
      
    // because of some bug open the window inside not possible
    const _classicAccessRequestPage = () => {
      window.open(props.siteCollectionURL + '/Access%20Requests/pendingreq.aspx');
    };

    const menu = useConst<IContextualMenuProps>(() => ({
      shouldFocusOnMount: true,
      items: [
        { key: 'Classic permissions page', text: 'Classic permissions page', onClick: (()=>window.open(props.siteCollectionURL + '/_layouts/15/user.aspx')) },
        { key: 'Classic permissions levels page', text: 'Classic permissions levels page', onClick: (()=>window.open(props.siteCollectionURL + '/_layouts/15/role.aspx?Source=[SiteUrl]/_layouts/15/user.aspx'))},
        { key: 'Classic groups page', text: 'Classic groups page', onClick: (()=>window.open(props.siteCollectionURL + '/_layouts/15/groups.aspx'))},
        { key: 'Classic default groups page', text: 'Classic default groups page', onClick: (()=>window.open(props.siteCollectionURL + '/_layouts/15/permsetup.aspx'))},
        { key: 'Classic all site users page', text: 'Classic all site users page', onClick: (()=>window.open(props.siteCollectionURL + '/_layouts/15/people.aspx?MembershipGroupId=0'))},
        { key: 'Classic access request page', text: 'Classic access request page', onClick: (()=>_classicAccessRequestPage())},
      ],
    }));
    
    return <DefaultButton text="SharePoint" menuProps={menu} />;

  } catch (error) {
    // console.log(error);
    if (props.config.throwErrors) {throw error;}
  }
}; 

export default memo(SpMenu);
