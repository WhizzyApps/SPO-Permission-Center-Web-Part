import * as React  from 'react';
import {memo} from 'react';

import cssStyles from './PermissionCenter.module.scss';
import UserContainer from './UserContainer';

const showLogs = false;
type Props = {
    state: any;
    props;
    isGroupsLoading: boolean;
};

const AllUsers: React.FC<Props> = ({ state, props, isGroupsLoading }) => {
    
  try {
    
    // make array of user names
    let userNameArray = [];
    Object.keys(state.users).map((userEntry) => {
      userNameArray.push(state.users[userEntry].name);
    });
    
    // sort userNameArray by user name
    userNameArray.sort();

    // Jsx Array of UserContainer
    const userContainerArray = userNameArray.map((userName) => {
      let userEntry;
      // get userEntry of user name
      Object.keys(state.users).map(
        (userEntryItem) => {
          if (state.users[userEntryItem].name === userName) {
            userEntry = userEntryItem;
          }
        }
      );
      return <UserContainer state={state} props={props} userEntry={userEntry} showPermissions={true} showPermissionsDirectAccess={false}/>;
    });
    
    const userLoader = <div className={cssStyles.placeholderItem}>Loading users</div>;

    return <div className={cssStyles.allUserContainer}>
      {isGroupsLoading ? userLoader : userContainerArray}
    </div>;

  } catch (error) {
    if (showLogs) {console.log(error);}
    if (props.throwErrors) {throw error;}
  }
};

export default memo(AllUsers); 