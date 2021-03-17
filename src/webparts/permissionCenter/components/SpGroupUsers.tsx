import * as React from 'react';
import {memo} from 'react';

import cssStyles from './PermissionCenter.module.scss';
import UserContainer from './UserContainer';

const showLogs = false;

type Props = {
    spGroupEntry: string;
    state: any;
    props;
};

const SpGroupContainer: React.FC<Props> = ({ spGroupEntry, state, props }) => {
  
  try {
      
    // get userEntries of spGroup.users
    const userEntryArray = state.spGroups[spGroupEntry].users ;
    
    // make array of userobjects from state.users to filter and sort them after 
    // and keep the connection from userEntry to get the order of userEntries and pass it to userContainer
    const userObjectArray = Object.keys(state.users).reduce((arr, key)=>{
      const subObj = {[key]: state.users[key]};
      return arr.concat(subObj);
    }, []);
  
    // filter userObjectArray with userEntries of userEntryArray
    const usersFilteredArray = userObjectArray.filter( i => userEntryArray.includes( Object.keys(i)[0] ) );
    
    // sort usersArray by user name
    const usersSortedArray = usersFilteredArray.sort((a, b) => {
      const item1 = Object.values(a)[0]["name"];
      const item2 = Object.values(b)[0]["name"];
      return item1.localeCompare(item2);
    });

    // make userContainerArray (Jsx Array)
    const userContainerArray = usersSortedArray.map(
      (userObjectItem)=> {
        const userEntry = Object.keys(userObjectItem)[0];
        return <UserContainer state={state} props={props} userEntry={userEntry} showPermissions={false} showPermissionsDirectAccess={state.spGroups[spGroupEntry].displayName==="Access given directly"} />;
      }
    );

    // check if usersArray has users
    const noUsers =  <div>No users</div>;
    let usersArrayHasContent = false;
    if (userEntryArray[0]) {
      usersArrayHasContent = true;
    }

    // check if user has access to view members of sp group
    let hasNoAccess = false;
    if (state.spGroups[spGroupEntry].users == 'no access') {
      hasNoAccess = true;
    }
    const noAccess = <div>No access</div>;
    
    // userLoader
    const userLoader = <div className={cssStyles.placeholderItem}>Loading users</div>;
    let isLoading = false;
    if (state.spGroups[spGroupEntry].users[0] === null) {
      isLoading = true;
    }
    
    // return userContainerArray or noUsers or userLoader
    return <div className={cssStyles.userContainer}>
      {hasNoAccess ? noAccess : usersArrayHasContent ? userContainerArray : isLoading ? userLoader : noUsers }
    </div>;
    
  } catch (error) {
    if (showLogs) {console.log(error);}
    if (props.throwErrors) {throw error;}
  }
};

export default memo(SpGroupContainer);