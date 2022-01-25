import * as React  from 'react';
import { useState, memo } from "react";
import AnimateHeight from 'react-animate-height';

import UserCard from './UserCard';
import cssStyles from './PermissionCenter.module.scss';

type Props = {
  state: any;
  props;
  userEntry: string;
  showPermissions: boolean;
  showPermissionsDirectAccess: boolean;
};

const UserContainer: React.FC<Props> = ({ state, props, userEntry, showPermissions, showPermissionsDirectAccess}) => {
  try {
    // for expanding user card
    const [isUserCardExpanded, setUserCardExpanded] = useState(false);
    const [cardHeight, setCardHeight] = useState('0');
    const _toggleExpandUserCard = () => {
      if (isUserCardExpanded) {setTimeout(()=>setUserCardExpanded (false), 100);}
      else {setUserCardExpanded (true);}
      setCardHeight(cardHeight == '0' ? 'auto' : '0' );
    };

    let showUserCard = true;
    /* 
    if (state.users[userEntry].name.includes("Administrator")) {
      showUserCard = false;
    }  */

    const userCardComponent = 
      <UserCard state={state} props={props} userEntry={userEntry} />;

    const _toggleExpandUserCardAlternative = () => {
      setUserCardExpanded(!isUserCardExpanded);
    };
    return (
      <div className={cssStyles.userRowContainer}>
        
        <div className={cssStyles.userRow}  > 
          {props.config.showCardUser
            ? <div 
              className={showUserCard ? isUserCardExpanded ? cssStyles.userActive : cssStyles.user : cssStyles.userInvalid} 
              onClick={showUserCard && _toggleExpandUserCard}
              title = {showUserCard && (isUserCardExpanded ? "Collapse user card" : "Expand user card") }
              aria-expanded={ cardHeight !== '0' }
              aria-controls= 'userCard'
              >
              {state.users[userEntry].name}
              {/* for user tab */}
              {showPermissions 
                && (state.users[userEntry].permissionLevel[0] 
                  ? (<span> {
                      state.users[userEntry].permissionLevel[0]._values // for IE11, because "...new Set" creates a different object (in PermissionCenter.tsx)
                      ? '(' + state.users[userEntry].permissionLevel[0]._values.join(", ") + ')'
                      : '(' + state.users[userEntry].permissionLevel.join(", ") + ')'
                    } </span>)
                  : ""
                )
              }
              {/* for groups tab and hidden groups tab */}
              {showPermissionsDirectAccess 
                && ( <span> {
                  state.users[userEntry].permissionLevelDirectAccess && state.users[userEntry].permissionLevelDirectAccess[0]
                  && state.users[userEntry].permissionLevelDirectAccess[0]._values // of IE11 because "...new Set" creates a different object (in PermissionCenter.tsx)
                    ? '(' + state.users[userEntry].permissionLevelDirectAccess[0]._values.join(", ") + ')'
                    : '(' + state.users[userEntry].permissionLevelDirectAccess.join(", ") + ')'
                }</span>) 
              }
            </div>

            : <div>
              {state.users[userEntry].name}
            </div>
          }
        </div>
        
        {showUserCard && (
          props.config.disableAnimateHeightUserCard
          ? (
            <div onClick={_toggleExpandUserCardAlternative}>
              {isUserCardExpanded && (
                props.config.logComponentVars && console.log("isUserCardExpanded", isUserCardExpanded),
                props.config.logComponentVars && console.log("userCardComponent", userCardComponent),
                userCardComponent)}
            </div>
          )
          : <AnimateHeight 
            id='userCard'
            duration={100}
            height={cardHeight}
            >
            {isUserCardExpanded && (
              props.config.logComponentVars && console.log("isUserCardExpanded", isUserCardExpanded, "cardHeight", cardHeight),
              props.config.logComponentVars && console.log("userCardComponent", userCardComponent),
              userCardComponent
            )}

          </AnimateHeight>
        )}
      </div>
    );
  }
  catch (error) {
    // console.log(error);
    if (props.config.throwErrors) {throw error;}
  }
};

export default memo(UserContainer);