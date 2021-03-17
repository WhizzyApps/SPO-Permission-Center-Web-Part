import * as React  from 'react';
import {memo, useState} from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { SPHttpClient, ISPHttpClientOptions} from '@microsoft/sp-http';

import AnimateHeight from 'react-animate-height';

import SpGroupCard from './SpGroupCard';
import SpGroupUsers from './SpGroupUsers';
import cssStyles from './PermissionCenter.module.scss';

const showLogs = false;
const logErrors = false;

type Props = {
  spGroupEntry;
  state;
  props;
  hideGroup;
  isGroupsLoading;
};


const SpGroupContainer: React.FC<Props> = ({ spGroupEntry, state, props, hideGroup, isGroupsLoading}) => {
  
  try {
      
    // get data from SharePoint REST Api
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
        if (logErrors) {console.log(error);}
        if (props.throwErrors) {throw error;}
        error['value'] = [];
        error['status'] = "error";
        return error;
      }
    };

    // do not display group card for admins
    let isNoAdminsGroup = true;
    if (state.spGroups[spGroupEntry].displayName === "Site Admins") {
      isNoAdminsGroup = false;
    }
    // do not display group card for admins
    let isNoDirectAccessGroup = true;
    if (state.spGroups[spGroupEntry].displayName === "Access given directly") {
      isNoDirectAccessGroup = false;
    }
    // case of Admins, click opens permissions page
    const _permissionsPage = (event) => {
        window.open(props.siteCollectionURL + '/_layouts/15/user.aspx');
    };

    // open document / folder
    const _openItem = () => {
      // decide if folder or item
      const pathSplit = state.spGroups[spGroupEntry].item.parentLink.split('/');
      const pathLastPart = pathSplit[pathSplit.length-1];
      const isFolder = pathLastPart.includes('.');

      // if folder, open folder
      if (isFolder) {
        const pathFolder = `${props.siteCollectionURL}/${state.spGroups[spGroupEntry].item.path}` ;
        const rootPath = pathFolder.slice(0,-state.spGroups[spGroupEntry].item.title.length);
        const urlFolder = `${rootPath}Forms/AllItems.aspx?id=${encodeURIComponent(pathFolder)}`;
        window.open(urlFolder);
      } 
      // if document, open document
      else {
        window.open(`${props.siteCollectionURL}/_layouts/15/Doc.aspx?sourcedoc=%7B${state.spGroups[spGroupEntry].item.guid}%7D`);
      }
    };

    // open item permissions
    const _openItemPermissions = async () => {
      const guid = state.spGroups[spGroupEntry].item.guid;
      const listId = state.spGroups[spGroupEntry].item.listId;
      const requestUrl = `${props.siteCollectionURL}/_api/web/Lists('${listId}')/GetItemByUniqueId('${guid}')?$Select=Id`;
      const getPermResult = await _spApiGet(requestUrl);
      const spId = getPermResult["Id"];
      const permissionsPage = `${props.siteCollectionURL}/_layouts/15/user.aspx?obj=%7B${listId}%7D,${spId},LISTITEM`;
      window.open(permissionsPage);
    };

    // for expanding group content
    const [contentIsExpanded, setContentIsExpanded] = useState(true);
    const [contentHeight, setContentHeight] = useState('auto');
    const _toggleExpandGroupContent = () => {
      setContentIsExpanded (!contentIsExpanded);
      setContentHeight(contentHeight == '0' ? 'auto' : '0' );
    };
    // toggle between expanding group card and users
    const [isCardExpanded, setCardIsExpanded] = useState(false);
    const [cardHeight, setCardHeight] = useState('0');
    const [usersHeight, setUsersHeight] = useState('auto');
    const _toggleExpandCard = () => {
      if (isCardExpanded) 
        {setTimeout(()=>setCardIsExpanded (false),100);}
      else {setCardIsExpanded (true);}
      setCardHeight(cardHeight == '0' ? 'auto' : '0' );
      setUsersHeight(usersHeight == '0' ? 'auto' : '0' );
    };
    // hide group because of config
    if (
      ((spGroupEntry == 'spGroup1') && (props.config.showAdmins == false))
      || ((spGroupEntry == 'spGroup2') && (props.config.showOwners == false))
      || ((spGroupEntry == 'spGroup3') && (props.config.showMembers == false))
      || ((state.spGroups[spGroupEntry].displayName == 'Access given directly') && (props.config.showDirectAccess == false))
    ) {
      hideGroup = true;
    }

    return (
      <> {!hideGroup && (
        <div>
          <div className={ cssStyles.groupHeader }>
              
            {/* content expand icon */}
            <div className={cssStyles.iconContainer} >
              <Icon 
                iconName={ "ChevronDownMed"} 
                className={`${cssStyles.icon} ${contentIsExpanded? cssStyles.iconInit: cssStyles.iconTurned}` } 
                aria-expanded={ contentHeight !== '0' }
                aria-controls='content'
                onClick= {()=>_toggleExpandGroupContent()}
              />
            </div>

            {/* sp group title */}
            <div 
              className={cssStyles.groupTitleContainer } 
              >

              {/* group name */}
              <span 
                className={cssStyles.groupTitle } 
                title={state.spGroups[spGroupEntry].description}
                >
                {state.spGroups[spGroupEntry].displayName}
              </span>

              {/* subtitle (just in case of sharing groups) */}
              {/* state.spGroups[spGroupEntry].subTitle exists only on sharing groups */}
              <span className={cssStyles.groupSubTitle }> 
                {state.spGroups[spGroupEntry].subTitle && state.spGroups[spGroupEntry].subTitle}
              </span>

              {/* for sharing group: link to shared item */}
              {/* state.spGroups[spGroupEntry].item exists only on sharing groups */}
              
              {state.spGroups[spGroupEntry].item && (
                <>
                  <div 
                    onClick={()=> _openItem()}
                    className={cssStyles.groupTitleLinkOpen}
                    title= 'Open Document/Folder'
                  >Open item</div>
                  <div 
                    onClick={()=> _openItemPermissions()}
                    className={cssStyles.groupTitleLinkPerm}
                    title= 'Open permissions page of item'
                  >Permissions</div>
                </>
              )}

              {/* permissions */}
              {state.spGroups[spGroupEntry].permissionLevel && state.spGroups[spGroupEntry].permissionLevel[0] && (

                <div className={ cssStyles.permissions } > 
                  ({
                    state.spGroups[spGroupEntry].permissionLevel.map(
                      (permissionLevelItem, permissionLevelIndex)=>
                        <span 
                          className={((permissionLevelItem !== "Administrator") && (permissionLevelItem !== "Limited Access")) ? cssStyles.permissionLevel : null} 
                          onClick={
                            (permissionLevelItem !== "Administrator") && (permissionLevelItem !== "Limited Access")
                            && (()=>window.open(`${props.siteCollectionURL}/_layouts/15/editrole.aspx?role=${permissionLevelItem}`))
                          }
                          title = {((permissionLevelItem !== "Administrator") && (permissionLevelItem !== "Limited Access")) && "Open classic permission level page"}
                          >
                          {permissionLevelItem +( permissionLevelIndex < state.spGroups[spGroupEntry].permissionLevel.length -1 ? ", " : '')}
                        </span>
                    )
                  })
                </div>
              )}

            </div>

            {/* cardOrUsers expand icon */}
            {props.config.showCardGroup &&
              <div
                className={` ${contentIsExpanded ? isCardExpanded ? cssStyles.editIconActive : cssStyles.editIcon : cssStyles.editIconHidden}`}
                onClick= {(event)=>{(isNoAdminsGroup && isNoDirectAccessGroup) ? _toggleExpandCard() : _permissionsPage(event);}} // open permission page instead if admins or Access given directly
                aria-expanded={ cardHeight !== '0' }
                aria-controls= 'cardAndUsers'
                title = {isNoAdminsGroup ? (isNoDirectAccessGroup ? (isCardExpanded ? "Show users" : "Show group card to edit group") : 'Open classic permissions page to manage members with Access given directly') : "Open classic permissions page to manage admins"} // change tooltip if admins or Access given directly
                >
                  <Icon 
                    className={cssStyles.cardAndUsersIcon} 
                    iconName={(isNoAdminsGroup && isNoDirectAccessGroup) ? "SingleColumnEdit" : "FileSymlink"} // change icon if admins or Access given directly
                  /> 
              </div>
            }
          </div>
  
          {/* group content */}
          <AnimateHeight 
            id='content'
            duration={100}
            height={contentHeight}
          >
            {/* group card */}
            {(isNoAdminsGroup && isNoDirectAccessGroup) && ( // do not display group card if admins or Access given directly
              <AnimateHeight 
                id='cardAndUsers'
                duration={100}
                height={cardHeight}
              >
                <div className={ cssStyles.groupCard }>
                  {isCardExpanded && <SpGroupCard spGroup={state.spGroups[spGroupEntry] } state={state} props={props} />}
                </div>
                
              </AnimateHeight>
            )}
            
            {/* group users */}
            <AnimateHeight 
              id='cardAndUsers'
              duration={100}
              height={usersHeight}
            >
              <SpGroupUsers spGroupEntry={spGroupEntry} state={state} props={props} /> 
            </AnimateHeight>

          </AnimateHeight>
        </div>
      )}</>
    );
    
  } catch (error) {
    if (showLogs) {console.log(error);}
    if (props.throwErrors) {throw error;}
  }
};

export default memo(SpGroupContainer);