import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './PeopleExplorer.module.scss';
import { IPeopleExplorerProps } from './IPeopleExplorerProps';
import { WebPartTitle } from '@pnp/spfx-controls-react';
import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { DisplayMode } from '@microsoft/sp-core-library';
import { PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { Icon, IconButton, IPersonaProps, ISize, Shimmer } from 'office-ui-fabric-react';
import { graph } from '@pnp/graph';
import '@pnp/graph/users';
import * as Handlebars from 'handlebars';
import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";
import { PnPClientStorage } from '@pnp/common';


export const PeopleExplorer:React.FC<IPeopleExplorerProps> = (props) => {

  const CACHE_KEY:string = `peopleExplorerData|${props.context.instanceId}|${props.context.pageContext.listItem.id}`;

  const [displayMode, setDisplayMode] = useState<DisplayMode>(props.displayMode);
  const [showPicker, setShowPicker] = useState<boolean>(true);
  const [people, setPeople] = useState<any[]>(props.people);
  const [peopleInfo, setPeopleInfo] = useState<any[]>([]);

  const handleBarTemplate = Handlebars.compile(props.template);
  const _storage:PnPClientStorage = new PnPClientStorage();

  useEffect(() => {
    setDisplayMode(props.displayMode);
    let _p = [...people];
    setPeople(_p);
  }, [props.displayMode, props.template]);

  useEffect(() => {
    // Get person details from Graph API
    // const persons = await graph.me.people.search(_person.secondaryText).get();
    // const user = persons && persons.length > 0 ? persons[0] : {};

    // Get person details from SP User Profile
    // Check if data available in session storage
    _storage.session.getOrPut(CACHE_KEY, () => {
      return new Promise((resolve, reject) => {
        let _peopleInfo = people.map(p => ({ ...p, loading:true}));
        setPeopleInfo([..._peopleInfo]);
        Promise.all(people.map((p, index) => {
          return sp.profiles.getPropertiesFor(p["loginName"]).then(user => {
            const userProps = [...user.UserProfileProperties];
            userProps.forEach(p => {
              user[p.Key] = p.Value;
            });
            user.UserProfileProperties = [];
    
    
            const _p = {
              imageInitials: p.imageInitials,
              imageUrl: p.imageUrl,
              mail: p.secondaryText,
              id: p.secondaryText,
              ...user
            }
            _peopleInfo[index] = _p;
          });
        })).then(() => {
          resolve(_peopleInfo);
        });
      });
    }).then((peopleInfo:any[]) => {
      setPeopleInfo([...peopleInfo]);
    });
  }, [])

  const onPeopleSelected = async (values:IPersonaProps[]) => {
    let _currentPeople = [...people];
    let _currentPeopleInfo = [...peopleInfo];
    if(values && values.length > 0) {
      const _person = values[0];
      const index = _currentPeople.push(values[0]);
      _currentPeopleInfo.push({...values[0], loading:true});
      setPeopleInfo(_currentPeopleInfo);
      setShowPicker(false);

      // Get person details from Graph API
      // const persons = await graph.me.people.search(_person.secondaryText).get();
      // const user = persons && persons.length > 0 ? persons[0] : {};

      // Get person details from SP User Profile
      let user = await sp.profiles.getPropertiesFor(_person["loginName"]);
      const userProps = [...user.UserProfileProperties];
      userProps.forEach(p => {
        user[p.Key] = p.Value;
      });
      user.UserProfileProperties = [];

      const _p = {
        imageInitials: _person.imageInitials,
        imageUrl: _person.imageUrl,
        mail: _person.secondaryText,
        id: _person.secondaryText,
        ...user
      }

      _currentPeople.splice(index -1, 1, _person);
      setPeople(_currentPeople);
      props.updatePeople(_currentPeople);

      _currentPeopleInfo.splice(index -1, 1, _p);
      setPeopleInfo(_currentPeopleInfo);

      // Clear the cache
      _storage.session.delete(CACHE_KEY);

      setShowPicker(true);
    }
  }

  const removePerson = (id:string) => {
    let _p = [...people];
    let _peopleInfo = [...peopleInfo];

    let _index = -1;
    _peopleInfo.forEach((p, index) => {
      if(p.id == id) {
        _index = index
      }
    });
    if(_index != -1) {
      _p.splice(_index, 1);
      _peopleInfo.splice(_index, 1);
      setPeople(_p);
      setPeopleInfo(_peopleInfo);
      props.updatePeople(_p);
      
      // Clear the cache
      _storage.session.delete(CACHE_KEY);

    }
  }

  const onRenderGridItem = (item: any, finalSize: ISize, isCompact: boolean): JSX.Element => {
    if(item == "newitem") {
      return <div
      className={styles.peopleCard}
      data-is-focusable={true}
      role="listitem"
      aria-label={item.text}
    ><PeoplePicker
      defaultSelectedUsers={[]}
      context={props.context}
      placeholder="Name or email address"
      personSelectionLimit={1}
      showtooltip={true}
      required={true}
      onChange={onPeopleSelected}
      showHiddenInUI={false}
      principalTypes={[PrincipalType.User]}
      resolveDelay={1000} /> </div>
    }
    return <div
      className={styles.peopleCard}
      data-is-focusable={true}
      role="listitem"
      aria-label={item.text}
    >
      {displayMode == DisplayMode.Edit && (
      <IconButton 
        onClick={() => {removePerson(item.id)}}
        className={styles.closeButton}>
          <Icon iconName="ChromeClose"></Icon>
        </IconButton>
      )}
      <div style={{marginBottom:"20px"}} dangerouslySetInnerHTML={{ __html:handleBarTemplate({...item, styles:styles}) }} />
      {item.loading && (
        <Shimmer style={{marginTop:"1rem"}}></Shimmer>
      )}
    </div>;
    
  }

  return(<div className={styles.peopleExplorer}>
    <WebPartTitle displayMode={props.displayMode}
      title={props.title}
      updateProperty={props.updateTitle} />

      <GridLayout
        listProps={{
          className: styles.listStyle,
        }}
        ariaLabel="List of content, use right and left arrow keys to navigate, arrow down to access details."
        items={displayMode == DisplayMode.Edit && showPicker ? [...peopleInfo, "newitem"] : peopleInfo}
        onRenderGridItem={(item: any, finalSize: ISize, isCompact: boolean) => onRenderGridItem(item, finalSize, isCompact)}
      >
      </GridLayout>
  </div>);
};