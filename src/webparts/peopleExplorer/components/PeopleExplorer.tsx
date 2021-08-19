import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from './PeopleExplorer.module.scss';
import { IPeopleExplorerProps } from './IPeopleExplorerProps';
import { WebPartTitle } from '@pnp/spfx-controls-react';
import { PeoplePicker } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { DisplayMode } from '@microsoft/sp-core-library';
import { PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { GridLayout } from "@pnp/spfx-controls-react/lib/GridLayout";
import { Icon, IconButton, IPersonaProps, ISize, Shimmer, Text } from 'office-ui-fabric-react';
import { graph } from '@pnp/graph';
import '@pnp/graph/users';
import * as Handlebars from 'handlebars';


export const PeopleExplorer:React.FC<IPeopleExplorerProps> = (props) => {

  const [displayMode, setDisplayMode] = useState<DisplayMode>(props.displayMode);
  const [showPicker, setShowPicker] = useState<boolean>(true);
  const [people, setPeople] = useState<any[]>(props.people);

  const handleBarTemplate = Handlebars.compile(props.template);

  useEffect(() => {
    setDisplayMode(props.displayMode);
    let _p = [...people];
    setPeople(_p);
  }, [props.displayMode, props.template]);

  const onPeopleSelected = async (values:IPersonaProps[]) => {
    if(values && values.length > 0) {
      let _allPeople = [...people];
      const _person = values[0];
      const index = _allPeople.push({...values[0], loading:true});
      setPeople(_allPeople);
      setShowPicker(false);

      const persons = await graph.me.people.search(_person.secondaryText).get();
      const user = persons && persons.length > 0 ? persons[0] : {};
      const _p = {
        imageInitials: _person.imageInitials,
        imageUrl: _person.imageUrl,
        mail: _person.secondaryText,
        ...user
      }

      _allPeople.splice(index -1, 1, _p);
      setPeople(_allPeople);
      props.updatePeople(_allPeople);
      setShowPicker(true);
    }
  }

  const removePerson = (id:string) => {
    let _p = [...people];
    let _index = -1;
    _p.forEach((p, index) => {
      if(p.id == id) {
        _index = index
      }
    });
    if(_index != -1) {
      _p.splice(_index, 1);
      setPeople(_p);
      props.updatePeople(_p);
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
      <div dangerouslySetInnerHTML={{ __html:handleBarTemplate({...item, styles:styles}) }} />
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
        ariaLabel="List of content, use right and left arrow keys to navigate, arrow down to access details."
        items={displayMode == DisplayMode.Edit && showPicker ? [...people, "newitem"] : people}
        onRenderGridItem={(item: any, finalSize: ISize, isCompact: boolean) => onRenderGridItem(item, finalSize, isCompact)}
      >
      </GridLayout>
  </div>);
};