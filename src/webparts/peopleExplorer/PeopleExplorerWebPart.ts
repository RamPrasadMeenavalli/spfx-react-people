import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneChoiceGroup,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'PeopleExplorerWebPartStrings';
import  { PeopleExplorer } from './components/PeopleExplorer';
import { IPeopleExplorerProps } from './components/IPeopleExplorerProps';
import { graph } from "@pnp/graph";
import { initializeIcons } from 'office-ui-fabric-react';
import { HandleBarTemplates } from './HandleBarTemplates';

initializeIcons();

export interface IPeopleExplorerWebPartProps {
  title: string;
  people: any[];
  template: string;
  customTemplate: string;
}

export default class PeopleExplorerWebPart extends BaseClientSideWebPart<IPeopleExplorerWebPartProps> {

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      graph.setup({
        spfxContext: this.context
      });
    });
  }

  public render(): void {
    const template = this.properties.template == "custom" ? 
    this.properties.customTemplate : 
    HandleBarTemplates.getTemplate(this.properties.template);
    console.debug(template);
    const element: React.ReactElement<IPeopleExplorerProps> = React.createElement(
      PeopleExplorer,
      {
        title: this.properties.title,
        displayMode: this.displayMode,
        context: this.context,
        updateTitle: (value:string) => {
          this.properties.title = value;
        },
        people:this.properties.people,
        updatePeople: (values:any[]) => {
          this.properties.people = values;
        },
        template: template,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneChoiceGroup('template', {
                  label:"Template type",
                  options:[{
                    key: 'simple',
                    text: 'Simple',
                    checked: this.properties.template == 'simple'
                  },
                  {
                    key: 'detailed',
                    text: 'Detailed',
                    checked: this.properties.template == 'detailed'
                  },
                  {
                    key: 'custom',
                    text: 'Custom',
                    checked: this.properties.template == 'custom'
                  }
                ]
                }),

                PropertyPaneTextField('customTemplate',{
                  disabled: this.properties.template != "custom",
                  multiline: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
