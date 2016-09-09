import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-client-preview';

import * as strings from 'spfxReactTodoListStrings';
import SpfxReactTodoList, { ISpfxReactTodoListProps } from './components/SpfxReactTodoList';
import { ISpfxReactTodoListWebPartProps } from './ISpfxReactTodoListWebPartProps';

export default class SpfxReactTodoListWebPart extends BaseClientSideWebPart<ISpfxReactTodoListWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    const element: React.ReactElement<ISpfxReactTodoListProps> = React.createElement(SpfxReactTodoList, {
      listName: this.properties.listName,
      hideFinishedTasks: this.properties.hideFinishedTasks,
      context: this.context
    });

    ReactDom.render(element, this.domElement);
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'Type the name of your TODO list'
                }),
                PropertyPaneToggle('hideFinishedTasks', {
                  label: 'Hide finished tasks'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
