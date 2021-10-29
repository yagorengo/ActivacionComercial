import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { setup as pnpSetup } from '@pnp/common';
import * as strings from 'AddCuitWebPartStrings';
import AddCuit from './components/AddCuit';
import { IAddCuitProps } from './components/IAddCuitProps';

export interface IAddCuitWebPartProps {
  description: string;
}

export default class AddCuitWebPart extends BaseClientSideWebPart<IAddCuitWebPartProps> {
  public onInit(): Promise<void> {

   

    pnpSetup({
      spfxContext: this.context
    });

    return Promise.resolve();
  }
  public render(): void {
    const element: React.ReactElement<IAddCuitProps> = React.createElement(
      AddCuit,
      {
        description: this.properties.description,
        context:this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

 

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
