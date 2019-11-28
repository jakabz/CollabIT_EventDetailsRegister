import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'EventRegistrationWebPartStrings';
import EventRegistration from './components/EventRegistration';
import { IEventRegistrationProps } from './components/IEventRegistrationProps';
import { sp } from "@pnp/sp";

export interface IEventRegistrationWebPartProps {
  eventId: string;
  eventRegistrationItem: any;
  eventItem: any;
  lang: string;
}


export default class EventRegistrationWebPart extends BaseClientSideWebPart<IEventRegistrationWebPartProps> {

  private eventListId:string;
  private eventListItems:any[];
  private eventRegistrationItem:any;
  private eventItem:any;
  private _dropdownOptionsItems: IPropertyPaneDropdownOption[] = [];
  private _dropdownOptionsLang: IPropertyPaneDropdownOption[] = [
    {
      key: 'hu-HU',
      text: 'Hungarian'
    },{
      key: 'en-US',
      text: 'English'
    }
  ];

  public render(): void {
    
    // GET EVENT LIST DATA
    sp.web.lists.getByTitle("Events").get().then((list:any[]) => {
      this.eventListId = list['Id'];
      
      // GET EVENT LIST ITEMS
      sp.web.lists.getByTitle("Events").items.filter("EventPage eq '"+location.pathname+"'").get().then((items:any[]) => {
        this.eventListItems = items;
        items.map((item,i) => {
          if(item.Id == this.properties.eventId){
            this.eventItem = item;
          }
        });

        sp.web.lists.getByTitle("Event Registration").items.filter("EventID eq "+this.properties.eventId).get().then((regItem:any[]) => {
          this.eventRegistrationItem = regItem[0];

          const element: React.ReactElement<IEventRegistrationProps > = React.createElement(
            EventRegistration,
            {
              eventId: this.properties.eventId,
              eventListId: this.eventListId,
              eventRegistrationItem: this.eventRegistrationItem,
              eventItem: this.eventItem,
              self: this,
              refresh: this._refresh,
              lang: this.properties.lang
            }
          );
          this._dropdownOptionsItems = this.eventListItems.map((listItem:any) => {
            return {
              key: listItem.Id,
              text: listItem.Title
            };
          });
          ReactDom.render(element, this.domElement);
        });
      });
    });
  }

  public _refresh(n) {
    if(n){
      this.render();
    }
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
                PropertyPaneDropdown('eventId', {
                  label: strings.listItemFieldLabel,
                  options: this._dropdownOptionsItems
                }),
                PropertyPaneDropdown('lang', {
                  label: strings.langItemFieldLabel,
                  options: this._dropdownOptionsLang
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
