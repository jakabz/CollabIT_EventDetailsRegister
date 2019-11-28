import * as React from 'react';
import * as ReactDom from 'react-dom';
import styles from './EventRegistration.module.scss';
import { IEventRegistrationProps } from './IEventRegistrationProps';
import { IIconProps, PrimaryButton } from 'office-ui-fabric-react';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as strings from 'EventRegistrationWebPartStrings';
import { sp } from "@pnp/sp";

export default class EventRegistration extends React.Component<IEventRegistrationProps, {}> {

  public sign(self):void {
    ReactDom.render( <div className={ styles.eventRegistration }><Spinner className={styles.spinner} label="Please wait..." size={SpinnerSize.large} ariaLive="assertive" labelPosition="right" /></div>, self.domElement);
    sp.web.currentUser.get().then((user) => {
      const listPath = `${self.context.pageContext.web.serverRelativeUrl}/Lists/EventRegistration/EventID_${self.properties.eventId}_${self.eventItem.Title}`;
      sp.web.lists.getByTitle('Event Registration').addValidateUpdateItemUsingPath([
        { FieldName: 'Title', FieldValue: self.eventItem.Title },
        { FieldName: 'EventID', FieldValue: String(self.properties.eventId) },
        { FieldName: 'Person', FieldValue: JSON.stringify([{ "Key": user.LoginName }]) }
      ], listPath).then((resp) => {
        self.render();
      });
    });
    
  }
  public unsign(self):void {
    ReactDom.render( <div className={ styles.eventRegistration }><Spinner className={styles.spinner} label="Please wait..." size={SpinnerSize.large} ariaLive="assertive" labelPosition="right" /></div>, self.domElement);
    sp.web.lists.getByTitle('Event Registration').items.getById(self.eventRegistrationItem.Id).recycle().then(resp => {
      self.render();
    });
    alert('Successful unsubscribe! Please, delete the event from your outlook calendar.');
  }

  public saveEvent(self):void {
    var icsUrl = self.props.self.context.pageContext.web.serverRelativeUrl+'/_vti_bin/owssvr.dll?CS=109&Cmd=Display&List='+self.props.eventListId+'&CacheControl=1&ID='+self.props.eventId+'&Using=event.ics';
    window.open(icsUrl);
  }
  
  public render(): React.ReactElement<IEventRegistrationProps> {
    //console.info(this.props);
    let month:string;
    let day:string;
    let year:string;
    let start:string;
    let end:string;
    let sing:IIconProps;
    let unsing:IIconProps;
    let save:IIconProps;

    if(this.props.eventItem){
      month = new Date(this.props.eventItem.EventDate).toLocaleString(this.props.lang,{month:'short'});
      day = new Date(this.props.eventItem.EventDate).toLocaleString(this.props.lang,{day:'2-digit'});
      year = new Date(this.props.eventItem.EventDate).toLocaleString(this.props.lang,{year:'numeric'});
      start = new Date(this.props.eventItem.EventDate).toLocaleString(this.props.lang,{ year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
      end = new Date(this.props.eventItem.EventDate).toLocaleString(this.props.lang,{ year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' });
      sing = { iconName: 'UserFollowed' };
      unsing = { iconName: 'UserRemove' };
      save = { iconName: 'CalendarReply' };
    }

    return (
      <div className={ styles.eventRegistration }>
        {this.props.eventItem ?
        <div>
          <div className={ styles.DateBox }>
            <div className={ styles.month }>{month}</div>
            <div className={ styles.day }>{day}</div>
            <div className={ styles.year }>{year}</div>
          </div>
          <div className={ styles.EventDetails }>
            <div className={ styles.EventTitle }>{this.props.eventItem.Title}</div>
            <div className={ styles.EventLocation }>{this.props.eventItem.Location}</div>
            <div className={ styles.EventDates }>
              <div className={ styles.EventStart }>{start}</div><span> - </span>
              <div className={ styles.EventEnd }>{end}</div>
            </div>
            <div className={ styles.EventTools }>
              {!this.props.eventRegistrationItem ?
              <span><PrimaryButton text={strings.SignUp} iconProps={sing} onClick={() => this.sign(this.props.self)} /> </span>
              :
              <span><PrimaryButton text={strings.UnSubscribe} iconProps={unsing} onClick={() => this.unsign(this.props.self)} /> </span>
              }
              <PrimaryButton text={strings.SaveEvent} iconProps={save} onClick={() => this.saveEvent(this)} />
            </div>
          </div>
          <div className={ styles.clear }></div>
        </div>
        : <div>Please select event...</div>
        } 
      </div>
    );
  }
}
