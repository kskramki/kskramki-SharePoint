import * as React from 'react';
import styles from './MsTeamsHandler.module.scss';
import { IMsTeamsHandlerProps} from './IMsTeamsHandlerProps';
import {IMsTeamsHandlerState} from './IMsTeamsHandlerState';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClientFactory } from '@microsoft/sp-http';
import { IGroupItem } from './IGroupItem';
import {DetailsList,  DetailsListLayoutMode,CheckboxVisibility,SelectionMode} from 'office-ui-fabric-react'

export default class MsTeamsHandler extends React.Component<IMsTeamsHandlerProps, IMsTeamsHandlerState> {
//Constructor
  public constructor(props)
  {
    super(props);
    this.state={
  Teamstitle:"",
  groups:[]
};
this.handleChange = this.handleChange.bind(this);
this.CreateTeam =this.CreateTeam.bind(this);
this.getUsers = this.getUsers.bind(this);

  }
  //Compopnent Mount 
public componentDidMount ()
{
//Joine Group  need User.ReadWrite.All permision
  this.props.client.api("/users/49ba6e73-6df7-441b-98be-8cd747f2c631/joinedTeams").get().then(response=>{
    console.log(response);
  });
  
 //this.props.client.api("/teams").version("beta").post(content,this.SuccessFailureCallBack);
 
 //this.props.client.api("/groups?$select=id,resourceProvisioningOptions").get(this.SuccessFailureCallBack);
 //need User.ReadBasic.All
this.props.client.api("/users/").get().then(response=>{
  console.log(response)
  if(response['@odata.nextLink']!=null)
  {
    this.getUsers(response['@odata.nextLink'])
  }
}
  );
   
 this.props.client.api("/groups").get().then(response=>{
console.log(response)
var groups: Array<IGroupItem> = new Array<IGroupItem>();
response.value.map(((item:any)=>{
  groups.push({displayName:item.displayName,id:item.id});
}))
this.setState(
  {
    groups: groups,
  }
); 
});
  

}
//Get all users with paging.
private getUsers(nexturl:string):void
{
  this.props.client.api(nexturl).get().then(response=>{
console.log(response)
if(response['@odata.nextLink']!=null)
  {
    this.getUsers(response['@odata.nextLink']);
  }
});
}

private SuccessFailureCallBack(err:any,response:any,rawresponse?:any)
{
 // console.log("First : ",err ,"Second ",response,"thirs", rawresponse);
  if(rawresponse!=null && rawresponse.status =="202")
  alert("Created Succesfully")
  else{
    console.log("Error:", err, "Response :" ,response, "Rawresponse",rawresponse )
    alert("Erro:Please chec the console")
  }
}

private handleChange(event):void  {
  this.setState({Teamstitle: event.target.value});
}

  private CreateTeam()
{
  let TT= this.state.Teamstitle;
  var content;
  if(TT!=null && TT!="")
  {
    content =`{
    "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('standard')",
    "displayName": "` +TT +`",
    "description": "Created using GraphAPI from SPFx"
  }`;
  this.props.client.api("/teams").version("beta").post(content,this.SuccessFailureCallBack);
}
console.log(content);
}

  public render(): React.ReactElement<IMsTeamsHandlerProps> {
    let _usersListColumns = [
      {
        key: 'displayName',
        name: 'Group Display name',
        fieldName: 'displayName',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      },
      {
        key: 'id',
        name: 'ID',
        fieldName: 'id',
        minWidth: 50,
        maxWidth: 100,
        isResizable: true
      },
     
    ];
    return (
      <div className={ styles.msTeamsHandler }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
            
              <span className={ styles.title }>Creating Team  </span> 
              
              <input placeholder="Please enter Team Title" type="text" id="txtteamtitle" value= {this.state.Teamstitle} onChange={this.handleChange}></input>
              <button value="Create" id="btnCreateTeam" onClick={this.CreateTeam}>Create Team</button>
              </div>
          </div>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Check Console logs for  more</span>
              </a>
            </div>
          </div>
        </div>
        <DetailsList
                      items={ this.state.groups }
                      columns={ _usersListColumns }
                      setKey='set'
                      checkboxVisibility={ CheckboxVisibility.hidden }
                      selectionMode={ SelectionMode.none }
                      layoutMode={ DetailsListLayoutMode.fixedColumns }
                      compact={ true }
                  />
        </div>
    );
  }
}
