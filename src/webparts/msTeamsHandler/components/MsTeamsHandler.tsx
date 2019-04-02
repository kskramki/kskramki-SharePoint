import * as React from 'react';
import styles from './MsTeamsHandler.module.scss';
import { IMsTeamsHandlerProps} from './IMsTeamsHandlerProps';
import {IMsTeamsHandlerState} from './IMsTeamsHandlerState';
import { escape } from '@microsoft/sp-lodash-subset';
import { AadHttpClientFactory } from '@microsoft/sp-http';
import { IGroupItem,IUserItem } from './IGroupItem';
import {DetailsList,  DetailsListLayoutMode,CheckboxVisibility,SelectionMode,Dropdown,DropdownMenuItemType,IDropdownOption} from 'office-ui-fabric-react'

// const dropdownStyles: Partial<IDropdownStyles> = {
//   dropdown: { width: 300 }
// };
var dpoptions :Array<IDropdownOption> = new Array<IDropdownOption>();
var officeUsers: Array<IDropdownOption> = new Array<IDropdownOption>();
export default class MsTeamsHandler extends React.Component<IMsTeamsHandlerProps, IMsTeamsHandlerState> {
//Constructor
  public constructor(props)
  {
    super(props);
    this.state={
  Teamstitle:"",
  groups:[],
  doptions:[],
 users:[],
 selectedGroup:'',
 selectedUser:''


};
this.handleChange = this.handleChange.bind(this);
this.CreateTeam =this.CreateTeam.bind(this);
this.getUsers = this.getUsers.bind(this);
this._onChange = this._onChange.bind(this);
this._onUserNameChange = this._onUserNameChange.bind(this);

  }
  //Compopnent Mount 
public componentDidMount ()
{
//Joine Group  need User.ReadWrite.All permision
 // this.props.client.api("/users/49ba6e73-6df7-441b-98be-8cd747f2c631/joinedTeams").get().then(response=>{
 ////   console.log("Joined Team :",response);
 // });
  
 //this.props.client.api("/teams").version("beta").post(content,this.SuccessFailureCallBack);
 
 //this.props.client.api("/groups?$select=id,resourceProvisioningOptions").get(this.SuccessFailureCallBack);
 //need User.ReadBasic.All
this.props.client.api("/users/").get().then(response=>{
  console.log(response);
 response.value.map((item:any)=>{officeUsers.push({key:item.id,text:item.displayName})})
  if(response['@odata.nextLink']!=null)
  {
    this.getUsers(response['@odata.nextLink'])
  }
this.setState({
  users:officeUsers
})
}
  );
  //Get all groups in the tenant which is having Teams in group.
  //V1.0 - /groups?$select=id,resourceProvisioningOptions
 this.props.client.api("/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')").version('beta').get().then(response=>{
//console.log(response)
var groups: Array<IGroupItem> = new Array<IGroupItem>();
var dpoptions :Array<IDropdownOption> = new Array<IDropdownOption>();
response.value.map(((item:any)=>{
  groups.push({displayName:item.displayName,id:item.id});
  dpoptions.push({key:item.id,text:item.displayName})
}))
this.setState(
  {
    groups: groups,
    doptions:dpoptions
  }
); 
});
 
}


//Get all users with paging.
private getUsers(nexturl:string):void
{
  this.props.client.api(nexturl).get().then(response=>{
//console.log(response)
response.value.map((item:any)=>{officeUsers.push({key:item.id,text:item.displayName})})
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
//Create MS team using GraphAPI beta version
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
 private _onSelect(event: React.FormEvent<HTMLDivElement>, item?: IDropdownOption):void{
   alert(item.text)
 }
 private _onChange = (item: IDropdownOption,index:Number): void => {
  console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
  this.setState({selectedGroup:item.key})
  
};
private _onUserNameChange = (item: IDropdownOption,index:Number): void => {
  console.log(`Selection change: ${item.text} ${item.selected ? 'selected' : 'unselected'}`);
  this.setState({selectedUser:item.key})
};

  public render(): React.ReactElement<IMsTeamsHandlerProps> {

    console.log(dpoptions);
    let _usersListColumns = [
      {
        key: 'displayName',
        name: 'Group Display name',
        fieldName: 'displayName',
        minWidth: 50,
        maxWidth: 150,
        isResizable: true
      },
      {
        key: 'id',
        name: 'ID',
        fieldName: 'id',
        minWidth: 150,
        maxWidth: 250,
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
       <div>
        <Dropdown
        label="Group Names"
        options={this.state.doptions} 
       onChanged = {this._onChange}
        style={ { width: 300 }}
        
         disabled={false}
      />
      <Dropdown
        label="User Names"
        options={this.state.users} 
       onChanged = {this._onChange}
        style={ { width: 300 }}
        
         disabled={false}
      />
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
