import { IGroupItem } from "./IGroupItem";

export interface IMsTeamsHandlerState
{
Teamstitle :string;
groups:Array<IGroupItem>;
doptions:Array<any>;
users:Array<any>;
//selectedGroup:string | number;
//selectedUser:string | number;
isHidden:boolean;
}