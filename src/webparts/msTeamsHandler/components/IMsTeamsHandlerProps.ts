import { MSGraphClient } from "@microsoft/sp-http";

export interface IMsTeamsHandlerProps {
  TeamTitle: string;
  client :MSGraphClient;
}

