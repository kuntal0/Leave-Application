import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ILeaveApplicationProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  listname:string;
  disabled?: boolean;
  checked?: boolean;
  //color: string;
}
