// import { SPHttpClient } from "@microsoft/sp-http";

export interface ITodoListProps {
  description: string;
  websiteUrl: string;
  spHttpClient: any;
  sp: any;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
