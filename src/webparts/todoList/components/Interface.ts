import { SPHttpClient } from "@microsoft/sp-http";

export interface ISPListItem {
  Id: string;
  Title: string;
  Status: string;
}

export interface IItemListProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
}

export interface ITodoFormValues {
  title: string;
  status: string;
}
