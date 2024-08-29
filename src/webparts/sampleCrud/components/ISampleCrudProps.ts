import { WebPartContext } from "@microsoft/sp-webpart-base";
export interface ISampleCrudProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  websiteUrl: string;
  spcontext: WebPartContext
}
