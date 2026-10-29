import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICaraghAssistantProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  context: WebPartContext;
}