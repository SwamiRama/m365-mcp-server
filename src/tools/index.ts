import { GraphClient } from '../graph/client.js';
import { MailTools, mailToolDefinitions, type ListMessagesInput, type GetMessageInput, type ListFoldersInput, type GetAttachmentInput, type UserContext } from './mail.js';
import { SharePointTools, sharePointToolDefinitions, type ListSitesInput, type ListDrivesInput, type ListChildrenInput, type GetFileInput, type SearchFilesInput, type SearchAndReadInput } from './sharepoint.js';
import { CalendarTools, calendarToolDefinitions, type ListEventsInput, type GetEventInput } from './calendar.js';
import { OneDriveTools, oneDriveToolDefinitions, type ListFilesInput, type OdGetFileInput, type SearchInput, type RecentInput, type SharedWithMeInput } from './onedrive.js';

// Re-export tool definitions and types
export { mailToolDefinitions } from './mail.js';
export { sharePointToolDefinitions } from './sharepoint.js';
export { calendarToolDefinitions } from './calendar.js';
export { oneDriveToolDefinitions } from './onedrive.js';
export type { UserContext } from './mail.js';

// Combined tool definitions
export const allToolDefinitions = [
  ...mailToolDefinitions,
  ...sharePointToolDefinitions,
  ...oneDriveToolDefinitions,
  ...calendarToolDefinitions,
];

// Tool executor class that wraps all tools
export class ToolExecutor {
  private mailTools: MailTools;
  private sharePointTools: SharePointTools;
  private calendarTools: CalendarTools;
  private oneDriveTools: OneDriveTools;

  constructor(graphClient: GraphClient, userContext?: UserContext) {
    this.mailTools = new MailTools(graphClient, userContext);
    this.sharePointTools = new SharePointTools(graphClient);
    this.calendarTools = new CalendarTools(graphClient);
    this.oneDriveTools = new OneDriveTools(graphClient);
  }

  async execute(toolName: string, args: Record<string, unknown>): Promise<object> {
    switch (toolName) {
      // Mail tools
      case 'mail_list_messages':
        return this.mailTools.listMessages(args as ListMessagesInput);
      case 'mail_get_message':
        return this.mailTools.getMessage(args as GetMessageInput);
      case 'mail_list_folders':
        return this.mailTools.listFolders(args as ListFoldersInput);
      case 'mail_get_attachment':
        return this.mailTools.getAttachment(args as GetAttachmentInput);

      // SharePoint/Files tools
      case 'sp_search_read':
        return this.sharePointTools.searchAndRead(args as SearchAndReadInput);
      case 'sp_search':
        return this.sharePointTools.searchFiles(args as SearchFilesInput);
      case 'sp_list_sites':
        return this.sharePointTools.listSites(args as ListSitesInput);
      case 'sp_list_drives':
        return this.sharePointTools.listDrives(args as ListDrivesInput);
      case 'sp_list_children':
        return this.sharePointTools.listChildren(args as ListChildrenInput);
      case 'sp_get_file':
        return this.sharePointTools.getFile(args as GetFileInput);

      // OneDrive tools
      case 'od_my_drive':
        return this.oneDriveTools.myDrive();
      case 'od_list_files':
        return this.oneDriveTools.listFiles(args as ListFilesInput);
      case 'od_get_file':
        return this.oneDriveTools.getFile(args as OdGetFileInput);
      case 'od_search':
        return this.oneDriveTools.search(args as SearchInput);
      case 'od_recent':
        return this.oneDriveTools.recent(args as RecentInput);
      case 'od_shared_with_me':
        return this.oneDriveTools.sharedWithMe(args as SharedWithMeInput);

      // Calendar tools
      case 'cal_list_calendars':
        return this.calendarTools.listCalendars();
      case 'cal_list_events':
        return this.calendarTools.listEvents(args as ListEventsInput);
      case 'cal_get_event':
        return this.calendarTools.getEvent(args as GetEventInput);

      default:
        throw new Error(`Unknown tool: ${toolName}`);
    }
  }
}
