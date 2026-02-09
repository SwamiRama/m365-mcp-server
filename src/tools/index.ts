import { GraphClient } from '../graph/client.js';
import { MailTools, mailToolDefinitions, type ListMessagesInput, type GetMessageInput, type UserContext } from './mail.js';
import { SharePointTools, sharePointToolDefinitions, type ListSitesInput, type ListDrivesInput, type ListChildrenInput, type GetFileInput, type SearchFilesInput } from './sharepoint.js';

// Re-export tool definitions and types
export { mailToolDefinitions } from './mail.js';
export { sharePointToolDefinitions } from './sharepoint.js';
export type { UserContext } from './mail.js';

// Combined tool definitions
export const allToolDefinitions = [
  ...mailToolDefinitions,
  ...sharePointToolDefinitions,
];

// Tool executor class that wraps all tools
export class ToolExecutor {
  private mailTools: MailTools;
  private sharePointTools: SharePointTools;

  constructor(graphClient: GraphClient, userContext?: UserContext) {
    this.mailTools = new MailTools(graphClient, userContext);
    this.sharePointTools = new SharePointTools(graphClient);
  }

  async execute(toolName: string, args: Record<string, unknown>): Promise<object> {
    switch (toolName) {
      // Mail tools
      case 'mail_list_messages':
        return this.mailTools.listMessages(args as ListMessagesInput);
      case 'mail_get_message':
        return this.mailTools.getMessage(args as GetMessageInput);
      case 'mail_list_folders':
        return this.mailTools.listFolders({});

      // SharePoint/Files tools
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

      default:
        throw new Error(`Unknown tool: ${toolName}`);
    }
  }
}
