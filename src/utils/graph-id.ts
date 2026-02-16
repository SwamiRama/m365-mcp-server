import { z } from 'zod';

// Safe pattern for Graph API resource IDs (alphanumeric, hyphens, dots, underscores, commas, colons, exclamation marks)
export const graphIdPattern = /^[a-zA-Z0-9\-._,!:]+$/;
export const graphIdSchema = z.string().min(1).regex(graphIdPattern, 'Invalid resource ID format');
