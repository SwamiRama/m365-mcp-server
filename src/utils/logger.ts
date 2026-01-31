import pino from 'pino';
import { config } from './config.js';

// PII patterns to redact
const PII_PATTERNS = [
  /Bearer\s+[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]*/gi, // JWT tokens
  /access_token['":\s]+['"]?[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]+\.[A-Za-z0-9\-_]*/gi,
  /refresh_token['":\s]+['"]?[A-Za-z0-9\-_]+/gi,
  /client_secret['":\s]+['"]?[A-Za-z0-9\-_]+/gi,
  /password['":\s]+['"]?[^\s'"]+/gi,
  /[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g, // Email addresses
];

function redactPII(obj: unknown): unknown {
  if (typeof obj === 'string') {
    let result = obj;
    for (const pattern of PII_PATTERNS) {
      result = result.replace(pattern, '[REDACTED]');
    }
    return result;
  }

  if (Array.isArray(obj)) {
    return obj.map(redactPII);
  }

  if (obj !== null && typeof obj === 'object') {
    const redacted: Record<string, unknown> = {};
    for (const [key, value] of Object.entries(obj)) {
      // Redact sensitive keys entirely
      if (/token|secret|password|authorization|cookie/i.test(key)) {
        redacted[key] = '[REDACTED]';
      } else {
        redacted[key] = redactPII(value);
      }
    }
    return redacted;
  }

  return obj;
}

const baseOptions: pino.LoggerOptions = {
  level: config.logLevel,
  formatters: {
    level: (label) => ({ level: label }),
    bindings: () => ({}),
  },
  hooks: {
    logMethod(inputArgs, method) {
      // Redact PII from all log arguments
      const redactedArgs = inputArgs.map(redactPII) as Parameters<typeof method>;
      return method.apply(this, redactedArgs);
    },
  },
  timestamp: pino.stdTimeFunctions.isoTime,
};

const devOptions: pino.LoggerOptions = {
  ...baseOptions,
  transport: {
    target: 'pino-pretty',
    options: {
      colorize: true,
      translateTime: 'SYS:standard',
      ignore: 'pid,hostname',
    },
  },
};

export const logger = pino(
  config.nodeEnv === 'development' ? devOptions : baseOptions
);

// Create child logger with correlation ID
export function createRequestLogger(correlationId: string) {
  return logger.child({ correlationId });
}

export type Logger = typeof logger;
