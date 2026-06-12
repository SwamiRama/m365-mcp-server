import pino from 'pino';
import { config } from './config.js';
import { redactPII } from './redact.js';

export { redactPII };

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
