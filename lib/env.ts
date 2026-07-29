import { z } from 'zod';

const envSchema = z.object({
  AZURE_CLIENT_ID: z.string().default('mock-client-id'),
  AZURE_CLIENT_SECRET: z.string().default('mock-client-secret'),
  AZURE_TENANT_ID: z.string().default('common'),
  NEXTAUTH_SECRET: z.string().default('default-secret-at-least-32-chars-long!'),
  NEXTAUTH_URL: z.string().default('http://localhost:3000'),
});

export const env = envSchema.parse(process.env);
