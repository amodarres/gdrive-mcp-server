import express from 'express';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StreamableHTTPServerTransport } from '@modelcontextprotocol/sdk/server/streamableHttp.js';
import { google } from 'googleapis';
import { z } from 'zod';

const PORT = process.env.PORT || 3000;

function getAuthClient() {
  const keyJson = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!keyJson) throw new Error('GOOGLE_SERVICE_ACCOUNT_KEY environment variable is required');
  const credentials = JSON.parse(keyJson);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/presentations',
      'https://www.googleapis.com/auth/drive',
    ],
  });
}

function buildServer() {
  const server = new McpServer({ name: 'gdrive-slides', version: '1.0.0' });

  server.tool(
    'create_presentation',
    'Creates a new Google Slides presentation with the given title',
    { title: z.string().describe('Title of the new presentation') },
    async ({ title }) => {
      try {
        const slides = google.slides({ version: 'v1', auth: getAuthClient() });
        const res = await slides.presentations.create({ requestBody: { title } });
        return {
          content: [{
            type: 'text',
            text: JSON.stringify({
              presentationId: res.data.presentationId,
              title: res.data.title,
            }),
          }],
        };
      } catch (error) {
        console.error(JSON.stringify(error, null, 2));
        throw error;
      }
    }
  );

  server.tool(
    'add_slide',
    'Adds a new slide with a title and body text to an existing presentation',
    {
      presentationId: z.string().describe('ID of the presentation'),
      title: z.string().describe('Title text for the new slide'),
      body: z.string().describe('Body text for the new slide'),
    },
    async ({ presentationId, title, body }) => {
      const slides = google.slides({ version: 'v1', auth: getAuthClient() });
      const slideObjectId = `slide_${Date.now()}`;
      const titleObjectId = `${slideObjectId}_title`;
      const bodyObjectId = `${slideObjectId}_body`;

      await slides.presentations.batchUpdate({
        presentationId,
        requestBody: {
          requests: [
            {
              createSlide: {
                objectId: slideObjectId,
                slideLayoutReference: { predefinedLayout: 'TITLE_AND_BODY' },
                placeholderIdMappings: [
                  {
                    layoutPlaceholder: { type: 'TITLE', index: 0 },
                    objectId: titleObjectId,
                  },
                  {
                    layoutPlaceholder: { type: 'BODY', index: 0 },
                    objectId: bodyObjectId,
                  },
                ],
              },
            },
            { insertText: { objectId: titleObjectId, text: title } },
            { insertText: { objectId: bodyObjectId, text: body } },
          ],
        },
      });

      return {
        content: [{
          type: 'text',
          text: JSON.stringify({ slideObjectId, presentationId }),
        }],
      };
    }
  );

  server.tool(
    'get_presentation_url',
    'Returns the shareable Google Slides URL for a presentation',
    { presentationId: z.string().describe('ID of the presentation') },
    async ({ presentationId }) => {
      const url = `https://docs.google.com/presentation/d/${presentationId}/edit`;
      return { content: [{ type: 'text', text: url }] };
    }
  );

  return server;
}

const app = express();
app.use(express.json());

app.post('/mcp', async (req, res) => {
  const server = buildServer();
  const transport = new StreamableHTTPServerTransport({ sessionIdGenerator: undefined });
  await server.connect(transport);
  await transport.handleRequest(req, res, req.body);
});

app.get('/health', (_req, res) => res.json({ status: 'ok' }));

app.listen(PORT, '0.0.0.0', () => {
  console.log(`gdrive-mcp-server listening on 0.0.0.0:${PORT}`);
  try {
    const keyJson = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;
    if (!keyJson) {
      console.error('[startup] GOOGLE_SERVICE_ACCOUNT_KEY is not set');
    } else {
      const credentials = JSON.parse(keyJson);
      console.log(`[startup] service account email: ${credentials.client_email}`);
    }
  } catch (e) {
    console.error('[startup] failed to parse GOOGLE_SERVICE_ACCOUNT_KEY:', e.message);
  }
});
