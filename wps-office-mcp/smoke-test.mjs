import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const transport = new StdioClientTransport({
  command: 'node',
  args: [new URL('./dist/index.js', import.meta.url).pathname],
  cwd: new URL('.', import.meta.url).pathname,
});
const client = new Client({ name: 'smoke-test', version: '1.0.0' });
await client.connect(transport);

const tools = await client.listTools();
console.log('tools/list count:', tools.tools.length);
console.log('sample tools:', tools.tools.slice(0, 5).map((t) => t.name).join(', '));

const res = await client.callTool({ name: 'wps_check_connection', arguments: {} });
console.log('check_connection:', JSON.stringify(res).slice(0, 300));

await client.close();
console.log('MCP handshake OK');
