import { createServer } from "node:http";
import { readFile } from "node:fs/promises";

const server = createServer(async (request, response) => {
  if (request.url === "/health") { response.writeHead(200, { "content-type": "text/plain" }); response.end("ready"); return; }
  if (request.url === "/miku-md2pptx-worker.js") { try { response.writeHead(200, { "content-type": "text/javascript", "cache-control": "no-store" }); response.end(await readFile("workplace/browser-tests/miku-md2pptx-worker.js")); } catch (error) { response.writeHead(500); response.end(error instanceof Error ? error.message : String(error)); } return; }
  response.writeHead(404); response.end("not found");
});
server.listen(4176, "127.0.0.1");
