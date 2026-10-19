/**
 * Starts the render service: POST a .pptx, get a PNG back.
 *
 *   npm run render-server -- [--port 8787] [--host 127.0.0.1] [--max-mb 50] [--queue 16] [--width 480]
 *
 * It listens on the local machine only, unless --host says otherwise: it does no authentication.
 */
import { createRenderServer } from "../src/headless/render-server";

const option = (name: string, fallback: string) => {
  const at = process.argv.indexOf(`--${name}`);
  return at >= 0 ? process.argv[at + 1] : fallback;
};
const port = Number(option("port", process.env.PORT ?? "8787"));
const host = option("host", "127.0.0.1");

const server = createRenderServer({
  maxBodyBytes: Number(option("max-mb", "50")) * 1048576,
  maxQueue: Number(option("queue", "16")),
  thumbnailWidth: Number(option("width", "480"))
});
server.listen(port, host, () => console.log(`渲染服务已启动：http://${host}:${port}（POST /thumbnail、POST /render、GET /health）`));
for (const signal of ["SIGINT", "SIGTERM"] as const) process.on(signal, () => server.close(() => process.exit(0)));
