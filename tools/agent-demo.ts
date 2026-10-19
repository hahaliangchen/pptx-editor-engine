/**
 * Lets a model build or change a deck through the tools, and records what it did.
 * Works with any OpenAI-compatible chat endpoint (vLLM, Ollama, OpenAI ...).
 *
 *   npm run agent -- --base http://host:port/v1 --model name --task "做一份……" [--deck 底稿.pptx] [--out 目录]
 *                    [--key 密钥] [--max-turns 16] [--max-tokens 32768] [--no-think] [--no-vision] [--temperature 0.3]
 */
import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { HeadlessSession } from "../src/headless/session";
import { TOOL_GUIDE, TOOLS } from "../src/api/tools";
import { OpenAIMessage, callOpenAITool, toOpenAI } from "../src/api/adapters";

function option(name: string, fallback?: string): string | undefined {
  const at = process.argv.indexOf(`--${name}`);
  return at >= 0 ? process.argv[at + 1] : fallback;
}
const flag = (name: string) => process.argv.includes(`--${name}`);

const base = (option("base", process.env.OPENAI_BASE_URL) || "").replace(/\/$/, "");
const model = option("model", process.env.OPENAI_MODEL) || "";
const key = option("key", process.env.OPENAI_API_KEY);
const deck = option("deck", "tools/templates/blank.pptx")!;
const out = option("out", "agent-out")!;
const maxTurns = Number(option("max-turns", "16"));
const temperature = Number(option("temperature", "0.3"));
const maxTokens = Number(option("max-tokens", "32768"));
/** Some servers (vLLM with Qwen) can switch the model's thinking off with a template switch. */
const thinking = !flag("no-think");
const vision = !flag("no-vision");
const task = option("task", "")!;

if (!base || !model || !task) {
  console.log('用法：npm run agent -- --base http://host:port/v1 --model 模型名 --task "要做什么" [--deck 底稿.pptx] [--out 目录] [--max-tokens 32768] [--no-think] [--no-vision]');
  process.exit(1);
}

interface Usage { prompt_tokens?: number; completion_tokens?: number; completion_tokens_details?: { reasoning_tokens?: number } }

async function chat(messages: unknown[], tools: unknown[]): Promise<{ message: any; usage: Usage; finish: string }> {
  const response = await fetch(`${base}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json", ...(key ? { Authorization: `Bearer ${key}` } : {}) },
    body: JSON.stringify({
      model, messages, tools, temperature, max_tokens: maxTokens,
      ...(thinking ? {} : { chat_template_kwargs: { enable_thinking: false } })
    }),
    signal: AbortSignal.timeout(900_000)
  });
  const body = await response.json() as any;
  if (!response.ok || body.error) throw new Error(`模型接口出错 ${response.status}：${JSON.stringify(body.error ?? body).slice(0, 400)}`);
  return { message: body.choices[0].message, usage: body.usage ?? {}, finish: body.choices[0].finish_reason };
}

const brief = (text: string, n = 160) => (text.length > n ? `${text.slice(0, n)}…` : text);

async function main(): Promise<void> {
  mkdirSync(out, { recursive: true });
  const session = await HeadlessSession.open(readFileSync(deck), { quiet: true });
  const tools = toOpenAI(vision ? TOOLS : TOOLS.filter(tool => tool.name !== "render_slide"));
  const messages: any[] = [{ role: "system", content: TOOL_GUIDE }, { role: "user", content: task }];
  const stats = { turns: 0, calls: {} as Record<string, number>, editFailures: 0, promptTokens: 0, completionTokens: 0, reasoningTokens: 0, images: 0 };
  const started = Date.now();
  let answer = "";

  for (let turn = 1; turn <= maxTurns; turn++) {
    stats.turns = turn;
    const { message, usage, finish } = await chat(messages, tools);
    stats.promptTokens += usage.prompt_tokens ?? 0;
    stats.completionTokens += usage.completion_tokens ?? 0;
    stats.reasoningTokens += usage.completion_tokens_details?.reasoning_tokens ?? 0;
    const calls = message.tool_calls ?? [];
    console.log(`\n[第 ${turn} 轮] ${finish}，${usage.completion_tokens ?? "?"} token 输出，${calls.length} 个工具调用`);
    messages.push({ role: "assistant", content: message.content ?? "", ...(calls.length ? { tool_calls: calls } : {}) });
    if (finish === "length") {
      console.log(`  输出被 max_tokens（${maxTokens}）截断了；可以调大 --max-tokens，或用 --no-think 关掉思考。`);
      answer = message.content ?? "";
      break;
    }
    if (calls.length === 0) {
      answer = message.content ?? "";
      break;
    }
    for (const call of calls) {
      const name = call.function.name as string;
      stats.calls[name] = (stats.calls[name] ?? 0) + 1;
      const reply: OpenAIMessage[] = await callOpenAITool(session, call);
      messages.push(...reply);
      const text = (reply[0] as { content: string }).content;
      let line = `  ${name}`;
      try {
        const args = JSON.parse(call.function.arguments || "{}");
        if (name === "apply_edits") line += `（${args.commands?.length ?? "?"} 条指令）`;
        if (name === "render_slide") line += `（第 ${args.slide} 页）`;
      } catch { line += "（参数不是合法 JSON）"; }
      if (name === "apply_edits") {
        const result = JSON.parse(text);
        if (result.ok) line += " → 成功";
        else {
          stats.editFailures++;
          line += ` → 失败：${result.failedAt === undefined ? "整批被拒" : `第 ${result.failedAt + 1} 条`} ${result.error?.code}：${brief(result.error?.message ?? "", 220)} [${result.error?.path ?? ""}]`;
        }
      } else if (name !== "render_slide") {
        line += ` → ${brief(text, 80)}`;
      }
      const image = reply.find(message => message.role === "user") as { content: Array<{ type: string; image_url?: { url: string } }> } | undefined;
      const url = image?.content.find(part => part.type === "image_url")?.image_url?.url;
      if (url) {
        stats.images++;
        const file = `${out}/turn${String(turn).padStart(2, "0")}-render${stats.images}.png`;
        writeFileSync(file, Buffer.from(url.split(",")[1], "base64"));
        line += ` → ${file}`;
      }
      console.log(line);
    }
  }

  const pages = (await session.outline()).slides.length;
  writeFileSync(`${out}/result.pptx`, await session.export());
  for (let page = 1; page <= pages; page++) writeFileSync(`${out}/final-${page}.png`, await session.render(page, { width: 960 }));
  const redacted = messages.map(message => JSON.parse(JSON.stringify(message, (field, value) =>
    typeof value === "string" && value.startsWith("data:image") ? "[图片]" : value)));
  writeFileSync(`${out}/transcript.json`, JSON.stringify({ model, task, deck, stats, answer, messages: redacted }, null, 2));
  session.dispose();

  console.log(`\n结果：${pages} 页，${out}/result.pptx；用时 ${Math.round((Date.now() - started) / 1000)} 秒`);
  console.log(`轮数 ${stats.turns}，工具调用 ${JSON.stringify(stats.calls)}，编辑失败 ${stats.editFailures} 次，`
    + `输入 ${stats.promptTokens} / 输出 ${stats.completionTokens}（其中推理 ${stats.reasoningTokens}）token，看了 ${stats.images} 张渲染图`);
  if (answer) console.log(`\n模型最后说：\n${brief(answer, 800)}`);
}

main().catch(error => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});
