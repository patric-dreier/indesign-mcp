import { z } from "zod";
import { defineTool } from "./registry.js";
import { runScriptWithResultFile } from "../transport/result-file.js";
import { wrapExtendScript, lit, prelude } from "../compose.js";
import { findDocumentById, findFrameById } from "../script-helpers.js";

const InputSchema = z
  .object({
    frame_id: z.string(),
    document_id: z.string().optional(),
  })
  .strict();

type Input = z.infer<typeof InputSchema>;

const ScriptResultSchema = z.object({
  frame_id: z.string(),
  text: z.string(),
  character_count: z.number().int().nonnegative(),
  footnote_count: z.number().int().nonnegative(),
});

type ScriptResult = z.infer<typeof ScriptResultSchema>;

type Result = ScriptResult;

function buildScriptBody(input: Input): string {
  const docExpr =
    input.document_id !== undefined
      ? `findDocumentById(${lit(input.document_id)})`
      : "app.activeDocument";

  return `
${prelude(findDocumentById, findFrameById)}
var doc = ${docExpr};
var item = findFrameById(doc, ${lit(input.frame_id)});

if (item.constructor.name !== "TextFrame") {
  throw { name: "invalid_argument", message: "frame " + ${lit(input.frame_id)} + " is not a text frame", entity: "frame", id: ${lit(input.frame_id)} };
}

var story = item.parentStory;
var fullText = String(story.contents || "");

var footnoteCount = 0;
try {
  footnoteCount = story.footnotes.length;
} catch (e) {
  footnoteCount = 0;
}

return {
  frame_id: ${lit(input.frame_id)},
  text: fullText,
  character_count: fullText.length,
  footnote_count: footnoteCount
};
`;
}

export const getFrameTextTool = defineTool<Input, Result>({
  name: "get_frame_text",
  description:
    "Returns the complete, untruncated text content of a text frame's story, along with its character count and the number of footnotes it contains.",
  inputSchema: InputSchema,
  async handler(input) {
    return runScriptWithResultFile<ScriptResult>({
      language: "JavaScript",
      scriptTemplate: wrapExtendScript(buildScriptBody(input)),
      resultSchema: ScriptResultSchema,
    });
  },
});
