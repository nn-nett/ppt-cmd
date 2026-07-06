/* global PowerPoint */

import { HandlerResult } from "../types";
import { getSelectedShape } from "./shared";

// Toggle generico de propriedade booleana da fonte (bold, italic). Segue o
// contrato AD-4: PowerPoint.run, seleciona o objeto afetado, retorna HandlerResult.
function toggleFontBoolean(property: "bold" | "italic", label: string) {
  return async function (): Promise<HandlerResult> {
    try {
      let affectedId = "";
      await PowerPoint.run(async (context) => {
        const shape = await getSelectedShape(context);
        if (!shape) {
          throw new Error("nenhuma forma selecionada no slide");
        }
        const font = shape.textFrame.textRange.font;
        font.load(property);
        await context.sync();

        font[property] = !font[property];
        await context.sync();

        shape.load("id");
        await context.sync();
        affectedId = shape.id;
        context.presentation.getSelectedSlides().getItemAt(0).setSelectedShapes([affectedId]);
        await context.sync();
      });
      return { ok: true, message: `${label} alternado`, affectedIds: [affectedId] };
    } catch (error) {
      return {
        ok: false,
        error: `falha ao alternar ${label}`,
        hint: error instanceof Error ? error.message : String(error),
      };
    }
  };
}

export const executeBold = toggleFontBoolean("bold", "negrito");
export const executeItalic = toggleFontBoolean("italic", "itálico");

export async function executeUnderline(): Promise<HandlerResult> {
  try {
    let affectedId = "";
    await PowerPoint.run(async (context) => {
      const shape = await getSelectedShape(context);
      if (!shape) {
        throw new Error("nenhuma forma selecionada no slide");
      }
      const font = shape.textFrame.textRange.font;
      font.load("underline");
      await context.sync();

      const isUnderlined = font.underline !== PowerPoint.ShapeFontUnderlineStyle.none;
      font.underline = isUnderlined
        ? PowerPoint.ShapeFontUnderlineStyle.none
        : PowerPoint.ShapeFontUnderlineStyle.single;
      await context.sync();

      shape.load("id");
      await context.sync();
      affectedId = shape.id;
      context.presentation.getSelectedSlides().getItemAt(0).setSelectedShapes([affectedId]);
      await context.sync();
    });
    return { ok: true, message: "sublinhado alternado", affectedIds: [affectedId] };
  } catch (error) {
    return {
      ok: false,
      error: "falha ao alternar sublinhado",
      hint: error instanceof Error ? error.message : String(error),
    };
  }
}
