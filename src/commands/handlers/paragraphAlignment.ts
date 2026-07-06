/* global PowerPoint */

import { HandlerResult } from "../types";
import { getSelectedShape } from "./shared";

function setAlignment(alignment: PowerPoint.ParagraphHorizontalAlignment, label: string) {
  return async function (): Promise<HandlerResult> {
    try {
      let affectedId = "";
      await PowerPoint.run(async (context) => {
        const shape = await getSelectedShape(context);
        if (!shape) {
          throw new Error("nenhuma forma selecionada no slide");
        }
        shape.textFrame.textRange.paragraphFormat.horizontalAlignment = alignment;
        await context.sync();

        shape.load("id");
        await context.sync();
        affectedId = shape.id;
        context.presentation.getSelectedSlides().getItemAt(0).setSelectedShapes([affectedId]);
        await context.sync();
      });
      return { ok: true, message: `alinhamento ${label} aplicado`, affectedIds: [affectedId] };
    } catch (error) {
      return {
        ok: false,
        error: `falha ao alinhar (${label})`,
        hint: error instanceof Error ? error.message : String(error),
      };
    }
  };
}

export const executeAlignLeft = setAlignment(PowerPoint.ParagraphHorizontalAlignment.left, "à esquerda");
export const executeAlignCenter = setAlignment(PowerPoint.ParagraphHorizontalAlignment.center, "centralizado");
export const executeAlignRight = setAlignment(PowerPoint.ParagraphHorizontalAlignment.right, "à direita");
export const executeJustify = setAlignment(PowerPoint.ParagraphHorizontalAlignment.justify, "justificado");
