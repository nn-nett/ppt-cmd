/* global PowerPoint */

import { HandlerResult } from "../types";

export async function executeNewSlide(): Promise<HandlerResult> {
  try {
    await PowerPoint.run(async (context) => {
      context.presentation.slides.add();
      await context.sync();
    });
    return { ok: true, message: "novo slide inserido", affectedIds: [] };
  } catch (error) {
    return {
      ok: false,
      error: "falha ao inserir slide",
      hint: error instanceof Error ? error.message : String(error),
    };
  }
}
