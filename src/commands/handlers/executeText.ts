/* global PowerPoint */

import { HandlerResult } from "../types";

// Contrato de handler (Arquitetura AD-4): executa via PowerPoint.run/context.sync(),
// sempre seleciona todos os objetos criados/afetados, retorna HandlerResult -- nunca
// manipula UI ou estado de sessao diretamente.
export async function executeText(): Promise<HandlerResult> {
  try {
    let newShapeId = "";

    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const textBox = slide.shapes.addTextBox("");
      textBox.load("id");
      await context.sync();

      newShapeId = textBox.id;
      slide.setSelectedShapes([newShapeId]);
      await context.sync();
    });

    return {
      ok: true,
      message: "caixa de texto inserida no slide",
      affectedIds: [newShapeId],
    };
  } catch (error) {
    return {
      ok: false,
      error: "falha ao inserir caixa de texto",
      hint: error instanceof Error ? error.message : String(error),
    };
  }
}
