/* global PowerPoint */

import { HandlerResult } from "../types";
import { getSelectedShape } from "./shared";

export async function executeBullets(): Promise<HandlerResult> {
  try {
    let affectedId = "";
    await PowerPoint.run(async (context) => {
      const shape = await getSelectedShape(context);
      if (!shape) {
        throw new Error("nenhuma forma selecionada no slide");
      }
      const bulletFormat = shape.textFrame.textRange.paragraphFormat.bulletFormat;
      bulletFormat.load("visible");
      await context.sync();

      const nowVisible = !bulletFormat.visible;
      bulletFormat.visible = nowVisible;
      if (nowVisible) {
        bulletFormat.type = PowerPoint.BulletType.unnumbered;
      }
      await context.sync();

      shape.load("id");
      await context.sync();
      affectedId = shape.id;
      context.presentation.getSelectedSlides().getItemAt(0).setSelectedShapes([affectedId]);
      await context.sync();
    });
    return { ok: true, message: "marcadores alternados", affectedIds: [affectedId] };
  } catch (error) {
    return {
      ok: false,
      error: "falha ao alternar marcadores",
      hint: error instanceof Error ? error.message : String(error),
    };
  }
}
