/* global PowerPoint */

import { HandlerResult } from "../types";

export async function executeShape(): Promise<HandlerResult> {
  try {
    let newShapeId = "";
    await PowerPoint.run(async (context) => {
      const slide = context.presentation.getSelectedSlides().getItemAt(0);
      const rectangle = slide.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle, {
        left: 100,
        top: 100,
        width: 150,
        height: 150,
      });
      rectangle.load("id");
      await context.sync();

      newShapeId = rectangle.id;
      slide.setSelectedShapes([newShapeId]);
      await context.sync();
    });
    return { ok: true, message: "forma inserida no slide", affectedIds: [newShapeId] };
  } catch (error) {
    return {
      ok: false,
      error: "falha ao inserir forma",
      hint: error instanceof Error ? error.message : String(error),
    };
  }
}
