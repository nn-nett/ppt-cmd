/* global PowerPoint */

// Helper compartilhado: pega a primeira forma atualmente selecionada no slide.
// Usado por handlers que modificam uma selecao existente (ex: bold, alinhamento),
// diferente de handlers que criam elementos novos (ex: executeText, executeShape).
export async function getSelectedShape(
  context: PowerPoint.RequestContext
): Promise<PowerPoint.Shape | undefined> {
  const selectedShapes = context.presentation.getSelectedShapes();
  selectedShapes.load("items");
  await context.sync();

  if (selectedShapes.items.length === 0) {
    return undefined;
  }
  return selectedShapes.getItemAt(0);
}
