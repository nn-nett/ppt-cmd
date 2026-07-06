import { resolveCommand } from "../commands/registry";
import { executeText } from "../commands/handlers/executeText";
import { HandlerResult, CommandHandler } from "../commands/types";

// Mapa handlerId -> funcao handler. Cresce conforme o Registro cresce (Story 1.6+).
const handlers: Record<string, CommandHandler> = {
  executeText,
};

// Pipeline unico do parser (Arquitetura AD-6): resolve alias/palavra completa
// contra o Registro (AD-3) e despacha pro handler. A UI nunca resolve alias
// diretamente -- so chama parseAndDispatch e usa o resultado.
export async function parseAndDispatch(input: string): Promise<HandlerResult> {
  const entry = resolveCommand(input);

  if (!entry) {
    return {
      ok: false,
      error: "comando não encontrado",
      hint: input,
    };
  }

  const handler = handlers[entry.handlerId];
  if (!handler) {
    return {
      ok: false,
      error: "handler não implementado",
      hint: entry.handlerId,
    };
  }

  const result = await handler();
  if (result.ok === true) {
    return { ...result, command: entry.palavraCompleta };
  }
  return result;
}
