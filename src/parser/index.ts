import { resolveCommand } from "../commands/registry";
import { executeText } from "../commands/handlers/executeText";
import { executeBold, executeItalic, executeUnderline } from "../commands/handlers/textFormatting";
import {
  executeAlignLeft,
  executeAlignCenter,
  executeAlignRight,
  executeJustify,
} from "../commands/handlers/paragraphAlignment";
import { executeNewSlide } from "../commands/handlers/executeNewSlide";
import { executeShape } from "../commands/handlers/executeShape";
import { executeBullets } from "../commands/handlers/executeBullets";
import { HandlerResult, CommandHandler } from "../commands/types";

// Mapa handlerId -> funcao handler. Cresce conforme o Registro cresce (Story 1.6).
const handlers: Record<string, CommandHandler> = {
  executeText,
  executeBold,
  executeItalic,
  executeUnderline,
  executeAlignLeft,
  executeAlignCenter,
  executeAlignRight,
  executeJustify,
  executeNewSlide,
  executeShape,
  executeBullets,
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
