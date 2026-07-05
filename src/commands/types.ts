export type HandlerResult =
  | { ok: true; message: string; affectedIds: string[] }
  | { ok: false; error: string; hint: string };

export interface CommandRegistryEntry {
  alias: string;
  palavraCompleta: string;
  handlerId: string;
  icone?: string;
  animacaoId?: string;
  schemaSubPrompt?: unknown;
  schemaAtributos?: unknown;
}

export type CommandHandler = () => Promise<HandlerResult>;
