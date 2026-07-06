import { CommandRegistryEntry } from "./types";

// Registro de Comandos declarativo (Arquitetura AD-3) -- fonte unica de verdade
// do catalogo, consumida pelo parser e (nas proximas historias) pela UI/catalog-list.
export const commandRegistry: CommandRegistryEntry[] = [
  {
    alias: "t",
    palavraCompleta: "text",
    handlerId: "executeText",
  },
];

export function resolveCommand(input: string): CommandRegistryEntry | undefined {
  const normalized = input.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }
  return commandRegistry.find(
    (entry) => entry.alias.toLowerCase() === normalized || entry.palavraCompleta.toLowerCase() === normalized
  );
}

// Sugestao de autocomplete por substring (FR6, fuzzy/mid-string match -- nao so prefixo).
// Usada pro ghost-text: retorna a primeira entrada cuja palavraCompleta contem a
// substring digitada. Nao confundir com resolveCommand (match exato, usado no Enter).
export function findGhostSuggestion(input: string): CommandRegistryEntry | undefined {
  const normalized = input.trim().toLowerCase();
  if (!normalized) {
    return undefined;
  }
  return commandRegistry.find((entry) => entry.palavraCompleta.toLowerCase().includes(normalized));
}
