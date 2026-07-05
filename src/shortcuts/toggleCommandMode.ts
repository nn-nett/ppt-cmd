/* global document, Office */

let commandModeActive = false;

function setCommandModeVisual(active: boolean): void {
  const chip = document.getElementById("mode-chip");
  const rest = document.getElementById("command-bar-rest");
  const prompt = document.getElementById("command-bar-prompt");
  const input = document.getElementById("command-bar-input") as HTMLInputElement | null;

  if (chip) {
    chip.style.display = active ? "block" : "none";
  }
  if (rest) {
    rest.style.borderColor = active ? "#3DDC97" : "#333333";
  }
  if (prompt) {
    prompt.style.color = active ? "#3DDC97" : "#6A6A6A";
  }
  if (input) {
    if (active) {
      input.focus();
    } else {
      input.value = "";
      input.blur();
    }
  }
}

export function activateCommandMode(): void {
  commandModeActive = true;
  setCommandModeVisual(true);
}

export function deactivateCommandMode(): void {
  commandModeActive = false;
  setCommandModeVisual(false);
}

export function toggleCommandMode(): void {
  if (commandModeActive) {
    deactivateCommandMode();
  } else {
    activateCommandMode();
  }
}

export function registerToggleCommandMode(): void {
  // Fora do callback de Office.onReady -- associate nao espera onReady resolver.
  // NOTA: gatilho de teclado ainda nao confirmado funcionando neste ambiente
  // (ver Dev Agent Record da Story 1.2) -- mantido registrado, nao bloqueia
  // o clique manual, que e o caminho principal por enquanto.
  Office.actions.associate("ToggleCommandMode", () => {
    return Office.addin
      .showAsTaskpane()
      .then(() => {
        toggleCommandMode();
        return;
      })
      .catch((error) => {
        return error.code;
      });
  });

  document.addEventListener("DOMContentLoaded", () => {
    const rest = document.getElementById("command-bar-rest");
    const input = document.getElementById("command-bar-input") as HTMLInputElement | null;

    if (rest) {
      rest.addEventListener("click", () => {
        activateCommandMode();
      });
    }

    if (input) {
      input.addEventListener("focus", () => {
        activateCommandMode();
      });
      input.addEventListener("keydown", (event: KeyboardEvent) => {
        if (event.key === "Escape") {
          deactivateCommandMode();
        }
      });
    }
  });
}
