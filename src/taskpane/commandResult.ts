/* global document */

// Exibicao do feedback de sucesso/erro (FR7, DESIGN.md components.success-state/error-state,
// EXPERIENCE.md Voice and Tone). Sucesso: glow sutil unico, ~1.6s. Erro: cartao estatico.

let hideTimer: ReturnType<typeof setTimeout> | undefined;

export function showSuccess(command: string, message: string): void {
  const el = document.getElementById("command-result");
  const text = document.getElementById("command-result-text");
  const hint = document.getElementById("command-result-hint");
  if (!el || !text || !hint) {
    return;
  }

  el.className = "command-result command-result--success";
  text.textContent = `✓ OK ${command} — ${message}`;
  hint.textContent = "";
  el.style.display = "block";

  scheduleHide();
}

export function showError(error: string, hint: string): void {
  const el = document.getElementById("command-result");
  const text = document.getElementById("command-result-text");
  const hintEl = document.getElementById("command-result-hint");
  if (!el || !text || !hintEl) {
    return;
  }

  el.className = "command-result command-result--error";
  text.textContent = `✗ ${error}: "${hint}"`;
  hintEl.textContent = "nenhum alias corresponde — pressione Esc ou continue digitando";
  el.style.display = "block";

  scheduleHide();
}

function scheduleHide(): void {
  if (hideTimer) {
    clearTimeout(hideTimer);
  }
  hideTimer = setTimeout(() => {
    hideResult();
  }, 2200);
}

export function hideResult(): void {
  const el = document.getElementById("command-result");
  if (el) {
    el.style.display = "none";
  }
}
