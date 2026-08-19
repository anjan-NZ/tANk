// Pane polish, not protection: the source is public and a webview can always be inspected
// from outside the page. Inputs keep their own menu so paste still works.

const DEVTOOLS_KEYS = new Set(["I", "J", "C"]);

function editable(target: EventTarget | null): boolean {
  const el = target as HTMLElement | null;
  if (!el || !el.closest) return false;
  return !!el.closest("input, textarea, select, [contenteditable='true']");
}

export function lockDown(): void {
  document.addEventListener("contextmenu", (e) => {
    if (editable(e.target)) return;
    e.preventDefault();
  });

  document.addEventListener("keydown", (e) => {
    const devtools =
      e.key === "F12" ||
      (e.ctrlKey && e.shiftKey && DEVTOOLS_KEYS.has(e.key.toUpperCase())) ||
      (e.ctrlKey && !e.shiftKey && e.key.toLowerCase() === "u");
    if (devtools) e.preventDefault();
  });

  // a dropped file would navigate away from the add-in
  document.addEventListener("dragover", (e) => e.preventDefault());
  document.addEventListener("drop", (e) => e.preventDefault());
}
