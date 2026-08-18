export type HostApp = "excel" | "word" | "powerpoint" | "none";

/** Which Office app the pane is running inside. Decides which tools the model is offered. */
export function detectHost(): HostApp {
  if (typeof Office === "undefined" || !Office.context) return "none";
  switch (Office.context.host) {
    case Office.HostType.Excel:
      return "excel";
    case Office.HostType.Word:
      return "word";
    case Office.HostType.PowerPoint:
      return "powerpoint";
    default:
      return "none";
  }
}

export const HOST_LABEL: Record<HostApp, string> = {
  excel: "Excel",
  word: "Word",
  powerpoint: "PowerPoint",
  none: "no Office app",
};
