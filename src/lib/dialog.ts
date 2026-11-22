/* global Office, console, location */
/*
 * Message / confirmation dialog shown by the ribbon and keyboard commands.
 * The page (dialog.html) is opened without query parameters so the dialog title bar
 * shows a short URL; the content is sent with messageChild once the page reports
 * it is ready. Falls back to query parameters where DialogApi 1.2 is unavailable.
 */

export type DialogKind = "info" | "success" | "warning" | "error" | "confirm";

export interface DialogOptions {
  title: string;
  message: string;
  kind?: DialogKind;
  /** confirm dialogs only: label of the primary button */
  okLabel?: string;
}

function isDarkTheme(): boolean {
  try {
    const bg = Office.context.officeTheme?.bodyBackgroundColor;
    if (bg && /^#?[0-9a-f]{6}$/i.test(bg)) {
      const hex = bg.replace("#", "");
      const [r, g, b] = [0, 2, 4].map((k) => parseInt(hex.slice(k, k + 2), 16));
      return 0.2126 * r + 0.7152 * g + 0.0722 * b < 128;
    }
  } catch {
    /* officeTheme not available */
  }
  return typeof matchMedia === "function" && matchMedia("(prefers-color-scheme: dark)").matches;
}

/**
 * Resolves with "ok" (primary button), "cancel", "close" (OK on a message), or ""
 * when the dialog could not be shown or was dismissed with the X.
 */
export function showDialog(opts: DialogOptions): Promise<string> {
  const payload = JSON.stringify({ ...opts, kind: opts.kind ?? "info", dark: isDarkTheme() });
  let canMessageChild = false;
  try {
    canMessageChild = Office.context.requirements.isSetSupported("DialogApi", "1.2");
  } catch {
    /* assume not */
  }

  return new Promise((resolve) => {
    let url: string;
    try {
      url = new URL("dialog.html", location.href).href;
      if (!canMessageChild) url += "?data=" + encodeURIComponent(payload);
    } catch {
      resolve("");
      return;
    }
    const height = opts.kind === "confirm" || opts.message.length > 160 ? 30 : 24;
    Office.context.ui.displayDialogAsync(url, { height, width: 28, displayInIframe: true }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.warn("dialog failed:", result.error);
        resolve("");
        return;
      }
      const dialog = result.value;
      let settled = false;
      const finish = (value: string) => {
        if (settled) return;
        settled = true;
        try {
          dialog.close();
        } catch {
          /* already closed */
        }
        resolve(value);
      };
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
        const message = (arg as { message?: string }).message ?? "";
        if (message === "ready") {
          if (canMessageChild) dialog.messageChild(payload);
          return;
        }
        finish(message);
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, () => finish(""));
    });
  });
}
