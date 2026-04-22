import { getLocaleID, getString } from "../utils/locale";
import { getPref } from "../utils/prefs";
import {
  getStorageNovelBooks,
  ensureSelectedNovelPath,
  rememberReaderSelectionSnapshot,
  startReplaceFromSelection,
} from "./novel";

function getReadAsNovelMenuLabel() {
  return getPref("language") === "en-US" ? "Read as Novel" : "看小说";
}

function getNovelMenuStateTexts() {
  if (getPref("language") === "en-US") {
    return {
      notConfigured: "Set book folder in plugin settings",
      missingDirectory: "Configured folder does not exist",
      emptyFolder: "No EPUB files in folder",
      recentTag: "Recent",
    };
  }
  return {
    notConfigured: "请先在插件设置中配置图书目录",
    missingDirectory: "配置的目录不存在",
    emptyFolder: "目录中没有 EPUB 文件",
    recentTag: "最近阅读",
  };
}

function formatBookMenuLabel(
  menuLabel: string,
  fileName: string,
  isRecent: boolean,
  recentTag: string,
) {
  const maxNameLength = 56;
  const clippedName = fileName.length > maxNameLength
    ? `${fileName.slice(0, maxNameLength - 1)}...`
    : fileName;
  const recentPrefix = isRecent ? `[${recentTag}] ` : "";
  return `${menuLabel} > ${recentPrefix}${clippedName}`;
}

type NovelBookMenuPanelState = {
  root: HTMLDivElement;
  observer?: MutationObserver;
  closeTimer?: ReturnType<typeof globalThis.setTimeout>;
};

type NovelBookMenuHoverBridgeState = {
  reader: _ZoteroTypes.ReaderInstance;
  menuLabel: string;
  recentTag: string;
  books: Array<{ filePath: string; fileName: string; lastReadAt: number }>;
  selectedText: string;
  selectionRange?: Range;
  selectionRect?: [number, number, number, number];
  selectionPosition?: any;
  selectionAnnotation?: any;
  activeButton?: HTMLButtonElement;
  closeTimer?: ReturnType<typeof globalThis.setTimeout>;
};

const novelBookMenuPanelByReader = new WeakMap<_ZoteroTypes.ReaderInstance, NovelBookMenuPanelState>();
const novelBookMenuHoverBridgeByDocument = new WeakMap<Document, NovelBookMenuHoverBridgeState>();
const novelBookMenuHoverBridgeBoundDocuments = new WeakSet<Document>();
const novelReaderShortcutBridgeBoundDocuments = new WeakSet<Document>();
const novelReaderLiveSelectionBoundDocuments = new WeakSet<Document>();
const NOVEL_BOOK_PANEL_STYLE_ID = "zotero-thief-novel-book-panel-style";
const NOVEL_BOOK_MENU_SELECTOR = ".context-menu";
const NOVEL_BOOK_MENU_BUTTON_SELECTOR = "button.basic, button.row.basic";

function getReaderMenuDocument(reader?: _ZoteroTypes.ReaderInstance): Document | undefined {
  const windows: Array<Window | undefined> = [
    (reader as any)?._iframeWindow,
    (reader as any)?._window,
    (reader as any)?._iframe?.contentWindow,
    (reader as any)?._browser?.contentWindow,
    (reader as any)?.browser?.contentWindow,
    ztoolkit.getGlobal("window") as Window | undefined,
  ];

  for (const win of windows) {
    const doc = win?.document;
    if (doc?.body) {
      return doc;
    }
  }

  return undefined;
}

function normalizeMenuText(text: string) {
  return text.replace(/\s+/g, " ").trim();
}

function ensureNovelBookMenuStyles(doc: Document) {
  if (doc.getElementById(NOVEL_BOOK_PANEL_STYLE_ID)) {
    return;
  }

  const style = doc.createElement("style");
  style.id = NOVEL_BOOK_PANEL_STYLE_ID;
  style.textContent = `
    .zotero-thief-novel-book-panel {
      position: fixed;
      z-index: 2147483647;
      width: 340px;
      max-width: min(340px, calc(100vw - 24px));
      max-height: min(340px, calc(100vh - 24px));
      background: var(--material-background, #ffffff);
      color: var(--fill-primary, #151515);
      border: 1px solid rgba(0, 0, 0, 0.12);
      border-radius: 10px;
      box-shadow: 0 14px 32px rgba(0, 0, 0, 0.18);
      overflow: hidden;
      font-size: 13px;
      line-height: 1.35;
      user-select: none;
    }

    .zotero-thief-novel-book-panel__header {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 12px;
      padding: 10px 12px 8px;
      border-bottom: 1px solid rgba(0, 0, 0, 0.08);
      font-weight: 600;
      white-space: nowrap;
      background: rgba(127, 127, 127, 0.03);
    }

    .zotero-thief-novel-book-panel__list {
      max-height: 280px;
      overflow-y: auto;
      padding: 6px 0;
    }

    .zotero-thief-novel-book-panel__item {
      display: block;
      width: 100%;
      box-sizing: border-box;
      padding: 8px 12px;
      border: 0;
      background: transparent;
      color: inherit;
      text-align: left;
      font: inherit;
      cursor: pointer;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }

    .zotero-thief-novel-book-panel__item:disabled {
      cursor: default;
      opacity: 0.5;
    }

    .zotero-thief-novel-book-panel__item:hover,
    .zotero-thief-novel-book-panel__item:focus-visible {
      background: rgba(127, 127, 127, 0.14);
      outline: none;
    }
  `;
  const host = doc.head || doc.documentElement;
  host?.appendChild(style);
}

function closeNovelBookMenuPanel(reader?: _ZoteroTypes.ReaderInstance) {
  const state = reader ? novelBookMenuPanelByReader.get(reader) : undefined;
  if (!state) {
    return;
  }

  ztoolkit.log("novel-debug", "menu.panel.close", {
    hasReader: Boolean(reader),
    hasTimer: Boolean(state.closeTimer),
  });
  if (state.closeTimer) {
    globalThis.clearTimeout(state.closeTimer);
  }
  state.observer?.disconnect();
  state.root.remove();
  novelBookMenuPanelByReader.delete(reader!);
}

function scheduleCloseNovelBookMenuPanel(reader: _ZoteroTypes.ReaderInstance, delay = 120) {
  const state = novelBookMenuPanelByReader.get(reader);
  if (!state) {
    return;
  }

  ztoolkit.log("novel-debug", "menu.panel.scheduleClose", {
    delay,
    hasTimer: Boolean(state.closeTimer),
  });
  if (state.closeTimer) {
    globalThis.clearTimeout(state.closeTimer);
  }
  state.closeTimer = globalThis.setTimeout(() => {
    ztoolkit.log("novel-debug", "menu.panel.scheduleClose.fire", { delay });
    closeNovelBookMenuPanel(reader);
  }, delay);
}

function getNovelBookMenuPanelElement(doc: Document): HTMLDivElement | undefined {
  return doc.querySelector(".zotero-thief-novel-book-panel") as HTMLDivElement | undefined;
}

function getNovelBookMenuButtonFromTarget(
  target: EventTarget | null,
  menuLabel: string,
): HTMLButtonElement | undefined {
  const maybeElement = target as {
    closest?: (selector: string) => Element | null;
  } | null;
  if (!maybeElement || typeof maybeElement.closest !== "function") {
    return undefined;
  }

  const button = maybeElement.closest(NOVEL_BOOK_MENU_BUTTON_SELECTOR) as HTMLButtonElement | null;
  if (!button) {
    return undefined;
  }

  const buttonLabel = normalizeMenuText(button.textContent || "");
  if (!buttonLabel || !buttonLabel.includes(menuLabel)) {
    return undefined;
  }

  return button;
}

function getNovelBookMenuButtonByLabel(
  doc: Document,
  menuLabel: string,
): HTMLButtonElement | undefined {
  const buttons = Array.from(
    doc.querySelectorAll(NOVEL_BOOK_MENU_BUTTON_SELECTOR),
  ) as HTMLButtonElement[];
  return buttons.find((button) => {
    const text = normalizeMenuText(button.textContent || "");
    return Boolean(text && text.includes(menuLabel));
  });
}

function isNodeLike(value: unknown): value is { nodeType: number } {
  return Boolean(
    value
    && typeof value === "object"
    && typeof (value as { nodeType?: unknown }).nodeType === "number",
  );
}

function updateNovelBookMenuHoverBridge(options: {
  reader: _ZoteroTypes.ReaderInstance;
  menuLabel: string;
  books: Array<{ filePath: string; fileName: string; lastReadAt: number }>;
  recentTag: string;
  selectedText: string;
  selectionRange?: Range;
  selectionRect?: [number, number, number, number];
  selectionPosition?: any;
  selectionAnnotation?: any;
}) {
  const doc = getReaderMenuDocument(options.reader);
  if (!doc) {
    ztoolkit.log("novel-debug", "menu.hoverBridge.noDocument", {
      menuLabel: options.menuLabel,
    });
    return;
  }

  ztoolkit.log("novel-debug", "menu.hoverBridge.update", {
    menuLabel: options.menuLabel,
    booksCount: options.books.length,
    selectedLength: options.selectedText.length,
    hasSelectionRange: Boolean(options.selectionRange),
    hasSelectionRect: Boolean(options.selectionRect),
  });

  novelBookMenuHoverBridgeByDocument.set(doc, {
    reader: options.reader,
    menuLabel: options.menuLabel,
    recentTag: options.recentTag,
    books: options.books,
    selectedText: options.selectedText,
    selectionRange: options.selectionRange,
    selectionRect: options.selectionRect,
    selectionPosition: options.selectionPosition,
    selectionAnnotation: options.selectionAnnotation,
  });

  if (novelBookMenuHoverBridgeBoundDocuments.has(doc)) {
    ztoolkit.log("novel-debug", "menu.hoverBridge.alreadyBound", {
      menuLabel: options.menuLabel,
    });
    return;
  }
  novelBookMenuHoverBridgeBoundDocuments.add(doc);
  ztoolkit.log("novel-debug", "menu.hoverBridge.bindDocument", {
    menuLabel: options.menuLabel,
  });

  const clearPanelTimer = (reader: _ZoteroTypes.ReaderInstance) => {
    const state = novelBookMenuPanelByReader.get(reader);
    if (!state?.closeTimer) {
      return;
    }
    globalThis.clearTimeout(state.closeTimer);
    state.closeTimer = undefined;
  };

  const handlePointerOver = (event: MouseEvent) => {
    const state = novelBookMenuHoverBridgeByDocument.get(doc);
    if (!state) {
      return;
    }

    const button = getNovelBookMenuButtonFromTarget(event.target, state.menuLabel);
    ztoolkit.log("novel-debug", "menu.hoverBridge.mouseover", {
      menuLabel: state.menuLabel,
      targetTag: (event.target as Element | null)?.tagName || "unknown",
      relatedTag: (event.relatedTarget as Element | null)?.tagName || "unknown",
      matched: Boolean(button),
    });
    if (!button) {
      return;
    }

    state.activeButton = button;
    clearPanelTimer(state.reader);
    ztoolkit.log("novel-debug", "menu.hoverBridge.openPanel", {
      menuLabel: state.menuLabel,
      targetText: normalizeMenuText(button.textContent || ""),
      booksCount: state.books.length,
    });
    openNovelBookMenuPanel({
      reader: state.reader,
      anchor: button,
      menuLabel: state.menuLabel,
      recentTag: state.recentTag,
      books: state.books,
      selectedText: state.selectedText,
      selectionRange: state.selectionRange,
      selectionRect: state.selectionRect,
      selectionPosition: state.selectionPosition,
      selectionAnnotation: state.selectionAnnotation,
    });
  };

  const handlePointerOut = (event: MouseEvent) => {
    const state = novelBookMenuHoverBridgeByDocument.get(doc);
    if (!state) {
      return;
    }

    const button = getNovelBookMenuButtonFromTarget(event.target, state.menuLabel);
    ztoolkit.log("novel-debug", "menu.hoverBridge.mouseout", {
      menuLabel: state.menuLabel,
      targetTag: (event.target as Element | null)?.tagName || "unknown",
      relatedTag: (event.relatedTarget as Element | null)?.tagName || "unknown",
      matched: Boolean(button),
      activeButton: Boolean(state.activeButton),
    });
    if (!button) {
      return;
    }

    const panel = getNovelBookMenuPanelElement(doc);
    const relatedTarget = event.relatedTarget;
    if (panel && isNodeLike(relatedTarget) && panel.contains(relatedTarget as unknown as Node)) {
      ztoolkit.log("novel-debug", "menu.hoverBridge.pointerLeftToPanel", {
        menuLabel: state.menuLabel,
      });
      clearPanelTimer(state.reader);
      return;
    }

    if (button === state.activeButton) {
      ztoolkit.log("novel-debug", "menu.hoverBridge.scheduleCloseFromButton", {
        menuLabel: state.menuLabel,
      });
      scheduleCloseNovelBookMenuPanel(state.reader, 170);
    }
  };

  const handleKeydown = (event: KeyboardEvent) => {
    if (event.key !== "Escape") {
      return;
    }

    const state = novelBookMenuHoverBridgeByDocument.get(doc);
    if (!state) {
      return;
    }

    ztoolkit.log("novel-debug", "menu.hoverBridge.escape", {
      menuLabel: state.menuLabel,
    });
    closeNovelBookMenuPanel(state.reader);
  };

  const MutationObserverCtor = doc.defaultView?.MutationObserver || (globalThis as any).MutationObserver;
  const observer = MutationObserverCtor
    ? new MutationObserverCtor(() => {
      const state = novelBookMenuHoverBridgeByDocument.get(doc);
      if (!state) {
        return;
      }

      if (!doc.querySelector(NOVEL_BOOK_MENU_SELECTOR)) {
        ztoolkit.log("novel-debug", "menu.hoverBridge.menuGone", {
          menuLabel: state.menuLabel,
        });
        closeNovelBookMenuPanel(state.reader);
      }
    })
    : undefined;
  const observerHost = doc.body || doc.documentElement;
  if (observerHost && observer) {
    observer.observe(observerHost, { childList: true, subtree: true });
  }

  ztoolkit.log("novel-debug", "menu.hoverBridge.boundListeners", {
    menuLabel: options.menuLabel,
    hasObserverHost: Boolean(observerHost),
    hasMutationObserver: Boolean(observer),
  });

  doc.addEventListener("mouseover", handlePointerOver, true);
  doc.addEventListener("mouseout", handlePointerOut, true);
  doc.addEventListener("keydown", handleKeydown, true);
}

function bindReaderShortcutBridge(reader: _ZoteroTypes.ReaderInstance, doc: Document) {
  if (novelReaderShortcutBridgeBoundDocuments.has(doc)) {
    return;
  }
  novelReaderShortcutBridgeBoundDocuments.add(doc);

  const handleKeydown = (event: KeyboardEvent) => {
    const key = String(event.key || "").toLowerCase();
    if (key !== "r") {
      return;
    }

    const selectedText = ztoolkit.Reader.getSelectedText(reader).trim();
    const selectionRange = getReaderSelectionRangeSnapshot(reader);
    const popupSnapshot = getReaderSelectionPopupSnapshot(reader);
    ztoolkit.log("novel-debug", "readerShortcutBridge.keydown", {
      key,
      selectedLength: selectedText.length,
      hasRange: Boolean(selectionRange),
      hasPopupRect: Boolean(popupSnapshot?.rect),
    });
    if (!selectedText && !selectionRange) {
      return;
    }

    const epubPath = getPref("bookStoragePath") ? ensureSelectedNovelPath() : undefined;
    if (!epubPath) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    void startReplaceFromSelection({
      reader,
      selectedText,
      selectionRange,
      selectionRect: popupSnapshot?.rect,
      selectionPosition: popupSnapshot?.annotation?.position,
      selectionAnnotation: popupSnapshot?.annotation,
      epubPath,
    });
  };

  doc.addEventListener("keydown", handleKeydown, true);
}

function bindReaderLiveSelectionBridge(reader: _ZoteroTypes.ReaderInstance, doc: Document) {
  if (novelReaderLiveSelectionBoundDocuments.has(doc)) {
    return;
  }
  novelReaderLiveSelectionBoundDocuments.add(doc);

  const updateSnapshotFromLiveSelection = () => {
    const selectedText = ztoolkit.Reader.getSelectedText(reader).trim();
    const selectionRange = getReaderSelectionRangeSnapshot(reader);
    if (!selectedText && !selectionRange) {
      return;
    }

    const popupSnapshot = getReaderSelectionPopupSnapshot(reader);
    rememberReaderSelectionSnapshot(reader, {
      selectedText: selectedText || selectionRange?.toString().trim() || "",
      selectionRange,
      selectionRect: popupSnapshot?.rect,
      selectionPosition: popupSnapshot?.annotation?.position,
      selectionAnnotation: popupSnapshot?.annotation,
    });
  };

  const handleSelectionChange = () => {
    updateSnapshotFromLiveSelection();
  };

  const handleKeydown = (event: KeyboardEvent) => {
    const key = String(event.key || "").toLowerCase();
    if (key !== "r") {
      return;
    }

    const selectedText = ztoolkit.Reader.getSelectedText(reader).trim();
    const selectionRange = getReaderSelectionRangeSnapshot(reader);
    const popupSnapshot = getReaderSelectionPopupSnapshot(reader);
    if (!selectedText && !selectionRange) {
      return;
    }

    const epubPath = ensureSelectedNovelPath();
    if (!epubPath) {
      return;
    }

    event.preventDefault();
    event.stopPropagation();

    void startReplaceFromSelection({
      reader,
      selectedText,
      selectionRange,
      selectionRect: popupSnapshot?.rect,
      selectionPosition: popupSnapshot?.annotation?.position,
      selectionAnnotation: popupSnapshot?.annotation,
      epubPath,
    });
  };

  doc.addEventListener("selectionchange", handleSelectionChange, true);
  doc.addEventListener("keyup", handleSelectionChange, true);
  doc.addEventListener("mousedown", () => {
    // Selection often changes on mouseup, but mousedown gives us a cheap re-check path.
    updateSnapshotFromLiveSelection();
  }, true);
  doc.addEventListener("keydown", handleKeydown, true);

  updateSnapshotFromLiveSelection();
}

export function syncNovelReaderLiveSelectionBridges() {
  const readers = Array.isArray((Zotero.Reader as any)?._readers)
    ? ((Zotero.Reader as any)._readers as _ZoteroTypes.ReaderInstance[])
    : [];

  for (const reader of readers) {
    const doc = getReaderMenuDocument(reader);
    if (!doc) {
      continue;
    }
    bindReaderLiveSelectionBridge(reader, doc);
    bindReaderShortcutBridge(reader, doc);
  }
}

function positionNovelBookMenuPanel(
  panel: HTMLDivElement,
  anchor: HTMLElement | undefined,
  point?: { x: number; y: number },
) {
  const menuRoot = panel.ownerDocument?.querySelector(NOVEL_BOOK_MENU_SELECTOR) as HTMLElement | null;
  const menuRect = menuRoot?.getBoundingClientRect();
  const pointIsValid = Boolean(
    point
    && Number.isFinite(point.x)
    && Number.isFinite(point.y)
    && (Math.abs(point.x) > 1 || Math.abs(point.y) > 1),
  );
  const rect = anchor?.getBoundingClientRect() || {
    left: menuRect?.left || (pointIsValid ? point!.x : 12),
    right: menuRect?.right || (pointIsValid ? point!.x : 12),
    top: menuRect?.top || (pointIsValid ? point!.y : 12),
    bottom: menuRect?.bottom || (pointIsValid ? point!.y : 12),
  };
  const win = panel.ownerDocument?.defaultView;
  const viewportWidth = win?.innerWidth || 1024;
  const viewportHeight = win?.innerHeight || 768;
  const panelWidth = Math.min(340, Math.max(280, Math.round(viewportWidth * 0.26)));
  const panelHeight = 220;
  const gap = 4;

  let left = rect.right + gap;
  if (left + panelWidth > viewportWidth - 12) {
    left = Math.max(12, rect.left - panelWidth - gap);
  }
  let top = rect.top - 8;
  if (top + panelHeight > viewportHeight - 12) {
    top = Math.max(12, viewportHeight - panelHeight - 12);
  }

  panel.style.width = `${panelWidth}px`;
  panel.style.left = `${Math.max(12, left)}px`;
  panel.style.top = `${Math.max(12, top)}px`;
}

function openNovelBookMenuPanel(options: {
  reader: _ZoteroTypes.ReaderInstance;
  anchor?: HTMLElement;
  point?: { x: number; y: number };
  menuLabel: string;
  recentTag: string;
  books: Array<{ filePath: string; fileName: string; lastReadAt: number }>;
  selectedText: string;
  selectionRange?: Range;
  selectionRect?: [number, number, number, number];
  selectionPosition?: any;
  selectionAnnotation?: any;
}) {
  const doc = getReaderMenuDocument(options.reader);
  if (!doc) {
    ztoolkit.log("novel-debug", "menu.panel.noDocument", {
      menuLabel: options.menuLabel,
    });
    return;
  }

  if (!options.books.length) {
    ztoolkit.log("novel-debug", "menu.panel.noBooks", {
      menuLabel: options.menuLabel,
      selectedLength: options.selectedText.length,
    });
    return;
  }

  ztoolkit.log("novel-debug", "menu.panel.openRequest", {
    menuLabel: options.menuLabel,
    booksCount: options.books.length,
    selectedLength: options.selectedText.length,
    hasAnchor: Boolean(options.anchor),
    point: options.point,
  });
  closeNovelBookMenuPanel(options.reader);
  ensureNovelBookMenuStyles(doc);

  const root = doc.createElement("div");
  root.className = "zotero-thief-novel-book-panel";
  root.tabIndex = -1;

  const header = doc.createElement("div");
  header.className = "zotero-thief-novel-book-panel__header";
  header.textContent = options.menuLabel;
  root.appendChild(header);

  const list = doc.createElement("div");
  list.className = "zotero-thief-novel-book-panel__list";

  for (const book of options.books) {
    const button = doc.createElement("button");
    button.type = "button";
    button.className = "zotero-thief-novel-book-panel__item";
    button.textContent = formatBookMenuLabel(
      options.menuLabel,
      book.fileName,
      Number(book.lastReadAt) > 0,
      options.recentTag,
    );
    button.addEventListener("click", () => {
      closeNovelBookMenuPanel(options.reader);
      addon.hooks.onDialogEvents("readAsNovel", {
        reader: options.reader,
        selectedText: options.selectedText,
        selectionRange: options.selectionRange,
        selectionRect: options.selectionRect,
        selectionPosition: options.selectionPosition,
        selectionAnnotation: options.selectionAnnotation,
        epubPath: book.filePath,
      });
    });
    list.appendChild(button);
  }

  root.appendChild(list);
  const host = doc.body || doc.documentElement;
  host?.appendChild(root);

  const state: NovelBookMenuPanelState = {
    root,
    observer: undefined,
  };
  novelBookMenuPanelByReader.set(options.reader, state);

  const MutationObserverCtor = doc.defaultView?.MutationObserver || (globalThis as any).MutationObserver;
  if (MutationObserverCtor) {
    state.observer = new MutationObserverCtor(() => {
      if (!doc.querySelector(NOVEL_BOOK_MENU_SELECTOR)) {
        ztoolkit.log("novel-debug", "menu.panel.observerClose", {
          menuLabel: options.menuLabel,
        });
        closeNovelBookMenuPanel(options.reader);
      }
    });
  } else {
    ztoolkit.log("novel-debug", "menu.panel.noMutationObserver", {
      menuLabel: options.menuLabel,
    });
  }

  root.addEventListener("mouseenter", () => {
    ztoolkit.log("novel-debug", "menu.panel.mouseenter", {
      menuLabel: options.menuLabel,
    });
    if (state.closeTimer) {
      globalThis.clearTimeout(state.closeTimer);
    }
  });
  root.addEventListener("mouseleave", () => {
    ztoolkit.log("novel-debug", "menu.panel.mouseleave", {
      menuLabel: options.menuLabel,
    });
    scheduleCloseNovelBookMenuPanel(options.reader, 170);
  });

  if (host && state.observer) {
    state.observer.observe(host, { childList: true, subtree: true });
  }
  globalThis.setTimeout(() => positionNovelBookMenuPanel(root, options.anchor, options.point), 0);
  ztoolkit.log("novel-debug", "menu.panel.opened", {
    menuLabel: options.menuLabel,
  });
}

function bindNovelBookMenuHoverBridge(options: {
  reader: _ZoteroTypes.ReaderInstance;
  menuLabel: string;
  books: Array<{ filePath: string; fileName: string; lastReadAt: number }>;
  recentTag: string;
  selectedText: string;
  selectionRange?: Range;
  selectionRect?: [number, number, number, number];
  selectionPosition?: any;
  selectionAnnotation?: any;
  point?: { x: number; y: number };
}) {
  updateNovelBookMenuHoverBridge(options);
}

function example(
  target: any,
  propertyKey: string | symbol,
  descriptor: PropertyDescriptor,
) {
  const original = descriptor.value;
  descriptor.value = function (...args: any) {
    try {
      ztoolkit.log(`Calling example ${target.name}.${String(propertyKey)}`);
      return original.apply(this, args);
    } catch (e) {
      ztoolkit.log(`Error in example ${target.name}.${String(propertyKey)}`, e);
      throw e;
    }
  };
  return descriptor;
}

export class BasicExampleFactory {
  @example
  static registerNotifier() {
    const callback = {
      notify: async (
        event: string,
        type: string,
        ids: number[] | string[],
        extraData: { [key: string]: any },
      ) => {
        if (!addon?.data.alive) {
          this.unregisterNotifier(notifierID);
          return;
        }
        addon.hooks.onNotify(event, type, ids, extraData);
      },
    };

    // Register the callback in Zotero as an item observer
    const notifierID = Zotero.Notifier.registerObserver(callback, [
      "tab",
      "item",
      "file",
    ]);

    Zotero.Plugins.addObserver({
      shutdown: ({ id }) => {
        if (id === addon.data.config.addonID)
          this.unregisterNotifier(notifierID);
      },
    });
  }

  @example
  static exampleNotifierCallback() {
    ztoolkit.log("popup-muted", {
      stage: "exampleNotifierCallback",
      text: "Open Tab Detected!",
    });
  }

  @example
  private static unregisterNotifier(notifierID: string) {
    Zotero.Notifier.unregisterObserver(notifierID);
  }

  @example
  static registerPrefs() {
    Zotero.PreferencePanes.register({
      pluginID: addon.data.config.addonID,
      src: rootURI + "content/preferences.xhtml",
      label: getString("prefs-title"),
      image: `chrome://${addon.data.config.addonRef}/content/icons/favicon.png`,
    });
  }
}

export class KeyExampleFactory {
  @example
  static registerShortcuts() {
    // Register an event key for Alt+L
    ztoolkit.Keyboard.register((ev, keyOptions) => {
      ztoolkit.log(ev, keyOptions.keyboard);
      if (keyOptions.keyboard?.equals("shift,l")) {
        addon.hooks.onShortcuts("larger");
      }
      if (ev.shiftKey && ev.key === "S") {
        addon.hooks.onShortcuts("smaller");
      }
    });

    ztoolkit.log("popup-muted", {
      stage: "registerShortcuts",
      text: "Example Shortcuts: Alt+L/S/C",
    });
  }

  @example
  static exampleShortcutLargerCallback() {
    ztoolkit.log("popup-muted", {
      stage: "exampleShortcutLargerCallback",
      text: "Larger!",
    });
  }

  @example
  static exampleShortcutSmallerCallback() {
    ztoolkit.log("popup-muted", {
      stage: "exampleShortcutSmallerCallback",
      text: "Smaller!",
    });
  }
}

export class UIExampleFactory {
  @example
  static registerStyleSheet(win: _ZoteroTypes.MainWindow) {
    const doc = win.document;
    const styles = ztoolkit.UI.createElement(doc, "link", {
      properties: {
        type: "text/css",
        rel: "stylesheet",
        href: `chrome://${addon.data.config.addonRef}/content/zoteroPane.css`,
      },
    });
    doc.documentElement?.appendChild(styles);
    doc.getElementById("zotero-item-pane-content")?.classList.add("makeItRed");
  }

  @example
  static registerReaderSelectionMenu() {
    Zotero.Reader.registerEventListener(
      "renderTextSelectionPopup",
      ({ reader, params }) => {
        const popupSnapshot = getReaderSelectionPopupSnapshot(reader);
        const liveSelectedText = ztoolkit.Reader.getSelectedText(reader).trim();
        const liveSelectionRange = getReaderSelectionRangeSnapshot(reader);
        const readerDoc = getReaderMenuDocument(reader);
        if (readerDoc) {
          bindReaderLiveSelectionBridge(reader, readerDoc);
          bindReaderShortcutBridge(reader, readerDoc);
        }
        rememberReaderSelectionSnapshot(reader, {
          selectedText: liveSelectedText || params.annotation.text || "",
          selectionRange: liveSelectionRange,
          selectionRect: popupSnapshot?.rect,
          selectionPosition: params.annotation.position,
          selectionAnnotation: params.annotation,
        });
      },
      addon.data.config.addonID,
    );

    Zotero.Reader.registerEventListener(
      "createViewContextMenu",
      ({ reader, append, params }) => {
        const selectedText = ztoolkit.Reader.getSelectedText(reader);
        const selectionRange = getReaderSelectionRangeSnapshot(reader);
        const popupSnapshot = getReaderSelectionPopupSnapshot(reader);
        const readerDoc = getReaderMenuDocument(reader);
        if (readerDoc) {
          bindReaderLiveSelectionBridge(reader, readerDoc);
          bindReaderShortcutBridge(reader, readerDoc);
        }
        const fallbackRect = popupSnapshot?.rect || buildRectFromContextPoint(params?.x, params?.y);
        const fallbackSelectedText = selectedText.trim() || selectionRange?.toString().trim() || "";
        rememberReaderSelectionSnapshot(reader, {
          selectedText: fallbackSelectedText,
          selectionRange,
          selectionRect: fallbackRect,
          selectionPosition: (params as any)?.position,
          selectionAnnotation: popupSnapshot?.annotation,
        });
        ztoolkit.log("novel-debug", "menu.createViewContextMenu", {
          selectedLength: selectedText.length,
          fallbackLength: fallbackSelectedText.length,
          menuPoint: { x: Number(params?.x) || 0, y: Number(params?.y) || 0 },
          hasSelectionRange: Boolean(selectionRange),
          hasSelectionRect: Boolean(fallbackRect),
          hasSelectionPosition: Boolean((params as any)?.position),
          hasSelectionAnnotation: Boolean(popupSnapshot?.annotation),
          selectionPreview: selectionRange?.toString().slice(0, 80) || "",
        });
        if (!fallbackSelectedText) return;
        const storage = getStorageNovelBooks();
        const menuTexts = getNovelMenuStateTexts();
        const menuLabel = getReadAsNovelMenuLabel();
        const menuPoint = { x: Number(params?.x) || 0, y: Number(params?.y) || 0 };
        const noOp = () => { };
        closeNovelBookMenuPanel(reader);

        if (!storage.configured) {
          append({
            label: `${menuLabel}: ${menuTexts.notConfigured}`,
            disabled: true,
            persistent: true,
            onCommand: noOp,
          });
        } else if (!storage.dirExists) {
          append({
            label: `${menuLabel}: ${menuTexts.missingDirectory}`,
            disabled: true,
            persistent: true,
            onCommand: noOp,
          });
        } else if (storage.books.length === 0) {
          append({
            label: `${menuLabel}: ${menuTexts.emptyFolder}`,
            disabled: true,
            persistent: true,
            onCommand: () => { },
          });
        } else {
          ztoolkit.log("novel-debug", "menu.createViewContextMenu.appendHoverEntry", {
            menuLabel,
            booksCount: storage.books.length,
          });
          append({
            label: menuLabel,
            persistent: true,
            onCommand: () => {
              const doc = getReaderMenuDocument(reader);
              const anchorButton = doc ? getNovelBookMenuButtonByLabel(doc, menuLabel) : undefined;
              ztoolkit.log("novel-debug", "menu.onCommand.openPanel", {
                menuLabel,
                booksCount: storage.books.length,
                hasAnchorButton: Boolean(anchorButton),
              });
              openNovelBookMenuPanel({
                reader,
                anchor: anchorButton,
                point: menuPoint,
                menuLabel,
                recentTag: menuTexts.recentTag,
                books: storage.books,
                selectedText: fallbackSelectedText,
                selectionRange,
                selectionRect: fallbackRect,
                selectionPosition: (params as any)?.position,
                selectionAnnotation: popupSnapshot?.annotation,
              });
            },
          });

          globalThis.setTimeout(() => {
            ztoolkit.log("novel-debug", "menu.bindHoverBridge.timeout", {
              menuLabel,
              booksCount: storage.books.length,
            });
            bindNovelBookMenuHoverBridge({
              reader,
              menuLabel,
              books: storage.books,
              recentTag: menuTexts.recentTag,
              selectedText: fallbackSelectedText,
              selectionRange,
              selectionRect: fallbackRect,
              selectionPosition: (params as any)?.position,
              selectionAnnotation: popupSnapshot?.annotation,
              point: menuPoint,
            });
          }, 0);
        }
      },
      addon.data.config.addonID,
    );
  }

  @example
  static registerRightClickMenuPopup(win: Window) {
    (ztoolkit as any).Menu.register(
      "item",
      {
        tag: "menu",
      });
    // menu->File menuitem
    (ztoolkit as any).Menu.register("menuFile", {
      tag: "menuitem",
      label: getString("menuitem-filemenulabel"),
      oncommand: "alert('Hello World! File Menuitem.')",
    });
  }

  @example
  static async registerExtraColumn() {
    const field = "test1";
    await Zotero.ItemTreeManager.registerColumns({
      pluginID: addon.data.config.addonID,
      dataKey: field,
      label: "text column",
      dataProvider: (item: Zotero.Item, dataKey: string) => {
        return field + String(item.id);
      },
      iconPath: "chrome://zotero/skin/cross.png",
    });
  }

  @example
  static async registerExtraColumnWithCustomCell() {
    const field = "test2";
    await Zotero.ItemTreeManager.registerColumns({
      pluginID: addon.data.config.addonID,
      dataKey: field,
      label: "custom column",
      dataProvider: (item: Zotero.Item, dataKey: string) => {
        return field + String(item.id);
      },
      renderCell(index, data, column, isFirstColumn, doc) {
        ztoolkit.log("Custom column cell is rendered!");
        const span = doc.createElement("span");
        span.className = `cell ${column.className}`;
        span.style.background = "#0dd068";
        span.innerText = "⭐" + data;
        return span;
      },
    });
  }

  @example
  static registerItemPaneCustomInfoRow() {
    Zotero.ItemPaneManager.registerInfoRow({
      rowID: "example",
      pluginID: addon.data.config.addonID,
      editable: true,
      label: {
        l10nID: getLocaleID("item-info-row-example-label"),
      },
      position: "afterCreators",
      onGetData: ({ item }) => {
        return item.getField("title");
      },
      onSetData: ({ item, value }) => {
        item.setField("title", value);
      },
    });
  }

  @example
  static registerItemPaneSection() {
    Zotero.ItemPaneManager.registerSection({
      paneID: "example",
      pluginID: addon.data.config.addonID,
      header: {
        l10nID: getLocaleID("item-section-example1-head-text"),
        icon: "chrome://zotero/skin/16/universal/book.svg",
      },
      sidenav: {
        l10nID: getLocaleID("item-section-example1-sidenav-tooltip"),
        icon: "chrome://zotero/skin/20/universal/save.svg",
      },
      onRender: ({ body, item, editable, tabType }) => {
        body.textContent = JSON.stringify({
          id: item?.id,
          editable,
          tabType,
        });
      },
    });
  }

  @example
  static async registerReaderItemPaneSection() {
    Zotero.ItemPaneManager.registerSection({
      paneID: "reader-example",
      pluginID: addon.data.config.addonID,
      header: {
        l10nID: getLocaleID("item-section-example2-head-text"),
        // Optional
        l10nArgs: `{"status": "Initialized"}`,
        // Can also have a optional dark icon
        icon: "chrome://zotero/skin/16/universal/book.svg",
      },
      sidenav: {
        l10nID: getLocaleID("item-section-example2-sidenav-tooltip"),
        // Optional
        l10nArgs: `{"status": "Ready"}`,
        icon: "chrome://zotero/skin/20/universal/save.svg",
      },
      bodyXHTML:
        '<html:h1 id="test">THIS IS TEST</html:h1><browser disableglobalhistory="true" remote="true" maychangeremoteness="true" type="content" flex="1" id="browser" style="width: 180%; height: 280px"/>',
      onInit: async ({ body, item, setL10nArgs, setSectionSummary, setSectionButtonStatus }) => {
        ztoolkit.log("Section secondary render start!", item?.id);
        await Zotero.Promise.delay(1000);
        ztoolkit.log("Section secondary render finish!", item?.id);
        const title = body.querySelector("#test") as HTMLElement | null;
        if (title) {
          title.style.color = "green";
          title.textContent = item.getField("title");
        }
        setL10nArgs(`{ "status": "Loaded" }`);
        setSectionSummary("rendered!");
        setSectionButtonStatus("test", { hidden: false });
      },
      // Optional, Called when the section is toggled. Can happen anytime even if the section is not visible or not rendered
      onToggle: ({ item }) => {
        ztoolkit.log("Section toggled!", item?.id);
      },
      // Optional, Buttons to be shown in the section header
      sectionButtons: [
        {
          type: "test",
          icon: "chrome://zotero/skin/16/universal/empty-trash.svg",
          l10nID: getLocaleID("item-section-example2-button-tooltip"),
          onClick: ({ item, paneID }) => {
            ztoolkit.log("Section clicked!", item?.id);
            Zotero.ItemPaneManager.unregisterSection(paneID);
          },
        },
      ],
    });
  }
}

export class PromptExampleFactory {
  @example
  static registerNormalCommandExample() {
    ztoolkit.Prompt.register([
      {
        name: "Normal Command Test",
        label: "Plugin Template",
        callback(prompt) {
          ztoolkit.getGlobal("alert")("Command triggered!");
        },
      },
    ]);
  }

  @example
  static registerAnonymousCommandExample(window: Window) {
    ztoolkit.Prompt.register([
      {
        id: "search",
        callback: async (prompt) => {
          // https://github.com/zotero/zotero/blob/7262465109c21919b56a7ab214f7c7a8e1e63909/chrome/content/zotero/integration/quickFormat.js#L589
          function getItemDescription(item: Zotero.Item) {
            const nodes = [];
            let str = "";
            let author,
              authorDate = "";
            if (item.firstCreator) {
              author = authorDate = item.firstCreator;
            }
            let date = item.getField("date", true, true) as string;
            if (date && (date = date.substr(0, 4)) !== "0000") {
              authorDate += " (" + parseInt(date) + ")";
            }
            authorDate = authorDate.trim();
            if (authorDate) nodes.push(authorDate);

            const publicationTitle = item.getField(
              "publicationTitle",
              false,
              true,
            );
            if (publicationTitle) {
              nodes.push(`<i>${publicationTitle}</i>`);
            }
            let volumeIssue = item.getField("volume");
            const issue = item.getField("issue");
            if (issue) volumeIssue += "(" + issue + ")";
            if (volumeIssue) nodes.push(volumeIssue);

            const publisherPlace = [];
            let field;
            if ((field = item.getField("publisher")))
              publisherPlace.push(field);
            if ((field = item.getField("place"))) publisherPlace.push(field);
            if (publisherPlace.length) nodes.push(publisherPlace.join(": "));

            const pages = item.getField("pages");
            if (pages) nodes.push(pages);

            if (!nodes.length) {
              const url = item.getField("url");
              if (url) nodes.push(url);
            }

            // compile everything together
            for (let i = 0, n = nodes.length; i < n; i++) {
              const node = nodes[i];

              if (i != 0) str += ", ";

              if (typeof node === "object") {
                const label =
                  Zotero.getMainWindow().document.createElement("label");
                label.setAttribute("value", str);
                label.setAttribute("crop", "end");
                str = "";
              } else {
                str += node;
              }
            }
            if (str.length) str += ".";
            return str;
          }
          function filter(ids: number[]) {
            ids = ids.filter(async (id) => {
              const item = (await Zotero.Items.getAsync(id)) as Zotero.Item;
              return item.isRegularItem() && !(item as any).isFeedItem;
            });
            return ids;
          }
          const text = prompt.inputNode.value;
          prompt.showTip("Searching...");
          const s = new Zotero.Search();
          s.addCondition("quicksearch-titleCreatorYear", "contains", text);
          s.addCondition("itemType", "isNot", "attachment");
          let ids = await s.search();
          // prompt.exit will remove current container element.
          // @ts-expect-error ignore
          prompt.exit();
          const container = prompt.createCommandsContainer();
          container.classList.add("suggestions");
          ids = filter(ids);
          console.log(ids.length);
          if (ids.length == 0) {
            const s = new Zotero.Search();
            const operators = [
              "is",
              "isNot",
              "true",
              "false",
              "isInTheLast",
              "isBefore",
              "isAfter",
              "contains",
              "doesNotContain",
              "beginsWith",
            ];
            let hasValidCondition = false;
            let joinMode = "all";
            if (/\s*\|\|\s*/.test(text)) {
              joinMode = "any";
            }
            text.split(/\s*(&&|\|\|)\s*/g).forEach((conditinString: string) => {
              const conditions = conditinString.split(/\s+/g);
              if (
                conditions.length == 3 &&
                operators.indexOf(conditions[1]) != -1
              ) {
                hasValidCondition = true;
                s.addCondition(
                  "joinMode",
                  joinMode as _ZoteroTypes.Search.Operator,
                  "",
                );
                s.addCondition(
                  conditions[0] as string,
                  conditions[1] as _ZoteroTypes.Search.Operator,
                  conditions[2] as string,
                );
              }
            });
            if (hasValidCondition) {
              ids = await s.search();
            }
          }
          ids = filter(ids);
          console.log(ids.length);
          if (ids.length > 0) {
            ids.forEach((id: number) => {
              const item = Zotero.Items.get(id);
              const title = item.getField("title");
              const ele = ztoolkit.UI.createElement(window.document!, "div", {
                namespace: "html",
                classList: ["command"],
                listeners: [
                  {
                    type: "mousemove",
                    listener: function () {
                      // @ts-expect-error ignore
                      prompt.selectItem(this);
                    },
                  },
                  {
                    type: "click",
                    listener: () => {
                      prompt.promptNode.style.display = "none";
                      ztoolkit.getGlobal("Zotero_Tabs").select("zotero-pane");
                      ztoolkit.getGlobal("ZoteroPane").selectItem(item.id);
                    },
                  },
                ],
                styles: {
                  display: "flex",
                  flexDirection: "column",
                  justifyContent: "start",
                },
                children: [
                  {
                    tag: "span",
                    styles: {
                      fontWeight: "bold",
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    },
                    properties: {
                      innerText: title,
                    },
                  },
                  {
                    tag: "span",
                    styles: {
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                    },
                    properties: {
                      innerHTML: getItemDescription(item),
                    },
                  },
                ],
              });
              container.appendChild(ele);
            });
          } else {
            // @ts-expect-error ignore
            prompt.exit();
            prompt.showTip("Not Found.");
          }
        },
      },
    ]);
  }

  @example
  static registerConditionalCommandExample() {
    ztoolkit.Prompt.register([
      {
        name: "Conditional Command Test",
        label: "Plugin Template",
        // The when function is executed when Prompt UI is woken up by `Shift + P`, and this command does not display when false is returned.
        when: () => {
          const items = ztoolkit.getGlobal("ZoteroPane").getSelectedItems();
          return items.length > 0;
        },
        callback(prompt) {
          prompt.inputNode.placeholder = "Hello World!";
          const items = ztoolkit.getGlobal("ZoteroPane").getSelectedItems();
          ztoolkit.getGlobal("alert")(
            `You select ${items.length} items!\n\n${items
              .map(
                (item, index) =>
                  String(index + 1) + ". " + item.getDisplayTitle(),
              )
              .join("\n")}`,
          );
        },
      },
    ]);
  }
}

export class HelperExampleFactory {
  @example
  static async readAsNovelFromSelection(payload?: {
    reader?: _ZoteroTypes.ReaderInstance;
    selectedText?: string;
    epubPath?: string;
    selectionRange?: Range;
    selectionRect?: [number, number, number, number];
    selectionPosition?: any;
    selectionAnnotation?: any;
  }) {
    await startReplaceFromSelection(payload);
  }

  @example
  static async dialogExample() {
    const dialogData: { [key: string | number]: any } = {
      inputValue: "test",
      checkboxValue: true,
      loadCallback: () => {
        ztoolkit.log(dialogData, "Dialog Opened!");
      },
      unloadCallback: () => {
        ztoolkit.log(dialogData, "Dialog closed!");
      },
    };
    const dialogHelper = new ztoolkit.Dialog(10, 2)
      .addCell(0, 0, {
        tag: "h1",
        properties: { innerHTML: "Helper Examples" },
      })
      .addCell(1, 0, {
        tag: "h2",
        properties: { innerHTML: "Dialog Data Binding" },
      })
      .addCell(2, 0, {
        tag: "p",
        properties: {
          innerHTML:
            "Elements with attribute 'data-bind' are binded to the prop under 'dialogData' with the same name.",
        },
        styles: {
          width: "200px",
        },
      })
      .addCell(3, 0, {
        tag: "label",
        namespace: "html",
        attributes: {
          for: "dialog-checkbox",
        },
        properties: { innerHTML: "bind:checkbox" },
      })
      .addCell(
        3,
        1,
        {
          tag: "input",
          namespace: "html",
          id: "dialog-checkbox",
          attributes: {
            "data-bind": "checkboxValue",
            "data-prop": "checked",
            type: "checkbox",
          },
          properties: { label: "Cell 1,0" },
        },
        false,
      )
      .addCell(4, 0, {
        tag: "label",
        namespace: "html",
        attributes: {
          for: "dialog-input",
        },
        properties: { innerHTML: "bind:input" },
      })
      .addCell(
        4,
        1,
        {
          tag: "input",
          namespace: "html",
          id: "dialog-input",
          attributes: {
            "data-bind": "inputValue",
            "data-prop": "value",
            type: "text",
          },
        },
        false,
      )
      .addCell(5, 0, {
        tag: "h2",
        properties: { innerHTML: "Toolkit Helper Examples" },
      })
      .addCell(
        6,
        0,
        {
          tag: "button",
          namespace: "html",
          attributes: {
            type: "button",
          },
          listeners: [
            {
              type: "click",
              listener: (e: Event) => {
                addon.hooks.onDialogEvents("clipboardExample");
              },
            },
          ],
          children: [
            {
              tag: "div",
              styles: {
                padding: "2.5px 15px",
              },
              properties: {
                innerHTML: "example:clipboard",
              },
            },
          ],
        },
        false,
      )
      .addCell(
        7,
        0,
        {
          tag: "button",
          namespace: "html",
          attributes: {
            type: "button",
          },
          listeners: [
            {
              type: "click",
              listener: (e: Event) => {
                addon.hooks.onDialogEvents("filePickerExample");
              },
            },
          ],
          children: [
            {
              tag: "div",
              styles: {
                padding: "2.5px 15px",
              },
              properties: {
                innerHTML: "example:filepicker",
              },
            },
          ],
        },
        false,
      )
      .addCell(
        8,
        0,
        {
          tag: "button",
          namespace: "html",
          attributes: {
            type: "button",
          },
          listeners: [
            {
              type: "click",
              listener: (e: Event) => {
                addon.hooks.onDialogEvents("progressWindowExample");
              },
            },
          ],
          children: [
            {
              tag: "div",
              styles: {
                padding: "2.5px 15px",
              },
              properties: {
                innerHTML: "example:progressWindow",
              },
            },
          ],
        },
        false,
      )
      .addCell(
        9,
        0,
        {
          tag: "button",
          namespace: "html",
          attributes: {
            type: "button",
          },
          listeners: [
            {
              type: "click",
              listener: (e: Event) => {
                addon.hooks.onDialogEvents("vtableExample");
              },
            },
          ],
          children: [
            {
              tag: "div",
              styles: {
                padding: "2.5px 15px",
              },
              properties: {
                innerHTML: "example:virtualized-table",
              },
            },
          ],
        },
        false,
      )
      .addButton("Confirm", "confirm")
      .addButton("Cancel", "cancel")
      .addButton("Help", "help", {
        noClose: true,
        callback: (e) => {
          dialogHelper.window?.alert(
            "Help Clicked! Dialog will not be closed.",
          );
        },
      })
      .setDialogData(dialogData)
      .open("Dialog Example");
    addon.data.dialog = dialogHelper;
    await dialogData.unloadLock.promise;
    addon.data.dialog = undefined;
    if (addon.data.alive)
      ztoolkit.getGlobal("alert")(
        `Close dialog with ${dialogData._lastButtonId}.\nCheckbox: ${dialogData.checkboxValue}\nInput: ${dialogData.inputValue}.`,
      );
    ztoolkit.log(dialogData);
  }

  @example
  static clipboardExample() {
    new ztoolkit.Clipboard()
      .addText(
        "![Plugin Template](https://github.com/windingwind/zotero-plugin-template)",
        "text/unicode",
      )
      .addText(
        '<a href="https://github.com/windingwind/zotero-plugin-template">Plugin Template</a>',
        "text/html",
      )
      .copy();
    ztoolkit.getGlobal("alert")("Copied!");
  }

  @example
  static async filePickerExample() {
    const path = await new ztoolkit.FilePicker(
      "Import File",
      "open",
      [
        ["PNG File(*.png)", "*.png"],
        ["Any", "*.*"],
      ],
      "image.png",
    ).open();
    ztoolkit.getGlobal("alert")(`Selected ${path}`);
  }

  @example
  static progressWindowExample() {
    ztoolkit.log("popup-muted", {
      stage: "progressWindowExample",
      text: "ProgressWindow Example!",
    });
  }

  @example
  static vtableExample() {
    ztoolkit.getGlobal("alert")("See src/modules/preferenceScript.ts");
  }
}

function getReaderSelectionRangeSnapshot(
  reader?: _ZoteroTypes.ReaderInstance,
): Range | undefined {
  const candidateWindows: Array<Window | undefined> = [
    (reader as any)?._iframeWindow,
    (reader as any)?._window,
    (reader as any)?._iframe?.contentWindow,
    (reader as any)?._browser?.contentWindow,
    (reader as any)?.browser?.contentWindow,
    (reader as any)?._internalReader?._iframeWindow,
    (reader as any)?._internalReader?._iframe?.contentWindow,
    (reader as any)?._internalReader?._browser?.contentWindow,
    ztoolkit.getGlobal("window") as Window | undefined,
  ];

  for (const win of candidateWindows) {
    const selection = win?.getSelection?.();
    if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
      return selection.getRangeAt(0).cloneRange();
    }
  }

  return undefined;
}

function getReaderSelectionPopupSnapshot(
  reader?: _ZoteroTypes.ReaderInstance,
): {
  rect?: [number, number, number, number];
  annotation?: any;
} | undefined {
  const internalReader = (reader as any)?._internalReader;
  const primaryPopup = internalReader?._state?.primaryViewSelectionPopup;
  const secondaryPopup = internalReader?._state?.secondaryViewSelectionPopup;
  const lastPopup = internalReader?._lastView === "secondary"
    ? secondaryPopup
    : primaryPopup;
  const popup =
    lastPopup ||
    primaryPopup ||
    secondaryPopup ||
    internalReader?._primaryView?._selectionPopup ||
    internalReader?._secondaryView?._selectionPopup;

  if (!popup) {
    return undefined;
  }

  const rect = popup.rect;
  const normalizedRect =
    Array.isArray(rect) &&
      rect.length === 4 &&
      rect.every((value: unknown) => typeof value === "number" && Number.isFinite(value))
      ? (rect as [number, number, number, number])
      : undefined;

  return {
    rect: normalizedRect,
    annotation: popup.annotation,
  };
}

function buildRectFromContextPoint(
  x?: number,
  y?: number,
): [number, number, number, number] | undefined {
  if (typeof x !== "number" || typeof y !== "number") {
    return undefined;
  }

  // A small viewport rect around right-click point as a last-resort overlay anchor.
  return [x - 6, y - 18, x + 360, y + 42];
}
