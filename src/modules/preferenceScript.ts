import { config } from "../../package.json";
import {
  clearNovelCache,
  ensureSelectedNovelPath,
  getSelectedNovelBook,
  getSelectedNovelPath,
  getStorageNovelBooks,
  refreshNovelTypography,
  setSelectedNovelPath,
} from "./novel";
import { getPref, setPref } from "../utils/prefs";

type PrefLanguage = "zh-CN" | "en-US";

const DEFAULT_LANGUAGE: PrefLanguage = "zh-CN";
const DEFAULT_NOVEL_FONT_SIZE_PT = 16;
const MIN_NOVEL_FONT_SIZE_PT = 8;
const MAX_NOVEL_FONT_SIZE_PT = 72;

const PREF_TEXTS: Record<
  PrefLanguage,
  {
    title: string;
    enable: string;
    language: string;
    languageZh: string;
    languageEn: string;
    novelFontSize: string;
    bookStoragePath: string;
    bookStorageSelect: string;
    bookStorageOpen: string;
    bookStorageClear: string;
    bookShelfTitle: string;
    bookShelfHint: string;
    bookShelfSelected: string;
    bookShelfEmpty: string;
    bookShelfUnconfigured: string;
    bookShelfInvalid: string;
    bookShelfNone: string;
    bookShelfRecentTag: string;
    bookShelfSelectHint: string;
    shortcutsTitle: string;
    shortcutForward: string;
    shortcutBackward: string;
    shortcutBoss: string;
    shortcutShow: string;
  }
> = {
  "zh-CN": {
    title: "zotero-thief 设置",
    enable: "开启",
    language: "插件语言",
    languageZh: "中文",
    languageEn: "English",
    novelFontSize: "小说字体大小（pt）",
    bookStoragePath: "图书存储位置",
    bookStorageSelect: "选择目录",
    bookStorageOpen: "打开本地目录",
    bookStorageClear: "清空阅读缓存",
    bookShelfTitle: "书架",
    bookShelfHint: "点击下方列表中的小说即可选中；列表默认按最近阅读排序，可滚动浏览全部文件。",
    bookShelfSelected: "当前选择",
    bookShelfEmpty: "目录中没有 EPUB 文件",
    bookShelfUnconfigured: "请先配置图书存储目录",
    bookShelfInvalid: "配置的目录不存在",
    bookShelfNone: "尚未选择小说",
    bookShelfRecentTag: "最近阅读",
    bookShelfSelectHint: "r 键会优先使用这里选中的小说来初始化阅读界面。",
    shortcutsTitle: "快捷键设置",
    shortcutForward: "前进（默认: w）",
    shortcutBackward: "后退（默认: q）",
    shortcutBoss: "老板键/还原（默认: e）",
    shortcutShow: "显示替换（默认: r）",
  },
  "en-US": {
    title: "zotero-thief Preferences",
    enable: "Enable",
    language: "Plugin Language",
    languageZh: "Chinese",
    languageEn: "English",
    novelFontSize: "Novel Font Size (pt)",
    bookStoragePath: "Book Storage Location",
    bookStorageSelect: "Choose Folder",
    bookStorageOpen: "Open Local Folder",
    bookStorageClear: "Clear Reading Cache",
    bookShelfTitle: "Book Shelf",
    bookShelfHint: "Click a novel below to select it. The list is sorted by recent reading and can scroll through all EPUB files.",
    bookShelfSelected: "Current Selection",
    bookShelfEmpty: "No EPUB files in the folder",
    bookShelfUnconfigured: "Configure a book storage folder first",
    bookShelfInvalid: "Configured folder does not exist",
    bookShelfNone: "No novel selected yet",
    bookShelfRecentTag: "Recent",
    bookShelfSelectHint: "The r key will use the selected novel here to initialize the reading view.",
    shortcutsTitle: "Shortcut Settings",
    shortcutForward: "Forward (default: w)",
    shortcutBackward: "Backward (default: q)",
    shortcutBoss: "Boss Key / Restore (default: e)",
    shortcutShow: "Show Replacement (default: r)",
  },
};

export async function registerPrefsScripts(_window: Window) {
  if (!addon.data.prefs) {
    addon.data.prefs = {
      window: _window,
      columns: [],
      rows: [],
    };
  } else {
    addon.data.prefs.window = _window;
  }

  syncPrefsControlValues();
  applyPrefsLanguage(getNormalizedPrefLanguage());
  bindPrefEvents();
  renderBookShelf();
  setupBookShelfDropdownBehavior();
}

function getNormalizedPrefLanguage(rawValue?: string): PrefLanguage {
  return rawValue === "en-US" ? "en-US" : DEFAULT_LANGUAGE;
}

function getPrefsDocument(): Document | undefined {
  return addon.data.prefs?.window.document;
}

function getControl<T extends Element>(idSuffix: string): T | undefined {
  const doc = getPrefsDocument();
  if (!doc) {
    return undefined;
  }
  return doc.querySelector(`#zotero-prefpane-${config.addonRef}-${idSuffix}`) as
    | T
    | undefined;
}

function setText(idSuffix: string, text: string) {
  const element = getControl<HTMLElement>(idSuffix);
  if (!element) {
    return;
  }
  element.textContent = text;
}

function applyPrefsLanguage(language: PrefLanguage) {
  const text = PREF_TEXTS[language];
  setText("title", text.title);
  getControl<XUL.Checkbox>("enable")?.setAttribute("label", text.enable);
  setText("language-label", text.language);
  setText("novel-font-size-label", text.novelFontSize);
  setText("book-storage-path-label", text.bookStoragePath);
  setText("book-storage-select", text.bookStorageSelect);
  setText("book-storage-open", text.bookStorageOpen);
  setText("book-storage-clear", text.bookStorageClear);
  setText("book-shelf-title", text.bookShelfTitle);
  setText("book-shelf-hint", text.bookShelfHint);
  setText("book-shelf-select-hint", text.bookShelfSelectHint);
  setText("shortcuts-title", text.shortcutsTitle);
  setText("shortcut-forward-label", text.shortcutForward);
  setText("shortcut-backward-label", text.shortcutBackward);
  setText("shortcut-boss-label", text.shortcutBoss);
  setText("shortcut-show-label", text.shortcutShow);

  const languageSelect = getControl<HTMLSelectElement>("language");
  if (languageSelect) {
    const zhOption = languageSelect.querySelector(
      'option[value="zh-CN"]',
    ) as HTMLOptionElement | null;
    const enOption = languageSelect.querySelector(
      'option[value="en-US"]',
    ) as HTMLOptionElement | null;
    if (zhOption) {
      zhOption.textContent = text.languageZh;
    }
    if (enOption) {
      enOption.textContent = text.languageEn;
    }
  }
}

function syncPrefsControlValues() {
  const language = getNormalizedPrefLanguage(getPref("language"));
  const languageSelect = getControl<HTMLSelectElement>("language");
  if (languageSelect) {
    languageSelect.value = language;
  }

  const rawFontSize = Number(getPref("novelFontSize"));
  const fontSize = Number.isFinite(rawFontSize)
    ? Math.min(MAX_NOVEL_FONT_SIZE_PT, Math.max(MIN_NOVEL_FONT_SIZE_PT, Math.round(rawFontSize)))
    : DEFAULT_NOVEL_FONT_SIZE_PT;
  const fontSizeInput = getControl<HTMLInputElement>("novel-font-size");
  if (fontSizeInput) {
    fontSizeInput.value = String(fontSize);
  }

  const storagePathInput = getControl<HTMLInputElement>("book-storage-path");
  if (storagePathInput) {
    storagePathInput.value = String(getPref("bookStoragePath") || "");
  }
}

function getBookShelfSelect(): HTMLSelectElement | undefined {
  return getControl<HTMLSelectElement>("book-shelf");
}

function getBookShelfStatus(): HTMLElement | undefined {
  return getControl<HTMLElement>("book-shelf-status");
}

function collapseBookShelfDropdown() {
  const select = getBookShelfSelect();
  if (!select) {
    return;
  }
  select.size = 1;
}

function expandBookShelfDropdown() {
  const select = getBookShelfSelect();
  if (!select) {
    return;
  }
  const optionCount = Math.max(1, select.options.length);
  select.size = Math.min(5, optionCount);
}

function setupBookShelfDropdownBehavior() {
  const select = getBookShelfSelect();
  const doc = getPrefsDocument();
  if (!select || !doc) {
    return;
  }

  if (select.dataset.dropdownBound === "1") {
    return;
  }
  select.dataset.dropdownBound = "1";

  select.addEventListener("focus", () => {
    expandBookShelfDropdown();
  });

  select.addEventListener("mousedown", () => {
    expandBookShelfDropdown();
  });

  select.addEventListener("blur", () => {
    collapseBookShelfDropdown();
  });

  doc.addEventListener("mousedown", (event: MouseEvent) => {
    const target = event.target as Node | null;
    if (!target || !select.contains(target)) {
      collapseBookShelfDropdown();
    }
  });
}

function renderBookShelf() {
  const select = getBookShelfSelect();
  const status = getBookShelfStatus();
  if (!select) {
    return;
  }

  const doc = select.ownerDocument;
  if (!doc) {
    return;
  }

  const storage = getStorageNovelBooks();
  const language = getNormalizedPrefLanguage(getPref("language"));
  const text = PREF_TEXTS[language];
  const selectedPath = ensureSelectedNovelPath();

  select.textContent = "";

  if (!storage.configured) {
    const option = doc.createElement("option");
    option.disabled = true;
    option.selected = true;
    option.textContent = text.bookShelfUnconfigured;
    select.appendChild(option);
    if (status) {
      status.textContent = `${text.bookShelfSelected}: ${text.bookShelfNone}`;
    }
    return;
  }

  if (!storage.dirExists) {
    const option = doc.createElement("option");
    option.disabled = true;
    option.selected = true;
    option.textContent = text.bookShelfInvalid;
    select.appendChild(option);
    if (status) {
      status.textContent = `${text.bookShelfSelected}: ${text.bookShelfNone}`;
    }
    return;
  }

  if (storage.books.length === 0) {
    const option = doc.createElement("option");
    option.disabled = true;
    option.selected = true;
    option.textContent = text.bookShelfEmpty;
    select.appendChild(option);
    if (status) {
      status.textContent = `${text.bookShelfSelected}: ${text.bookShelfNone}`;
    }
    return;
  }

  for (const book of storage.books) {
    const option = doc.createElement("option");
    option.value = book.filePath;
    option.textContent = Number(book.lastReadAt) > 0
      ? `[${text.bookShelfRecentTag}] ${book.fileName}`
      : book.fileName;
    if (selectedPath && book.filePath === selectedPath) {
      option.selected = true;
    }
    select.appendChild(option);
  }

  if (!select.value && storage.books[0]) {
    select.value = storage.books[0].filePath;
    setSelectedNovelPath(storage.books[0].filePath);
  }

  if (status) {
    const selectedBook = getSelectedNovelBook();
    status.textContent = `${text.bookShelfSelected}: ${selectedBook?.fileName || text.bookShelfNone}`;
  }
}

async function chooseStorageDirectory(): Promise<string | undefined> {
  const prefsWindow = addon.data.prefs?.window;
  if (!prefsWindow) {
    return undefined;
  }

  const language = getNormalizedPrefLanguage(getPref("language"));
  const title = language === "en-US" ? "Choose Book Folder" : "选择图书目录";
  const currentPath = String(getPref("bookStoragePath") || "").trim();
  const selectedPath = await new ztoolkit.FilePicker(
    title,
    "folder",
    [],
    "",
    prefsWindow,
    undefined,
    currentPath || undefined,
  ).open();

  return selectedPath || undefined;
}

function openDirectoryByPath(path: string): boolean {
  const trimmed = String(path || "").trim();
  if (!trimmed) {
    return false;
  }

  try {
    const classesAny = Components.classes as any;
    const file = classesAny["@mozilla.org/file/local;1"].createInstance(
      Components.interfaces.nsIFile,
    ) as any;
    file.initWithPath(trimmed);
    if (!file.exists() || !file.isDirectory()) {
      return false;
    }

    try {
      file.reveal();
      return true;
    } catch (_error) {
      // Fallback for platforms where reveal is not available.
      file.launch();
      return true;
    }
  } catch (_error) {
    return false;
  }
}

function bindPrefEvents() {
  const languageSelect = getControl<HTMLSelectElement>("language");
  languageSelect?.addEventListener("change", (event: Event) => {
    const select = event.target as HTMLSelectElement;
    const language = getNormalizedPrefLanguage(select.value);
    select.value = language;
    setPref("language", language);

    // Fluent may update labels asynchronously after panel load/change, so re-apply in later ticks.
    applyPrefsLanguage(language);
    addon.data.prefs?.window.setTimeout(() => applyPrefsLanguage(language), 0);
    addon.data.prefs?.window.setTimeout(() => applyPrefsLanguage(language), 80);
    addon.data.prefs?.window.setTimeout(() => renderBookShelf(), 0);
    addon.data.prefs?.window.setTimeout(() => renderBookShelf(), 80);
  });

  const fontSizeInput = getControl<HTMLInputElement>("novel-font-size");
  fontSizeInput?.addEventListener("change", (event: Event) => {
    const input = event.target as HTMLInputElement;
    const parsed = Number.parseInt(input.value, 10);
    const fontSize = Number.isFinite(parsed)
      ? Math.min(MAX_NOVEL_FONT_SIZE_PT, Math.max(MIN_NOVEL_FONT_SIZE_PT, parsed))
      : DEFAULT_NOVEL_FONT_SIZE_PT;
    input.value = String(fontSize);
    setPref("novelFontSize", fontSize);
    refreshNovelTypography();
  });

  const storagePathInput = getControl<HTMLInputElement>("book-storage-path");
  const selectStorageButton = getControl<HTMLButtonElement>("book-storage-select");
  const openStorageButton = getControl<HTMLButtonElement>("book-storage-open");
  const clearStorageButton = getControl<HTMLButtonElement>("book-storage-clear");

  selectStorageButton?.addEventListener("click", async () => {
    const selectedPath = await chooseStorageDirectory();
    if (!selectedPath) {
      return;
    }
    setPref("bookStoragePath", selectedPath);
    if (storagePathInput) {
      storagePathInput.value = selectedPath;
    }
    renderBookShelf();
  });

  openStorageButton?.addEventListener("click", () => {
    const path = String(getPref("bookStoragePath") || "");
    if (!openDirectoryByPath(path)) {
      ztoolkit.log("book-storage.open.failed", { path });
    }
  });

  clearStorageButton?.addEventListener("click", () => {
    const path = String(getPref("bookStoragePath") || "");
    const cleared = clearNovelCache(path);
    ztoolkit.log("book-storage.clear", { path, cleared });
    renderBookShelf();
  });

  const bookShelfSelect = getBookShelfSelect();
  bookShelfSelect?.addEventListener("change", (event: Event) => {
    const select = event.target as HTMLSelectElement;
    const selectedPath = String(select.value || "").trim();
    if (!selectedPath) {
      collapseBookShelfDropdown();
      return;
    }
    setSelectedNovelPath(selectedPath);
    collapseBookShelfDropdown();
    renderBookShelf();
  });

  const shortcutMappings: Array<{
    idSuffix: string;
    pref: "shortcutForward" | "shortcutBackward" | "shortcutBoss" | "shortcutShow";
    fallback: string;
  }> = [
      { idSuffix: "shortcut-forward", pref: "shortcutForward", fallback: "w" },
      { idSuffix: "shortcut-backward", pref: "shortcutBackward", fallback: "q" },
      { idSuffix: "shortcut-boss", pref: "shortcutBoss", fallback: "e" },
      { idSuffix: "shortcut-show", pref: "shortcutShow", fallback: "r" },
    ];

  for (const item of shortcutMappings) {
    addon.data.prefs!.window.document
      ?.querySelector(`#zotero-prefpane-${config.addonRef}-${item.idSuffix}`)
      ?.addEventListener("change", (event: Event) => {
        const input = event.target as HTMLInputElement;
        const normalized = (input.value || "").trim().slice(0, 1).toLowerCase();
        const value = normalized || item.fallback;
        input.value = value;
        setPref(item.pref, value);
      });
  }
}
