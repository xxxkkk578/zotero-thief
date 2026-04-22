import { getPref, setPref } from "../utils/prefs";

type SelectionRect = [number, number, number, number];
type SelectionPositionLike = any;

type SelectionAnnotationLike = {
    text?: string;
    position?: SelectionPositionLike;
};

type ReaderPayload = {
    reader?: _ZoteroTypes.ReaderInstance;
    selectedText?: string;
    epubPath?: string;
    selectionRange?: Range;
    selectionRect?: SelectionRect;
    selectionPosition?: SelectionPositionLike;
    selectionAnnotation?: SelectionAnnotationLike;
};

type ReaderSelectionSnapshot = {
    selectedText: string;
    selectionRange?: Range;
    selectionRect?: SelectionRect;
    selectionPosition?: SelectionPositionLike;
    selectionAnnotation?: SelectionAnnotationLike;
};

type KeyboardSelectionSnapshot = {
    selectedText: string;
    selectionRange?: Range;
};

type ReplacementState = {
    reader?: _ZoteroTypes.ReaderInstance;
    originalText: string;
    selectedLength: number;
    initialRange?: Range;
    selectionRect?: SelectionRect;
    selectionPosition?: SelectionPositionLike;
    selectionAnnotation?: SelectionAnnotationLike;
    marker?: HTMLElement;
    pdfVisualContainers?: HTMLElement[];
    pdfVisualSyncCleanup?: () => void;
    pdfVisualHidden?: boolean;
    pdfVisualParts?: Array<{
        element: HTMLElement;
        pageIndex: number;
        expectedLength: number;
        maxWidth: number;
        softMaxWidth: number;
        font: string;
        letterSpacingPx: number;
        leftRatio: number;
        topRatio: number;
        widthRatio: number;
        heightRatio: number;
        basePageWidth: number;
        basePageHeight: number;
        baseFontSizePx?: number;
        baseLineHeightPx?: number;
        baseLetterSpacingPx?: number;
    }>;
    activeChunkLength?: number;
    textReplacementParts?: Array<{
        node: Text;
        startOffset: number;
        endOffset: number;
        originalNodeValue: string;
    }>;
};

type NovelState = {
    sourcePath: string;
    sourceHash: string;
    englishText: string;
    cursor: number;
    history: number[];
};

type NovelCacheBookRecord = {
    hash: string;
    filePath: string;
    fileName: string;
    cursor: number;
    lastReadAt: number;
};

type NovelCacheData = {
    version: 1;
    booksByHash: Record<string, NovelCacheBookRecord>;
    pathToHash: Record<string, string>;
};

export type NovelStorageBook = {
    filePath: string;
    fileName: string;
    lastReadAt: number;
};

export type NovelStorageBooksResult = {
    configured: boolean;
    dirExists: boolean;
    books: NovelStorageBook[];
};

const NOVEL_MARKER_CLASS = "myreader-novel-replacement";
const NOVEL_LATEX_FONT_STACK = '"Latin Modern Roman", "Computer Modern Serif", "CMU Serif", "STIX Two Text", "Times New Roman", serif';
const NOVEL_LATEX_FONT_SIZE_PT = 13;
const NOVEL_LATEX_MIN_FONT_SIZE_PT = 8;
const NOVEL_LATEX_MAX_FONT_SIZE_PT = 72;
const NOVEL_LATEX_LINE_HEIGHT_RATIO = 1.15;
const NOVEL_JUSTIFY_TIGHTEN_PX = 0.3;
const NOVEL_TEX_STRETCH_EM = 0.18;
const NOVEL_TEX_SHRINK_EM = 0.09;
const NOVEL_TEX_HYPHEN_PENALTY = 45;
const NOVEL_TEX_CONSECUTIVE_HYPHEN_PENALTY = 90;
const NOVEL_TEX_FITNESS_DEMERIT = 120;
const NOVEL_TEX_INFINITE_BADNESS = 10000;
const NOVEL_PDF_ROW_VERTICAL_PADDING_PX = 1.4;
const NOVEL_PDF_ROW_FONT_TO_BOX_RATIO = 0.88;
const NOVEL_PDF_ROW_LINE_HEIGHT_TO_BOX_RATIO = 0.98;
const NOVEL_CACHE_FILE_NAME = ".zotero-thief-cache.json";
const NOVEL_CACHE_WRITE_DEBOUNCE_MS = 450;
const DEFAULT_CHUNK_MIN = 80;
const TEXT_NODE_TYPE = 3;
const ELEMENT_NODE_TYPE = 1;

let replacementState: ReplacementState | undefined;
let novelState: NovelState | undefined;
let shortcutsRegistered = false;
let nativeShortcutFallbackRegistered = false;
let lastShortcutKey = "";
let lastShortcutTime = 0;
const readerSelectionSnapshots = new Map<string, ReaderSelectionSnapshot>();
const readerSelectionReaders = new Map<string, _ZoteroTypes.ReaderInstance>();
let lastSelectionReader: _ZoteroTypes.ReaderInstance | undefined;
const novelTextMeasurerByDoc = new WeakMap<Document, HTMLElement>();
let novelCachePersistTimer: number | undefined;
let novelCachePersistReason = "";

function getNovelFontSizePt() {
    const rawSize = Number(getPref("novelFontSize"));
    if (!Number.isFinite(rawSize)) {
        return NOVEL_LATEX_FONT_SIZE_PT;
    }
    return Math.min(
        NOVEL_LATEX_MAX_FONT_SIZE_PT,
        Math.max(NOVEL_LATEX_MIN_FONT_SIZE_PT, rawSize),
    );
}

function getNovelFontSizePx() {
    return ptToPx(getNovelFontSizePt());
}

function getNovelLineHeightPx() {
    return ptToPx(getNovelFontSizePt() * NOVEL_LATEX_LINE_HEIGHT_RATIO);
}

function applyNovelTypographyToMarker(marker: HTMLElement) {
    marker.style.fontFamily = NOVEL_LATEX_FONT_STACK;
    marker.style.fontSize = `${getNovelFontSizePx().toFixed(2)}px`;
    marker.style.lineHeight = `${getNovelLineHeightPx().toFixed(2)}px`;
    marker.style.letterSpacing = "0px";
    marker.style.wordSpacing = "normal";
    marker.style.fontWeight = "400";
    marker.style.fontStyle = "normal";
    marker.style.fontKerning = "normal";
    marker.style.fontVariantLigatures = "common-ligatures";
    marker.style.textAlign = "left";
}

export function refreshNovelTypography() {
    if (!replacementState || !novelState) {
        return;
    }

    if (replacementState.marker) {
        applyNovelTypographyToMarker(replacementState.marker);
    }

    if (replacementState.pdfVisualParts?.length) {
        for (const part of replacementState.pdfVisualParts) {
            applyNovelTypographyToMarker(part.element);
            part.element.style.font = `normal 400 ${part.element.style.fontSize} ${NOVEL_LATEX_FONT_STACK}`;
            const baseFontSizePx = Number.parseFloat(part.element.style.fontSize);
            const baseLineHeightPx = Number.parseFloat(part.element.style.lineHeight);
            const baseLetterSpacingPx = Number.parseFloat(part.element.style.letterSpacing);
            part.baseFontSizePx = Number.isFinite(baseFontSizePx) ? baseFontSizePx : undefined;
            part.baseLineHeightPx = Number.isFinite(baseLineHeightPx) ? baseLineHeightPx : undefined;
            part.baseLetterSpacingPx = Number.isFinite(baseLetterSpacingPx) ? baseLetterSpacingPx : undefined;
        }
        refreshPdfVisualOverlayGeometry(replacementState.reader);
    }

    applyChunkToActiveMarker(true);
}

export function rememberReaderSelectionSnapshot(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    snapshot: ReaderSelectionSnapshot,
) {
    if (!reader || !snapshot.selectedText.trim()) {
        return;
    }

    readerSelectionReaders.set(getReaderSnapshotKey(reader), reader);
    lastSelectionReader = reader;

    readerSelectionSnapshots.set(getReaderSnapshotKey(reader), {
        selectedText: snapshot.selectedText.trim(),
        selectionRange: cloneNonCollapsedRange(snapshot.selectionRange),
        selectionRect: snapshot.selectionRect,
        selectionPosition: snapshot.selectionPosition,
        selectionAnnotation: snapshot.selectionAnnotation,
    });
}

export async function startReplaceFromSelection(payload?: ReaderPayload) {
    try {
        debugNovel("startReplaceFromSelection.begin", {
            hasReader: Boolean(payload?.reader),
            payloadSelectedLength: payload?.selectedText?.length || 0,
            hasSelectionRange: Boolean(payload?.selectionRange),
            payloadSelectionRange: describeRange(payload?.selectionRange),
            hasSelectionRect: Boolean(payload?.selectionRect),
            hasSelectionPosition: Boolean(payload?.selectionPosition),
        });

        const rawSelectedText = payload?.selectedText ?? "";
        const selectedText = rawSelectedText.trim();
        const initialRange = await resolveInitialRange(
            payload?.reader,
            payload?.selectionRange,
            payload?.selectionPosition,
            payload?.selectionAnnotation,
            rawSelectedText || selectedText,
        );

        const remembered = getRememberedSelectionSnapshot(payload?.reader);

        const fallbackSelectedText =
            selectedText ||
            remembered?.selectedText ||
            initialRange?.toString().trim() ||
            (payload?.reader ? ztoolkit.Reader.getSelectedText(payload.reader).trim() : "");

        if (!fallbackSelectedText) {
            debugNovel("startReplaceFromSelection.noSelectedText", {
                rawSelectedLength: rawSelectedText.length,
                hasInitialRange: Boolean(initialRange),
            });
            showProgress("未检测到选中文本，请先在阅读器中选中内容。", "warning");
            return;
        }

        if (!initialRange) {
            debugNovel("startReplaceFromSelection.noInitialRange", {
                fallbackSelectedLength: fallbackSelectedText.length,
                selectedPreview: fallbackSelectedText.slice(0, 80),
            });
            showProgress("未能恢复阅读器选区，将继续选择 EPUB；如需立即替换，请重新选中文本后再试。", "warning");
        }

        const epubPath = payload?.epubPath || ensureSelectedNovelPath();

        debugNovel("startReplaceFromSelection.bookPath", {
            picked: Boolean(epubPath),
            path: epubPath || "",
        });

        if (!epubPath) {
            showProgress("请先在设置页选择一本小说。", "warning");
            return;
        }

        if (!/\.epub$/i.test(epubPath)) {
            showProgress("仅支持选择 EPUB 文件。", "warning");
            return;
        }

        const englishText = extractEnglishTextFromEpub(epubPath);
        debugNovel("startReplaceFromSelection.epubExtracted", {
            englishLength: englishText.length,
        });
        if (!englishText) {
            showProgress("未从 EPUB 中提取到可用英文文本。", "warning");
            return;
        }

        // Enforce single active replacement block: clear previous block before applying a new one.
        clearActiveReplacementBlock("startReplaceFromSelection.newSelection", payload?.reader);

        const bookHash = computeFileSha256(epubPath);
        const cachedCursor = getCachedCursorForBook(epubPath, bookHash);
        const initialCursor = clampCursor(cachedCursor, englishText.length);

        novelState = {
            sourcePath: epubPath,
            sourceHash: bookHash,
            englishText,
            cursor: initialCursor,
            history: [],
        };

        replacementState = {
            reader: payload?.reader,
            originalText: fallbackSelectedText,
            selectedLength: fallbackSelectedText.length,
            initialRange,
            selectionRect: payload?.selectionRect || remembered?.selectionRect,
            selectionPosition: payload?.selectionPosition || remembered?.selectionPosition,
            selectionAnnotation: payload?.selectionAnnotation || remembered?.selectionAnnotation,
        };

        debugNovel("startReplaceFromSelection.readyToApply", {
            selectedLength: fallbackSelectedText.length,
            initialRange: describeRange(initialRange),
        });

        if (!replacementState.initialRange) {
            const lateRange = await resolveInitialRange(
                payload?.reader,
                payload?.selectionRange,
                payload?.selectionPosition,
                payload?.selectionAnnotation,
                fallbackSelectedText,
            );
            if (lateRange) {
                replacementState.initialRange = lateRange;
                debugNovel("startReplaceFromSelection.lateRange", {
                    lateRange: describeRange(lateRange),
                });
            }
        }

        if (applyChunkToCurrentSelection()) {
            persistCurrentNovelProgress("startReplaceFromSelection.applied", true);
            showProgress("已加载小说并完成替换。可使用快捷键前进/后退/还原/显示。", "success");
        } else {
            showProgress("已加载 EPUB，暂未替换阅读器内容。请重新选中文本后按显示键。", "warning");
        }
    } catch (error) {
        ztoolkit.log("readAsNovel error", error);
        debugNovel("startReplaceFromSelection.exception", {
            error: String(error),
        });
        showProgress("替换失败，请重试。", "error");
    }
}

async function resolveInitialRange(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    selectionRange: Range | undefined,
    selectionPosition: SelectionPositionLike | undefined,
    selectionAnnotation: SelectionAnnotationLike | undefined,
    selectedText: string,
): Promise<Range | undefined> {
    const snapshot = getRememberedSelectionSnapshot(reader);
    const fromPayload = cloneNonCollapsedRange(selectionRange);
    const fromPayloadPosition = cloneNonCollapsedRange(
        findRangeBySelectionPosition(reader, selectionPosition),
    );
    const fromPayloadAnnotation = cloneNonCollapsedRange(
        findRangeBySelectionAnnotation(reader, selectionAnnotation),
    );
    const fromSnapshotRange = cloneNonCollapsedRange(snapshot?.selectionRange);
    const fromSnapshotPosition = cloneNonCollapsedRange(
        findRangeBySelectionPosition(reader, snapshot?.selectionPosition),
    );
    const fromSnapshotAnnotation = cloneNonCollapsedRange(
        findRangeBySelectionAnnotation(reader, snapshot?.selectionAnnotation),
    );
    const fromLiveSelection = cloneNonCollapsedRange(getReaderSelectionRange(reader));
    const fromTextSearch = cloneNonCollapsedRange(findRangeBySelectedText(reader, selectedText));

    let range =
        fromPayload ||
        fromPayloadPosition ||
        fromPayloadAnnotation ||
        fromSnapshotRange ||
        fromSnapshotPosition ||
        fromSnapshotAnnotation ||
        fromLiveSelection ||
        fromTextSearch;

    debugNovel("resolveInitialRange.firstPass", {
        selectedLength: selectedText.length,
        fromPayload: describeRange(fromPayload),
        fromPayloadPosition: describeRange(fromPayloadPosition),
        fromPayloadAnnotation: describeRange(fromPayloadAnnotation),
        fromSnapshotRange: describeRange(fromSnapshotRange),
        fromSnapshotPosition: describeRange(fromSnapshotPosition),
        fromSnapshotAnnotation: describeRange(fromSnapshotAnnotation),
        fromLiveSelection: describeRange(fromLiveSelection),
        fromTextSearch: describeRange(fromTextSearch),
        chosen: describeRange(range),
    });

    if (range) {
        return range;
    }

    // Context menu can temporarily clear selection; retry once after event loop tick.
    await Zotero.Promise.delay(60);
    const retrySnapshot = getRememberedSelectionSnapshot(reader);
    const retrySnapshotRange = cloneNonCollapsedRange(retrySnapshot?.selectionRange);
    const retrySnapshotPosition = cloneNonCollapsedRange(
        findRangeBySelectionPosition(reader, retrySnapshot?.selectionPosition),
    );
    const retrySnapshotAnnotation = cloneNonCollapsedRange(
        findRangeBySelectionAnnotation(reader, retrySnapshot?.selectionAnnotation),
    );
    const retryLiveSelection = cloneNonCollapsedRange(getReaderSelectionRange(reader));
    const retryTextSearch = cloneNonCollapsedRange(findRangeBySelectedText(reader, selectedText));
    range = retrySnapshotRange || retrySnapshotPosition || retrySnapshotAnnotation || retryLiveSelection || retryTextSearch;

    debugNovel("resolveInitialRange.retryPass", {
        retrySnapshotRange: describeRange(retrySnapshotRange),
        retrySnapshotPosition: describeRange(retrySnapshotPosition),
        retrySnapshotAnnotation: describeRange(retrySnapshotAnnotation),
        retryLiveSelection: describeRange(retryLiveSelection),
        retryTextSearch: describeRange(retryTextSearch),
        chosen: describeRange(range),
    });

    return range;
}

export function registerNovelShortcuts() {
    if (shortcutsRegistered) {
        return;
    }

    shortcutsRegistered = true;
    debugNovel("shortcut.register.begin", {
        hasMainWindows: Zotero.getMainWindows().length,
    });

    ztoolkit.Keyboard.register((event) => {
        handleNovelShortcutEvent(event as KeyboardEvent, "ztoolkit");
    });

    registerNativeShortcutFallback();
}

function registerNativeShortcutFallback() {
    if (nativeShortcutFallbackRegistered) {
        return;
    }

    nativeShortcutFallbackRegistered = true;
    const mainWindows = Zotero.getMainWindows();
    for (const win of mainWindows) {
        registerNativeShortcutFallbackOnWindow(win as Window);
    }

    debugNovel("shortcut.register.nativeFallback", {
        windowCount: mainWindows.length,
    });
}

function registerNativeShortcutFallbackOnWindow(win?: Window) {
    if (!win) {
        return;
    }

    const marker = "__myreaderNovelShortcutFallbackRegistered";
    if ((win as any)[marker]) {
        return;
    }
    (win as any)[marker] = true;

    win.addEventListener(
        "keydown",
        (event: KeyboardEvent) => {
            handleNovelShortcutEvent(event, "native");
        },
        true,
    );

    for (const frameWin of getChildFrameWindows(win)) {
        registerNativeShortcutFallbackOnWindow(frameWin);
    }
}

function handleNovelShortcutEvent(event: KeyboardEvent, source: "ztoolkit" | "native") {
    if (!addon.data.alive) {
        return;
    }

    const eventType = String(event.type || "").toLowerCase();
    if (eventType && eventType !== "keydown") {
        return;
    }

    const key = String(event.key || "").toLowerCase();
    const forwardKey = getShortcutKey("shortcutForward", "w");
    const backwardKey = getShortcutKey("shortcutBackward", "q");
    const bossKey = getShortcutKey("shortcutBoss", "e");
    const showKey = getShortcutKey("shortcutShow", "r");
    const isRelevantKey = key === forwardKey || key === backwardKey || key === bossKey || key === showKey || key === "r";

    if (!isRelevantKey) {
        return;
    }

    debugNovel("shortcut.keydown.candidate", {
        source,
        pressedKey: key,
        configuredForwardKey: forwardKey,
        configuredBackwardKey: backwardKey,
        configuredBossKey: bossKey,
        configuredShowKey: showKey,
        repeat: Boolean(event.repeat),
        metaKey: Boolean(event.metaKey),
        ctrlKey: Boolean(event.ctrlKey),
        altKey: Boolean(event.altKey),
        targetTag: (event.target as HTMLElement | null)?.tagName || "unknown",
        editableTarget: isEditableTarget(event.target),
        hasNovelState: Boolean(novelState),
        hasReplacementState: Boolean(replacementState),
    });

    if (event.repeat) {
        debugNovel("shortcut.keydown.abort.repeat", { source, pressedKey: key });
        return;
    }

    if (event.metaKey || event.ctrlKey || event.altKey) {
        debugNovel("shortcut.keydown.abort.modifier", { source, pressedKey: key });
        return;
    }

    if (isEditableTarget(event.target)) {
        debugNovel("shortcut.keydown.abort.editableTarget", { source, pressedKey: key });
        return;
    }

    const now = Date.now();
    if (key === lastShortcutKey && now - lastShortcutTime < 180) {
        debugNovel("shortcut.keydown.debounced", {
            source,
            pressedKey: key,
            deltaMs: now - lastShortcutTime,
        });
        return;
    }
    lastShortcutKey = key;
    lastShortcutTime = now;

    if (key === forwardKey) {
        event.preventDefault();
        debugNovel("shortcut.keydown.dispatchForward", { source, pressedKey: key });
        forwardChunk();
        return;
    }
    if (key === backwardKey) {
        event.preventDefault();
        debugNovel("shortcut.keydown.dispatchBackward", { source, pressedKey: key });
        backwardChunk();
        return;
    }
    if (key === bossKey) {
        event.preventDefault();
        debugNovel("shortcut.keydown.dispatchBoss", { source, pressedKey: key });
        restoreOriginalText();
        return;
    }
    if (key === showKey) {
        event.preventDefault();
        const preferredReader = lastSelectionReader || replacementState?.reader;
        const rememberedSnapshot = getRememberedSelectionSnapshot(preferredReader);
        const eventSnapshot = getKeyboardSelectionSnapshot(event);
        const pdfSelectionRange = getPdfSelectionRange(preferredReader);
        const hasSelection = Boolean(
            eventSnapshot.selectedText ||
            eventSnapshot.selectionRange ||
            pdfSelectionRange ||
            getReaderSelectionRange(preferredReader) ||
            rememberedSnapshot?.selectedText,
        );

        debugNovel("shortcut.keydown.dispatchShow", {
            source,
            pressedKey: key,
            configuredShowKey: showKey,
            hasSelection,
            hasPreferredReader: Boolean(preferredReader),
            eventSelectionLength: eventSnapshot.selectedText.length,
            hasEventSelectionRange: Boolean(eventSnapshot.selectionRange),
            hasPdfSelectionRange: Boolean(pdfSelectionRange),
        });

        if (hasSelection) {
            const epubPath = ensureSelectedNovelPath();
            if (!epubPath) {
                showProgress("请先在设置页选择一本小说。", "warning");
                return;
            }

            void startReplaceFromSelection({
                reader: preferredReader,
                selectedText: eventSnapshot.selectedText || rememberedSnapshot?.selectedText,
                selectionRange: eventSnapshot.selectionRange || pdfSelectionRange || rememberedSnapshot?.selectionRange,
                selectionRect: rememberedSnapshot?.selectionRect,
                selectionPosition: rememberedSnapshot?.selectionPosition,
                selectionAnnotation: rememberedSnapshot?.selectionAnnotation,
                epubPath,
            });
            return;
        }

        showProgress(`检测到显示快捷键（${showKey.toUpperCase()}），正在尝试恢复...`, "default");
        reShowChunk();
    }
}

function getKeyboardSelectionSnapshot(event: KeyboardEvent): KeyboardSelectionSnapshot {
    const windows: Array<Window | undefined> = [];
    const target = event.target as Node | null;
    const targetView = target?.ownerDocument?.defaultView as Window | undefined;

    windows.push(event.view as Window | undefined);
    windows.push(targetView);
    windows.push(ztoolkit.getGlobal("window") as Window | undefined);

    const visited = new Set<Window>();
    const candidates: Window[] = [];
    for (const win of windows) {
        if (win && !visited.has(win)) {
            visited.add(win);
            candidates.push(win);
        }
    }

    for (let i = 0; i < candidates.length; i++) {
        const win = candidates[i];
        for (const child of getChildFrameWindows(win)) {
            if (!visited.has(child)) {
                visited.add(child);
                candidates.push(child);
            }
        }
    }

    for (const win of candidates) {
        const selection = win?.getSelection?.();
        if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
            continue;
        }
        const range = selection.getRangeAt(0);
        const selectedText = range.toString().trim();
        if (!selectedText) {
            continue;
        }
        return {
            selectedText,
            selectionRange: range.cloneRange(),
        };
    }

    return {
        selectedText: "",
    };
}

export function forwardChunk() {
    if (!ensureNovelReady()) {
        return;
    }

    const step = getCurrentChunkLength();
    const currentCursor = clampCursor(novelState!.cursor, novelState!.englishText.length);
    const nextCursor = advanceCursorByDisplayed(
        novelState!.englishText,
        currentCursor,
        step,
    );
    if (nextCursor !== currentCursor) {
        const history = novelState!.history;
        if (history.length === 0 || history[history.length - 1] !== currentCursor) {
            history.push(currentCursor);
        }
        novelState!.cursor = nextCursor;
        persistCurrentNovelProgress("forwardChunk");
    }
    applyChunkToActiveMarker(true);
}

export function backwardChunk() {
    if (!ensureNovelReady()) {
        return;
    }

    const history = novelState!.history;
    if (history.length > 0) {
        const previousCursor = history.pop()!;
        novelState!.cursor = clampCursor(previousCursor, novelState!.englishText.length);
    } else {
        const step = getCurrentChunkLength();
        novelState!.cursor = clampCursor(novelState!.cursor - step, novelState!.englishText.length);
    }
    persistCurrentNovelProgress("backwardChunk");
    applyChunkToActiveMarker(true);
}

export function restoreOriginalText() {
    if (replacementState?.pdfVisualParts?.length) {
        if (replacementState.pdfVisualContainers?.length) {
            for (const container of replacementState.pdfVisualContainers) {
                container.style.display = "none";
            }
        } else if (replacementState.marker) {
            replacementState.marker.style.display = "none";
        }
        replacementState.pdfVisualHidden = true;
        showProgress("已还原为原论文文本。", "success");
        return;
    }

    if (replacementState?.textReplacementParts?.length) {
        replacementState.pdfVisualSyncCleanup?.();
        replacementState.pdfVisualSyncCleanup = undefined;
        for (const part of replacementState.textReplacementParts) {
            part.node.nodeValue = part.originalNodeValue;
        }
        replacementState.textReplacementParts = undefined;
        replacementState.marker = undefined;
        replacementState.pdfVisualContainers = undefined;
        replacementState.activeChunkLength = undefined;
        showProgress("已还原为原论文文本。", "success");
        return;
    }

    if (!replacementState?.marker) {
        showProgress("没有可还原的替换内容。", "warning");
        return;
    }

    if (replacementState.marker.dataset.novelMode === "overlay") {
        replacementState.marker.remove();
        replacementState.marker = undefined;
    } else {
        replacementState.marker.textContent = replacementState.originalText;
    }
    showProgress("已还原为原论文文本。", "success");
}

export function reShowChunk() {
    debugNovel("reShowChunk.enter", collectReShowDebugState());

    if (!ensureNovelReady()) {
        debugNovel("reShowChunk.abort.notReady", collectReShowDebugState());
        return;
    }

    if (replacementState?.pdfVisualParts?.length && replacementState.pdfVisualHidden) {
        debugNovel("reShowChunk.branch.restoreHiddenPdfVisual", collectReShowDebugState());
        if (replacementState.pdfVisualContainers?.length) {
            for (const container of replacementState.pdfVisualContainers) {
                container.style.display = "";
            }
        } else if (replacementState.marker) {
            replacementState.marker.style.display = "";
        }
        replacementState.pdfVisualHidden = false;
        refreshPdfVisualOverlayGeometry(replacementState?.reader);
        return;
    }

    const cachedPosition = replacementState?.selectionPosition || replacementState?.selectionAnnotation?.position;
    if (!replacementState?.marker && !replacementState?.pdfVisualParts?.length && cachedPosition) {
        debugNovel("reShowChunk.branch.tryPdfPosition", collectReShowDebugState());
        if (applyChunkByPdfPositionPreserveStyles(replacementState?.reader, cachedPosition)) {
            debugNovel("reShowChunk.branch.tryPdfPosition.success", collectReShowDebugState());
            return;
        }
        debugNovel("reShowChunk.branch.tryPdfPosition.failed", collectReShowDebugState());
    }

    if (replacementState?.textReplacementParts?.length) {
        debugNovel("reShowChunk.branch.textReplacementParts", collectReShowDebugState());
        applyChunkToActiveMarker();
        return;
    }

    if (replacementState?.pdfVisualParts?.length) {
        debugNovel("reShowChunk.branch.pdfVisualParts", collectReShowDebugState());
        applyChunkToActiveMarker();
        return;
    }

    if (replacementState?.marker) {
        debugNovel("reShowChunk.branch.marker", collectReShowDebugState());
        applyChunkToActiveMarker();
        return;
    }

    // If marker is missing, try to use the current selection again
    const range = getReaderSelectionRange(replacementState?.reader);
    const canReuseByPosition = Boolean(
        replacementState?.selectionPosition
        || replacementState?.selectionAnnotation?.position,
    );
    if (!range && !replacementState?.selectionRect && !canReuseByPosition) {
        debugNovel("reShowChunk.abort.noRangeNoRectNoPosition", collectReShowDebugState());
        showProgress("未找到可显示的位置，请先重新选中一段文本。", "warning");
        return;
    }

    debugNovel("reShowChunk.branch.rebuildState", {
        ...collectReShowDebugState(),
        rebuiltRange: describeRange(range),
    });
    replacementState = {
        reader: replacementState?.reader,
        originalText: range?.toString() || replacementState?.originalText || "",
        selectedLength: Math.max((range?.toString() || replacementState?.originalText || "").length, DEFAULT_CHUNK_MIN),
        selectionRect: replacementState?.selectionRect,
        selectionPosition: replacementState?.selectionPosition,
        selectionAnnotation: replacementState?.selectionAnnotation,
        initialRange: range || undefined,
        activeChunkLength: undefined,
    };

    if (!applyChunkToCurrentSelection()) {
        debugNovel("reShowChunk.rebuildState.applyFailed", collectReShowDebugState());
        showProgress("未找到可显示的位置，请先重新选中一段文本。", "warning");
        return;
    }

    debugNovel("reShowChunk.rebuildState.applySuccess", collectReShowDebugState());
}

function ensureNovelReady() {
    if (!novelState || !replacementState) {
        showProgress("请先在阅读器选中文本并点击“看小说”加载 EPUB。", "warning");
        return false;
    }
    return true;
}

function clearActiveReplacementBlock(reason: string, nextReader?: _ZoteroTypes.ReaderInstance) {
    if (!replacementState) {
        return;
    }

    persistCurrentNovelProgress(`clearActiveReplacementBlock:${reason}`, true);

    debugNovel("replacement.clear.begin", {
        reason,
        currentReaderKey: getReaderSnapshotKey(replacementState.reader),
        nextReaderKey: getReaderSnapshotKey(nextReader),
        hasMarker: Boolean(replacementState.marker),
        markerMode: replacementState.marker?.dataset?.novelMode || "none",
        pdfVisualContainerCount: replacementState.pdfVisualContainers?.length || 0,
        textReplacementPartCount: replacementState.textReplacementParts?.length || 0,
    });

    replacementState.pdfVisualSyncCleanup?.();
    replacementState.pdfVisualSyncCleanup = undefined;

    if (replacementState.textReplacementParts?.length) {
        for (const part of replacementState.textReplacementParts) {
            part.node.nodeValue = part.originalNodeValue;
        }
        replacementState.textReplacementParts = undefined;
    }

    if (replacementState.pdfVisualContainers?.length) {
        for (const container of replacementState.pdfVisualContainers) {
            container.remove();
        }
        replacementState.pdfVisualContainers = undefined;
    }

    if (replacementState.marker) {
        if (replacementState.marker.dataset.novelMode === "overlay") {
            replacementState.marker.remove();
        } else {
            replacementState.marker.textContent = replacementState.originalText;
        }
        replacementState.marker = undefined;
    }

    replacementState.pdfVisualParts = undefined;
    replacementState.pdfVisualHidden = undefined;
    replacementState.activeChunkLength = undefined;
    replacementState = undefined;

    debugNovel("replacement.clear.done", {
        reason,
    });
}

function applyChunkToCurrentSelection() {
    if (!replacementState || !novelState) return false;

    if (replacementState.pdfVisualParts?.length) {
        replacementState.pdfVisualHidden = false;
        applyChunkToPdfVisualParts();
        return true;
    }

    if (replacementState.textReplacementParts?.length) {
        applyChunkToTextReplacementParts();
        return true;
    }

    const range =
        cloneNonCollapsedRange(replacementState.initialRange) ||
        getReaderSelectionRange(replacementState.reader);

    if (!range) {
        const position = replacementState.selectionPosition || replacementState.selectionAnnotation?.position;
        if (applyChunkByPdfPositionPreserveStyles(replacementState.reader, position)) {
            replacementState.initialRange = undefined;
            return true;
        }

        const pdfRange = getPdfSelectionRange(replacementState.reader);
        if (pdfRange && applyChunkByPreservingTextLayerStyles(pdfRange)) {
            replacementState.initialRange = undefined;
            return true;
        }
        return applyChunkToOverlay();
    }

    if (applyChunkByPreservingTextLayerStyles(range)) {
        replacementState.initialRange = undefined;
        return true;
    }

    const chunk = getChunkFromCursor();
    const ownerDoc =
        range.startContainer.ownerDocument
        || range.endContainer.ownerDocument
        || getPreferredReaderDocument(replacementState.reader);
    if (!ownerDoc) {
        return false;
    }
    const marker = ownerDoc.createElement("span");
    marker.className = NOVEL_MARKER_CLASS;
    marker.dataset.novelMode = "inline";
    marker.textContent = chunk;
    applyNovelTypographyToMarker(marker);

    range.deleteContents();
    range.insertNode(marker);

    replacementState.marker = marker;
    replacementState.initialRange = undefined;
    replacementState.activeChunkLength = chunk.length;

    const selection = marker.ownerDocument?.defaultView?.getSelection();
    selection?.removeAllRanges();
    return true;
}

function applyChunkToOverlay() {
    if (!replacementState || !novelState) {
        return false;
    }

    const selectionRect = getSelectionRectForReplacement(
        replacementState.reader,
        replacementState.selectionRect,
        replacementState.selectionPosition,
    );
    const doc = getPreferredReaderDocument(replacementState.reader);
    if (!selectionRect || !doc?.body) {
        return false;
    }

    const chunk = getChunkFromCursor();
    const marker = doc.createElement("div");
    marker.className = NOVEL_MARKER_CLASS;
    marker.dataset.novelMode = "overlay";
    marker.textContent = chunk;

    const width = Math.max(48, selectionRect[2] - selectionRect[0]);
    const height = Math.max(20, selectionRect[3] - selectionRect[1]);
    marker.style.position = "fixed";
    marker.style.left = `${selectionRect[0]}px`;
    marker.style.top = `${selectionRect[1]}px`;
    marker.style.width = `${width}px`;
    marker.style.minHeight = `${height}px`;
    marker.style.maxHeight = `${Math.max(height * 3, 120)}px`;
    marker.style.overflow = "hidden";
    marker.style.background = "rgba(255,255,255,0.98)";
    marker.style.color = "inherit";
    applyNovelTypographyToMarker(marker);
    marker.style.padding = "2px 4px";
    marker.style.borderRadius = "2px";
    marker.style.zIndex = "2147483000";
    marker.style.pointerEvents = "none";
    marker.style.whiteSpace = "normal";

    doc.body.appendChild(marker);

    replacementState.marker = marker;
    replacementState.initialRange = undefined;
    replacementState.selectionRect = selectionRect;
    replacementState.activeChunkLength = chunk.length;

    debugNovel("applyChunkToOverlay.hit", {
        rect: selectionRect,
        chunkLength: chunk.length,
    });

    return true;
}

function applyChunkToActiveMarker(silent = false) {
    if (replacementState?.pdfVisualParts?.length && novelState) {
        applyChunkToPdfVisualParts();
        if (!silent) {
            showProgress(`已切换小说片段（${novelState.cursor}/${novelState.englishText.length}）`, "default");
        }
        return;
    }

    if (replacementState?.textReplacementParts?.length && novelState) {
        applyChunkToTextReplacementParts();
        if (!silent) {
            showProgress(`已切换小说片段（${novelState.cursor}/${novelState.englishText.length}）`, "default");
        }
        return;
    }

    if (!replacementState?.marker || !novelState) {
        showProgress("未检测到已替换文本，按 R 可重新显示到当前选区。", "warning");
        return;
    }

    const chunk = getChunkFromCursor();
    applyNovelTypographyToMarker(replacementState.marker);
    replacementState.marker.textContent = chunk;
    if (!silent) {
        showProgress(`已切换小说片段（${novelState.cursor}/${novelState.englishText.length}）`, "default");
    }
}

function getChunkFromCursor() {
    const text = novelState!.englishText;
    if (!text) return "";

    const wantedLength = Math.max(
        replacementState?.selectedLength || DEFAULT_CHUNK_MIN,
        DEFAULT_CHUNK_MIN,
    );
    const start = clampCursor(novelState!.cursor, text.length);
    const end = Math.min(start + wantedLength, text.length);
    const chunk = text.slice(start, end).trim();

    if (chunk.length === 0 && start > 0) {
        // Fallback to the last available fragment near EOF
        const fallbackStart = Math.max(0, text.length - wantedLength);
        return text.slice(fallbackStart, text.length).trim();
    }
    return chunk;
}

function getCurrentChunkLength() {
    if (replacementState?.activeChunkLength && replacementState.activeChunkLength > 0) {
        return replacementState.activeChunkLength;
    }
    return Math.max(replacementState?.selectedLength || DEFAULT_CHUNK_MIN, DEFAULT_CHUNK_MIN);
}

function getReaderSelectionRange(reader?: _ZoteroTypes.ReaderInstance): Range | undefined {
    const selection = getReaderSelection(reader);
    if (!selection || selection.rangeCount === 0) {
        return undefined;
    }
    const range = selection.getRangeAt(0);
    return range.collapsed ? undefined : range;
}

function findRangeBySelectionAnnotation(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    annotation?: SelectionAnnotationLike,
): Range | undefined {
    return findRangeBySelectionPosition(reader, annotation?.position);
}

function findRangeBySelectionPosition(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    position?: SelectionPositionLike,
): Range | undefined {
    if (!reader || !position) {
        return undefined;
    }

    const internalReader = (reader as any)?._internalReader;
    const views = [
        internalReader?._primaryView,
        internalReader?._secondaryView,
    ].filter(Boolean) as Array<{ toDisplayedRange?: (selector: any) => Range | null }>;

    for (const view of views) {
        if (typeof view.toDisplayedRange !== "function") {
            continue;
        }
        try {
            const range = view.toDisplayedRange(position);
            const clone = cloneNonCollapsedRange(range || undefined);
            if (clone) {
                debugNovel("findRangeBySelectionPosition.hit", {
                    positionType: typeof position,
                });
                return clone;
            }
        } catch {
            // Ignore selector mismatch between view types
        }
    }

    debugNovel("findRangeBySelectionPosition.notFound", {
        positionType: typeof position,
    });
    return undefined;
}

function getPdfSelectionRange(reader?: _ZoteroTypes.ReaderInstance): Range | undefined {
    const view = getActivePdfView(reader);
    if (!view) {
        return undefined;
    }

    const selectionRanges = Array.isArray(view._selectionRanges) ? view._selectionRanges : [];
    if (selectionRanges.length === 0) {
        return undefined;
    }

    const win = view._iframeWindow as Window | undefined;
    if (!win?.document) {
        return undefined;
    }

    return buildRangeFromPdfSelectionRanges(win, selectionRanges);
}

function getActivePdfView(reader?: _ZoteroTypes.ReaderInstance): any {
    const internal = (reader as any)?._internalReader;
    const views = [
        internal?._lastView === "secondary" ? internal?._secondaryView : internal?._primaryView,
        internal?._primaryView,
        internal?._secondaryView,
    ].filter(Boolean);

    for (const view of views) {
        if (Array.isArray(view?._selectionRanges) && view?._iframeWindow?.document) {
            return view;
        }
    }

    return undefined;
}

function buildRangeFromPdfSelectionRanges(win: Window, selectionRanges: any[]): Range | undefined {
    const first = selectionRanges[0];
    const last = selectionRanges[selectionRanges.length - 1];
    if (!first || !last) {
        return undefined;
    }

    const startOffset = Math.min(first.anchorOffset, first.headOffset);
    const endOffset = Math.max(last.anchorOffset, last.headOffset);

    const startContainer = win.document.querySelector(
        `[data-page-number="${first.position.pageIndex + 1}"] .textLayer`,
    );
    const endContainer = win.document.querySelector(
        `[data-page-number="${last.position.pageIndex + 1}"] .textLayer`,
    );

    if (!startContainer || !endContainer) {
        return undefined;
    }

    const startPoint = getTextLayerNodeOffset(startContainer, startOffset);
    const endPoint = getTextLayerNodeOffset(endContainer, endOffset);
    if (!startPoint || !endPoint) {
        return undefined;
    }

    const range = win.document.createRange();
    range.setStart(startPoint.node, startPoint.offset);
    range.setEnd(endPoint.node, endPoint.offset);
    return range.collapsed ? undefined : range;
}

function getTextLayerNodeOffset(container: Element, offset: number): { node: Text; offset: number } | undefined {
    let textIndex = 0;
    const stack: Node[] = [container];

    while (stack.length) {
        const node = stack.pop()!;
        if (node.nodeType === TEXT_NODE_TYPE) {
            const textNode = node as Text;
            const value = textNode.nodeValue || "";
            let local = 0;
            for (let i = 0; i < value.length; i++) {
                if (value[i].trim()) {
                    if (textIndex === offset) {
                        return { node: textNode, offset: local };
                    }
                    textIndex += 1;
                }
                local += 1;
            }
            if (textIndex === offset) {
                return { node: textNode, offset: local };
            }
            continue;
        }

        if (node.nodeType === ELEMENT_NODE_TYPE) {
            for (let i = node.childNodes.length - 1; i >= 0; i--) {
                stack.push(node.childNodes[i]);
            }
        }
    }

    return undefined;
}

function applyChunkByPreservingTextLayerStyles(range: Range): boolean {
    if (!replacementState || !novelState) {
        return false;
    }

    const parts = collectTextReplacementParts(range);
    if (parts.length === 0) {
        return false;
    }

    replacementState.textReplacementParts = parts;
    replacementState.pdfVisualSyncCleanup?.();
    replacementState.pdfVisualSyncCleanup = undefined;
    replacementState.pdfVisualParts = undefined;
    replacementState.pdfVisualContainers = undefined;
    replacementState.marker = undefined;
    applyChunkToTextReplacementParts();
    return true;
}

function applyChunkByPdfPositionPreserveStyles(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    position?: SelectionPositionLike,
): boolean {
    const visualParts = collectPdfVisualPartsByPosition(reader, position);
    if (visualParts.length === 0 || !replacementState) {
        return false;
    }

    const containerByPage = new Map<number, HTMLElement>();
    const overlayContainers: HTMLElement[] = [];

    const overlayParts: Array<{
        element: HTMLElement;
        pageIndex: number;
        expectedLength: number;
        maxWidth: number;
        softMaxWidth: number;
        font: string;
        letterSpacingPx: number;
        leftRatio: number;
        topRatio: number;
        widthRatio: number;
        heightRatio: number;
        basePageWidth: number;
        basePageHeight: number;
        baseFontSizePx?: number;
        baseLineHeightPx?: number;
        baseLetterSpacingPx?: number;
    }> = [];
    const dominantStyle = getDominantVisualStyle(visualParts);

    replacementState.marker?.remove();
    if (replacementState.pdfVisualContainers?.length) {
        for (const old of replacementState.pdfVisualContainers) {
            old.remove();
        }
    }

    for (const part of visualParts) {
        const styleSource = dominantStyle || part;
        let container = containerByPage.get(part.pageIndex);
        if (!container) {
            container = part.doc.createElement("div");
            container.className = NOVEL_MARKER_CLASS;
            container.dataset.novelMode = "pdf-visual-overlay";
            container.style.position = "absolute";
            container.style.left = "0";
            container.style.top = "0";
            container.style.width = "100%";
            container.style.height = "100%";
            container.style.pointerEvents = "none";
            container.style.zIndex = "50";

            const pageStyle = part.doc.defaultView?.getComputedStyle(part.pageElement);
            if (pageStyle?.position === "static") {
                part.pageElement.style.position = "relative";
            }

            part.pageElement.appendChild(container);
            containerByPage.set(part.pageIndex, container);
            overlayContainers.push(container);
        }

        const element = part.doc.createElement("div");
        element.style.position = "absolute";
        element.style.left = `${part.rect.left}px`;
        const baseRowHeight = Math.max(1, part.rect.height);
        const verticalPaddingPx = Math.min(
            NOVEL_PDF_ROW_VERTICAL_PADDING_PX,
            baseRowHeight * 0.2,
        );
        const rowTop = Math.max(0, part.rect.top - verticalPaddingPx);
        const rowHeight = baseRowHeight + verticalPaddingPx * 2;
        element.style.top = `${rowTop}px`;
        const viewportWidth = part.doc.documentElement?.clientWidth || part.doc.defaultView?.innerWidth || 0;
        const rowWidth = Math.max(1, part.rect.width);
        const remainingRightSpace = Math.max(0, viewportWidth - (part.rect.left + rowWidth) - 12);
        const overflowAllowance = Math.min(4, remainingRightSpace);
        const renderWidth = rowWidth + overflowAllowance;
        const pageBounds = part.pageElement.getBoundingClientRect();
        const pageWidth = Math.max(1, pageBounds.width);
        const pageHeight = Math.max(1, pageBounds.height);
        element.style.width = `${Math.max(1, renderWidth)}px`;
        element.style.height = `${Math.max(1, rowHeight)}px`;
        element.style.overflow = "hidden";
        element.style.whiteSpace = "nowrap";
        element.style.background = part.backgroundColor;
        element.style.color = styleSource.color;
        const latexBodyFontSizePx = Math.min(
            getNovelFontSizePx(),
            rowHeight * NOVEL_PDF_ROW_FONT_TO_BOX_RATIO,
        );
        const latexBodyLineHeightPx = Math.min(
            getNovelLineHeightPx(),
            rowHeight * NOVEL_PDF_ROW_LINE_HEIGHT_TO_BOX_RATIO,
        );
        const effectiveFontSize = `${latexBodyFontSizePx.toFixed(2)}px`;
        const effectiveLineHeight = `${latexBodyLineHeightPx.toFixed(2)}px`;
        const fontFamily = NOVEL_LATEX_FONT_STACK;
        element.style.font = `normal 400 ${effectiveFontSize} ${fontFamily}`;
        element.style.fontFamily = fontFamily;
        element.style.fontSize = effectiveFontSize;
        element.style.fontWeight = "400";
        element.style.fontStyle = "normal";
        element.style.lineHeight = effectiveLineHeight;
        element.style.letterSpacing = "0px";
        element.style.wordSpacing = "normal";
        element.style.textTransform = "none";
        element.style.fontStretch = "normal";
        element.style.fontKerning = "normal";
        element.style.fontVariantLigatures = "common-ligatures";
        element.style.textAlign = "left";
        element.style.padding = "0";
        element.style.margin = "0";
        element.style.textRendering = "geometricPrecision";
        container.appendChild(element);

        const baseFontSizePx = Number.parseFloat(element.style.fontSize);
        const baseLineHeightPx = Number.parseFloat(element.style.lineHeight);
        const baseLetterSpacingPx = Number.parseFloat(element.style.letterSpacing);
        overlayParts.push({
            element,
            pageIndex: part.pageIndex,
            expectedLength: part.expectedLength,
            maxWidth: rowWidth,
            softMaxWidth: renderWidth,
            font: element.style.font,
            letterSpacingPx: parseLetterSpacingPx(element.style.letterSpacing),
            leftRatio: part.rect.left / pageWidth,
            topRatio: rowTop / pageHeight,
            widthRatio: rowWidth / pageWidth,
            heightRatio: Math.max(1, rowHeight) / pageHeight,
            basePageWidth: pageWidth,
            basePageHeight: pageHeight,
            baseFontSizePx: Number.isFinite(baseFontSizePx) ? baseFontSizePx : undefined,
            baseLineHeightPx: Number.isFinite(baseLineHeightPx) ? baseLineHeightPx : undefined,
            baseLetterSpacingPx: Number.isFinite(baseLetterSpacingPx) ? baseLetterSpacingPx : undefined,
        });
    }

    replacementState.marker = overlayContainers[0];
    replacementState.pdfVisualContainers = overlayContainers;
    replacementState.pdfVisualParts = overlayParts;
    replacementState.pdfVisualHidden = false;
    replacementState.textReplacementParts = undefined;
    replacementState.pdfVisualSyncCleanup?.();
    replacementState.pdfVisualSyncCleanup = registerPdfVisualSync(reader);
    refreshPdfVisualOverlayGeometry(reader);
    applyChunkToPdfVisualParts();
    debugNovel("applyChunkByPdfPositionPreserveStyles.hit", {
        partCount: overlayParts.length,
        pageIndex: position?.pageIndex,
    });
    return true;
}

function getDominantVisualStyle(parts: Array<{
    expectedLength: number;
    color: string;
    fontFamily: string;
    fontSize: string;
    font: string;
    fontWeight: string;
    fontStyle: string;
    lineHeight: string;
    letterSpacing: string;
    textTransform: string;
    fontStretch: string;
    preferredFontName?: string;
    preferredFontSizePx?: number;
}>) {
    if (!parts.length) {
        return undefined;
    }

    let dominant = parts[0];
    let maxWeight = Math.max(1, parts[0].expectedLength || 1);
    for (const part of parts) {
        const weight = Math.max(1, part.expectedLength || 1);
        if (weight > maxWeight) {
            maxWeight = weight;
            dominant = part;
        }
    }
    return dominant;
}

function getOrCreatePdfPageOverlayContainer(pageElement: HTMLElement, doc: Document): HTMLElement {
    let container = pageElement.querySelector(`.${NOVEL_MARKER_CLASS}[data-novel-mode="pdf-visual-overlay"]`) as HTMLElement | null;
    if (container) {
        return container;
    }

    container = doc.createElement("div");
    container.className = NOVEL_MARKER_CLASS;
    container.dataset.novelMode = "pdf-visual-overlay";
    container.style.position = "absolute";
    container.style.left = "0";
    container.style.top = "0";
    container.style.width = "100%";
    container.style.height = "100%";
    container.style.pointerEvents = "none";
    container.style.zIndex = "50";

    const pageStyle = doc.defaultView?.getComputedStyle(pageElement);
    if (pageStyle?.position === "static") {
        pageElement.style.position = "relative";
    }

    pageElement.appendChild(container);
    return container;
}

function refreshPdfVisualOverlayGeometry(reader: _ZoteroTypes.ReaderInstance | undefined) {
    if (!replacementState?.pdfVisualParts?.length) {
        return;
    }

    const view = getActivePdfView(reader || replacementState.reader);
    const pdfViewer = view?._iframeWindow?.PDFViewerApplication?.pdfViewer;
    if (!pdfViewer) {
        return;
    }

    const overlayContainerSet = new Set<HTMLElement>();
    for (const part of replacementState.pdfVisualParts) {
        const pageView = pdfViewer.getPageView?.(part.pageIndex);
        const pageElement = pageView?.div as HTMLElement | undefined;
        if (!pageElement) {
            continue;
        }

        const doc = pageElement.ownerDocument;
        if (!doc) {
            continue;
        }
        const container = getOrCreatePdfPageOverlayContainer(pageElement, doc);
        overlayContainerSet.add(container);

        if (part.element.parentElement !== container) {
            container.appendChild(part.element);
        }

        const pageBounds = pageElement.getBoundingClientRect();
        const pageWidth = Math.max(1, pageBounds.width);
        const pageHeight = Math.max(1, pageBounds.height);

        const left = part.leftRatio * pageWidth;
        const top = part.topRatio * pageHeight;
        const width = Math.max(1, part.widthRatio * pageWidth);
        const height = Math.max(1, part.heightRatio * pageHeight);
        const remainingRightSpace = Math.max(0, pageWidth - (left + width) - 12);
        const overflowAllowance = Math.min(4, remainingRightSpace);

        part.element.style.left = `${left}px`;
        part.element.style.top = `${top}px`;
        part.element.style.width = `${width + overflowAllowance}px`;
        part.element.style.height = `${height}px`;

        part.maxWidth = width;
        part.softMaxWidth = width + overflowAllowance;

        const scaleY = pageHeight / Math.max(1, part.basePageHeight);
        const scaleX = pageWidth / Math.max(1, part.basePageWidth);

        if (part.baseFontSizePx && Number.isFinite(part.baseFontSizePx)) {
            part.element.style.fontSize = `${Math.max(8, part.baseFontSizePx * scaleY).toFixed(2)}px`;
        }
        if (part.baseLineHeightPx && Number.isFinite(part.baseLineHeightPx)) {
            part.element.style.lineHeight = `${Math.max(8, part.baseLineHeightPx * scaleY).toFixed(2)}px`;
        }
        if (part.baseLetterSpacingPx && Number.isFinite(part.baseLetterSpacingPx)) {
            const letterSpacing = part.baseLetterSpacingPx * scaleX;
            part.element.style.letterSpacing = `${letterSpacing.toFixed(2)}px`;
            part.letterSpacingPx = letterSpacing;
        }
    }

    replacementState.pdfVisualContainers = Array.from(overlayContainerSet);
    replacementState.marker = replacementState.pdfVisualContainers[0];
}

function registerPdfVisualSync(reader: _ZoteroTypes.ReaderInstance | undefined): (() => void) | undefined {
    const view = getActivePdfView(reader || replacementState?.reader);
    const iframeWindow = view?._iframeWindow as Window | undefined;
    const eventBus = iframeWindow?.PDFViewerApplication?.eventBus;
    if (!iframeWindow || !eventBus?.on || !eventBus?.off) {
        return undefined;
    }

    let rafId = 0;
    const schedule = () => {
        if (rafId) {
            iframeWindow.cancelAnimationFrame(rafId);
        }
        rafId = iframeWindow.requestAnimationFrame(() => {
            rafId = 0;
            refreshPdfVisualOverlayGeometry(reader || replacementState?.reader);
        });
    };

    eventBus.on("scalechanging", schedule);
    eventBus.on("pagerendered", schedule);
    eventBus.on("rotationchanging", schedule);
    iframeWindow.addEventListener("resize", schedule);

    return () => {
        if (rafId) {
            iframeWindow.cancelAnimationFrame(rafId);
            rafId = 0;
        }
        eventBus.off("scalechanging", schedule);
        eventBus.off("pagerendered", schedule);
        eventBus.off("rotationchanging", schedule);
        iframeWindow.removeEventListener("resize", schedule);
    };
}

function applyChunkToPdfVisualParts() {
    if (!replacementState?.pdfVisualParts || !novelState) {
        return;
    }

    const start = clampCursor(novelState.cursor, novelState.englishText.length);
    const fallbackLength = Math.max(getCurrentChunkLength(), DEFAULT_CHUNK_MIN);
    const sourceWindowLength = Math.min(
        novelState.englishText.length - start,
        Math.max(fallbackLength * 4, 800),
    );
    const sourceWindow = novelState.englishText.slice(start, start + sourceWindowLength);
    const layout = layoutPdfChunkByRows(sourceWindow, replacementState.pdfVisualParts);

    for (let i = 0; i < replacementState.pdfVisualParts.length; i++) {
        const part = replacementState.pdfVisualParts[i];
        const line = layout.lines[i] || "";
        part.element.textContent = line;
        const spacing = layout.wordSpacings[i];
        if (typeof spacing === "number") {
            part.element.style.wordSpacing = `${spacing.toFixed(3)}px`;
        } else {
            applyFullJustification(part.element, line, part.maxWidth);
        }
    }

    replacementState.activeChunkLength = layout.consumedChars > 0
        ? layout.consumedChars
        : Math.min(fallbackLength, sourceWindow.length);
}

function layoutPdfChunkByRows(
    sourceWindow: string,
    parts: Array<{
        element: HTMLElement;
        expectedLength: number;
        maxWidth: number;
        softMaxWidth: number;
        font: string;
        letterSpacingPx: number;
    }>,
): { lines: string[]; consumedChars: number; wordSpacings: number[] } {
    if (parts.length === 0) {
        return { lines: [], consumedChars: 0, wordSpacings: [] };
    }

    const source = sourceWindow.replace(/\s+/g, " ").trimStart();
    if (!source) {
        return { lines: parts.map(() => ""), consumedChars: 0, wordSpacings: parts.map(() => 0) };
    }

    const spans = getWordSpans(source);
    const words = spans.map((span) => source.slice(span.start, span.end));
    const lines = parts.map(() => "");
    const wordSpacings = parts.map(() => 0);

    const doc = parts[0].element.ownerDocument;
    if (!doc) {
        return { lines: parts.map(() => ""), consumedChars: 0, wordSpacings: parts.map(() => 0) };
    }
    type TexOption = {
        line: string;
        nextWordIndex: number;
        nextWordOffset: number;
        consumedChars: number;
        hyphenated: boolean;
        wordSpacingPx: number;
        lineCost: number;
        fitnessClass: number;
    };

    type TexResult = {
        lines: string[];
        wordSpacings: number[];
        consumedChars: number;
        score: number;
        lastHyphenated: boolean;
        lastFitnessClass: number;
    };

    const emPx = getNovelFontSizePx();
    const stretchPerGap = emPx * NOVEL_TEX_STRETCH_EM;
    const shrinkPerGap = emPx * NOVEL_TEX_SHRINK_EM;

    const consumedFromCursor = (wordIndex: number, wordOffset: number): number => {
        if (wordIndex <= 0 && wordOffset <= 0) {
            return 0;
        }
        if (wordIndex >= spans.length) {
            return source.length;
        }
        if (wordOffset > 0) {
            return spans[wordIndex].start + wordOffset;
        }
        return spans[wordIndex - 1].end;
    };

    const buildOptions = (
        row: number,
        startWordIndex: number,
        startWordOffset: number,
    ): TexOption[] => {
        if (startWordIndex >= words.length) {
            return [];
        }

        const part = parts[row];
        const widthLimit = part.maxWidth;
        const options: TexOption[] = [];

        let line = "";
        let wi = startWordIndex;
        let offset = startWordOffset;
        let bestFull: TexOption | undefined;

        while (wi < words.length) {
            const fullWord = words[wi];
            const word = fullWord.slice(offset);
            const candidate = line ? `${line} ${word}` : word;
            if (textFitsPartWidth(part.element, candidate, widthLimit)) {
                line = candidate;
                const spacing = computeJustifySpacing(part.element, line, widthLimit);
                const diff = computeLineWidthDiff(part.element, line, widthLimit, spacing);
                const lineCost = computeTexLineCost(diff, line, false, stretchPerGap, shrinkPerGap);
                const fitnessClass = getTexFitnessClass(diff, line, stretchPerGap, shrinkPerGap);
                bestFull = {
                    line,
                    nextWordIndex: wi + 1,
                    nextWordOffset: 0,
                    consumedChars: spans[wi].end,
                    hyphenated: false,
                    wordSpacingPx: spacing,
                    lineCost,
                    fitnessClass,
                };
                wi += 1;
                offset = 0;
                continue;
            }

            const splitCandidates = findHyphenSplitCandidates(part.element, line, word, widthLimit);
            for (const splitAt of splitCandidates) {
                const hyphenated = `${word.slice(0, splitAt)}-`;
                const hyphenLine = line ? `${line} ${hyphenated}` : hyphenated;
                const spacing = computeJustifySpacing(part.element, hyphenLine, widthLimit);
                const diff = computeLineWidthDiff(part.element, hyphenLine, widthLimit, spacing);
                const lineCost = computeTexLineCost(diff, hyphenLine, true, stretchPerGap, shrinkPerGap) + NOVEL_TEX_HYPHEN_PENALTY;
                const fitnessClass = getTexFitnessClass(diff, hyphenLine, stretchPerGap, shrinkPerGap);
                options.push({
                    line: hyphenLine,
                    nextWordIndex: wi,
                    nextWordOffset: offset + splitAt,
                    consumedChars: spans[wi].start + offset + splitAt,
                    hyphenated: true,
                    wordSpacingPx: spacing,
                    lineCost,
                    fitnessClass,
                });
            }
            break;
        }

        if (bestFull) {
            options.push(bestFull);
        }

        if (options.length === 0) {
            const fallbackWord = words[startWordIndex].slice(startWordOffset);
            const spacing = computeJustifySpacing(part.element, fallbackWord, widthLimit);
            const diff = computeLineWidthDiff(part.element, fallbackWord, widthLimit, spacing);
            const fitnessClass = getTexFitnessClass(diff, fallbackWord, stretchPerGap, shrinkPerGap);
            options.push({
                line: fallbackWord,
                nextWordIndex: startWordIndex + 1,
                nextWordOffset: 0,
                consumedChars: spans[startWordIndex].end,
                hyphenated: false,
                wordSpacingPx: spacing,
                lineCost: computeTexLineCost(diff, fallbackWord, false, stretchPerGap, shrinkPerGap),
                fitnessClass,
            });
        }

        return options;
    };

    const memo = new Map<string, TexResult | undefined>();
    const solve = (
        row: number,
        wordIndex: number,
        wordOffset: number,
        prevHyphenated: boolean,
        prevFitnessClass: number,
    ): TexResult => {
        if (row >= parts.length || wordIndex >= words.length) {
            return {
                lines: [],
                wordSpacings: [],
                consumedChars: consumedFromCursor(wordIndex, wordOffset),
                score: 0,
                lastHyphenated: prevHyphenated,
                lastFitnessClass: prevFitnessClass,
            };
        }

        const key = `${row}|${wordIndex}|${wordOffset}|${prevHyphenated ? 1 : 0}|${prevFitnessClass}`;
        const cached = memo.get(key);
        if (cached) {
            return cached;
        }

        const options = buildOptions(row, wordIndex, wordOffset);
        let best: TexResult | undefined;

        for (const option of options) {
            const tail = solve(
                row + 1,
                option.nextWordIndex,
                option.nextWordOffset,
                option.hyphenated,
                option.fitnessClass,
            );
            let score = option.lineCost + tail.score;
            if (prevHyphenated && option.hyphenated) {
                score += NOVEL_TEX_CONSECUTIVE_HYPHEN_PENALTY;
            }
            if (prevFitnessClass >= 0 && Math.abs(prevFitnessClass - option.fitnessClass) > 1) {
                score += NOVEL_TEX_FITNESS_DEMERIT;
            }

            const candidate: TexResult = {
                lines: [option.line, ...tail.lines],
                wordSpacings: [option.wordSpacingPx, ...tail.wordSpacings],
                consumedChars: tail.consumedChars,
                score,
                lastHyphenated: tail.lastHyphenated,
                lastFitnessClass: tail.lastFitnessClass,
            };

            if (!best) {
                best = candidate;
                continue;
            }

            if (candidate.consumedChars > best.consumedChars) {
                best = candidate;
                continue;
            }

            if (candidate.consumedChars === best.consumedChars && candidate.score < best.score) {
                best = candidate;
            }
        }

        const resolved = best || {
            lines: [],
            wordSpacings: [],
            consumedChars: consumedFromCursor(wordIndex, wordOffset),
            score: Number.POSITIVE_INFINITY,
            lastHyphenated: prevHyphenated,
            lastFitnessClass: prevFitnessClass,
        };
        memo.set(key, resolved);
        return resolved;
    };

    const best = solve(0, 0, 0, false, -1);
    for (let i = 0; i < parts.length; i++) {
        lines[i] = best.lines[i] || "";
        wordSpacings[i] = Number.isFinite(best.wordSpacings[i]) ? best.wordSpacings[i] : 0;
    }

    return {
        lines,
        consumedChars: Math.max(0, best.consumedChars),
        wordSpacings,
    };
}

function isHyphenatableWord(word: string): boolean {
    return /^[A-Za-z]{6,}$/.test(word);
}

function findHyphenSplitCandidates(
    partElement: HTMLElement,
    currentLine: string,
    word: string,
    widthLimit: number,
): number[] {
    if (!isHyphenatableWord(word)) {
        return [];
    }

    const minHead = 3;
    const minTail = 2;
    const maxHead = word.length - minTail;
    if (maxHead < minHead) {
        return [];
    }

    const result: Array<{ split: number; score: number }> = [];

    for (let headLen = maxHead; headLen >= minHead; headLen--) {
        const hyphenated = `${word.slice(0, headLen)}-`;
        const candidate = currentLine ? `${currentLine} ${hyphenated}` : hyphenated;
        if (textFitsPartWidth(partElement, candidate, widthLimit)) {
            result.push({ split: headLen, score: scoreHyphenSplit(word, headLen) });
        }
    }

    return result
        .sort((a, b) => b.score - a.score)
        .slice(0, 5)
        .map((item) => item.split);
}

function scoreHyphenSplit(word: string, headLen: number): number {
    const head = word.slice(0, headLen).toLowerCase();
    const tail = word.slice(headLen).toLowerCase();
    const vowels = /[aeiouy]/;
    const headHasVowel = vowels.test(head);
    const tailHasVowel = vowels.test(tail);
    const boundary = `${head.slice(-1)}${tail[0] || ""}`;
    const likelyGoodBoundary = /[aeiouy][^aeiouy]|[^aeiouy][aeiouy]/.test(boundary);
    const balanced = 1 - Math.abs(head.length - tail.length) / Math.max(word.length, 1);

    let score = 0;
    if (headHasVowel) score += 1.5;
    if (tailHasVowel) score += 1.5;
    if (likelyGoodBoundary) score += 2;
    score += balanced;
    return score;
}

function computeTexLineCost(
    widthDiff: number,
    line: string,
    hyphenated: boolean,
    stretchPerGap: number,
    shrinkPerGap: number,
): number {
    const gaps = countWordGaps(line);
    if (gaps <= 0) {
        const ratio = Math.abs(widthDiff) / Math.max(1, measurePlainTextLength(line));
        let cost = 100 + 400 * ratio * ratio;
        if (hyphenated) {
            cost += NOVEL_TEX_HYPHEN_PENALTY;
        }
        return cost;
    }

    const capacity = widthDiff >= 0
        ? gaps * Math.max(0.01, stretchPerGap)
        : gaps * Math.max(0.01, shrinkPerGap);

    const ratio = Math.abs(widthDiff) / capacity;
    if (ratio > 2.5) {
        return NOVEL_TEX_INFINITE_BADNESS;
    }
    let badness = 100 * Math.pow(ratio, 3);
    if (ratio > 1) {
        badness += 1500 * Math.pow(ratio - 1, 2);
    }
    if (hyphenated) {
        badness += NOVEL_TEX_HYPHEN_PENALTY;
    }
    return badness;
}

function getTexFitnessClass(
    widthDiff: number,
    line: string,
    stretchPerGap: number,
    shrinkPerGap: number,
): number {
    const gaps = Math.max(1, countWordGaps(line));
    const capacity = widthDiff >= 0
        ? gaps * Math.max(0.01, stretchPerGap)
        : gaps * Math.max(0.01, shrinkPerGap);
    const ratio = widthDiff / capacity;

    if (ratio < -0.5) return 0; // tight
    if (ratio < 0.5) return 1;  // decent
    if (ratio < 1.0) return 2;  // loose
    return 3;                   // very loose
}

function computeLineWidthDiff(
    partElement: HTMLElement,
    line: string,
    widthLimit: number,
    wordSpacingPx: number,
): number {
    const measured = measurePartTextWidth(partElement, line, `${wordSpacingPx.toFixed(3)}px`);
    return widthLimit - measured;
}

function measurePlainTextLength(text: string): number {
    return Math.max(1, text.replace(/\s+/g, "").length);
}

function computeJustifySpacing(partElement: HTMLElement, line: string, maxWidth: number): number {
    const gaps = countWordGaps(line);
    if (gaps <= 0) {
        return 0;
    }

    const baseWidth = measurePartTextWidth(partElement, line, "normal");
    const remaining = maxWidth - baseWidth;
    if (!Number.isFinite(remaining)) {
        return 0;
    }

    const emPx = getNovelFontSizePx();
    const maxStretch = emPx * NOVEL_TEX_STRETCH_EM;
    const maxShrink = emPx * NOVEL_TEX_SHRINK_EM;
    const raw = remaining / gaps - NOVEL_JUSTIFY_TIGHTEN_PX;
    return Math.max(-maxShrink, Math.min(maxStretch, raw));
}

function getWordSpans(text: string): Array<{ start: number; end: number }> {
    const spans: Array<{ start: number; end: number }> = [];
    const regex = /\S+/g;
    let match = regex.exec(text);
    while (match) {
        spans.push({
            start: match.index,
            end: match.index + match[0].length,
        });
        match = regex.exec(text);
    }
    return spans;
}

function advanceCursorByDisplayed(text: string, start: number, consumedChars: number): number {
    if (!text) {
        return 0;
    }

    const safeStart = clampCursor(start, text.length);
    const step = Math.max(0, consumedChars);
    if (step <= 0) {
        return safeStart;
    }

    let next = Math.min(safeStart + step, text.length);
    while (next < text.length && /\s/.test(text[next])) {
        next += 1;
    }

    if (next >= text.length) {
        return Math.max(text.length - 1, 0);
    }
    return next;
}

function measureTextWidth(ctx: CanvasRenderingContext2D, text: string, letterSpacingPx: number): number {
    const base = ctx.measureText(text).width;
    if (!letterSpacingPx || text.length <= 1) {
        return base;
    }
    return base + (text.length - 1) * letterSpacingPx;
}

function getOrCreateNovelTextMeasurer(doc: Document): HTMLElement {
    const existing = novelTextMeasurerByDoc.get(doc);
    if (existing && existing.isConnected) {
        return existing;
    }

    const node = doc.createElement("div");
    node.className = `${NOVEL_MARKER_CLASS}-measurer`;
    node.style.position = "fixed";
    node.style.left = "-100000px";
    node.style.top = "-100000px";
    node.style.visibility = "hidden";
    node.style.pointerEvents = "none";
    node.style.whiteSpace = "nowrap";
    const host = doc.body || doc.documentElement;
    if (host) {
        host.appendChild(node);
    }
    novelTextMeasurerByDoc.set(doc, node);
    return node;
}

function textFitsPartWidth(partElement: HTMLElement, text: string, maxWidth: number): boolean {
    const width = measurePartTextWidth(partElement, text, "normal");
    return width <= Math.max(1, maxWidth - 0.5);
}

function countWordGaps(text: string): number {
    const matches = text.match(/\s+/g);
    if (!matches) {
        return 0;
    }
    return matches.reduce((sum, token) => sum + (token.length > 0 ? 1 : 0), 0);
}

function applyFullJustification(partElement: HTMLElement, line: string, maxWidth: number) {
    if (!line.trim()) {
        partElement.style.wordSpacing = "normal";
        return;
    }

    const gaps = countWordGaps(line);
    if (gaps <= 0) {
        partElement.style.wordSpacing = "normal";
        return;
    }

    const baseWidth = measurePartTextWidth(partElement, line, "normal");
    const remaining = maxWidth - baseWidth;
    if (remaining <= 0.25) {
        partElement.style.wordSpacing = "normal";
        return;
    }

    const spacing = Math.max(0, remaining / gaps - NOVEL_JUSTIFY_TIGHTEN_PX);
    partElement.style.wordSpacing = `${spacing.toFixed(3)}px`;
}

function measurePartTextWidth(partElement: HTMLElement, text: string, wordSpacing: string): number {
    const doc = partElement.ownerDocument;
    if (!doc) {
        return 0;
    }
    const measurer = getOrCreateNovelTextMeasurer(doc);

    measurer.style.font = partElement.style.font;
    measurer.style.fontFamily = partElement.style.fontFamily;
    measurer.style.fontSize = partElement.style.fontSize;
    measurer.style.fontWeight = partElement.style.fontWeight;
    measurer.style.fontStyle = partElement.style.fontStyle;
    measurer.style.fontStretch = partElement.style.fontStretch;
    measurer.style.letterSpacing = partElement.style.letterSpacing;
    measurer.style.wordSpacing = wordSpacing;
    measurer.style.textTransform = partElement.style.textTransform;
    measurer.style.fontKerning = partElement.style.fontKerning;
    measurer.style.fontVariantLigatures = partElement.style.fontVariantLigatures;
    measurer.textContent = text;

    return measurer.getBoundingClientRect().width;
}

function parseLetterSpacingPx(value: string): number {
    const parsed = Number.parseFloat(value);
    return Number.isFinite(parsed) ? parsed : 0;
}

function ptToPx(pt: number): number {
    return pt * (96 / 72);
}

function normalizeFontSize(fontSize: string, rectHeight: number): string {
    const parsed = Number.parseFloat(fontSize);
    if (!Number.isFinite(parsed) || parsed <= 0) {
        return `${Math.max(10, Math.round(rectHeight * 0.72))}px`;
    }

    return `${parsed}px`;
}

function normalizeLineHeight(lineHeight: string, normalizedFontSize: string): string {
    const lineParsed = Number.parseFloat(lineHeight);
    const fontParsed = Number.parseFloat(normalizedFontSize);

    if (!Number.isFinite(lineParsed) || lineParsed <= 0 || !Number.isFinite(fontParsed) || fontParsed <= 0) {
        return "normal";
    }

    const ratio = lineParsed / fontParsed;
    if (ratio < 0.9 || ratio > 2.2) {
        return "normal";
    }

    return `${lineParsed}px`;
}

function isTransparentColor(color: string): boolean {
    const value = color.replace(/\s+/g, "").toLowerCase();
    return value === "transparent" || value === "rgba(0,0,0,0)" || value === "hsla(0,0%,0%,0)";
}

function collectPdfVisualPartsByPosition(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    position?: SelectionPositionLike,
): Array<{
    doc: Document;
    pageElement: HTMLElement;
    pageIndex: number;
    rect: { left: number; top: number; width: number; height: number };
    expectedLength: number;
    backgroundColor: string;
    color: string;
    fontFamily: string;
    fontSize: string;
    font: string;
    fontWeight: string;
    fontStyle: string;
    lineHeight: string;
    letterSpacing: string;
    textTransform: string;
    fontStretch: string;
    transform: string;
    transformOrigin: string;
    preferredFontName?: string;
    preferredFontSizePx?: number;
}> {
    const view = getActivePdfView(reader);
    if (!view || !position || typeof position.pageIndex !== "number" || !Array.isArray(position.rects)) {
        return [];
    }

    const pageRects: Array<{ pageIndex: number; rect: number[] }> = [];
    for (const rect of position.rects) {
        if (Array.isArray(rect) && rect.length >= 4) {
            pageRects.push({ pageIndex: position.pageIndex, rect });
        }
    }
    if (Array.isArray(position.nextPageRects)) {
        for (const rect of position.nextPageRects) {
            if (Array.isArray(rect) && rect.length >= 4) {
                pageRects.push({ pageIndex: position.pageIndex + 1, rect });
            }
        }
    }

    const entries = new Map<HTMLElement, {
        doc: Document;
        pageElement: HTMLElement;
        pageIndex: number;
        rect: { left: number; top: number; width: number; height: number };
        length: number;
        backgroundColor: string;
        color: string;
        fontFamily: string;
        fontSize: string;
        font: string;
        fontWeight: string;
        fontStyle: string;
        lineHeight: string;
        letterSpacing: string;
        textTransform: string;
        fontStretch: string;
        transform: string;
        transformOrigin: string;
    }>();

    const pdfViewer = view._iframeWindow?.PDFViewerApplication?.pdfViewer;

    for (const pageRect of pageRects) {
        const pageView = pdfViewer?.getPageView?.(pageRect.pageIndex);
        const textLayer = pageView?.div?.querySelector?.(".textLayer") as HTMLElement | null;
        if (!textLayer || typeof view.getClientRect !== "function") {
            continue;
        }

        let clientRect: number[] | undefined;
        try {
            clientRect = view.getClientRect(pageRect.rect, pageRect.pageIndex);
        } catch {
            continue;
        }
        if (!clientRect || clientRect.length < 4) {
            continue;
        }

        const [left, top, right, bottom] = clientRect;
        const layerDoc = textLayer.ownerDocument;
        if (!layerDoc) {
            continue;
        }

        const pageBounds = (pageView.div as HTMLElement).getBoundingClientRect();
        const pageStyle = layerDoc.defaultView?.getComputedStyle(pageView.div as Element);
        const backgroundColor = pageStyle?.backgroundColor || "rgba(255,255,255,0.98)";
        const pageColor = pageStyle?.color || "#111";

        const nodeWalker = layerDoc.createTreeWalker(textLayer, 4);
        let node = nodeWalker.nextNode();
        while (node) {
            const textNode = node as Text;
            const value = textNode.nodeValue || "";
            if (value.trim()) {
                const parent = textNode.parentElement as HTMLElement | null;
                if (parent) {
                    const rect = parent.getBoundingClientRect();
                    const intersects = rect.right >= left && rect.left <= right && rect.bottom >= top && rect.top <= bottom;
                    if (intersects) {
                        const style = layerDoc.defaultView?.getComputedStyle(parent);
                        const rawColor = style?.color || pageColor;
                        const color = isTransparentColor(rawColor) ? pageColor : rawColor;
                        const existing = entries.get(parent);
                        if (existing) {
                            existing.length += value.length;
                        } else {
                            entries.set(parent, {
                                doc: layerDoc,
                                pageElement: pageView.div as HTMLElement,
                                pageIndex: pageRect.pageIndex,
                                rect: {
                                    left: rect.left - pageBounds.left,
                                    top: rect.top - pageBounds.top,
                                    width: rect.width,
                                    height: rect.height,
                                },
                                length: value.length,
                                backgroundColor,
                                color,
                                fontFamily: style?.fontFamily || "inherit",
                                fontSize: style?.fontSize || "inherit",
                                font: style?.font || "",
                                fontWeight: style?.fontWeight || "normal",
                                fontStyle: style?.fontStyle || "normal",
                                lineHeight: style?.lineHeight || "normal",
                                letterSpacing: style?.letterSpacing || "normal",
                                textTransform: style?.textTransform || "none",
                                fontStretch: style?.fontStretch || "normal",
                                transform: style?.transform || "none",
                                transformOrigin: style?.transformOrigin || "left top",
                            });
                        }
                    }
                }
            }
            node = nodeWalker.nextNode();
        }
    }

    const sorted = Array.from(entries.values())
        .filter((entry) => entry.length > 0)
        .sort((a, b) => {
            if (Math.abs(a.rect.top - b.rect.top) > 1) {
                return a.rect.top - b.rect.top;
            }
            return a.rect.left - b.rect.left;
        });

    const charStyleHint = collectPdfSelectionCharStyleHint(view, position, pageRects);

    const rows: Array<{
        doc: Document;
        pageElement: HTMLElement;
        pageIndex: number;
        rect: { left: number; top: number; width: number; height: number };
        expectedLength: number;
        backgroundColor: string;
        color: string;
        fontFamily: string;
        fontSize: string;
        font: string;
        fontWeight: string;
        fontStyle: string;
        lineHeight: string;
        letterSpacing: string;
        textTransform: string;
        fontStretch: string;
        transform: string;
        transformOrigin: string;
        preferredFontName?: string;
        preferredFontSizePx?: number;
        baselineTop: number;
        styleWeight: number;
    }> = [];

    for (const entry of sorted) {
        const last = rows[rows.length - 1];
        const sameDoc = Boolean(last && last.doc === entry.doc);
        const samePage = Boolean(last && last.pageIndex === entry.pageIndex);
        const sameRow = sameDoc
            && samePage
            && Math.abs(last.baselineTop - entry.rect.top) <= Math.max(2, entry.rect.height * 0.35);

        if (!sameRow || !last) {
            rows.push({
                doc: entry.doc,
                pageElement: entry.pageElement,
                pageIndex: entry.pageIndex,
                rect: { ...entry.rect },
                expectedLength: entry.length,
                backgroundColor: entry.backgroundColor,
                color: entry.color,
                fontFamily: entry.fontFamily,
                fontSize: entry.fontSize,
                font: entry.font,
                fontWeight: entry.fontWeight,
                fontStyle: entry.fontStyle,
                lineHeight: entry.lineHeight,
                letterSpacing: entry.letterSpacing,
                textTransform: entry.textTransform,
                fontStretch: entry.fontStretch,
                transform: entry.transform,
                transformOrigin: entry.transformOrigin,
                preferredFontName: charStyleHint.fontName,
                preferredFontSizePx: charStyleHint.fontSizePx,
                baselineTop: entry.rect.top,
                styleWeight: entry.length,
            });
            continue;
        }

        const mergedLeft = Math.min(last.rect.left, entry.rect.left);
        const mergedTop = Math.min(last.rect.top, entry.rect.top);
        const mergedRight = Math.max(last.rect.left + last.rect.width, entry.rect.left + entry.rect.width);
        const mergedBottom = Math.max(last.rect.top + last.rect.height, entry.rect.top + entry.rect.height);

        last.rect.left = mergedLeft;
        last.rect.top = mergedTop;
        last.rect.width = mergedRight - mergedLeft;
        last.rect.height = mergedBottom - mergedTop;
        last.expectedLength += entry.length;
        last.baselineTop = (last.baselineTop * last.styleWeight + entry.rect.top * entry.length) / (last.styleWeight + entry.length);

        if (entry.length >= last.styleWeight) {
            last.backgroundColor = entry.backgroundColor;
            last.color = entry.color;
            last.fontFamily = entry.fontFamily;
            last.fontSize = entry.fontSize;
            last.font = entry.font;
            last.fontWeight = entry.fontWeight;
            last.fontStyle = entry.fontStyle;
            last.lineHeight = entry.lineHeight;
            last.letterSpacing = entry.letterSpacing;
            last.textTransform = entry.textTransform;
            last.fontStretch = entry.fontStretch;
            last.transform = entry.transform;
            last.transformOrigin = entry.transformOrigin;
            last.styleWeight = entry.length;
        } else {
            last.styleWeight += entry.length;
        }

        if (charStyleHint.fontName) {
            last.preferredFontName = charStyleHint.fontName;
        }
        if (charStyleHint.fontSizePx && Number.isFinite(charStyleHint.fontSizePx)) {
            last.preferredFontSizePx = charStyleHint.fontSizePx;
        }
    }

    return rows.map(({ baselineTop: _baselineTop, styleWeight: _styleWeight, ...row }) => row);
}

function collectPdfSelectionCharStyleHint(
    view: any,
    position: SelectionPositionLike | undefined,
    pageRects: Array<{ pageIndex: number; rect: number[] }>,
): { fontName?: string; fontSizePx?: number } {
    const pdfPages = view?._pdfPages;
    if (!pdfPages) {
        return {};
    }

    const fontCounts = new Map<string, number>();
    const fontHeights: number[] = [];

    const consumeChar = (char: any, pageIndex: number) => {
        if (!char || char.ignorable) {
            return;
        }

        if (typeof char.fontName === "string" && char.fontName.trim()) {
            const key = char.fontName.trim();
            fontCounts.set(key, (fontCounts.get(key) || 0) + 1);
        }

        if (Array.isArray(char.rect) && typeof view.getClientRect === "function") {
            try {
                const clientRect = view.getClientRect(char.rect, pageIndex);
                if (Array.isArray(clientRect) && clientRect.length >= 4) {
                    const h = Math.abs(clientRect[3] - clientRect[1]);
                    if (Number.isFinite(h) && h > 0) {
                        fontHeights.push(h);
                    }
                }
            } catch {
                // Ignore invalid page/rect combinations
            }
        }
    };

    const selectionRanges = Array.isArray(view?._selectionRanges) ? view._selectionRanges : [];
    if (selectionRanges.length > 0) {
        for (const range of selectionRanges) {
            const pageIndex = range?.position?.pageIndex;
            if (typeof pageIndex !== "number") {
                continue;
            }
            const page = pdfPages[pageIndex];
            const chars = page?.chars;
            if (!Array.isArray(chars) || chars.length === 0) {
                continue;
            }

            const anchorOffset = Number(range?.anchorOffset);
            const headOffset = Number(range?.headOffset);
            if (!Number.isFinite(anchorOffset) || !Number.isFinite(headOffset)) {
                continue;
            }

            const from = Math.max(0, Math.min(anchorOffset, headOffset));
            const to = Math.min(chars.length, Math.max(anchorOffset, headOffset));
            for (let i = from; i < to; i++) {
                consumeChar(chars[i], pageIndex);
            }
        }
    }

    if (fontCounts.size === 0 && position && pageRects.length > 0) {
        for (const pageRect of pageRects) {
            const page = pdfPages[pageRect.pageIndex];
            const chars = page?.chars;
            if (!Array.isArray(chars) || chars.length === 0) {
                continue;
            }

            for (const char of chars) {
                if (!Array.isArray(char?.rect) || char.rect.length < 4) {
                    continue;
                }
                const intersects = !(
                    char.rect[2] < pageRect.rect[0]
                    || char.rect[0] > pageRect.rect[2]
                    || char.rect[3] < pageRect.rect[1]
                    || char.rect[1] > pageRect.rect[3]
                );
                if (intersects) {
                    consumeChar(char, pageRect.pageIndex);
                }
            }
        }
    }

    let fontName: string | undefined;
    let maxCount = 0;
    for (const [name, count] of fontCounts) {
        if (count > maxCount) {
            maxCount = count;
            fontName = name;
        }
    }

    let fontSizePx: number | undefined;
    if (fontHeights.length > 0) {
        const sorted = fontHeights.slice().sort((a, b) => a - b);
        const median = sorted[Math.floor(sorted.length / 2)];
        fontSizePx = median * 0.86;
    }

    return { fontName, fontSizePx };
}

function collectPdfTextReplacementPartsByPosition(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    position?: SelectionPositionLike,
): Array<{
    node: Text;
    startOffset: number;
    endOffset: number;
    originalNodeValue: string;
}> {
    const view = getActivePdfView(reader);
    if (!view || !position || typeof position.pageIndex !== "number" || !Array.isArray(position.rects)) {
        return [];
    }

    const pageRects: Array<{ pageIndex: number; rect: number[] }> = [];
    for (const rect of position.rects) {
        if (Array.isArray(rect) && rect.length >= 4) {
            pageRects.push({ pageIndex: position.pageIndex, rect });
        }
    }
    if (Array.isArray(position.nextPageRects)) {
        for (const rect of position.nextPageRects) {
            if (Array.isArray(rect) && rect.length >= 4) {
                pageRects.push({ pageIndex: position.pageIndex + 1, rect });
            }
        }
    }
    if (pageRects.length === 0) {
        return [];
    }

    const textNodes = new Set<Text>();
    const pdfViewer = view._iframeWindow?.PDFViewerApplication?.pdfViewer;

    for (const pageRect of pageRects) {
        const pageView = pdfViewer?.getPageView?.(pageRect.pageIndex);
        const textLayer = pageView?.div?.querySelector?.(".textLayer") as HTMLElement | null;
        if (!textLayer || typeof view.getClientRect !== "function") {
            continue;
        }

        let clientRect: number[] | undefined;
        try {
            clientRect = view.getClientRect(pageRect.rect, pageRect.pageIndex);
        } catch {
            continue;
        }
        if (!clientRect || clientRect.length < 4) {
            continue;
        }

        const [left, top, right, bottom] = clientRect;
        const layerDoc = textLayer.ownerDocument;
        if (!layerDoc) {
            continue;
        }
        const nodeWalker = layerDoc.createTreeWalker(textLayer, 4);
        let node = nodeWalker.nextNode();
        while (node) {
            const textNode = node as Text;
            const value = textNode.nodeValue || "";
            if (value.trim()) {
                const parent = textNode.parentElement as HTMLElement | null;
                if (parent) {
                    const rect = parent.getBoundingClientRect();
                    const intersects = rect.right >= left && rect.left <= right && rect.bottom >= top && rect.top <= bottom;
                    if (intersects) {
                        textNodes.add(textNode);
                    }
                }
            }
            node = nodeWalker.nextNode();
        }
    }

    const parts = Array.from(textNodes)
        .map((node) => ({
            node,
            startOffset: 0,
            endOffset: (node.nodeValue || "").length,
            originalNodeValue: node.nodeValue || "",
        }))
        .filter((part) => part.endOffset > part.startOffset)
        .sort((a, b) => {
            const ra = a.node.parentElement?.getBoundingClientRect();
            const rb = b.node.parentElement?.getBoundingClientRect();
            if (!ra || !rb) return 0;
            if (Math.abs(ra.top - rb.top) > 1) {
                return ra.top - rb.top;
            }
            return ra.left - rb.left;
        });

    return parts;
}

function collectTextReplacementParts(range: Range): Array<{
    node: Text;
    startOffset: number;
    endOffset: number;
    originalNodeValue: string;
}> {
    const doc = range.startContainer.ownerDocument || range.endContainer.ownerDocument;
    if (!doc) {
        return [];
    }
    const walker = doc.createTreeWalker(range.commonAncestorContainer, 4);

    const parts: Array<{
        node: Text;
        startOffset: number;
        endOffset: number;
        originalNodeValue: string;
    }> = [];

    let current = walker.nextNode();
    while (current) {
        const textNode = current as Text;
        const value = textNode.nodeValue || "";
        if (!value) {
            current = walker.nextNode();
            continue;
        }

        if (range.intersectsNode(textNode)) {
            const startOffset = textNode === range.startContainer ? range.startOffset : 0;
            const endOffset = textNode === range.endContainer ? range.endOffset : value.length;
            if (endOffset > startOffset) {
                parts.push({
                    node: textNode,
                    startOffset,
                    endOffset,
                    originalNodeValue: value,
                });
            }
        }
        current = walker.nextNode();
    }

    return parts;
}

function applyChunkToTextReplacementParts() {
    if (!replacementState?.textReplacementParts || !novelState) {
        return;
    }

    const chunk = getChunkFromCursor();
    const totalLength = replacementState.textReplacementParts.reduce(
        (sum, part) => sum + (part.endOffset - part.startOffset),
        0,
    );

    const text = chunk.padEnd(totalLength, " ").slice(0, totalLength);
    let cursor = 0;

    for (const part of replacementState.textReplacementParts) {
        const partLength = part.endOffset - part.startOffset;
        const replacement = text.slice(cursor, cursor + partLength);
        cursor += partLength;

        const currentValue = part.node.nodeValue || "";
        const prefix = currentValue.slice(0, part.startOffset);
        const suffix = currentValue.slice(part.endOffset);
        part.node.nodeValue = `${prefix}${replacement}${suffix}`;
    }
}

function getSelectionRectForReplacement(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    fallbackRect?: SelectionRect,
    selectionPosition?: SelectionPositionLike,
): SelectionRect | undefined {
    if (fallbackRect) {
        return fallbackRect;
    }

    const fromPosition = getSelectionRectByPosition(reader, selectionPosition);
    if (fromPosition) {
        return fromPosition;
    }

    const popup = getReaderSelectionPopup(reader);
    const rect = popup?.rect;
    if (
        Array.isArray(rect) &&
        rect.length === 4 &&
        rect.every((value) => typeof value === "number" && Number.isFinite(value))
    ) {
        return rect as SelectionRect;
    }

    return undefined;
}

function getReaderSelectionPopup(reader?: _ZoteroTypes.ReaderInstance): any {
    const internal = (reader as any)?._internalReader;
    const primaryPopup = internal?._state?.primaryViewSelectionPopup;
    const secondaryPopup = internal?._state?.secondaryViewSelectionPopup;
    const lastPopup = internal?._lastView === "secondary" ? secondaryPopup : primaryPopup;
    const candidates = [
        lastPopup,
        primaryPopup,
        secondaryPopup,
        internal?._primaryView?._selectionPopup,
        internal?._secondaryView?._selectionPopup,
    ];

    for (const popup of candidates) {
        if (popup?.rect) {
            return popup;
        }
    }

    return undefined;
}

function getSelectionRectByPosition(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    position?: SelectionPositionLike,
): SelectionRect | undefined {
    if (!reader || !position) {
        return undefined;
    }

    const internal = (reader as any)?._internalReader;
    const views = [internal?._primaryView, internal?._secondaryView].filter(Boolean);

    for (const view of views) {
        if (typeof (view as any).getClientRectForPopup === "function") {
            try {
                const rect = (view as any).getClientRectForPopup(position);
                if (
                    Array.isArray(rect)
                    && rect.length === 4
                    && rect.every((value: unknown) => typeof value === "number" && Number.isFinite(value))
                ) {
                    return rect as SelectionRect;
                }
            } catch {
                // Ignore view/position mismatch
            }
        }

        if (typeof (view as any).toDisplayedRange === "function") {
            try {
                const range = (view as any).toDisplayedRange(position);
                if (!range || range.collapsed) {
                    continue;
                }

                const innerRect = range.getBoundingClientRect();
                const iframeRect = (view as any)._iframe?.getBoundingClientRect?.();
                if (!iframeRect) {
                    continue;
                }

                const rect: SelectionRect = [
                    iframeRect.x + innerRect.left,
                    iframeRect.y + innerRect.top,
                    iframeRect.x + innerRect.right,
                    iframeRect.y + innerRect.bottom,
                ];
                if (rect.every((value) => Number.isFinite(value))) {
                    return rect;
                }
            } catch {
                // Ignore selector mismatch
            }
        }
    }

    return undefined;
}

function getPreferredReaderDocument(
    reader?: _ZoteroTypes.ReaderInstance,
): Document | undefined {
    const docs = getReaderDocuments(reader);
    return docs[0];
}

function getRememberedSelectionSnapshot(
    reader?: _ZoteroTypes.ReaderInstance,
): ReaderSelectionSnapshot | undefined {
    if (!reader) {
        return undefined;
    }

    return readerSelectionSnapshots.get(getReaderSnapshotKey(reader));
}

function cloneNonCollapsedRange(range?: Range): Range | undefined {
    if (!range || range.collapsed) {
        return undefined;
    }
    return range.cloneRange();
}

function getReaderSelection(reader?: _ZoteroTypes.ReaderInstance): Selection | undefined {
    for (const win of getReaderWindows(reader)) {
        const selection = win?.getSelection?.();
        if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
            debugNovel("getReaderSelection.hit", {
                href: safeWindowHref(win),
                rangeCount: selection.rangeCount,
            });
            return selection;
        }
    }
    debugNovel("getReaderSelection.miss", {
        candidateWindowCount: getReaderWindows(reader).length,
    });
    return undefined;
}

function findRangeBySelectedText(
    reader: _ZoteroTypes.ReaderInstance | undefined,
    selectedText: string,
): Range | undefined {
    const rawSelectedText = selectedText.trim();
    if (!reader || !rawSelectedText) {
        return undefined;
    }

    const docs = getReaderDocuments(reader)
        .map((doc) => ({ doc, textLength: (doc.body?.innerText || doc.body?.textContent || "").length }))
        .sort((a, b) => b.textLength - a.textLength);

    if (docs.length === 0) {
        debugNovel("findRangeBySelectedText.noBody", {
            selectedLength: rawSelectedText.length,
        });
        return undefined;
    }

    for (const { doc, textLength } of docs) {
        const found = findRangeBySelectedTextInDocument(doc, rawSelectedText);
        if (found) {
            debugNovel("findRangeBySelectedText.hitInDoc", {
                selectedLength: rawSelectedText.length,
                docLength: textLength,
            });
            return found;
        }
    }

    debugNovel("findRangeBySelectedText.notFound", {
        selectedLength: rawSelectedText.length,
        testedDocs: docs.length,
        largestDocLength: docs[0]?.textLength || 0,
    });
    return undefined;
}

function findRangeBySelectedTextInDocument(doc: Document, rawSelectedText: string): Range | undefined {
    const body = doc.body;
    if (!body) {
        return undefined;
    }

    const nodes: Array<{ node: Text; start: number; end: number }> = [];
    let fullText = "";
    collectTextNodesDeep(body, nodes, (text) => {
        const start = fullText.length;
        fullText += text;
        return start;
    });

    let startIndex = fullText.indexOf(rawSelectedText);
    let targetLength = rawSelectedText.length;
    if (startIndex < 0) {
        const normalizedDoc = normalizeWhitespaceWithMap(fullText);
        const normalizedSelected = normalizeWhitespace(rawSelectedText);
        if (!normalizedSelected) {
            return undefined;
        }

        const normalizedStart = normalizedDoc.text.indexOf(normalizedSelected);
        if (normalizedStart < 0) {
            return undefined;
        }

        const mappedStart = normalizedDoc.map[normalizedStart];
        const normalizedEndIndex = normalizedStart + normalizedSelected.length - 1;
        const mappedEnd = normalizedDoc.map[normalizedEndIndex];
        if (mappedStart == null || mappedEnd == null) {
            return undefined;
        }

        startIndex = mappedStart;
        targetLength = mappedEnd - mappedStart + 1;

        // If whitespace-normalized hit is still unavailable, fallback to
        // alphanumeric-only matching to tolerate punctuation/symbol differences
        // (e.g., section sign, smart quotes, hyphenation artifacts in PDF text layers).
    }

    if (startIndex < 0) {
        const simplifiedDoc = simplifyTextWithMap(fullText);
        const simplifiedSelected = simplifyText(rawSelectedText);
        if (!simplifiedSelected) {
            return undefined;
        }

        const simpleStart = simplifiedDoc.text.indexOf(simplifiedSelected);
        if (simpleStart < 0) {
            return undefined;
        }

        const mappedStart = simplifiedDoc.map[simpleStart];
        const simpleEndIndex = simpleStart + simplifiedSelected.length - 1;
        const mappedEnd = simplifiedDoc.map[simpleEndIndex];
        if (mappedStart == null || mappedEnd == null) {
            return undefined;
        }

        startIndex = mappedStart;
        targetLength = mappedEnd - mappedStart + 1;
    }

    if (startIndex < 0) {
        const anchorResult = findRangeByAnchorsInText(fullText, nodes, rawSelectedText);
        if (anchorResult) {
            return anchorResult;
        }
    }

    if (startIndex < 0) {
        return undefined;
    }

    const endIndex = startIndex + targetLength;
    const startHit = nodes.find((entry) => entry.start <= startIndex && startIndex < entry.end);
    const endHit = nodes.find((entry) => entry.start < endIndex && endIndex <= entry.end);
    if (!startHit || !endHit) {
        return undefined;
    }

    const range = doc.createRange();
    range.setStart(startHit.node, startIndex - startHit.start);
    range.setEnd(endHit.node, endIndex - endHit.start);
    return range.collapsed ? undefined : range;
}

function collectTextNodesDeep(
    root: ParentNode | ShadowRoot,
    nodes: Array<{ node: Text; start: number; end: number }>,
    onText: (text: string) => number,
) {
    const rootNode = root as Node;
    for (const child of Array.from(rootNode.childNodes)) {
        if (!child) {
            continue;
        }
        if (child.nodeType === TEXT_NODE_TYPE) {
            const textNode = child as Text;
            const nodeText = textNode.nodeValue || "";
            if (nodeText) {
                const start = onText(nodeText);
                nodes.push({ node: textNode, start, end: start + nodeText.length });
            }
            continue;
        }

        if (child.nodeType === ELEMENT_NODE_TYPE) {
            const element = child as Element;
            if (element.shadowRoot) {
                collectTextNodesDeep(element.shadowRoot, nodes, onText);
            }
            collectTextNodesDeep(element, nodes, onText);
        }
    }
}

function findRangeByAnchorsInText(
    fullText: string,
    nodes: Array<{ node: Text; start: number; end: number }>,
    rawSelectedText: string,
): Range | undefined {
    const anchorLength = Math.min(80, Math.max(20, Math.floor(rawSelectedText.length / 6)));
    const prefix = rawSelectedText.slice(0, anchorLength);
    const suffix = rawSelectedText.slice(Math.max(0, rawSelectedText.length - anchorLength));

    const prefixCandidates = buildAnchorCandidates(prefix);
    const suffixCandidates = buildAnchorCandidates(suffix);

    let startIndex = findFirstAnchorIndex(fullText, prefixCandidates);
    let endIndex = findLastAnchorIndex(fullText, suffixCandidates);

    if (startIndex < 0 || endIndex < 0) {
        const simplifiedDoc = simplifyTextWithMap(fullText);
        const simplifiedPrefixCandidates = prefixCandidates
            .map((candidate) => simplifyText(candidate))
            .filter(Boolean);
        const simplifiedSuffixCandidates = suffixCandidates
            .map((candidate) => simplifyText(candidate))
            .filter(Boolean);

        const simplifiedStart = findFirstAnchorIndex(simplifiedDoc.text, simplifiedPrefixCandidates);
        const simplifiedEnd = findLastAnchorIndex(simplifiedDoc.text, simplifiedSuffixCandidates);
        if (simplifiedStart >= 0 && simplifiedEnd >= simplifiedStart) {
            startIndex = simplifiedDoc.map[simplifiedStart] ?? -1;
            const mappedEnd = simplifiedDoc.map[simplifiedEnd + simplifiedSuffixCandidates[0].length - 1] ?? -1;
            if (startIndex >= 0 && mappedEnd >= startIndex) {
                endIndex = mappedEnd + 1;
            }
        }
    }

    if (startIndex < 0 || endIndex < 0 || endIndex <= startIndex) {
        debugNovel("findRangeByAnchorsInText.notFound", {
            selectedLength: rawSelectedText.length,
            prefixLength: prefix.length,
            suffixLength: suffix.length,
        });
        return undefined;
    }

    const startHit = nodes.find((entry) => entry.start <= startIndex && startIndex < entry.end);
    const endHit = nodes.find((entry) => entry.start < endIndex && endIndex <= entry.end);
    if (!startHit || !endHit) {
        return undefined;
    }

    const ownerDoc = startHit.node.ownerDocument;
    if (!ownerDoc) {
        return undefined;
    }
    const range = ownerDoc.createRange();
    range.setStart(startHit.node, startIndex - startHit.start);
    range.setEnd(endHit.node, endIndex - endHit.start);
    debugNovel("findRangeByAnchorsInText.hit", {
        selectedLength: rawSelectedText.length,
        startIndex,
        endIndex,
        prefixLength: prefix.length,
        suffixLength: suffix.length,
    });
    return range.collapsed ? undefined : range;
}

function buildAnchorCandidates(fragment: string): string[] {
    const normalized = fragment.replace(/\s+/g, " ").trim();
    if (!normalized) {
        return [];
    }

    const candidates = new Set<string>();
    candidates.add(normalized);
    candidates.add(normalizeWhitespace(normalized));
    candidates.add(simplifyText(normalized));
    candidates.add(normalized.replace(/[\u2010-\u2015-\u2212-]/g, "-"));
    candidates.add(normalized.replace(/[\u2018\u2019\u201A\u201B]/g, "'").replace(/[\u201C\u201D\u201E\u201F]/g, '"'));
    return Array.from(candidates).filter(Boolean);
}

function findFirstAnchorIndex(text: string, candidates: string[]): number {
    for (const candidate of candidates) {
        if (!candidate) continue;
        const index = text.indexOf(candidate);
        if (index >= 0) {
            return index;
        }
    }
    return -1;
}

function findLastAnchorIndex(text: string, candidates: string[]): number {
    for (const candidate of candidates) {
        if (!candidate) continue;
        const index = text.lastIndexOf(candidate);
        if (index >= 0) {
            return index + candidate.length;
        }
    }
    return -1;
}

function getReaderDocuments(reader?: _ZoteroTypes.ReaderInstance): Document[] {
    const docs: Document[] = [];
    const seen = new Set<Document>();
    for (const win of getReaderWindows(reader)) {
        const doc = win?.document;
        if (!doc || seen.has(doc) || !doc.body) {
            continue;
        }
        seen.add(doc);
        docs.push(doc);
    }
    return docs;
}

function normalizeWhitespace(input: string): string {
    return input
        .replace(/[\u00A0\u2000-\u200B\u202F\u205F\u3000]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function normalizeWhitespaceWithMap(input: string): { text: string; map: number[] } {
    const outChars: string[] = [];
    const outMap: number[] = [];
    let previousWasSpace = true;

    for (let i = 0; i < input.length; i++) {
        const ch = input[i];
        const isSpace = /\s|\u00A0|\u2000|\u2001|\u2002|\u2003|\u2004|\u2005|\u2006|\u2007|\u2008|\u2009|\u200A|\u200B|\u202F|\u205F|\u3000/.test(ch);
        if (isSpace) {
            if (!previousWasSpace) {
                outChars.push(" ");
                outMap.push(i);
                previousWasSpace = true;
            }
            continue;
        }

        outChars.push(ch);
        outMap.push(i);
        previousWasSpace = false;
    }

    if (outChars.length > 0 && outChars[outChars.length - 1] === " ") {
        outChars.pop();
        outMap.pop();
    }

    return {
        text: outChars.join(""),
        map: outMap,
    };
}

function simplifyText(input: string): string {
    return input
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, "")
        .trim();
}

function simplifyTextWithMap(input: string): { text: string; map: number[] } {
    const outChars: string[] = [];
    const outMap: number[] = [];
    const lower = input.toLowerCase();

    for (let i = 0; i < lower.length; i++) {
        const ch = lower[i];
        if (/[a-z0-9]/.test(ch)) {
            outChars.push(ch);
            outMap.push(i);
        }
    }

    return {
        text: outChars.join(""),
        map: outMap,
    };
}

function getReaderWindows(reader?: _ZoteroTypes.ReaderInstance): Array<Window | undefined> {
    const seeds: Array<Window | undefined> = [
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

    const result: Array<Window | undefined> = [];
    const visited = new Set<Window>();
    const queue: Window[] = [];

    for (const seed of seeds) {
        if (seed && !visited.has(seed)) {
            visited.add(seed);
            queue.push(seed);
            result.push(seed);
        }
    }

    while (queue.length > 0) {
        const win = queue.shift()!;
        const frameWins = getChildFrameWindows(win);
        for (const child of frameWins) {
            if (!visited.has(child)) {
                visited.add(child);
                queue.push(child);
                result.push(child);
            }
        }
    }

    return result;
}

function getChildFrameWindows(win: Window): Window[] {
    const children: Window[] = [];
    try {
        const frameElements = win.document?.querySelectorAll("iframe, frame") || [];
        frameElements.forEach((el: Element) => {
            const frameWin = (el as HTMLIFrameElement).contentWindow as Window | null;
            if (frameWin) {
                children.push(frameWin);
            }
        });
    } catch {
        // ignore cross-origin and dead windows
    }
    return children;
}

function describeRange(range?: Range): string {
    if (!range) {
        return "none";
    }
    const text = range.toString().replace(/\s+/g, " ").trim();
    const preview = text.slice(0, 60);
    return `${range.startContainer?.nodeName || "?"}:${range.startOffset}-${range.endContainer?.nodeName || "?"}:${range.endOffset}|len=${text.length}|preview=${preview}`;
}

function getReaderSnapshotKey(reader?: _ZoteroTypes.ReaderInstance): string {
    return String((reader as any)?.tabID || (reader as any)?._instanceID || "global");
}

function safeWindowHref(win?: Window): string {
    try {
        return win?.location?.href || "unknown";
    } catch {
        return "unavailable";
    }
}

function debugNovel(stage: string, detail?: Record<string, unknown>) {
    const payload = detail || {};
    ztoolkit.log("novel-debug", stage, payload);
    try {
        const globalConsole = ztoolkit.getGlobal("console") as Console | undefined;
        (globalConsole || console).log("[novel-debug]", stage, payload);
    } catch {
        // Ignore environments where console is unavailable
    }
}

function collectReShowDebugState() {
    return {
        hasNovelState: Boolean(novelState),
        hasReplacementState: Boolean(replacementState),
        hasMarker: Boolean(replacementState?.marker),
        markerMode: replacementState?.marker?.dataset?.novelMode || "none",
        pdfVisualPartCount: replacementState?.pdfVisualParts?.length || 0,
        pdfVisualHidden: Boolean(replacementState?.pdfVisualHidden),
        textReplacementPartCount: replacementState?.textReplacementParts?.length || 0,
        hasSelectionRect: Boolean(replacementState?.selectionRect),
        hasSelectionPosition: Boolean(
            replacementState?.selectionPosition
            || replacementState?.selectionAnnotation?.position,
        ),
        hasInitialRange: Boolean(replacementState?.initialRange),
        initialRange: describeRange(replacementState?.initialRange),
    };
}

function isEditableTarget(target: EventTarget | null) {
    const element = target as HTMLElement | null;
    if (!element) return false;
    const tag = element.tagName?.toLowerCase();
    return (
        tag === "input" ||
        tag === "textarea" ||
        element.isContentEditable === true
    );
}

function getShortcutKey(prefKey: "shortcutForward" | "shortcutBackward" | "shortcutBoss" | "shortcutShow", fallback: string) {
    const value = String(getPref(prefKey) || fallback)
        .trim()
        .toLowerCase();
    return value.length === 1 ? value : fallback;
}

function clampCursor(value: number, totalLength: number) {
    if (Number.isNaN(value)) return 0;
    return Math.max(0, Math.min(value, Math.max(totalLength - 1, 0)));
}

function showProgress(
    text: string,
    type: "default" | "success" | "warning" | "error" = "default",
) {
    ztoolkit.log("novel-progress-muted", {
        addon: addon.data.config.addonName,
        text,
        type,
    });
}

export function getConfiguredBookStoragePath(): string {
    return String(getPref("bookStoragePath") || "").trim();
}

export function getStorageNovelBooks(): NovelStorageBooksResult {
    const storagePath = getConfiguredBookStoragePath();
    if (!storagePath) {
        return {
            configured: false,
            dirExists: false,
            books: [],
        };
    }

    const directory = makeLocalFile(storagePath);
    if (!directory || !directory.exists() || !directory.isDirectory()) {
        return {
            configured: true,
            dirExists: false,
            books: [],
        };
    }

    const rawBooks = listEpubFilesFromDirectory(directory.path);
    const cache = readNovelCache(directory.path);
    const books = rawBooks
        .map((book) => {
            const pathHash = cache.pathToHash[normalizeLocalPath(book.filePath)] || "";
            const record = pathHash ? cache.booksByHash[pathHash] : undefined;
            return {
                filePath: book.filePath,
                fileName: book.fileName,
                lastReadAt: Number(record?.lastReadAt || 0),
            };
        })
        .sort((a, b) => {
            if (a.lastReadAt !== b.lastReadAt) {
                return b.lastReadAt - a.lastReadAt;
            }
            return a.fileName.localeCompare(b.fileName, "en");
        });

    return {
        configured: true,
        dirExists: true,
        books,
    };
}

export function getSelectedNovelPath(): string {
    return String(getPref("selectedNovelPath") || "").trim();
}

export function setSelectedNovelPath(path: string) {
    setPref("selectedNovelPath", String(path || "").trim());
}

export function ensureSelectedNovelPath(): string {
    const selectedPath = getSelectedNovelPath();
    const storage = getStorageNovelBooks();

    if (selectedPath) {
        const matched = storage.books.find(
            (book) => normalizeLocalPath(book.filePath) === normalizeLocalPath(selectedPath),
        );
        if (matched) {
            return matched.filePath;
        }
    }

    const firstBook = storage.books[0];
    if (firstBook?.filePath) {
        setSelectedNovelPath(firstBook.filePath);
        return firstBook.filePath;
    }

    return "";
}

export function getSelectedNovelBook(): NovelStorageBook | undefined {
    const storage = getStorageNovelBooks();
    const selectedPath = getSelectedNovelPath();
    if (!selectedPath) {
        return storage.books[0];
    }

    return storage.books.find(
        (book) => normalizeLocalPath(book.filePath) === normalizeLocalPath(selectedPath),
    );
}

export function clearNovelCache(storagePath: string): boolean {
    const directory = makeLocalFile(storagePath);
    if (!directory || !directory.exists() || !directory.isDirectory()) {
        return false;
    }

    try {
        const cacheFile = directory.clone() as any;
        cacheFile.append(NOVEL_CACHE_FILE_NAME);
        if (cacheFile.exists()) {
            cacheFile.remove(false);
        }
        return true;
    } catch (error) {
        ztoolkit.log("novel.cache.clear.error", error);
        return false;
    }
}

function persistCurrentNovelProgress(reason: string, immediate = false) {
    if (immediate) {
        if (novelCachePersistTimer != null) {
            clearTimeout(novelCachePersistTimer);
            novelCachePersistTimer = undefined;
        }
        persistCurrentNovelProgressNow(reason);
        return;
    }

    novelCachePersistReason = reason;
    if (novelCachePersistTimer != null) {
        return;
    }
    novelCachePersistTimer = setTimeout(() => {
        novelCachePersistTimer = undefined;
        persistCurrentNovelProgressNow(novelCachePersistReason || "scheduled");
        novelCachePersistReason = "";
    }, NOVEL_CACHE_WRITE_DEBOUNCE_MS) as unknown as number;
}

function persistCurrentNovelProgressNow(reason: string) {
    if (!novelState?.sourcePath || !novelState?.sourceHash) {
        return;
    }

    try {
        const storagePath = getConfiguredBookStoragePath();
        if (!storagePath) {
            return;
        }
        const directory = makeLocalFile(storagePath);
        if (!directory || !directory.exists() || !directory.isDirectory()) {
            return;
        }

        const cache = readNovelCache(directory.path);
        const normalizedPath = normalizeLocalPath(novelState.sourcePath);
        const now = Date.now();
        const nextRecord: NovelCacheBookRecord = {
            hash: novelState.sourceHash,
            filePath: novelState.sourcePath,
            fileName: getLeafNameFromPath(novelState.sourcePath),
            cursor: clampCursor(novelState.cursor, novelState.englishText.length),
            lastReadAt: now,
        };

        cache.booksByHash[novelState.sourceHash] = nextRecord;
        cache.pathToHash[normalizedPath] = novelState.sourceHash;
        writeNovelCache(directory.path, cache);

        debugNovel("cache.persist", {
            reason,
            sourcePath: novelState.sourcePath,
            sourceHash: novelState.sourceHash,
            cursor: nextRecord.cursor,
        });
    } catch (error) {
        ztoolkit.log("novel.cache.persist.error", error);
    }
}

function getCachedCursorForBook(epubPath: string, hash: string): number {
    const storagePath = getConfiguredBookStoragePath();
    if (!storagePath) {
        return 0;
    }

    const directory = makeLocalFile(storagePath);
    if (!directory || !directory.exists() || !directory.isDirectory()) {
        return 0;
    }

    const cache = readNovelCache(directory.path);
    const normalizedPath = normalizeLocalPath(epubPath);
    const hashFromPath = cache.pathToHash[normalizedPath] || "";
    const candidate = cache.booksByHash[hash] || (hashFromPath ? cache.booksByHash[hashFromPath] : undefined);

    if (!candidate) {
        return 0;
    }

    if (candidate.hash !== hash) {
        return 0;
    }

    return Number.isFinite(candidate.cursor) ? Math.max(0, Math.floor(candidate.cursor)) : 0;
}

function readNovelCache(storageDirPath: string): NovelCacheData {
    const empty: NovelCacheData = {
        version: 1,
        booksByHash: {},
        pathToHash: {},
    };

    const cacheFile = makeLocalFile(storageDirPath);
    if (!cacheFile || !cacheFile.exists() || !cacheFile.isDirectory()) {
        return empty;
    }
    cacheFile.append(NOVEL_CACHE_FILE_NAME);

    if (!cacheFile.exists() || !cacheFile.isFile()) {
        return empty;
    }

    try {
        const content = readLocalTextFile(cacheFile.path);
        if (!content.trim()) {
            return empty;
        }
        const parsed = JSON.parse(content) as Partial<NovelCacheData>;
        const parsedVersion = Number((parsed as any)?.version || 0);
        const isLegacyWithoutVersion =
            parsedVersion === 0
            && parsed
            && typeof parsed === "object"
            && (typeof (parsed as any).booksByHash === "object" || typeof (parsed as any).pathToHash === "object");
        if (parsedVersion !== 1 && !isLegacyWithoutVersion) {
            return empty;
        }

        const booksByHashRaw = parsed?.booksByHash && typeof parsed.booksByHash === "object"
            ? parsed.booksByHash as Record<string, NovelCacheBookRecord>
            : {};
        const pathToHashRaw = parsed?.pathToHash && typeof parsed.pathToHash === "object"
            ? parsed.pathToHash as Record<string, string>
            : {};

        const booksByHash: Record<string, NovelCacheBookRecord> = {};
        for (const [hash, value] of Object.entries(booksByHashRaw)) {
            if (!value || typeof value !== "object") {
                continue;
            }
            booksByHash[hash] = {
                hash: String((value as any).hash || hash),
                filePath: String((value as any).filePath || ""),
                fileName: String((value as any).fileName || ""),
                cursor: Number.isFinite(Number((value as any).cursor)) ? Math.max(0, Math.floor(Number((value as any).cursor))) : 0,
                lastReadAt: Number.isFinite(Number((value as any).lastReadAt)) ? Math.max(0, Math.floor(Number((value as any).lastReadAt))) : 0,
            };
        }

        const pathToHash: Record<string, string> = {};
        for (const [k, v] of Object.entries(pathToHashRaw)) {
            pathToHash[String(k)] = String(v || "");
        }

        return {
            version: 1,
            booksByHash,
            pathToHash,
        };
    } catch (error) {
        ztoolkit.log("novel.cache.read.error", error);
        return empty;
    }
}

function writeNovelCache(storageDirPath: string, cache: NovelCacheData) {
    const directory = makeLocalFile(storageDirPath);
    if (!directory || !directory.exists() || !directory.isDirectory()) {
        return;
    }

    const cacheFile = directory.clone() as any;
    cacheFile.append(NOVEL_CACHE_FILE_NAME);
    const nextContent = JSON.stringify(cache, null, 2);
    writeLocalTextFile(cacheFile.path, `${nextContent}\n`);
}

function listEpubFilesFromDirectory(storageDirPath: string): Array<{ filePath: string; fileName: string; }> {
    const directory = makeLocalFile(storageDirPath);
    if (!directory || !directory.exists() || !directory.isDirectory()) {
        return [];
    }

    const result: Array<{ filePath: string; fileName: string; }> = [];
    const iterator = directory.directoryEntries;
    while (iterator.hasMoreElements()) {
        const entry = iterator.getNext().QueryInterface(Components.interfaces.nsIFile) as any;
        if (!entry.isFile()) {
            continue;
        }
        if (!/\.epub$/i.test(entry.leafName)) {
            continue;
        }
        result.push({
            filePath: entry.path,
            fileName: entry.leafName,
        });
    }

    result.sort((a, b) => a.fileName.localeCompare(b.fileName, "en"));
    return result;
}

function computeFileSha256(path: string): string {
    const classesAny = Components.classes as any;
    const file = makeLocalFile(path);
    if (!file || !file.exists() || !file.isFile()) {
        throw new Error(`Book file not found: ${path}`);
    }

    const stream = classesAny["@mozilla.org/network/file-input-stream;1"].createInstance(
        Components.interfaces.nsIFileInputStream,
    );
    stream.init(file, 0x01, 0o444, 0);

    const crypto = classesAny["@mozilla.org/security/hash;1"].createInstance(
        Components.interfaces.nsICryptoHash,
    );
    crypto.init(Components.interfaces.nsICryptoHash.SHA256);
    crypto.updateFromStream(stream, file.fileSize);
    const digest = crypto.finish(true) as string;
    stream.close();

    return digest;
}

function makeLocalFile(path: string): any | undefined {
    const normalized = String(path || "").trim();
    if (!normalized) {
        return undefined;
    }

    try {
        const classesAny = Components.classes as any;
        const file = classesAny["@mozilla.org/file/local;1"].createInstance(
            Components.interfaces.nsIFile,
        ) as any;
        file.initWithPath(normalized);
        return file;
    } catch (_error) {
        return undefined;
    }
}

function readLocalTextFile(path: string): string {
    const classesAny = Components.classes as any;
    const file = makeLocalFile(path);
    if (!file || !file.exists() || !file.isFile()) {
        return "";
    }

    const fileInputStream = classesAny["@mozilla.org/network/file-input-stream;1"].createInstance(
        Components.interfaces.nsIFileInputStream,
    );
    fileInputStream.init(file, 0x01, 0o444, 0);

    const converter = classesAny["@mozilla.org/intl/converter-input-stream;1"].createInstance(
        Components.interfaces.nsIConverterInputStream,
    );
    converter.init(fileInputStream, "UTF-8", 0, 0);

    let output = "";
    const chunk = { value: "" };
    while (converter.readString(0xffffffff, chunk) !== 0) {
        output += chunk.value;
    }

    converter.close();
    fileInputStream.close();
    return output;
}

function writeLocalTextFile(path: string, content: string) {
    const classesAny = Components.classes as any;
    const file = makeLocalFile(path);
    if (!file) {
        return;
    }

    if (!file.exists()) {
        file.create(Components.interfaces.nsIFile.NORMAL_FILE_TYPE, 0o644);
    }

    const fileOutputStream = classesAny["@mozilla.org/network/file-output-stream;1"].createInstance(
        Components.interfaces.nsIFileOutputStream,
    );
    fileOutputStream.init(file, 0x02 | 0x08 | 0x20, 0o644, 0);

    const converter = classesAny["@mozilla.org/intl/converter-output-stream;1"].createInstance(
        Components.interfaces.nsIConverterOutputStream,
    );
    converter.init(fileOutputStream, "UTF-8", 0, 0);
    converter.writeString(content);
    converter.close();
    fileOutputStream.close();
}

function normalizeLocalPath(path: string): string {
    return String(path || "").replace(/\\/g, "/").trim();
}

function getLeafNameFromPath(path: string): string {
    const file = makeLocalFile(path);
    return file?.leafName || normalizeLocalPath(path).split("/").pop() || path;
}

function extractEnglishTextFromEpub(epubPath: string): string {
    const classesAny = Components.classes as any;
    const zip = classesAny["@mozilla.org/libjar/zip-reader;1"].createInstance(
        Components.interfaces.nsIZipReader,
    );
    const file = classesAny["@mozilla.org/file/local;1"].createInstance(
        Components.interfaces.nsIFile,
    );
    file.initWithPath(epubPath);
    zip.open(file);

    try {
        const opfPath = getPackageDocumentPath(zip);
        const spineEntries = getSpineHtmlEntries(zip, opfPath);
        const parser = new DOMParser();

        const chunks: string[] = [];
        for (const entry of spineEntries) {
            if (!zip.hasEntry(entry)) {
                continue;
            }
            const html = readZipEntryAsUTF8(zip, entry);
            if (!html) continue;
            const doc = parser.parseFromString(html, "text/html");
            doc.querySelectorAll("script,style,noscript,svg,math,img,figure").forEach((el: Element) => el.remove());
            const title = extractChapterTitleText(doc);
            const text = doc.body?.textContent || doc.documentElement?.textContent || "";
            const englishOnly = toEnglishOnly(text);

            if (title) {
                const titleEnglish = toEnglishOnly(title);
                if (titleEnglish && !englishOnly.toLowerCase().startsWith(titleEnglish.toLowerCase())) {
                    chunks.push(titleEnglish);
                }
            }

            if (englishOnly) {
                chunks.push(englishOnly);
            }
        }

        return chunks.join(" ").replace(/\s+/g, " ").trim();
    } finally {
        zip.close();
    }
}

function extractChapterTitleText(doc: Document): string {
    const heading = doc.querySelector("h1, h2, h3, h4, h5, h6");
    if (heading?.textContent?.trim()) {
        return heading.textContent.trim();
    }

    const title = doc.querySelector("title")?.textContent?.trim();
    if (title) {
        return title;
    }

    return "";
}

function getPackageDocumentPath(zip: any): string {
    const containerEntry = "META-INF/container.xml";
    if (!zip.hasEntry(containerEntry)) {
        throw new Error("Invalid EPUB: missing META-INF/container.xml");
    }

    const containerXml = readZipEntryAsUTF8(zip, containerEntry);
    const xml = new DOMParser().parseFromString(containerXml, "application/xml");
    const rootFile = xml.querySelector("rootfile");
    const fullPath = rootFile?.getAttribute("full-path");
    if (!fullPath) {
        throw new Error("Invalid EPUB: unable to locate package document");
    }
    return normalizeZipPath(fullPath);
}

function getSpineHtmlEntries(zip: any, opfPath: string): string[] {
    const opfXmlText = readZipEntryAsUTF8(zip, opfPath);
    const xml = new DOMParser().parseFromString(opfXmlText, "application/xml");

    const manifestMap = new Map<string, string>();
    const manifestItems = Array.from(xml.querySelectorAll("manifest > item")) as Element[];
    for (const item of manifestItems) {
        const id = item.getAttribute("id") || "";
        const href = item.getAttribute("href") || "";
        const mediaType = item.getAttribute("media-type") || "";
        if (!id || !href) continue;
        if (!/xhtml|html/i.test(mediaType) && !/\.(xhtml|html|htm)$/i.test(href)) {
            continue;
        }
        manifestMap.set(id, resolveZipRelativePath(opfPath, href));
    }

    const spineItems = Array.from(xml.querySelectorAll("spine > itemref")) as Element[];
    const result: string[] = [];
    for (const itemRef of spineItems) {
        const idRef = itemRef.getAttribute("idref") || "";
        const entry = manifestMap.get(idRef);
        if (entry) {
            result.push(entry);
        }
    }

    if (result.length > 0) {
        return result;
    }

    // Fallback for EPUBs with malformed spine
    const fallbackEntries: string[] = [];
    const iter = zip.findEntries("*.xhtml");
    while (iter.hasMore()) {
        fallbackEntries.push(String(iter.getNext()));
    }
    fallbackEntries.sort();
    return fallbackEntries;
}

function readZipEntryAsUTF8(zip: any, entry: string): string {
    const classesAny = Components.classes as any;
    const inputStream = zip.getInputStream(entry);
    const converter = classesAny["@mozilla.org/intl/converter-input-stream;1"].createInstance(
        Components.interfaces.nsIConverterInputStream,
    );
    converter.init(inputStream, "UTF-8", 0, 0);

    let output = "";
    const strObj = { value: "" };
    while (converter.readString(0xffffffff, strObj) !== 0) {
        output += strObj.value;
    }
    converter.close();
    inputStream.close();
    return output;
}

function toEnglishOnly(input: string): string {
    return input
        .replace(/\r\n/g, " ")
        .replace(/[^A-Za-z0-9\s.,;:!?"'()\[\]{}<>\-_/\\@#$%^&*+=~`]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

function resolveZipRelativePath(baseFilePath: string, relativePath: string): string {
    const baseSegments = normalizeZipPath(baseFilePath).split("/");
    baseSegments.pop();

    const relSegments = normalizeZipPath(relativePath).split("/");
    for (const segment of relSegments) {
        if (!segment || segment === ".") continue;
        if (segment === "..") {
            baseSegments.pop();
        } else {
            baseSegments.push(segment);
        }
    }

    return normalizeZipPath(baseSegments.join("/"));
}

function normalizeZipPath(path: string): string {
    return path.replace(/\\/g, "/").replace(/^\/+/, "");
}
