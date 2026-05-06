const crypto = require("crypto");
const { execFile } = require("child_process");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { Notice, Plugin, PluginSettingTab, Setting, TFile, normalizePath } = require("obsidian");

const TLDRAW_PLUGIN_ID = "tldraw";
const INK_PLUGIN_ID = "ink";
const EXCALIDRAW_PLUGIN_ID = "obsidian-excalidraw-plugin";
const MARKER_PLUGIN_ID = "marker-api";
const TEXT_EXTRACTOR_PLUGIN_ID = "text-extractor";
const TLDRAW_VIEW_TYPE = "tldraw-view";
const INK_WRITING_VIEW_TYPE = "ink_writing-view";
const INK_DRAWING_VIEW_TYPE = "ink_drawing-view";
const EXCALIDRAW_VIEW_TYPE = "excalidraw";
const INK_TLDRAW_EDITOR_PROPERTY = "__surfacePenBridgeTldrawEditor";
const ACTIVE_INK_EDITOR_PROPERTY = "__surfacePenBridgeActiveInkEditor";
const ACTIVE_INK_HOST_PROPERTY = "__surfacePenBridgeActiveInkHost";
const ACTIVE_INK_EVENT = "surface-pen-bridge-ink-active";
const TEMP_EXPORT_FOLDER = "_writing-bridge-cache";
const TEMP_EXPORT_FILE = normalizePath(`${TEMP_EXPORT_FOLDER}/selection.png`);
const STROKE_REQUEST_FILE = normalizePath(`${TEMP_EXPORT_FOLDER}/stroke-request.json`);
const DIRECT_TEXT_SHAPE_TYPES = new Set(["arrow", "geo", "note", "text"]);
const EXCALIDRAW_DIRECT_TEXT_TYPES = new Set(["text"]);
const COPY_DEBOUNCE_MS = 600;
const SELECTION_MENU_POLL_MS = 150;
const STYLE_PANEL_SYNC_MS = 500;
const TLDRAW_SIZE_STYLE = { id: "tldraw:size", defaultValue: "m" };
const TLDRAW_DRAW_SIZE_OPTIONS = ["default", "s", "m", "l", "xl"];
const RULED_PAGE_GRID_MULTIPLIER = 5;
const SIDEBAR_PEN_TAP_MAX_TRAVEL_PX = 18;
const SIDEBAR_PEN_TAP_MAX_DURATION_MS = 1200;
const SIDEBAR_PEN_REPLAY_DELAY_MS = 32;
const RULED_PAGE_MIN_SCREEN_STEP_PX = 12;
const RULED_PAGE_FULL_OPACITY_STEP_PX = 20;
const RULED_PAGE_MIN_OPACITY = 0.55;
const RULED_PAGE_LINE_THICKNESS_PX = 1;
const RULED_PAGE_EXTENT_PX = 24000;
const SCREENSHOT_EXPORT_PADDING = 8;
const SCREENSHOT_EXPORT_SCALE = 1;
const SCREENSHOT_EXPORT_PIXEL_RATIO = 1;
const OCR_EXPORT_PADDING = 24;
const OCR_EXPORT_SCALE = 3;
const OCR_EXPORT_PIXEL_RATIO = 2;
const INK_HELPER_EXE = "InkStrokeRecognizer.exe";
const REGION_SCREENSHOT_HELPER = "RegionScreenshot.ps1";
const OPENROUTER_API_URL = "https://openrouter.ai/api/v1/chat/completions";
const GEMINI_API_URL_BASE = "https://generativelanguage.googleapis.com/v1beta/models";
const VISION_PROVIDER_OPENROUTER = "openrouter";
const VISION_PROVIDER_GEMINI = "gemini";
const VISION_PROVIDER_CUSTOM = "custom";
const VISION_PROVIDERS = [
  VISION_PROVIDER_OPENROUTER,
  VISION_PROVIDER_GEMINI,
  VISION_PROVIDER_CUSTOM,
];
const GEMINI_DEFAULT_MODEL = "gemini-2.5-flash";
const OPENROUTER_DEFAULT_MODELS = [
  "moonshotai/kimi-vl-a3b-thinking:free",
  "qwen/qwen2.5-vl-32b-instruct:free",
];
const DEFAULT_VISION_TRANSCRIPTION_PROMPT = [
  "Transcribe the handwritten note or math in this image exactly.",
  "Preserve line breaks where they are visually obvious.",
  "Prefer the literal reading over cleanup or interpretation.",
  "If the selection is math, format it as readable linear math for pasting into a normal text box such as Excalidraw.",
  "Do not use LaTeX commands, LaTeX delimiters, or markdown math fencing.",
  "Prefer common math symbols such as ≤, ≥, ≠, ≈, √, ×, ÷, π, θ, ∑, ∫, and → when they are clearly intended.",
  "Use simple unicode superscripts like ² and ³ when obvious and compact; otherwise use ^ and _ only when needed.",
  "Keep fractions, exponents, limits, and equations easy to read in plain pasted text.",
  "Do not explain, solve, normalize, or summarize.",
  "Output only the transcription.",
].join(" ");
const TLDRAW_EMBED_PREVIEW_FORMAT_OPTIONS = ["png", "svg", "default"];
const FLOAT16_POW2 = Array.from({ length: 31 }, (_, index) => Math.pow(2, index - 15));
const FLOAT16_POW2_SUBNORMAL = Math.pow(2, -14) / 1024;
const FLOAT16_MANTISSA = Array.from({ length: 1024 }, (_, index) => 1 + index / 1024);
const DEFAULT_SETTINGS = {
  openRouterEnabled: true,
  visionLlmProvider: VISION_PROVIDER_OPENROUTER,
  openRouterApiKey: "",
  openRouterModels: OPENROUTER_DEFAULT_MODELS.join("\n"),
  geminiApiKey: "",
  geminiModel: GEMINI_DEFAULT_MODEL,
  customOpenAiBaseUrl: "",
  customOpenAiApiKey: "",
  customOpenAiModel: "",
  visionTranscriptionPrompt: DEFAULT_VISION_TRANSCRIPTION_PROMPT,
  toolCommandsEnabled: true,
  copySelectedTextEnabled: true,
  selectionScreenshotEnabled: true,
  regionScreenshotEnabled: true,
  selectionMenuEnabled: true,
  stylePanelToggleEnabled: true,
  ruledPageFeatureEnabled: true,
  sidebarPenAssistEnabled: true,
  tldrawDefaultPenSizeEnabled: true,
  excalidrawDefaultZoomEnabled: true,
  tldrawEmbedPreviewBridgeEnabled: true,
  tldrawDefaultDrawSize: "s",
  tldrawEmbedPreviewFormat: "png",
  stylePanelCollapsed: true,
  ruledPageEnabled: true,
  excalidrawDefaultZoom: "",
};

const FEATURE_TOGGLE_SETTING_KEYS = [
  "toolCommandsEnabled",
  "copySelectedTextEnabled",
  "selectionScreenshotEnabled",
  "regionScreenshotEnabled",
  "selectionMenuEnabled",
  "stylePanelToggleEnabled",
  "ruledPageFeatureEnabled",
  "sidebarPenAssistEnabled",
  "tldrawDefaultPenSizeEnabled",
  "excalidrawDefaultZoomEnabled",
  "tldrawEmbedPreviewBridgeEnabled",
];

const TOOL_COMMANDS = [
  {
    id: "surface-pen-draw",
    name: "Writing Bridge: Draw",
    hotkeyKey: "D",
    toolKey: "d",
    toolCode: "KeyD",
    toolId: "draw",
  },
  {
    id: "surface-pen-select",
    name: "Writing Bridge: Select",
    hotkeyKey: "S",
    toolKey: "v",
    toolCode: "KeyV",
    toolId: "select",
  },
  {
    id: "surface-pen-hand",
    name: "Writing Bridge: Hand",
    hotkeyKey: "H",
    toolKey: "h",
    toolCode: "KeyH",
    toolId: "hand",
  },
];

const COPY_COMMAND = {
  id: "surface-pen-copy-selected-text",
  name: "Writing Bridge: Copy Selected As Text",
  hotkeyKey: "C",
};

const SCREENSHOT_COMMAND = {
  id: "surface-pen-copy-selected-screenshot",
  name: "Writing Bridge: Copy Selected Screenshot",
  hotkeyKey: "P",
};

const REGION_SCREENSHOT_COMMAND = {
  id: "surface-pen-region-screenshot",
  name: "Writing Bridge: Region Screenshot to Clipboard",
};

const REGION_SCREENSHOT_PASTE_COMMAND = {
  id: "surface-pen-region-screenshot-paste",
  name: "Writing Bridge: Region Screenshot and Paste",
};

const REGION_SCROLLING_SCREENSHOT_COMMAND = {
  id: "surface-pen-region-scrolling-screenshot",
  name: "Writing Bridge: Scrolling Region Screenshot to Clipboard",
};

const REGION_SCROLLING_SCREENSHOT_PASTE_COMMAND = {
  id: "surface-pen-region-scrolling-screenshot-paste",
  name: "Writing Bridge: Scrolling Region Screenshot and Paste",
};

const TOGGLE_RULED_PAGE_COMMAND = {
  id: "surface-pen-toggle-ruled-page",
  name: "Writing Bridge: Toggle Ruled Page",
};

module.exports = class WritingBridgePlugin extends Plugin {
  async onload() {
    await this.loadSettings();
    this.copyInFlight = false;
    this.screenshotInFlight = false;
    this.regionScreenshotInFlight = false;
    this.lastCopyStartedAt = 0;
    this.selectionMenuEl = null;
    this.lastActiveCanvasTarget = null;
    this.lastActiveTldrawTarget = null;
    this.appliedTldrawDrawSizeByRootId = new Map();
    this.appliedExcalidrawZoomLeaves = new WeakSet();
    this.stylePanelToggleId = 0;
    this.excalidrawEmbedButtonId = 0;
    this.pendingExcalidrawFloatingUiSyncFrame = 0;
    this.pendingSidebarPenReplayTimeouts = new Set();
    this.tldrawEmbedPreviewObjectUrls = new Set();
    this.tldrawEmbedPreviewObjectUrlByImage = new WeakMap();
    this.tldrawEmbedPreviewSourceUrlByImage = new WeakMap();
    this.tldrawEmbedPreviewConversionByImage = new WeakMap();
    this.pendingTldrawEmbedPreviewSyncTimeout = 0;
    this.excalidrawFloatingUiTrackingUntil = 0;
    this.ruledPageStateByRootId = new Map();
    this.sidebarPenTapState = null;
    this.selectionMenuState = {
      dismissedSelectionKey: "",
      lastSelectionKey: "",
    };
    this.log("plugin loaded");

    for (const command of TOOL_COMMANDS) {
      this.registerToolCommand(command);
    }

    this.addSettingTab(new WritingBridgeSettingTab(this.app, this));
    this.registerCopyCommand();
    this.registerScreenshotCommand();
    this.registerRegionScreenshotCommands();
    this.installSelectionMenu();
    this.installStylePanelToggle();
    this.registerRuledPageCommand();
    this.installRuledPageOverlay();
    this.installSidebarPenAssist();
    this.installTldrawEmbedPreviewFormatBridge();
    this.registerInterval(window.setInterval(() => this.refreshSelectionMenu(), SELECTION_MENU_POLL_MS));
    const refreshSelectionMenu = () => this.refreshSelectionMenu();
    window.addEventListener(ACTIVE_INK_EVENT, refreshSelectionMenu);
    this.register(() => window.removeEventListener(ACTIVE_INK_EVENT, refreshSelectionMenu));
    this.register(() => {
      for (const timeoutId of this.pendingSidebarPenReplayTimeouts) {
        window.clearTimeout(timeoutId);
      }
      this.pendingSidebarPenReplayTimeouts.clear();
      if (this.pendingTldrawEmbedPreviewSyncTimeout) {
        window.clearTimeout(this.pendingTldrawEmbedPreviewSyncTimeout);
        this.pendingTldrawEmbedPreviewSyncTimeout = 0;
      }
      for (const url of this.tldrawEmbedPreviewObjectUrls) {
        URL.revokeObjectURL(url);
      }
      this.tldrawEmbedPreviewObjectUrls.clear();
    });
    this.register(() => this.destroySelectionMenu());
  }

  async loadSettings() {
    const loaded = (await this.loadData()) ?? {};
    this.settings = {
      ...DEFAULT_SETTINGS,
      ...loaded,
    };

    if (Array.isArray(this.settings.openRouterModels)) {
      this.settings.openRouterModels = this.settings.openRouterModels.join("\n");
    }

    if (!VISION_PROVIDERS.includes(this.settings.visionLlmProvider)) {
      this.settings.visionLlmProvider = DEFAULT_SETTINGS.visionLlmProvider;
    }

    if (!TLDRAW_DRAW_SIZE_OPTIONS.includes(this.settings.tldrawDefaultDrawSize)) {
      this.settings.tldrawDefaultDrawSize = DEFAULT_SETTINGS.tldrawDefaultDrawSize;
    }

    if (!TLDRAW_EMBED_PREVIEW_FORMAT_OPTIONS.includes(this.settings.tldrawEmbedPreviewFormat)) {
      this.settings.tldrawEmbedPreviewFormat = DEFAULT_SETTINGS.tldrawEmbedPreviewFormat;
    }

    if (typeof this.settings.visionTranscriptionPrompt !== "string") {
      this.settings.visionTranscriptionPrompt = DEFAULT_SETTINGS.visionTranscriptionPrompt;
    }

    if (typeof this.settings.excalidrawDefaultZoom !== "string") {
      this.settings.excalidrawDefaultZoom =
        this.settings.excalidrawDefaultZoom == null
          ? ""
          : `${this.settings.excalidrawDefaultZoom}`;
    }

    for (const key of FEATURE_TOGGLE_SETTING_KEYS) {
      if (typeof this.settings[key] !== "boolean") {
        this.settings[key] = DEFAULT_SETTINGS[key];
      }
    }
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }

  isFeatureEnabled(settingKey) {
    return this.settings?.[settingKey] !== false;
  }

  showFeatureDisabledNotice(featureLabel) {
    new Notice(`${featureLabel} is disabled in Writing Bridge settings`);
  }

  registerToolCommand(command) {
    this.addCommand({
      id: command.id,
      name: command.name,
      hotkeys: [
        {
          modifiers: ["Mod", "Alt", "Shift"],
          key: command.hotkeyKey,
        },
      ],
      checkCallback: (checking) => {
        if (!this.isFeatureEnabled("toolCommandsEnabled")) {
          if (!checking) {
            this.showFeatureDisabledNotice("Canvas tool commands");
          }
          return false;
        }

        const target = this.getCanvasTarget();
        if (!target) {
          if (!checking) {
            this.log(`command ${command.id}: no active canvas target`);
          }
          return false;
        }

        if (!checking) {
          this.log(`command ${command.id}: activating ${command.toolId} on ${target.kind}`);
          this.activateTool(target, command);
        }

        return true;
      },
    });
  }

  registerCopyCommand() {
    this.addCommand({
      id: COPY_COMMAND.id,
      name: COPY_COMMAND.name,
      hotkeys: [
        {
          modifiers: ["Mod", "Alt", "Shift"],
          key: COPY_COMMAND.hotkeyKey,
        },
      ],
      callback: () => {
        if (!this.isFeatureEnabled("copySelectedTextEnabled")) {
          this.showFeatureDisabledNotice("Copy selected as text");
          return;
        }

        const target = this.getCanvasTarget();
        if (!target && !this.getActiveDomTextSelection(null)) {
          this.log(`command ${COPY_COMMAND.id}: no active canvas target`);
        }
        void this.copySelectedAsText(target);
      },
    });
  }

  registerScreenshotCommand() {
    this.addCommand({
      id: SCREENSHOT_COMMAND.id,
      name: SCREENSHOT_COMMAND.name,
      hotkeys: [
        {
          modifiers: ["Mod", "Alt", "Shift"],
          key: SCREENSHOT_COMMAND.hotkeyKey,
        },
      ],
      callback: () => {
        if (!this.isFeatureEnabled("selectionScreenshotEnabled")) {
          this.showFeatureDisabledNotice("Copy selected screenshot");
          return;
        }

        const target = this.getCanvasTarget();
        if (!target) {
          this.log(`command ${SCREENSHOT_COMMAND.id}: no active canvas target`);
        }
        void this.copySelectedAsScreenshot(target);
      },
    });
  }

  registerRegionScreenshotCommands() {
    this.addCommand({
      id: REGION_SCREENSHOT_COMMAND.id,
      name: REGION_SCREENSHOT_COMMAND.name,
      callback: () => {
        if (!this.isFeatureEnabled("regionScreenshotEnabled")) {
          this.showFeatureDisabledNotice("Region screenshots");
          return;
        }

        void this.captureScreenRegion({ paste: false });
      },
    });

    this.addCommand({
      id: REGION_SCREENSHOT_PASTE_COMMAND.id,
      name: REGION_SCREENSHOT_PASTE_COMMAND.name,
      callback: () => {
        if (!this.isFeatureEnabled("regionScreenshotEnabled")) {
          this.showFeatureDisabledNotice("Region screenshots");
          return;
        }

        void this.captureScreenRegion({ paste: true });
      },
    });

    this.addCommand({
      id: REGION_SCROLLING_SCREENSHOT_COMMAND.id,
      name: REGION_SCROLLING_SCREENSHOT_COMMAND.name,
      callback: () => {
        if (!this.isFeatureEnabled("regionScreenshotEnabled")) {
          this.showFeatureDisabledNotice("Region screenshots");
          return;
        }

        void this.captureScrollingScreenRegion({ paste: false });
      },
    });

    this.addCommand({
      id: REGION_SCROLLING_SCREENSHOT_PASTE_COMMAND.id,
      name: REGION_SCROLLING_SCREENSHOT_PASTE_COMMAND.name,
      callback: () => {
        if (!this.isFeatureEnabled("regionScreenshotEnabled")) {
          this.showFeatureDisabledNotice("Region screenshots");
          return;
        }

        void this.captureScrollingScreenRegion({ paste: true });
      },
    });
  }

  async captureScreenRegion({ paste = false } = {}) {
    if (!this.isFeatureEnabled("regionScreenshotEnabled")) {
      this.showFeatureDisabledNotice("Region screenshots");
      return;
    }

    if (this.regionScreenshotInFlight) {
      this.log("region screenshot skipped: already running");
      return;
    }

    this.regionScreenshotInFlight = true;
    const targetBeforeCapture = this.getCanvasTarget();
    const activeElementBeforeCapture = document.activeElement instanceof HTMLElement ? document.activeElement : null;

    try {
      await this.runRegionScreenshotHelper();
      new Notice(paste ? "Region screenshot copied; pasting..." : "Region screenshot copied");
      if (paste) {
        this.pasteClipboardImageToActiveTarget(targetBeforeCapture, activeElementBeforeCapture);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message !== "region screenshot cancelled") {
        this.log(`region screenshot failed: ${message}`);
        new Notice(this.getFailureNoticeMessage("Region screenshot", error));
      }
    } finally {
      this.regionScreenshotInFlight = false;
    }
  }

  async runRegionScreenshotHelper() {
    const helperPath = this.getRegionScreenshotHelperPath();
    if (!fs.existsSync(helperPath)) {
      throw new Error("Region screenshot helper is missing");
    }

    await new Promise((resolve, reject) => {
      execFile(
        "powershell.exe",
        ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", helperPath],
        { windowsHide: false },
        (error, stdout, stderr) => {
          if (!error) {
            resolve();
            return;
          }

          if (error.code === 2) {
            reject(new Error("region screenshot cancelled"));
            return;
          }

          const detail = String(stderr || stdout || error.message || "").trim();
          reject(new Error(detail || "Region screenshot helper failed"));
        }
      );
    });
  }

  async captureScrollingScreenRegion({ paste = false } = {}) {
    if (!this.isFeatureEnabled("regionScreenshotEnabled")) {
      this.showFeatureDisabledNotice("Region screenshots");
      return;
    }

    if (this.regionScreenshotInFlight) {
      this.log("scrolling region screenshot skipped: already running");
      return;
    }

    this.regionScreenshotInFlight = true;
    const targetBeforeCapture = this.getCanvasTarget();
    const activeElementBeforeCapture = document.activeElement instanceof HTMLElement ? document.activeElement : null;
    const scrollElement = this.getActiveScreenshotScrollElement();

    try {
      const screenRect = await this.selectScreenRegion();
      if (!scrollElement) {
        await this.runRegionScreenshotHelperForRect(screenRect);
        new Notice(paste ? "Region screenshot copied; pasting..." : "Region screenshot copied");
        if (paste) {
          this.pasteClipboardImageToActiveTarget(targetBeforeCapture, activeElementBeforeCapture);
        }
        return;
      }

      const blob = await this.captureScrollingRegionToBlob(screenRect, scrollElement);
      await this.writeClipboardImage(blob);
      new Notice(paste ? "Scrolling screenshot copied; pasting..." : "Scrolling screenshot copied");
      if (paste) {
        this.pasteClipboardImageToActiveTarget(targetBeforeCapture, activeElementBeforeCapture);
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (message !== "region screenshot cancelled") {
        this.log(`scrolling region screenshot failed: ${message}`);
        new Notice(this.getFailureNoticeMessage("Scrolling screenshot", error));
      }
    } finally {
      this.regionScreenshotInFlight = false;
    }
  }

  async selectScreenRegion() {
    const helperPath = this.getRegionScreenshotHelperPath();
    if (!fs.existsSync(helperPath)) {
      throw new Error("Region screenshot helper is missing");
    }

    const stdout = await new Promise((resolve, reject) => {
      execFile(
        "powershell.exe",
        ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", helperPath, "-SelectOnly"],
        { windowsHide: false },
        (error, stdout, stderr) => {
          if (!error) {
            resolve(String(stdout || "").trim());
            return;
          }

          if (error.code === 2) {
            reject(new Error("region screenshot cancelled"));
            return;
          }

          const detail = String(stderr || stdout || error.message || "").trim();
          reject(new Error(detail || "Region screenshot selection failed"));
        }
      );
    });

    let rect = null;
    try {
      rect = JSON.parse(stdout);
    } catch (error) {
      throw this.wrapErrorWithContext("Region screenshot selection returned invalid JSON", error);
    }

    if (!this.isValidScreenRect(rect)) {
      throw new Error("Region screenshot selection returned an invalid rectangle");
    }

    return {
      x: Math.round(rect.x),
      y: Math.round(rect.y),
      width: Math.round(rect.width),
      height: Math.round(rect.height),
    };
  }

  isValidScreenRect(rect) {
    return (
      rect &&
      Number.isFinite(Number(rect.x)) &&
      Number.isFinite(Number(rect.y)) &&
      Number.isFinite(Number(rect.width)) &&
      Number.isFinite(Number(rect.height)) &&
      Number(rect.width) >= 4 &&
      Number(rect.height) >= 4
    );
  }

  async runRegionScreenshotHelperForRect(screenRect, outputFile = "") {
    const helperPath = this.getRegionScreenshotHelperPath();
    if (!fs.existsSync(helperPath)) {
      throw new Error("Region screenshot helper is missing");
    }

    const args = [
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      helperPath,
      "-CaptureRectJson",
      JSON.stringify(screenRect),
    ];
    if (outputFile) {
      args.push("-OutputFile", outputFile);
    }

    await new Promise((resolve, reject) => {
      execFile("powershell.exe", args, { windowsHide: true }, (error, stdout, stderr) => {
        if (!error) {
          resolve();
          return;
        }

        const detail = String(stderr || stdout || error.message || "").trim();
        reject(new Error(detail || "Region screenshot capture failed"));
      });
    });
  }

  getActiveScreenshotScrollElement() {
    const activeLeaf = document.querySelector(".workspace-leaf.mod-active");
    if (!(activeLeaf instanceof HTMLElement)) {
      return null;
    }

    const selectors = [
      ".cm-scroller",
      ".markdown-preview-view",
      ".markdown-source-view .cm-scroller",
      ".view-content",
    ];

    for (const selector of selectors) {
      const candidate = activeLeaf.querySelector(selector);
      if (candidate instanceof HTMLElement && this.canElementScrollVertically(candidate)) {
        return candidate;
      }
    }

    return this.canElementScrollVertically(activeLeaf) ? activeLeaf : null;
  }

  canElementScrollVertically(element) {
    return element instanceof HTMLElement && element.scrollHeight - element.clientHeight > 8;
  }

  async captureScrollingRegionToBlob(screenRect, scrollElement) {
    const startScrollTop = scrollElement.scrollTop;
    const maxScrollTop = Math.max(0, scrollElement.scrollHeight - scrollElement.clientHeight);
    const scrollStep = Math.max(80, Math.floor(screenRect.height * 0.85));
    const maxOutputHeight = 16000;
    const captures = [];

    try {
      for (let index = 0; index < 80; index += 1) {
        const currentScrollTop = scrollElement.scrollTop;
        const yOffset = Math.max(0, Math.round(currentScrollTop - startScrollTop));
        if (yOffset >= maxOutputHeight) {
          break;
        }

        const imageBlob = await this.captureScreenRectToBlob(screenRect);
        captures.push({ blob: imageBlob, y: yOffset });

        if (currentScrollTop >= maxScrollTop - 1 || yOffset + screenRect.height >= maxOutputHeight) {
          break;
        }

        const nextScrollTop = Math.min(maxScrollTop, currentScrollTop + scrollStep);
        if (Math.abs(nextScrollTop - currentScrollTop) < 1) {
          break;
        }

        scrollElement.scrollTop = nextScrollTop;
        await this.waitForScrollSettle();
      }
    } finally {
      scrollElement.scrollTop = startScrollTop;
    }

    if (captures.length === 0) {
      throw new Error("No screenshot strips were captured");
    }

    return this.stitchScreenshotStrips(captures, screenRect.width, screenRect.height, maxOutputHeight);
  }

  async captureScreenRectToBlob(screenRect) {
    const outputFile = path.join(os.tmpdir(), `writing-bridge-region-${Date.now()}-${crypto.randomUUID()}.png`);
    try {
      await this.runRegionScreenshotHelperForRect(screenRect, outputFile);
      return new Blob([await fs.promises.readFile(outputFile)], { type: "image/png" });
    } finally {
      fs.promises.unlink(outputFile).catch(() => {});
    }
  }

  waitForScrollSettle() {
    return new Promise((resolve) => window.setTimeout(resolve, 180));
  }

  async stitchScreenshotStrips(captures, width, stripHeight, maxOutputHeight) {
    const outputHeight = Math.min(
      maxOutputHeight,
      Math.max(...captures.map((capture) => capture.y + stripHeight))
    );
    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = outputHeight;
    const context = canvas.getContext("2d");
    if (!context) {
      throw new Error("Could not create screenshot stitch canvas");
    }

    for (const capture of captures) {
      const bitmap = await createImageBitmap(capture.blob);
      try {
        context.drawImage(bitmap, 0, capture.y);
      } finally {
        bitmap.close?.();
      }
    }

    const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
    if (!blob) {
      throw new Error("Could not encode scrolling screenshot");
    }

    return blob;
  }

  getRegionScreenshotHelperPath() {
    return this.getVaultAbsolutePath(
      normalizePath(`${this.app.vault.configDir}/plugins/${this.manifest.id}/runtime/${REGION_SCREENSHOT_HELPER}`)
    );
  }

  pasteClipboardImageToActiveTarget(targetBeforeCapture = null, activeElementBeforeCapture = null) {
    const target = targetBeforeCapture ?? this.getCanvasTarget();
    const pasteTarget = this.getPasteTargetElement(target) ?? activeElementBeforeCapture ?? document.activeElement;
    if (!(pasteTarget instanceof HTMLElement)) {
      const error = new Error("No active paste target was available");
      this.log(`region screenshot paste failed: ${error.message}`);
      new Notice(this.getFailureNoticeMessage("Paste screenshot", error));
      return false;
    }

    pasteTarget.focus?.({ preventScroll: true });
    window.setTimeout(() => {
      try {
        this.sendShortcutKey(pasteTarget, "v", "KeyV", {
          ctrlKey: true,
        });
        this.dispatchClipboardCommand(pasteTarget, "paste");
      } catch (error) {
        const message = this.getErrorMessage(error);
        this.log(`region screenshot paste failed: ${message}`);
        new Notice(this.getFailureNoticeMessage("Paste screenshot", error));
      }
    }, 100);
    return true;
  }

  getPasteTargetElement(target) {
    if (target?.kind === "tldraw" && target.container instanceof HTMLElement) {
      return target.container;
    }

    if (target?.kind === "excalidraw") {
      const root = this.getExcalidrawRootForView(target.view) ?? target.container;
      return root instanceof HTMLElement ? root : null;
    }

    const activeEditorEl = document.querySelector(".workspace-leaf.mod-active .cm-editor.cm-focused");
    if (activeEditorEl instanceof HTMLElement) {
      return activeEditorEl;
    }

    const activeLeafContent = document.querySelector(".workspace-leaf.mod-active .workspace-leaf-content");
    return activeLeafContent instanceof HTMLElement ? activeLeafContent : null;
  }

  registerRuledPageCommand() {
    this.addCommand({
      id: TOGGLE_RULED_PAGE_COMMAND.id,
      name: TOGGLE_RULED_PAGE_COMMAND.name,
      callback: () => {
        if (!this.isFeatureEnabled("ruledPageFeatureEnabled")) {
          this.showFeatureDisabledNotice("Ruled page overlay");
          return;
        }

        void this.toggleRuledPageEnabled();
      },
    });
  }

  getCanvasTarget() {
    const activeElement = document.activeElement;
    const activeTldrawContainer = this.getOwningTldrawContainer(activeElement);
    if (activeTldrawContainer) {
      return this.normalizeCanvasTarget(this.getTldrawTarget());
    }

    const activeExcalidrawContainer = this.getOwningExcalidrawContainer(activeElement);
    if (activeExcalidrawContainer) {
      return this.normalizeCanvasTarget(this.getExcalidrawTarget());
    }

    const activeLeafType = this.app.workspace.activeLeaf?.view?.getViewType?.();
    if (activeLeafType === TLDRAW_VIEW_TYPE) {
      return this.normalizeCanvasTarget(this.getTldrawTarget());
    }

    if (this.isInkViewType(activeLeafType)) {
      return this.normalizeCanvasTarget(this.getTldrawTarget());
    }

    if (activeLeafType === EXCALIDRAW_VIEW_TYPE) {
      return this.normalizeCanvasTarget(this.getExcalidrawTarget());
    }

    const selectedTldrawTarget = this.getSelectedTldrawTarget();
    if (selectedTldrawTarget) {
      return this.normalizeCanvasTarget(selectedTldrawTarget);
    }

    const visibleSelectedTldrawTarget = this.getVisibleSelectedTldrawTarget();
    if (visibleSelectedTldrawTarget) {
      return this.normalizeCanvasTarget(visibleSelectedTldrawTarget);
    }

    return (
      this.normalizeCanvasTarget(this.getTldrawTarget()) ??
      this.normalizeCanvasTarget(this.getExcalidrawTarget()) ??
      this.normalizeCanvasTarget(this.lastActiveCanvasTarget)
    );
  }

  normalizeCanvasTarget(target) {
    if (!target?.kind) {
      return null;
    }

    if (target.kind === "tldraw") {
      const normalized = this.normalizeTldrawTarget(target);
      return normalized ? { ...normalized, kind: "tldraw" } : null;
    }

    if (target.kind === "excalidraw") {
      const normalized = this.normalizeExcalidrawTarget(target);
      return normalized ? { ...normalized, kind: "excalidraw" } : null;
    }

    return null;
  }

  getTldrawTarget() {
    const activeContainer = this.getActiveTldrawContainer();
    const editor = this.getCurrentTldrawEditor(activeContainer);

    if (editor || activeContainer) {
      return { kind: "tldraw", editor, container: activeContainer };
    }

    const activeLeaf = this.app.workspace.activeLeaf;
    const activeLeafViewType = activeLeaf?.view?.getViewType?.();

    if (activeLeafViewType === TLDRAW_VIEW_TYPE && editor) {
      return { kind: "tldraw", editor, container: null };
    }

    return null;
  }

  getExcalidrawTarget() {
    const activeContainer = this.getActiveExcalidrawContainer();
    const view = this.getCurrentExcalidrawView(activeContainer);
    const api = this.getExcalidrawApi(view);

    if (view || activeContainer) {
      return { kind: "excalidraw", view, api, container: activeContainer };
    }

    return null;
  }

  getCurrentTldrawEditor(activeContainer) {
    const attachedEditor =
      this.getAttachedTldrawEditor(activeContainer) ??
      this.getAttachedTldrawEditor(document.querySelector(".workspace-leaf.mod-active"));
    if (attachedEditor) {
      return attachedEditor;
    }

    const activeInk = this.getWindowActiveInkTarget();
    if (activeInk && this.isWindowActiveInkTargetRelevant(activeContainer)) {
      return activeInk.editor;
    }

    const tldrawPlugin = this.app?.plugins?.plugins?.[TLDRAW_PLUGIN_ID];
    const editor = tldrawPlugin?.currTldrawEditor;
    if (!editor || typeof editor.setCurrentTool !== "function") {
      return null;
    }

    const activeLeaf = this.app.workspace.activeLeaf;
    const activeLeafViewType = activeLeaf?.view?.getViewType?.();
    if (this.isInkViewType(activeLeafViewType) || this.isInkTldrawContainer(activeContainer)) {
      return null;
    }

    if (activeLeafViewType === TLDRAW_VIEW_TYPE) {
      return editor;
    }

    if (activeContainer) {
      return editor;
    }

    return null;
  }

  isInkViewType(viewType) {
    return viewType === INK_WRITING_VIEW_TYPE || viewType === INK_DRAWING_VIEW_TYPE;
  }

  getSelectedTldrawTarget() {
    const roots = [
      document.querySelector(".workspace-leaf.mod-active"),
      this.app.workspace.activeLeaf?.view?.containerEl ?? null,
    ];
    for (const root of roots) {
      const target = this.findSelectedTldrawTarget(root);
      if (target) {
        return target;
      }
    }

    return null;
  }

  getVisibleSelectedTldrawTarget() {
    const selector = [
      ".tldraw-view-root",
      ".ptl-tldraw-image",
      ".ddc_ink_writing-editor",
      ".ddc_ink_drawing-editor",
      ".ddc_ink_embed",
      `.workspace-leaf-content[data-type="${TLDRAW_VIEW_TYPE}"]`,
      `.workspace-leaf-content[data-type="${INK_WRITING_VIEW_TYPE}"]`,
      `.workspace-leaf-content[data-type="${INK_DRAWING_VIEW_TYPE}"]`,
    ].join(", ");
    const roots = document.querySelectorAll(selector);

    for (const root of roots) {
      if (!(root instanceof HTMLElement) || !this.isElementProbablyVisible(root)) {
        continue;
      }

      const target = this.findSelectedTldrawTarget(root);
      if (target) {
        return target;
      }
    }

    return null;
  }

  findSelectedTldrawTarget(root) {
    if (!(root instanceof HTMLElement)) {
      return null;
    }

    const selector = [
      ".tl-container.tl-container__focused",
      ".tldraw-view-root",
      ".ptl-tldraw-image",
      ".ddc_ink_writing-editor",
      ".ddc_ink_drawing-editor",
      ".ddc_ink_embed",
      `.workspace-leaf-content[data-type="${TLDRAW_VIEW_TYPE}"]`,
      `.workspace-leaf-content[data-type="${INK_WRITING_VIEW_TYPE}"]`,
      `.workspace-leaf-content[data-type="${INK_DRAWING_VIEW_TYPE}"]`,
      ".tl-container",
    ].join(", ");
    const candidates = [root, ...root.querySelectorAll(selector)];
    const seenEditors = new Set();

    for (const candidate of candidates) {
      if (!(candidate instanceof HTMLElement)) {
        continue;
      }

      const target = this.normalizeTldrawTarget({
        editor: this.getAttachedTldrawEditor(candidate),
        container: candidate,
      });
      const editor = target?.editor;
      if (!editor || seenEditors.has(editor)) {
        continue;
      }

      seenEditors.add(editor);
      if (this.getTldrawSelectionIds(editor).length > 0) {
        return { kind: "tldraw", ...target };
      }
    }

    return null;
  }

  isElementProbablyVisible(element) {
    if (!(element instanceof HTMLElement) || !element.isConnected) {
      return false;
    }

    const rect = element.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
  }

  isInkTldrawContainer(element) {
    return (
      element instanceof HTMLElement &&
      Boolean(
        element.closest(
          [
            ".ddc_ink_writing-editor",
            ".ddc_ink_drawing-editor",
            ".ddc_ink_embed",
            `.workspace-leaf-content[data-type="${INK_WRITING_VIEW_TYPE}"]`,
            `.workspace-leaf-content[data-type="${INK_DRAWING_VIEW_TYPE}"]`,
          ].join(", ")
        )
      )
    );
  }

  getWindowActiveInkTarget() {
    if (typeof window === "undefined") {
      return null;
    }

    const editor = window[ACTIVE_INK_EDITOR_PROPERTY];
    const host = window[ACTIVE_INK_HOST_PROPERTY];
    if (
      !editor ||
      typeof editor.getSelectedShapeIds !== "function" ||
      typeof editor.setCurrentTool !== "function" ||
      !(host instanceof HTMLElement) ||
      !host.isConnected
    ) {
      return null;
    }

    return { editor, host };
  }

  isWindowActiveInkTargetRelevant(activeContainer) {
    const activeInk = this.getWindowActiveInkTarget();
    if (!activeInk) {
      return false;
    }

    if (activeContainer instanceof HTMLElement && this.isInkTldrawContainer(activeContainer)) {
      return true;
    }

    const activeLeafViewType = this.app.workspace.activeLeaf?.view?.getViewType?.();
    if (this.isInkViewType(activeLeafViewType)) {
      return true;
    }

    const activeLeafContainer = this.app.workspace.activeLeaf?.view?.containerEl;
    if (activeLeafContainer instanceof HTMLElement && activeLeafContainer.contains(activeInk.host)) {
      return true;
    }

    const activeLeaf = document.querySelector(".workspace-leaf.mod-active");
    if (activeLeaf instanceof HTMLElement && activeLeaf.contains(activeInk.host)) {
      return true;
    }

    const activeElement = document.activeElement;
    return activeElement instanceof Node && activeInk.host.contains(activeElement);
  }

  getAttachedTldrawEditor(node) {
    if (!(node instanceof HTMLElement)) {
      return null;
    }

    const hosts = [
      node,
      node.closest(".ddc_ink_writing-editor"),
      node.closest(".ddc_ink_drawing-editor"),
      node.closest(".ddc_ink_embed"),
      node.closest(`.workspace-leaf-content[data-type="${INK_WRITING_VIEW_TYPE}"]`),
      node.closest(`.workspace-leaf-content[data-type="${INK_DRAWING_VIEW_TYPE}"]`),
      node.closest(".tl-container"),
      node.querySelector(".ddc_ink_writing-editor"),
      node.querySelector(".ddc_ink_drawing-editor"),
      node.querySelector(".ddc_ink_embed"),
      node.querySelector(`.workspace-leaf-content[data-type="${INK_WRITING_VIEW_TYPE}"]`),
      node.querySelector(`.workspace-leaf-content[data-type="${INK_DRAWING_VIEW_TYPE}"]`),
      node.querySelector(".tl-container"),
    ];

    for (const host of hosts) {
      const editor = host?.[INK_TLDRAW_EDITOR_PROPERTY];
      if (this.isUsableTldrawEditor(editor)) {
        return editor;
      }

      const reactEditor = this.getInkEditorFromReactTree(host);
      if (this.isUsableTldrawEditor(reactEditor)) {
        return reactEditor;
      }
    }

    return null;
  }

  isUsableTldrawEditor(editor) {
    return (
      !!editor &&
      typeof editor.getSelectedShapeIds === "function" &&
      typeof editor.setCurrentTool === "function"
    );
  }

  getInkEditorFromReactTree(node) {
    if (!(node instanceof HTMLElement)) {
      return null;
    }

    const hosts = [
      node,
      node.closest(".ddc_ink_writing-editor"),
      node.closest(".ddc_ink_drawing-editor"),
      node.closest(".ddc_ink_embed"),
      node.closest(".ddc_ink_widget-root"),
      node.closest(".ink-writing-view-host"),
      node.closest(".ink-drawing-view-host"),
      node.querySelector(".ddc_ink_writing-editor"),
      node.querySelector(".ddc_ink_drawing-editor"),
      node.querySelector(".ddc_ink_embed"),
      node.querySelector(".ddc_ink_widget-root"),
      node.querySelector(".ink-writing-view-host"),
      node.querySelector(".ink-drawing-view-host"),
    ].filter((host) => host instanceof HTMLElement);

    for (const host of new Set(hosts)) {
      const editor = this.findTldrawEditorInReactNode(host);
      if (this.isUsableTldrawEditor(editor)) {
        return editor;
      }
    }

    return null;
  }

  findTldrawEditorInReactNode(node) {
    const fibers = this.getReactFibersForNode(node);
    for (const fiber of fibers) {
      const editor = this.findTldrawEditorInReactFiber(fiber);
      if (this.isUsableTldrawEditor(editor)) {
        return editor;
      }
    }

    return null;
  }

  getReactFibersForNode(node) {
    if (!(node instanceof HTMLElement)) {
      return [];
    }

    const fibers = [];
    for (const key of Object.keys(node)) {
      if (key.startsWith("__reactFiber$")) {
        fibers.push(node[key]);
        continue;
      }

      if (key.startsWith("__reactContainer$")) {
        const container = node[key];
        const current =
          container?._internalRoot?.current ??
          container?.current ??
          container?.stateNode?.current ??
          null;
        if (current) {
          fibers.push(current);
        }
      }
    }

    return fibers.filter(Boolean);
  }

  findTldrawEditorInReactFiber(rootFiber) {
    if (!rootFiber || typeof rootFiber !== "object") {
      return null;
    }

    const queue = [rootFiber];
    const seen = new Set();
    let visited = 0;

    while (queue.length && visited < 3000) {
      const fiber = queue.shift();
      if (!fiber || typeof fiber !== "object" || seen.has(fiber)) {
        continue;
      }

      seen.add(fiber);
      visited += 1;

      const propEditor = this.extractTldrawEditorFromReactProps(fiber.memoizedProps ?? fiber.pendingProps);
      if (this.isUsableTldrawEditor(propEditor)) {
        return propEditor;
      }

      const hookEditor = this.extractTldrawEditorFromReactHooks(fiber.memoizedState);
      if (this.isUsableTldrawEditor(hookEditor)) {
        return hookEditor;
      }

      const stateNodeEditor = this.extractTldrawEditorFromArbitraryValue(fiber.stateNode);
      if (this.isUsableTldrawEditor(stateNodeEditor)) {
        return stateNodeEditor;
      }

      if (fiber.child) {
        queue.push(fiber.child);
      }
      if (fiber.sibling) {
        queue.push(fiber.sibling);
      }
    }

    return null;
  }

  extractTldrawEditorFromReactProps(props) {
    if (!props || typeof props !== "object") {
      return null;
    }

    if (typeof props.getTlEditor === "function") {
      try {
        const editor = props.getTlEditor();
        if (this.isUsableTldrawEditor(editor)) {
          return editor;
        }
      } catch (_error) {
      }
    }

    return this.extractTldrawEditorFromArbitraryValue(props);
  }

  extractTldrawEditorFromReactHooks(hook) {
    const seen = new Set();
    let current = hook;
    let depth = 0;

    while (current && typeof current === "object" && !seen.has(current) && depth < 200) {
      seen.add(current);
      const editor = this.extractTldrawEditorFromArbitraryValue(current.memoizedState);
      if (this.isUsableTldrawEditor(editor)) {
        return editor;
      }
      current = current.next;
      depth += 1;
    }

    return null;
  }

  extractTldrawEditorFromArbitraryValue(value, seen = new Set(), depth = 0) {
    if (this.isUsableTldrawEditor(value)) {
      return value;
    }

    if (!value || typeof value !== "object" || seen.has(value) || depth > 4) {
      return null;
    }

    seen.add(value);

    if (typeof value.getTlEditor === "function") {
      try {
        const editor = value.getTlEditor();
        if (this.isUsableTldrawEditor(editor)) {
          return editor;
        }
      } catch (_error) {
      }
    }

    const directKeys = [
      "current",
      "editor",
      "tlEditor",
      "memoizedState",
      "stateNode",
      "props",
    ];
    for (const key of directKeys) {
      if (!(key in value)) {
        continue;
      }
      const editor = this.extractTldrawEditorFromArbitraryValue(value[key], seen, depth + 1);
      if (this.isUsableTldrawEditor(editor)) {
        return editor;
      }
    }

    if (Array.isArray(value)) {
      for (const entry of value) {
        const editor = this.extractTldrawEditorFromArbitraryValue(entry, seen, depth + 1);
        if (this.isUsableTldrawEditor(editor)) {
          return editor;
        }
      }
    }

    return null;
  }

  getActiveTldrawContainer() {
    const doc = document;
    const activeElement = doc.activeElement;

    const activeContainer = this.getOwningTldrawContainer(activeElement);
    if (activeContainer) {
      return activeContainer;
    }

    const activeInk = this.getWindowActiveInkTarget();
    if (activeInk && this.isWindowActiveInkTargetRelevant(null)) {
      const activeInkContainer = this.findTldrawContainer(activeInk.host) ?? this.getOwningTldrawContainer(activeInk.host);
      if (activeInkContainer) {
        return activeInkContainer;
      }
    }

    const activeLeaf = doc.querySelector(".workspace-leaf.mod-active");
    if (activeLeaf instanceof HTMLElement) {
      const activeLeafContainer = this.findTldrawContainer(activeLeaf);
      if (activeLeafContainer) {
        return activeLeafContainer;
      }
    }

    return null;
  }

  getCurrentExcalidrawView(activeContainer) {
    const activeLeaf = this.app.workspace.activeLeaf;
    const activeView = activeLeaf?.view;
    if (this.isExcalidrawView(activeView)) {
      return activeView;
    }

    if (activeContainer instanceof HTMLElement) {
      const leaves = this.app.workspace.getLeavesOfType(EXCALIDRAW_VIEW_TYPE) ?? [];
      for (const leaf of leaves) {
        const view = leaf?.view;
        if (!this.isExcalidrawView(view)) {
          continue;
        }

        const root = this.getExcalidrawRootForView(view);
        if (root && root.contains(activeContainer)) {
          return view;
        }
      }
    }

    return null;
  }

  getExcalidrawApi(view) {
    const api = view?.excalidrawAPI;
    return api && typeof api.getAppState === "function" ? api : null;
  }

  getConfiguredExcalidrawDefaultZoom() {
    if (!this.isFeatureEnabled("excalidrawDefaultZoomEnabled")) {
      return null;
    }

    const raw =
      typeof this.settings?.excalidrawDefaultZoom === "string"
        ? this.settings.excalidrawDefaultZoom.trim()
        : "";
    if (raw === "") {
      return null;
    }

    const normalized = raw.endsWith("%") ? raw.slice(0, -1) : raw;
    const parsed = Number(normalized);
    if (!Number.isFinite(parsed) || parsed <= 0) {
      return null;
    }

    const ratio = raw.endsWith("%") || parsed > 10 ? parsed / 100 : parsed;
    if (!Number.isFinite(ratio) || ratio <= 0) {
      return null;
    }

    return Math.min(30, Math.max(0.05, ratio));
  }

  applyPreferredExcalidrawZoom(view) {
    const zoom = this.getConfiguredExcalidrawDefaultZoom();
    if (zoom === null) {
      return;
    }

    const leaf = view?.leaf;
    if (!leaf || this.appliedExcalidrawZoomLeaves.has(leaf)) {
      return;
    }

    const api = this.getExcalidrawApi(view);
    if (!api || typeof api.updateScene !== "function") {
      return;
    }

    try {
      api.updateScene({
        appState: {
          zoom: { value: zoom },
        },
      });
      this.appliedExcalidrawZoomLeaves.add(leaf);
      this.log(`applied default Excalidraw zoom ${zoom} to leaf`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`failed to apply default Excalidraw zoom: ${message}`);
    }
  }

  syncExcalidrawDefaultZoom() {
    if (this.getConfiguredExcalidrawDefaultZoom() === null) {
      return;
    }

    const leaves = this.app.workspace.getLeavesOfType(EXCALIDRAW_VIEW_TYPE) ?? [];
    for (const leaf of leaves) {
      const view = leaf?.view;
      if (this.isExcalidrawView(view)) {
        this.applyPreferredExcalidrawZoom(view);
      }
    }
  }

  isExcalidrawView(view) {
    return Boolean(view && typeof view.getViewType === "function" && view.getViewType() === EXCALIDRAW_VIEW_TYPE);
  }

  getActiveExcalidrawContainer() {
    const doc = document;
    const activeElement = doc.activeElement;

    const activeContainer = this.getOwningExcalidrawContainer(activeElement);
    if (activeContainer) {
      return activeContainer;
    }

    const activeLeaf = doc.querySelector('.workspace-leaf.mod-active .workspace-leaf-content[data-type="excalidraw"]');
    if (activeLeaf instanceof HTMLElement) {
      const activeLeafContainer = this.findExcalidrawContainer(activeLeaf);
      if (activeLeafContainer) {
        return activeLeafContainer;
      }
    }

    return null;
  }

  activateTool(target, command) {
    if (target?.kind === "excalidraw") {
      this.activateExcalidrawTool(target, command.toolId);
      return;
    }

    if (target?.editor && this.setEditorTool(target.editor, command.toolId)) {
      return;
    }

    if (target?.container) {
      this.sendToolKey(target.container, command.toolKey, command.toolCode);
    }
  }

  activateExcalidrawTool(target, toolId) {
    const api = target?.api ?? this.getExcalidrawApi(target?.view);
    if (!api || typeof api.setActiveTool !== "function") {
      this.log(`excalidraw tool change failed: missing API for ${toolId}`);
      return;
    }

    const excalidrawTool =
      toolId === "draw" ? "freedraw" : toolId === "select" ? "selection" : toolId === "hand" ? "hand" : null;
    if (!excalidrawTool) {
      return;
    }

    try {
      api.setActiveTool({ type: excalidrawTool });
      this.focusExcalidrawTarget(target);
      this.log(`excalidraw tool set to ${excalidrawTool}`);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`excalidraw tool change failed: ${message}`);
    }
  }

  setEditorTool(editor, toolId) {
    try {
      editor.setCurrentTool(toolId);
      if (typeof editor.focus === "function") {
        editor.focus();
      }
      this.log(`editor tool set to ${toolId}`);
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`editor tool change failed: ${message}`);
      return false;
    }
  }

  installSelectionMenu() {
    const styleEl = document.createElement("style");
    styleEl.setAttribute("data-writing-bridge", "selection-menu");
    styleEl.textContent = `
      .writing-bridge-selection-menu {
        position: fixed;
        z-index: 100000;
        display: none;
        align-items: center;
        gap: 8px;
        padding: 8px;
        border-radius: 14px;
        border: 1px solid rgba(255, 255, 255, 0.14);
        background: rgba(24, 24, 28, 0.92);
        box-shadow: 0 10px 28px rgba(0, 0, 0, 0.28);
        backdrop-filter: blur(14px);
        touch-action: manipulation;
      }

      .writing-bridge-selection-menu.is-visible {
        display: flex;
      }

      .writing-bridge-selection-menu button {
        appearance: none;
        border: 0;
        border-radius: 10px;
        min-height: 40px;
        min-width: 72px;
        padding: 10px 14px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font: 600 12px/1 "Segoe UI", sans-serif;
        color: #f6f7fb;
        background: rgba(255, 255, 255, 0.08);
        cursor: pointer;
        touch-action: manipulation;
        user-select: none;
      }

      .writing-bridge-selection-menu button:hover {
        background: rgba(255, 255, 255, 0.16);
      }

      .writing-bridge-selection-menu button.writing-bridge-selection-menu__primary {
        background: #2f7cf6;
        color: white;
      }
    `;
    document.head.appendChild(styleEl);
    this.register(() => styleEl.remove());
  }

  installStylePanelToggle() {
    const styleEl = document.createElement("style");
    styleEl.setAttribute("data-writing-bridge", "style-panel-toggle");
    styleEl.textContent = `
      .tldraw-view-root.writing-bridge-style-panel-collapsed .tlui-style-panel {
        display: none !important;
      }

      .workspace-leaf-content[data-type="excalidraw"].writing-bridge-style-panel-collapsed .selected-shape-actions {
        display: none !important;
      }

      .writing-bridge-style-toggle {
        position: fixed;
        z-index: 1400;
        appearance: none;
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 10px;
        width: 22px;
        height: 42px;
        padding: 0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font: 700 15px/1 "Segoe UI", sans-serif;
        color: #f6f7fb;
        background: rgba(24, 24, 28, 0.72);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(12px);
        cursor: pointer;
        touch-action: manipulation;
        user-select: none;
        pointer-events: auto;
      }

      .writing-bridge-style-toggle:hover {
        background: rgba(36, 36, 42, 0.92);
      }

      .writing-bridge-style-toggle:active {
        transform: translateY(1px);
      }

      .writing-bridge-style-toggle.is-open {
        background: #2f7cf6;
        border-color: #2f7cf6;
        color: white;
      }
    `;
    document.head.appendChild(styleEl);
    this.register(() => styleEl.remove());

    this.syncStylePanelToggles();
    this.registerInterval(window.setInterval(() => this.syncStylePanelToggles(), STYLE_PANEL_SYNC_MS));
    this.registerEvent(this.app.workspace.on("layout-change", () => this.syncStylePanelToggles()));
    this.register(() => this.destroyStylePanelToggles());
  }

  installRuledPageOverlay() {
    const styleEl = document.createElement("style");
    styleEl.setAttribute("data-writing-bridge", "ruled-page");
    styleEl.textContent = `
      .writing-bridge-ruled-page {
        position: absolute;
        left: ${-RULED_PAGE_EXTENT_PX}px;
        top: ${-RULED_PAGE_EXTENT_PX}px;
        width: ${RULED_PAGE_EXTENT_PX * 2}px;
        height: ${RULED_PAGE_EXTENT_PX * 2}px;
        pointer-events: none;
        z-index: 0;
        --writing-bridge-ruled-visibility: 0;
        --writing-bridge-ruled-density-opacity: 1;
        opacity: calc(
          var(--writing-bridge-ruled-visibility) * var(--writing-bridge-ruled-density-opacity)
        );
        background-repeat: repeat;
        background-image: repeating-linear-gradient(
          to bottom,
          var(--writing-bridge-ruled-line) 0 var(--writing-bridge-ruled-line-thickness, 1px),
          transparent var(--writing-bridge-ruled-line-thickness, 1px) var(--writing-bridge-ruled-step, 40px)
        );
        transition: opacity 120ms ease;
      }

      .theme-light .writing-bridge-ruled-page,
      .theme-light .writing-bridge-excalidraw-ruled-page,
      body:not(.theme-dark) .writing-bridge-ruled-page,
      body:not(.theme-dark) .writing-bridge-excalidraw-ruled-page {
        --writing-bridge-ruled-line: rgba(74, 108, 184, 0.48);
      }

      .theme-dark .writing-bridge-ruled-page,
      .theme-dark .writing-bridge-excalidraw-ruled-page {
        --writing-bridge-ruled-line: rgba(158, 190, 255, 0.44);
      }

      .writing-bridge-ruled-page.is-visible {
        --writing-bridge-ruled-visibility: 1;
      }

      .writing-bridge-excalidraw-ruled-page {
        position: absolute;
        inset: 0;
        pointer-events: none;
        z-index: var(--zIndex-svgLayer, 3);
        --writing-bridge-ruled-visibility: 0;
        --writing-bridge-ruled-density-opacity: 1;
        opacity: calc(
          var(--writing-bridge-ruled-visibility) * var(--writing-bridge-ruled-density-opacity)
        );
        background-repeat: repeat;
        background-image: repeating-linear-gradient(
          to bottom,
          var(--writing-bridge-ruled-line) 0 var(--writing-bridge-ruled-line-thickness, 1px),
          transparent var(--writing-bridge-ruled-line-thickness, 1px) var(--writing-bridge-ruled-step, 40px)
        );
        transition: opacity 120ms ease;
      }

      .writing-bridge-excalidraw-ruled-page.is-visible {
        --writing-bridge-ruled-visibility: 1;
      }


      .writing-bridge-ruled-toggle {
        position: fixed;
        z-index: 1400;
        appearance: none;
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 999px;
        min-width: 54px;
        height: 28px;
        padding: 0 10px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font: 600 12px/1 "Segoe UI", sans-serif;
        color: #f6f7fb;
        background: rgba(24, 24, 28, 0.72);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.2);
        backdrop-filter: blur(12px);
        cursor: pointer;
        touch-action: manipulation;
        user-select: none;
      }

      .writing-bridge-ruled-toggle:hover {
        background: rgba(36, 36, 42, 0.92);
      }

      .writing-bridge-ruled-toggle.is-on {
        background: rgba(47, 124, 246, 0.92);
        border-color: rgba(47, 124, 246, 1);
        color: white;
      }

      .writing-bridge-ruled-toggle.is-off {
        opacity: 0.72;
      }

      .writing-bridge-ruled-toggle[data-target-kind="excalidraw"] {
        position: absolute;
        z-index: 24;
      }

      .writing-bridge-excalidraw-ruled-host {
        position: relative !important;
        overflow: hidden;
      }

      /* Keep this cheap.
         If the note already has an exported Excalidraw image, CSS can hang the
         edit affordance there without us stitching extra DOM into the page. */
      .writing-bridge-excalidraw-embed-actions {
        margin-top: 6px;
        margin-bottom: 4px;
        display: flex;
        justify-content: flex-start;
        align-items: center;
        gap: 8px;
        font: 500 12px/1.35 "Segoe UI", sans-serif;
        color: var(--text-muted);
      }

      .writing-bridge-excalidraw-embed-label {
        color: var(--text-faint);
        letter-spacing: 0.01em;
      }

      .writing-bridge-excalidraw-embed-edit-link {
        position: static;
        appearance: none;
        border: 0;
        padding: 0;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font: 600 12px/1.35 "Segoe UI", sans-serif;
        color: var(--text-accent);
        background: transparent;
        text-decoration: underline;
        text-underline-offset: 2px;
        cursor: pointer;
        touch-action: manipulation;
        user-select: none;
        pointer-events: auto;
      }

      .writing-bridge-excalidraw-embed-edit-link:hover {
        color: var(--text-accent-hover, var(--text-accent));
      }

      .writing-bridge-excalidraw-embed-edit-link:active {
        color: var(--text-normal);
      }
    `;
    document.head.appendChild(styleEl);
    this.register(() => styleEl.remove());

    this.syncRuledPageOverlays();
    this.syncRuledPageToggles();
    this.clearExcalidrawThemeMatching();
    this.syncExcalidrawDefaultZoom();
    this.registerInterval(window.setInterval(() => this.syncRuledPageOverlays(), STYLE_PANEL_SYNC_MS));
    this.registerInterval(window.setInterval(() => this.syncRuledPageToggles(), STYLE_PANEL_SYNC_MS));
    this.registerInterval(window.setInterval(() => this.syncExcalidrawDefaultZoom(), STYLE_PANEL_SYNC_MS));
    this.registerDomEvent(document, "scroll", () => this.scheduleExcalidrawFloatingUiSync(), {
      capture: true,
      passive: true,
    });
    this.registerDomEvent(document, "wheel", () => this.beginExcalidrawFloatingUiTracking(), {
      capture: true,
      passive: true,
    });
    this.registerDomEvent(document, "touchmove", () => this.beginExcalidrawFloatingUiTracking(), {
      capture: true,
      passive: true,
    });
    this.registerDomEvent(
      document,
      "pointerdown",
      (event) => this.handleExcalidrawViewportInteraction(event, 360),
      {
        capture: true,
        passive: true,
      }
    );
    this.registerDomEvent(
      document,
      "pointermove",
      (event) => this.handleExcalidrawViewportInteraction(event, 360),
      {
        capture: true,
        passive: true,
      }
    );
    this.registerDomEvent(
      document,
      "pointerup",
      (event) => this.handleExcalidrawViewportInteraction(event, 180),
      {
        capture: true,
        passive: true,
      }
    );
    this.registerDomEvent(window, "resize", () => this.scheduleExcalidrawFloatingUiSync(), {
      passive: true,
    });
    this.registerEvent(this.app.workspace.on("layout-change", () => {
      this.syncRuledPageOverlays();
      this.syncRuledPageToggles();
      this.syncExcalidrawDefaultZoom();
    }));
    this.registerEvent(this.app.workspace.on("active-leaf-change", (leaf) => {
      if (leaf?.view?.getViewType?.() !== EXCALIDRAW_VIEW_TYPE) {
        return;
      }

      this.scheduleExcalidrawFloatingUiSync();
      this.applyPreferredExcalidrawZoom(leaf.view);
    }));
    this.register(() => {
      if (this.pendingExcalidrawFloatingUiSyncFrame) {
        window.cancelAnimationFrame(this.pendingExcalidrawFloatingUiSyncFrame);
        this.pendingExcalidrawFloatingUiSyncFrame = 0;
      }
    });
    this.register(() => this.destroyRuledPageUi());
  }

  installSidebarPenAssist() {
    const styleEl = document.createElement("style");
    styleEl.setAttribute("data-writing-bridge", "sidebar-pen-assist");
    styleEl.textContent = `
      .workspace-ribbon .clickable-icon,
      .workspace-ribbon button,
      .workspace-tab-header,
      .workspace-tab-header-inner,
      .workspace-split.mod-left-split button,
      .workspace-split.mod-right-split button,
      .workspace-split.mod-left-split [role="button"],
      .workspace-split.mod-right-split [role="button"],
      .workspace-leaf-content[data-type="calculite"] button {
        touch-action: manipulation;
      }
    `;
    document.head.appendChild(styleEl);
    this.register(() => styleEl.remove());

    this.registerDomEvent(document, "pointerdown", (event) => this.handleSidebarPenPointerDown(event), {
      capture: true,
      passive: true,
    });
    this.registerDomEvent(document, "pointerup", (event) => this.handleSidebarPenPointerUp(event), {
      capture: true,
      passive: true,
    });
    this.registerDomEvent(document, "pointercancel", (event) => this.handleSidebarPenPointerCancel(event), {
      capture: true,
      passive: true,
    });
    this.registerDomEvent(document, "click", (event) => this.handleSidebarPenNativeClick(event), {
      capture: true,
      passive: true,
    });
  }

  handleSidebarPenPointerDown(event) {
    if (!this.isFeatureEnabled("sidebarPenAssistEnabled")) {
      this.sidebarPenTapState = null;
      return;
    }

    if (!this.isSidebarPenPointerEvent(event)) {
      return;
    }

    const target = this.getSidebarPenInteractiveTarget(event.target);
    if (!(target instanceof HTMLElement)) {
      this.sidebarPenTapState = null;
      return;
    }

    this.sidebarPenTapState = {
      pointerId: event.pointerId,
      target,
      startedAt: Date.now(),
      clientX: Number(event.clientX ?? 0),
      clientY: Number(event.clientY ?? 0),
      nativeClickSeen: false,
    };
  }

  handleSidebarPenPointerUp(event) {
    if (!this.isFeatureEnabled("sidebarPenAssistEnabled")) {
      return;
    }

    if (!this.isSidebarPenPointerEvent(event)) {
      return;
    }

    const state = this.sidebarPenTapState;
    if (!state || state.pointerId !== event.pointerId) {
      return;
    }

    const target =
      this.getSidebarPenInteractiveTarget(event.target) ??
      this.getSidebarPenInteractiveTarget(document.elementFromPoint(event.clientX, event.clientY));
    if (!(target instanceof HTMLElement) || !this.isSameSidebarPenTarget(state.target, target)) {
      if (this.sidebarPenTapState === state) {
        this.sidebarPenTapState = null;
      }
      return;
    }

    const travelPx = Math.hypot(
      Number(event.clientX ?? 0) - state.clientX,
      Number(event.clientY ?? 0) - state.clientY
    );
    const durationMs = Date.now() - state.startedAt;
    if (travelPx > SIDEBAR_PEN_TAP_MAX_TRAVEL_PX || durationMs > SIDEBAR_PEN_TAP_MAX_DURATION_MS) {
      if (this.sidebarPenTapState === state) {
        this.sidebarPenTapState = null;
      }
      return;
    }

    const replayEventState = {
      clientX: Number(event.clientX ?? 0),
      clientY: Number(event.clientY ?? 0),
      ctrlKey: Boolean(event.ctrlKey),
      altKey: Boolean(event.altKey),
      shiftKey: Boolean(event.shiftKey),
      metaKey: Boolean(event.metaKey),
    };
    const timeoutId = window.setTimeout(() => {
      this.pendingSidebarPenReplayTimeouts.delete(timeoutId);
      if (this.sidebarPenTapState === state) {
        this.sidebarPenTapState = null;
      }
      if (state.nativeClickSeen) {
        return;
      }

      this.replaySidebarPenTap(state.target, replayEventState);
    }, SIDEBAR_PEN_REPLAY_DELAY_MS);
    this.pendingSidebarPenReplayTimeouts.add(timeoutId);
  }

  handleSidebarPenPointerCancel(event) {
    if (!this.isFeatureEnabled("sidebarPenAssistEnabled")) {
      this.sidebarPenTapState = null;
      return;
    }

    if (this.sidebarPenTapState?.pointerId === event.pointerId) {
      this.sidebarPenTapState = null;
    }
  }

  handleSidebarPenNativeClick(event) {
    if (!this.isFeatureEnabled("sidebarPenAssistEnabled")) {
      return;
    }

    const state = this.sidebarPenTapState;
    if (!state?.target || !(event.target instanceof Node)) {
      return;
    }

    if (this.isSameSidebarPenTarget(state.target, event.target)) {
      state.nativeClickSeen = true;
    }
  }

  isSidebarPenPointerEvent(event) {
    return (
      event instanceof PointerEvent &&
      (event.pointerType === "pen" || event.pointerType === "touch") &&
      event.isPrimary !== false &&
      Number(event.button ?? 0) === 0
    );
  }

  getSidebarPenInteractiveTarget(target) {
    if (!(target instanceof Element)) {
      return null;
    }

    if (
      target.closest(
        [
          ".tldraw-view-root",
          ".tl-container",
          ".excalidraw-wrapper",
          ".excalidraw-view",
          ".cm-editor",
          ".modal",
          ".menu",
          ".suggestion-container",
        ].join(", ")
      )
    ) {
      return null;
    }

    const sidebarRegion = target.closest(
      [
        ".workspace-ribbon",
        ".workspace-side-dock",
        ".workspace-split.mod-left-split",
        ".workspace-split.mod-right-split",
        ".workspace-drawer",
      ].join(", ")
    );
    if (!(sidebarRegion instanceof HTMLElement)) {
      return null;
    }

    if (target.closest('input, textarea, select, [contenteditable="true"], [contenteditable=""]')) {
      return null;
    }

    const selectors = [
      '.workspace-leaf-content[data-type="calculite"] button',
      ".workspace-tab-header-status-icon",
      ".workspace-tab-header",
      ".workspace-tab-header-inner",
      ".workspace-tab-container .clickable-icon",
      ".workspace-ribbon-collapse-btn",
      ".workspace-ribbon-tab",
      ".side-dock-ribbon-action",
      ".clickable-icon",
      ".view-action",
      ".nav-action-button",
      ".tree-item-icon",
      ".tree-item-inner",
      ".tree-item-self",
      "a",
      '[role="button"]',
      "button",
    ];

    for (const selector of selectors) {
      const candidate = target.closest(selector);
      if (candidate instanceof HTMLElement && sidebarRegion.contains(candidate)) {
        return candidate;
      }
    }

    return null;
  }

  isSameSidebarPenTarget(expectedTarget, actualTarget) {
    return (
      expectedTarget instanceof Node &&
      actualTarget instanceof Node &&
      (expectedTarget === actualTarget ||
        expectedTarget.contains(actualTarget) ||
        actualTarget.contains(expectedTarget))
    );
  }

  replaySidebarPenTap(target, sourceEvent) {
    if (!(target instanceof HTMLElement) || !target.isConnected) {
      return;
    }

    const eventInit = {
      bubbles: true,
      cancelable: true,
      composed: true,
      clientX: Number(sourceEvent?.clientX ?? 0),
      clientY: Number(sourceEvent?.clientY ?? 0),
      ctrlKey: Boolean(sourceEvent?.ctrlKey),
      altKey: Boolean(sourceEvent?.altKey),
      shiftKey: Boolean(sourceEvent?.shiftKey),
      metaKey: Boolean(sourceEvent?.metaKey),
      button: 0,
    };

    target.focus?.({ preventScroll: true });
    this.dispatchSidebarPenSyntheticPointerEvent(target, "pointerdown", {
      ...eventInit,
      buttons: 1,
      pointerId: 1,
      pointerType: "mouse",
      isPrimary: true,
    });
    target.dispatchEvent(new MouseEvent("mousedown", { ...eventInit, buttons: 1 }));
    this.dispatchSidebarPenSyntheticPointerEvent(target, "pointerup", {
      ...eventInit,
      buttons: 0,
      pointerId: 1,
      pointerType: "mouse",
      isPrimary: true,
    });
    target.dispatchEvent(new MouseEvent("mouseup", { ...eventInit, buttons: 0 }));

    if (typeof target.click === "function") {
      target.click();
    } else {
      target.dispatchEvent(new MouseEvent("click", { ...eventInit, buttons: 0 }));
    }
  }

  dispatchSidebarPenSyntheticPointerEvent(target, type, eventInit) {
    if (!(target instanceof HTMLElement) || typeof PointerEvent === "undefined") {
      return;
    }

    target.dispatchEvent(new PointerEvent(type, eventInit));
  }

  syncRuledPageOverlays() {
    const roots = document.querySelectorAll(".tldraw-view-root");
    for (const root of roots) {
      this.syncRuledPageOverlay(root);
    }

    this.syncExcalidrawRuledPageOverlays();
    this.syncExcalidrawThemeMatching();
  }

  syncRuledPageOverlay(root) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    const target = this.getTldrawTarget();
    const targetForRoot = this.getTldrawTargetForRoot(root, target);
    this.applyPreferredTldrawDrawSize(root, targetForRoot);

    const canvas = root.querySelector(".tl-canvas");
    const existing = root.querySelector(".writing-bridge-ruled-page");
    if (
      !(canvas instanceof HTMLElement) ||
      !this.isFeatureEnabled("ruledPageFeatureEnabled") ||
      !this.settings.ruledPageEnabled
    ) {
      if (existing instanceof HTMLElement) {
        existing.remove();
      }
      return;
    }

    const shapesLayer = canvas.querySelector(".tl-html-layer.tl-shapes");
    const backgroundWrapper = canvas.querySelector(".tl-background__wrapper");
    const overlay =
      existing instanceof HTMLDivElement ? existing : this.createRuledPageOverlay();
    const host =
      shapesLayer instanceof HTMLElement
        ? shapesLayer
        : backgroundWrapper instanceof HTMLElement
          ? backgroundWrapper
          : canvas;
    if (!overlay.isConnected) {
      host.prepend(overlay);
    } else if (overlay.parentElement !== host) {
      overlay.remove();
      host.prepend(overlay);
    }

    const rootId = this.ensureStyleToggleRootId(root);
    const nextState = this.getRuledPageOverlayState(targetForRoot);
    const overlayState = nextState ?? this.ruledPageStateByRootId.get(rootId) ?? {
      stepPx: 40,
      offsetPx: 0,
      lineThicknessPx: 1,
      visible: true,
    };

    if (nextState) {
      this.ruledPageStateByRootId.set(rootId, overlayState);
    }

    this.applyRuledPageOverlayState(overlay, overlayState);
  }

  createRuledPageOverlay() {
    const overlay = document.createElement("div");
    overlay.className = "writing-bridge-ruled-page";
    overlay.setAttribute("aria-hidden", "true");
    return overlay;
  }

  applyRuledPageOverlayState(overlay, { stepPx, offsetPx, lineThicknessPx, visible, opacity }) {
    overlay.style.setProperty("--writing-bridge-ruled-step", `${stepPx}px`);
    overlay.style.setProperty(
      "--writing-bridge-ruled-line-thickness",
      `${lineThicknessPx ?? 1}px`
    );
    overlay.style.setProperty(
      "--writing-bridge-ruled-density-opacity",
      `${Math.max(0, Math.min(1, Number(opacity ?? 1) || 1))}`
    );
    overlay.style.backgroundPosition = `0 ${offsetPx}px`;
    overlay.classList.toggle("is-visible", Boolean(visible));
  }

  getRuledPageOpacityForScreenStep(screenStepPx) {
    const safeStepPx = Number(screenStepPx);
    if (!Number.isFinite(safeStepPx) || safeStepPx >= RULED_PAGE_FULL_OPACITY_STEP_PX) {
      return 1;
    }

    if (safeStepPx <= RULED_PAGE_MIN_SCREEN_STEP_PX) {
      return RULED_PAGE_MIN_OPACITY;
    }

    const progress =
      (safeStepPx - RULED_PAGE_MIN_SCREEN_STEP_PX) /
      (RULED_PAGE_FULL_OPACITY_STEP_PX - RULED_PAGE_MIN_SCREEN_STEP_PX);
    return RULED_PAGE_MIN_OPACITY + progress * (1 - RULED_PAGE_MIN_OPACITY);
  }

  async toggleRuledPageEnabled() {
    if (!this.isFeatureEnabled("ruledPageFeatureEnabled")) {
      this.showFeatureDisabledNotice("Ruled page overlay");
      return;
    }

    this.settings.ruledPageEnabled = !this.settings.ruledPageEnabled;
    await this.saveSettings();
    this.syncRuledPageOverlays();
    this.syncRuledPageToggles();
    new Notice(this.settings.ruledPageEnabled ? "Ruled page enabled" : "Ruled page disabled");
  }

  destroyRuledPageDecorations() {
    for (const overlay of document.querySelectorAll(".writing-bridge-ruled-page")) {
      if (overlay instanceof HTMLElement) {
        overlay.remove();
      }
    }

    for (const overlay of document.querySelectorAll(".writing-bridge-excalidraw-ruled-page")) {
      if (overlay instanceof HTMLElement) {
        overlay.remove();
      }
    }

    for (const button of document.querySelectorAll(".writing-bridge-ruled-toggle")) {
      if (button instanceof HTMLElement) {
        button.remove();
      }
    }

    this.ruledPageStateByRootId.clear();
  }

  destroyRuledPageUi() {
    for (const overlay of document.querySelectorAll(".writing-bridge-ruled-page")) {
      if (overlay instanceof HTMLElement) {
        overlay.remove();
      }
    }

    for (const overlay of document.querySelectorAll(".writing-bridge-excalidraw-ruled-page")) {
      if (overlay instanceof HTMLElement) {
        overlay.remove();
      }
    }

    for (const button of document.querySelectorAll(".writing-bridge-ruled-toggle")) {
      if (button instanceof HTMLElement) {
        button.remove();
      }
    }

    this.appliedTldrawDrawSizeByRootId.clear();

    for (const el of document.querySelectorAll(".writing-bridge-edit-shadow-host")) {
      if (el instanceof HTMLElement) {
        el.remove();
      }
    }

    for (const el of document.querySelectorAll(".writing-bridge-excalidraw-embed-actions")) {
      if (el instanceof HTMLElement) {
        el.remove();
      }
    }

    this.ruledPageStateByRootId.clear();
    this.clearExcalidrawThemeMatching();
  }

  syncRuledPageToggles() {
    if (!this.isFeatureEnabled("ruledPageFeatureEnabled")) {
      this.destroyRuledPageDecorations();
      return;
    }

    const activeRoot = this.getTldrawTargetRoot(
      this.normalizeTldrawTarget(this.getTldrawTarget())
    );
    const activeRootId =
      activeRoot instanceof HTMLElement ? this.ensureStyleToggleRootId(activeRoot) : "";
    const roots = document.querySelectorAll(".tldraw-view-root");
    for (const root of roots) {
      this.syncRuledPageToggle(root, activeRoot);
    }

    for (const button of document.querySelectorAll(".writing-bridge-ruled-toggle")) {
      if (!(button instanceof HTMLElement)) {
        continue;
      }

      if (!activeRootId || button.dataset.rootId !== activeRootId) {
        button.remove();
      }
    }

    this.syncExcalidrawRuledPageToggles();
  }

  syncRuledPageToggle(root, activeRoot = null) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    const canvas = root.querySelector(".tl-canvas");
    const rootId = this.ensureStyleToggleRootId(root);
    const existing = document.querySelector(
      `.writing-bridge-ruled-toggle[data-root-id="${rootId}"]`
    );

    if (!(canvas instanceof HTMLElement) || activeRoot !== root) {
      if (existing instanceof HTMLElement) {
        existing.remove();
      }
      return;
    }

    const button =
      existing instanceof HTMLButtonElement ? existing : this.createRuledPageToggleButton();

    if (!button.dataset.rootId) {
      button.dataset.rootId = rootId;
    }

    if (!button.isConnected) {
      document.body.appendChild(button);
    }

    const enabled = Boolean(this.settings.ruledPageEnabled);
    button.classList.toggle("is-on", enabled);
    button.classList.toggle("is-off", !enabled);
    button.textContent = "Lines";
    button.setAttribute("aria-label", enabled ? "Hide ruled lines" : "Show ruled lines");
    button.setAttribute("aria-pressed", String(enabled));
    button.title = enabled ? "Hide ruled lines" : "Show ruled lines";
    this.positionRuledPageToggle(button, root);
  }

  createRuledPageToggleButton() {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "writing-bridge-ruled-toggle";
    button.addEventListener("pointerup", (event) => {
      event.stopPropagation();
      event.preventDefault();
      void this.toggleRuledPageEnabled();
    });
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
    return button;
  }

  positionRuledPageToggle(button, root) {
    if (!(button instanceof HTMLElement) || !(root instanceof HTMLElement)) {
      return;
    }

    const rootRect = root.getBoundingClientRect();
    if (rootRect.width <= 0 || rootRect.height <= 0) {
      return;
    }

    const margin = 10;
    const top = Math.max(margin, Math.round(rootRect.top + 102));
    const left = Math.max(
      margin,
      Math.min(
        Math.round(rootRect.right - button.offsetWidth - margin),
        window.innerWidth - button.offsetWidth - margin
      )
    );

    button.style.top = `${top}px`;
    button.style.left = `${left}px`;
  }

  syncExcalidrawRuledPageOverlays() {
    const roots = document.querySelectorAll('.workspace-leaf-content[data-type="excalidraw"]');
    for (const root of roots) {
      this.syncExcalidrawRuledPageOverlay(root);
    }
  }

  // Theme sync removed — Excalidraw handles its own theme now.
  // This stub only cleans up stale classes from older sessions.
  syncExcalidrawThemeMatching() {
    this.clearExcalidrawThemeMatching();
  }

  syncExcalidrawFloatingUi() {
    this.syncExcalidrawRuledPageOverlays();
    this.syncExcalidrawRuledPageToggles();
  }

  scheduleExcalidrawFloatingUiSync() {
    if (this.pendingExcalidrawFloatingUiSyncFrame) {
      return;
    }

    this.pendingExcalidrawFloatingUiSyncFrame = window.requestAnimationFrame(() => {
      this.pendingExcalidrawFloatingUiSyncFrame = 0;
      this.syncExcalidrawFloatingUi();
    });
  }

  beginExcalidrawFloatingUiTracking(durationMs = 280) {
    const now = typeof performance !== "undefined" ? performance.now() : Date.now();
    this.excalidrawFloatingUiTrackingUntil = Math.max(
      this.excalidrawFloatingUiTrackingUntil,
      now + durationMs
    );

    const tick = () => {
      this.pendingExcalidrawFloatingUiSyncFrame = 0;
      this.syncExcalidrawFloatingUi();

      const currentTime = typeof performance !== "undefined" ? performance.now() : Date.now();
      if (currentTime < this.excalidrawFloatingUiTrackingUntil) {
        this.pendingExcalidrawFloatingUiSyncFrame = window.requestAnimationFrame(tick);
      }
    };

    if (!this.pendingExcalidrawFloatingUiSyncFrame) {
      this.pendingExcalidrawFloatingUiSyncFrame = window.requestAnimationFrame(tick);
    }
  }

  handleExcalidrawViewportInteraction(event, durationMs = 320) {
    const target = event?.target;
    if (!(target instanceof Element)) {
      return;
    }

    if (!target.closest('.workspace-leaf-content[data-type="excalidraw"], .excalidraw-view')) {
      return;
    }

    if (
      typeof PointerEvent !== "undefined" &&
      event instanceof PointerEvent &&
      event.type === "pointermove" &&
      Number(event.buttons ?? 0) === 0
    ) {
      return;
    }

    this.beginExcalidrawFloatingUiTracking(durationMs);
  }

  queueExcalidrawThemeSyncBurst(retries = 10, delayMs = 70) {
    const totalRetries = Math.max(1, Number.isFinite(retries) ? retries : 1);
    const nextDelayMs = Math.max(0, Number.isFinite(delayMs) ? delayMs : 0);

    for (let attempt = 0; attempt < totalRetries; attempt += 1) {
      const timeoutId = window.setTimeout(() => {
        this.pendingExcalidrawThemeSyncTimeouts.delete(timeoutId);
        this.syncExcalidrawThemeMatching();
      }, attempt * nextDelayMs);
      this.pendingExcalidrawThemeSyncTimeouts.add(timeoutId);
    }
  }

  syncExcalidrawThemeForRoot(root, desiredTheme = this.getObsidianThemeMode()) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    // Bail out early when the feature is off – clean up any stale classes
    // that a previous session might have left behind.
    if (!this.isFeatureEnabled("excalidrawThemeSyncEnabled")) {
      this.applyExcalidrawThemeClasses(root, desiredTheme, false);
      return;
    }

    const rootId = this.ensureStyleToggleRootId(root);
    const wrapper = this.findExcalidrawContainer(root);
    const host = this.findExcalidrawRuledPageHost(root);
    const backgroundColor = this.getObsidianBackgroundColor(root);
    const shouldMatch = Boolean(this.settings.matchExcalidrawThemeToObsidian);
    this.applyExcalidrawThemeClasses(root, desiredTheme, shouldMatch);

    if (root instanceof HTMLElement) {
      root.classList.toggle("writing-bridge-excalidraw-theme-match", shouldMatch);
      if (shouldMatch) {
        root.style.setProperty("--writing-bridge-excalidraw-bg", backgroundColor);
      } else {
        root.style.removeProperty("--writing-bridge-excalidraw-bg");
      }
    }

    if (wrapper instanceof HTMLElement) {
      wrapper.classList.toggle("writing-bridge-excalidraw-theme-match", shouldMatch);
      if (shouldMatch) {
        wrapper.style.setProperty("--writing-bridge-excalidraw-bg", backgroundColor);
      } else {
        wrapper.style.removeProperty("--writing-bridge-excalidraw-bg");
      }
    }

    if (host instanceof HTMLElement) {
      host.classList.toggle("writing-bridge-excalidraw-theme-match", shouldMatch);
      if (shouldMatch) {
        host.style.setProperty("--writing-bridge-excalidraw-bg", backgroundColor);
      } else {
        host.style.removeProperty("--writing-bridge-excalidraw-bg");
      }
    }

    // Let the Excalidraw startup script own updateScene.
    // Calling it from here was one of those "looks fine until it wedges the
    // loading scene" bugs, so this layer only nudges the shell styling.
  }

  syncExcalidrawEmbedThemeMatching() {
    const hosts = document.querySelectorAll(".excalidraw-md-host");
    const shouldMatch = Boolean(this.settings.matchExcalidrawThemeToObsidian);
    const backgroundColor = this.getObsidianBackgroundColor(document.body);
    for (const host of hosts) {
      if (!(host instanceof HTMLElement)) {
        continue;
      }

      host.classList.toggle("writing-bridge-excalidraw-theme-match", shouldMatch);
      if (shouldMatch) {
        host.style.setProperty("--writing-bridge-excalidraw-bg", backgroundColor);
      } else {
        host.style.removeProperty("--writing-bridge-excalidraw-bg");
      }
    }
  }

  syncExcalidrawExportImageEmbeds(desiredTheme = this.getObsidianThemeMode()) {
    if (!this.settings.matchExcalidrawThemeToObsidian) {
      return;
    }

    const selectors = [
      'img[src*=".excalidraw."]',
      'img[srcset*=".excalidraw."]',
      'img[src*=".excalidraw-"]',
      'img[src*=".excalidraw"]',
      'source[srcset*=".excalidraw."]',
      'image[href*=".excalidraw."]',
      'image[xlink\\:href*=".excalidraw."]',
    ];

    for (const element of document.querySelectorAll(selectors.join(","))) {
      this.syncExcalidrawExportImageElement(element, desiredTheme);
    }
  }

  syncExcalidrawExportImageElement(element, desiredTheme) {
    if (!(element instanceof Element)) {
      return;
    }

    const attrName = this.getExcalidrawExportImageAttrName(element);

    if (!attrName) {
      return;
    }

    const currentUrl = element.getAttribute(attrName)?.trim() ?? "";
    if (!currentUrl) {
      return;
    }

    const themedUrl = this.getThemedExcalidrawExportUrl(currentUrl, desiredTheme);
    if (!themedUrl || themedUrl === currentUrl) {
      return;
    }

    element.setAttribute(attrName, themedUrl);
    if (element instanceof HTMLImageElement && element.srcset) {
      element.srcset = this.getThemedExcalidrawExportUrl(element.srcset, desiredTheme) ?? element.srcset;
    }
  }

  getExcalidrawExportImageAttrName(element) {
    if (!(element instanceof Element)) {
      return "";
    }

    const candidateAttrs =
      element instanceof HTMLImageElement
        ? ["src", "srcset"]
        : ["srcset", "href", "xlink:href"];

    for (const attrName of candidateAttrs) {
      const value = element.getAttribute(attrName)?.trim() ?? "";
      if (value.includes(".excalidraw")) {
        return attrName;
      }
    }

    return "";
  }

  getThemedExcalidrawExportUrl(url, desiredTheme) {
    if (typeof url !== "string" || !url.includes(".excalidraw")) {
      return null;
    }

    const visualTheme = desiredTheme === "dark" ? "dark" : "light";
    const invertedVariants = this.settings.invertExcalidrawExportThemeVariants !== false;
    const variantTheme =
      visualTheme === "dark"
        ? invertedVariants
          ? "light"
          : "dark"
        : invertedVariants
          ? "dark"
          : "light";
    const otherVariantTheme = variantTheme === "dark" ? "light" : "dark";
    const themedVariantPattern = new RegExp(`\\.excalidraw\\.${otherVariantTheme}\\.(svg|png)`, "gi");
    if (themedVariantPattern.test(url)) {
      return url.replace(themedVariantPattern, `.excalidraw.${variantTheme}.$1`);
    }

    const unthemedVariantPattern = /\.excalidraw\.(svg|png)/gi;
    if (unthemedVariantPattern.test(url)) {
      return url.replace(unthemedVariantPattern, `.excalidraw.${variantTheme}.$1`);
    }

    return null;
  }

  clearExcalidrawThemeMatching() {
    for (const element of document.querySelectorAll(".writing-bridge-excalidraw-theme-match")) {
      if (!(element instanceof HTMLElement)) {
        continue;
      }

      element.classList.remove("writing-bridge-excalidraw-theme-match");
      element.style.removeProperty("--writing-bridge-excalidraw-bg");
    }

    for (const element of document.querySelectorAll(".writing-bridge-excalidraw-force-theme")) {
      if (!(element instanceof HTMLElement)) {
        continue;
      }

      element.classList.remove("writing-bridge-excalidraw-force-theme", "theme--dark", "theme--light");
      element.style.removeProperty("color-scheme");
      element.removeAttribute("data-writing-bridge-theme");
    }

  }

  applyExcalidrawThemeClasses(root, desiredTheme, shouldMatch) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    const wrapper = this.findExcalidrawContainer(root);
    const host = this.findExcalidrawRuledPageHost(root);
    // Only target the bridge's own shell elements – the workspace-leaf root,
    // the excalidraw-wrapper, and the ruled-page host.  Do NOT reach into
    // Excalidraw's internal React-managed elements (.excalidraw,
    // .excalidraw-container, .layer-ui__wrapper, etc.) because forcing
    // theme--dark / color-scheme on those elements fights with Excalidraw's
    // own theme state and blacks out the canvas.
    const targets = new Set();
    const addTarget = (element) => {
      if (element instanceof HTMLElement) {
        targets.add(element);
      }
    };

    addTarget(root);
    addTarget(wrapper);
    addTarget(host);

    const themeClass = desiredTheme === "dark" ? "theme--dark" : "theme--light";
    const oppositeThemeClass = desiredTheme === "dark" ? "theme--light" : "theme--dark";

    for (const element of targets) {
      if (shouldMatch) {
        element.classList.remove(oppositeThemeClass);
        element.classList.add("writing-bridge-excalidraw-force-theme", themeClass);
        element.dataset.writingBridgeTheme = desiredTheme;
      } else {
        element.classList.remove(
          "writing-bridge-excalidraw-force-theme",
          themeClass,
          oppositeThemeClass
        );
        element.removeAttribute("data-writing-bridge-theme");
      }
    }
  }

  syncExcalidrawEmbedEditButtons() {
    for (const shadowHost of document.querySelectorAll(".writing-bridge-edit-shadow-host")) {
      if (shadowHost instanceof HTMLElement) {
        shadowHost.remove();
      }
    }

    for (const overlayRoot of document.querySelectorAll(".writing-bridge-excalidraw-embed-overlay-root")) {
      if (overlayRoot instanceof HTMLElement) {
        overlayRoot.remove();
      }
    }

    const candidates = this.getExcalidrawEmbedEditCandidates();
    const activeAnchorIds = new Set(candidates.map((candidate) => candidate.anchorId));

    for (const row of document.querySelectorAll(".writing-bridge-excalidraw-embed-actions")) {
      if (
        row instanceof HTMLElement &&
        (!row.dataset.anchorId || !activeAnchorIds.has(row.dataset.anchorId))
      ) {
        row.remove();
      }
    }

    if (!this.isFeatureEnabled("excalidrawThemeSyncEnabled")) {
      return;
    }

    for (const candidate of candidates) {
      this.ensureExcalidrawEmbedActionRow(candidate.host, candidate.anchorId, candidate.file, candidate.anchor);
    }
  }

  getExcalidrawEmbedEditCandidates() {
    const selectors = [
      'img[src*=".excalidraw."]',
      'img[srcset*=".excalidraw."]',
      'source[srcset*=".excalidraw."]',
      'object[data*=".excalidraw."]',
      'embed[src*=".excalidraw."]',
    ];
    const processedVisualElements = new Set();
    const candidates = [];

    for (const element of document.querySelectorAll(selectors.join(","))) {
      if (!(element instanceof Element) || this.isCanvasSurfaceElement(element)) {
        continue;
      }

      const visualElement = this.getExcalidrawEmbedVisualElement(element) ?? element;
      if (!(visualElement instanceof Element) || processedVisualElements.has(visualElement)) {
        continue;
      }
      processedVisualElements.add(visualElement);

      const file = this.resolveExcalidrawSourceFileFromElement(element);
      if (!(file instanceof TFile)) {
        continue;
      }

      const host = this.getExcalidrawEmbedActionHost(visualElement) ?? this.getExcalidrawEmbedActionHost(element);
      if (!(host instanceof HTMLElement)) {
        continue;
      }

      const anchorId = this.ensureExcalidrawEmbedAnchorId(host);
      const anchor = this.getExcalidrawEmbedActionAnchor(visualElement, host);
      candidates.push({ anchorId, file, host, anchor });
    }

    return candidates;
  }

  ensureExcalidrawEmbedAnchorId(element) {
    if (!(element instanceof HTMLElement)) {
      return "";
    }

    if (typeof element.dataset.tldrawPenBridgeEmbedId === "string" && element.dataset.tldrawPenBridgeEmbedId) {
      return element.dataset.tldrawPenBridgeEmbedId;
    }

    this.excalidrawEmbedButtonId += 1;
    const anchorId = `writing-bridge-embed-${this.excalidrawEmbedButtonId}`;
    element.dataset.tldrawPenBridgeEmbedId = anchorId;
    return anchorId;
  }

  isCanvasSurfaceElement(element) {
    if (!(element instanceof Element)) {
      return false;
    }

    return Boolean(
      element.closest(
        '.workspace-leaf-content[data-type="excalidraw"], .tldraw-view-root, .ptl-tldraw-image'
      )
    );
  }

  getExcalidrawEmbedActionHost(element) {
    if (!(element instanceof Element)) {
      return null;
    }

    const host = element.closest(
      ".internal-embed, .markdown-embed, .image-embed, .el-embed-image, .excalidraw-md-host"
    );
    if (host instanceof HTMLElement && !this.isCanvasSurfaceElement(host)) {
      return host;
    }

    const picture = element.closest("picture");
    if (picture instanceof HTMLElement && !this.isCanvasSurfaceElement(picture)) {
      const pictureParent = picture.parentElement;
      if (pictureParent instanceof HTMLElement && !this.isCanvasSurfaceElement(pictureParent)) {
        return pictureParent;
      }
    }

    const parent = element.parentElement;
    if (parent instanceof HTMLElement && !this.isCanvasSurfaceElement(parent) && parent.tagName !== "PICTURE") {
      return parent;
    }

    const grandparent = parent?.parentElement;
    return grandparent instanceof HTMLElement && !this.isCanvasSurfaceElement(grandparent)
      ? grandparent
      : null;
  }

  getExcalidrawEmbedOverlayHost(element) {
    return document.body instanceof HTMLElement ? document.body : null;
  }

  getExcalidrawEmbedActionAnchor(element, host) {
    if (!(host instanceof HTMLElement)) {
      return null;
    }

    if (element instanceof Element) {
      const picture = element.closest("picture");
      if (picture instanceof HTMLElement && host.contains(picture)) {
        return picture;
      }

      if (element instanceof HTMLElement && host.contains(element)) {
        return element;
      }
    }

    return host;
  }

  ensureExcalidrawEmbedActionRow(host, anchorId, file, anchor = host) {
    if (
      !(host instanceof HTMLElement) ||
      typeof anchorId !== "string" ||
      anchorId === "" ||
      !(file instanceof TFile)
    ) {
      return null;
    }

    const selector = `.writing-bridge-excalidraw-embed-actions[data-anchor-id="${anchorId}"]`;
    const duplicates = Array.from(document.querySelectorAll(selector)).filter((el) => el instanceof HTMLDivElement);
    const existing = duplicates[0] instanceof HTMLDivElement ? duplicates[0] : null;
    const row = existing instanceof HTMLDivElement ? existing : document.createElement("div");
    row.className = "writing-bridge-excalidraw-embed-actions";
    row.dataset.anchorId = anchorId;
    row.dataset.sourcePath = file.path;

    for (const duplicate of duplicates.slice(1)) {
      duplicate.remove();
    }

    let label = row.querySelector(".writing-bridge-excalidraw-embed-label");
    if (!(label instanceof HTMLSpanElement)) {
      label = document.createElement("span");
      label.className = "writing-bridge-excalidraw-embed-label";
      label.textContent = "Excalidraw";
      row.appendChild(label);
    }

    let link = row.querySelector(".writing-bridge-excalidraw-embed-edit-link");
    if (!(link instanceof HTMLAnchorElement)) {
      link = document.createElement("a");
      link.href = "#";
      link.className = "writing-bridge-excalidraw-embed-edit-link";
      link.textContent = "Edit in Excalidraw";
      const stopEvent = (event) => {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation?.();
      };
      const openSource = (event) => {
        stopEvent(event);
        const sourcePath = row.dataset.sourcePath ?? "";
        void this.openExcalidrawSourcePath(sourcePath);
      };
      link.addEventListener("pointerdown", stopEvent, { capture: true });
      link.addEventListener("mousedown", stopEvent, { capture: true });
      link.addEventListener("touchstart", stopEvent, { capture: true, passive: false });
      link.addEventListener("pointerup", openSource, { capture: true });
      link.addEventListener("click", openSource, { capture: true });
      row.appendChild(link);
    }

    const desiredParent =
      anchor instanceof HTMLElement && anchor.parentElement instanceof HTMLElement && host.contains(anchor.parentElement)
        ? anchor.parentElement
        : host;
    const desiredPreviousSibling = anchor instanceof HTMLElement ? anchor : null;

    if (row.parentElement !== desiredParent || row.previousElementSibling !== desiredPreviousSibling) {
      row.remove();
      if (desiredPreviousSibling instanceof HTMLElement && desiredParent.contains(desiredPreviousSibling)) {
        desiredPreviousSibling.insertAdjacentElement("afterend", row);
      } else {
        desiredParent.appendChild(row);
      }
    }

    return row;
  }

  ensureExcalidrawEmbedOverlayRoot(overlayHost) {
    if (!(overlayHost instanceof HTMLElement)) {
      return null;
    }

    const existing = overlayHost.querySelector(".writing-bridge-excalidraw-embed-overlay-root");
    if (existing instanceof HTMLDivElement) {
      return existing;
    }

    const overlayRoot = document.createElement("div");
    overlayRoot.className = "writing-bridge-excalidraw-embed-overlay-root";
    overlayHost.appendChild(overlayRoot);
    return overlayRoot;
  }

  ensureExcalidrawGlobalEditButton(overlayHost, file) {
    if (
      !(overlayHost instanceof HTMLElement) ||
      (!(file instanceof TFile) && file !== null)
    ) {
      return null;
    }

    const overlayRoot = this.ensureExcalidrawEmbedOverlayRoot(overlayHost);
    if (!(overlayRoot instanceof HTMLDivElement)) {
      return null;
    }

    const existing = overlayRoot.querySelector(".writing-bridge-excalidraw-embed-edit-button");
    const button = existing instanceof HTMLButtonElement ? existing : document.createElement("button");
    button.type = "button";
    button.className = "writing-bridge-excalidraw-embed-edit-button";
    button.textContent = "Edit";
    button.dataset.sourcePath = file instanceof TFile ? file.path : "";
    if (!button.isConnected) {
      const stopEvent = (event) => {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation?.();
      };

      button.addEventListener("pointerdown", stopEvent, { capture: true });
      button.addEventListener("mousedown", stopEvent, { capture: true });
      button.addEventListener("touchstart", stopEvent, { capture: true, passive: false });
      button.addEventListener("pointerup", (event) => {
        stopEvent(event);
        const sourcePath = button.dataset.sourcePath ?? "";
        void this.openExcalidrawSourcePath(sourcePath);
      }, { capture: true });
      button.addEventListener("click", stopEvent, { capture: true });
    }

    if (button.parentElement !== overlayRoot) {
      button.remove();
      overlayRoot.appendChild(button);
    }

    return button;
  }

  getExcalidrawEmbedVisualElement(element) {
    if (!(element instanceof Element)) {
      return null;
    }

    if (
      element instanceof HTMLImageElement ||
      element instanceof HTMLEmbedElement ||
      element instanceof HTMLObjectElement
    ) {
      return element;
    }

    if (element instanceof HTMLSourceElement) {
      const picture = element.parentElement;
      const image = picture?.querySelector("img");
      if (image instanceof HTMLImageElement) {
        return image;
      }
      if (picture instanceof HTMLElement) {
        return picture;
      }
    }

    const picture = element.closest("picture");
    if (picture instanceof HTMLElement) {
      const image = picture.querySelector("img");
      if (image instanceof HTMLImageElement) {
        return image;
      }
      return picture;
    }

    return element instanceof HTMLElement ? element : null;
  }

  positionExcalidrawGlobalEditButton(button, candidate) {
    if (
      !(button instanceof HTMLButtonElement) ||
      !candidate ||
      !(candidate.visualElement instanceof Element)
    ) {
      return;
    }

    const targetRect = candidate.visualElement.getBoundingClientRect();
    const paneRect = candidate.leafContent instanceof HTMLElement
      ? candidate.leafContent.getBoundingClientRect()
      : { top: 0, right: window.innerWidth, left: 0 };
    const isVisible =
      targetRect.width >= 72 &&
      targetRect.height >= 48 &&
      targetRect.bottom >= 0 &&
      targetRect.right >= 0 &&
      targetRect.top <= window.innerHeight &&
      targetRect.left <= window.innerWidth;

    if (!isVisible) {
      button.style.display = "none";
      return;
    }

    const buttonWidth = button.offsetWidth || 52;
    const maxLeft = Math.max(8, Math.min(window.innerWidth, paneRect.right) - buttonWidth - 12);
    const left = Math.min(
      maxLeft,
      Math.max(8, Math.max(paneRect.left, 0) + 12)
    );
    const top = Math.min(
      Math.max(8, window.innerHeight - 36),
      Math.max(8, Math.max(paneRect.top, 0) + 12)
    );

    button.style.display = "inline-flex";
    button.style.left = `${left}px`;
    button.style.top = `${top}px`;
    button.style.right = "auto";
  }

  getBestVisibleExcalidrawEmbedCandidate() {
    const selectors = [
      'img[src*=".excalidraw."]',
      'img[srcset*=".excalidraw."]',
      'source[srcset*=".excalidraw."]',
      'object[data*=".excalidraw."]',
      'embed[src*=".excalidraw."]',
    ];
    const activeLeaf = document.querySelector(".workspace-leaf.mod-active");
    const processedVisualElements = new Set();
    let best = null;
    let bestScore = -Infinity;

    for (const element of document.querySelectorAll(selectors.join(","))) {
      if (!(element instanceof Element) || this.isCanvasSurfaceElement(element)) {
        continue;
      }

      const visualElement = this.getExcalidrawEmbedVisualElement(element) ?? element;
      if (!(visualElement instanceof Element) || processedVisualElements.has(visualElement)) {
        continue;
      }
      processedVisualElements.add(visualElement);

      const file = this.resolveExcalidrawSourceFileFromElement(element);
      if (!(file instanceof TFile)) {
        continue;
      }

      const rect = visualElement.getBoundingClientRect();
      const visibleWidth = Math.min(rect.right, window.innerWidth) - Math.max(rect.left, 0);
      const visibleHeight = Math.min(rect.bottom, window.innerHeight) - Math.max(rect.top, 0);
      if (visibleWidth < 72 || visibleHeight < 48) {
        continue;
      }

      const leaf = visualElement.closest(".workspace-leaf");
      const leafContent = visualElement.closest(".workspace-leaf-content");
      const isInActiveLeaf = activeLeaf instanceof Element && leaf === activeLeaf;
      const visibleArea = visibleWidth * visibleHeight;
      const centerY = rect.top + rect.height / 2;
      const distancePenalty = Math.abs(centerY - 120);
      const score = (isInActiveLeaf ? 1000000 : 0) + visibleArea - distancePenalty;

      if (score <= bestScore) {
        continue;
      }

      bestScore = score;
      best = {
        file,
        visualElement,
        leafContent: leafContent instanceof HTMLElement ? leafContent : null,
      };
    }

    return best;
  }

  ensureExcalidrawEmbedEditShadow(host, file) {
    if (!(host instanceof HTMLElement) || !(file instanceof TFile)) {
      return null;
    }

    const existing = host.querySelector(".writing-bridge-edit-shadow-host");
    if (existing instanceof HTMLElement) {
      existing.dataset.sourcePath = file.path;
      return existing;
    }

    // Break out of the image plugin click trap.
    // Closed shadow DOM keeps the Edit hitbox from bubbling back into whatever
    // is listening on the outer embed shell.
    const shadowHost = document.createElement("div");
    shadowHost.className = "writing-bridge-edit-shadow-host";
    shadowHost.dataset.sourcePath = file.path;

    const shadow = shadowHost.attachShadow({ mode: "closed" });
    const style = document.createElement("style");
    style.textContent = `
      :host {
        position: absolute;
        top: 8px;
        right: 8px;
        z-index: 200;
        pointer-events: auto;
      }
      button {
        appearance: none;
        border: 1px solid rgba(255, 255, 255, 0.14);
        border-radius: 999px;
        min-width: 52px;
        height: 28px;
        padding: 0 10px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        font: 600 12px/1 "Segoe UI", sans-serif;
        color: #f6f7fb;
        background: rgba(24, 24, 28, 0.82);
        box-shadow: 0 8px 24px rgba(0, 0, 0, 0.18);
        backdrop-filter: blur(10px);
        cursor: pointer;
        touch-action: manipulation;
        user-select: none;
      }
      button:hover {
        background: rgba(36, 36, 42, 0.94);
      }
      button:active {
        transform: translateY(1px);
      }
    `;

    const button = document.createElement("button");
    button.type = "button";
    button.textContent = "Edit";

    button.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
    });
    button.addEventListener("click", (e) => {
      e.stopPropagation();
      const sourcePath = shadowHost.dataset.sourcePath ?? "";
      void this.openExcalidrawSourcePath(sourcePath);
    });

    shadow.appendChild(style);
    shadow.appendChild(button);
    host.appendChild(shadowHost);
    return shadowHost;
  }


  async openExcalidrawSourcePath(sourcePath) {
    if (typeof sourcePath !== "string" || sourcePath === "") {
      this.log("embed edit open failed: missing source path");
      new Notice("Could not find the Excalidraw source file.");
      return;
    }

    this.log(`embed edit open requested: ${sourcePath}`);
    const file = this.app.vault.getAbstractFileByPath(sourcePath);
    if (!(file instanceof TFile)) {
      this.log(`embed edit open failed: source file missing for ${sourcePath}`);
      new Notice("Could not find the Excalidraw source file.");
      return;
    }

    if (typeof this.app.workspace.openLinkText === "function") {
      try {
        await this.app.workspace.openLinkText(file.path, "", false);
        this.log(`embed edit opened via openLinkText: ${file.path}`);
        return;
      } catch (error) {}
    }

    const leaf = this.app.workspace.activeLeaf ?? this.app.workspace.getMostRecentLeaf();
    if (!leaf) {
      this.log(`embed edit open failed: no leaf available for ${file.path}`);
      new Notice("Could not open the Excalidraw source file.");
      return;
    }

    await leaf.openFile(file);
    this.log(`embed edit opened via openFile: ${file.path}`);
  }

  resolveExcalidrawSourceFileFromElement(element) {
    const urls = this.getExcalidrawExportUrlsFromElement(element);
    if (!urls.length) {
      return null;
    }

    const markdownFiles = this.app.vault.getMarkdownFiles();
    for (const url of urls) {
      const candidatePaths = this.getExcalidrawSourcePathCandidates(url);
      for (const candidatePath of candidatePaths) {
        const direct = this.app.vault.getAbstractFileByPath(candidatePath);
        if (direct instanceof TFile) {
          return direct;
        }

        const normalizedCandidate = normalizePath(candidatePath).toLowerCase();
        const exactMatch = markdownFiles.find(
          (file) => normalizePath(file.path).toLowerCase() === normalizedCandidate
        );
        if (exactMatch instanceof TFile) {
          return exactMatch;
        }

        const fileName = path.posix.basename(candidatePath).toLowerCase();
        const suffixMatch = markdownFiles.find((file) => file.path.toLowerCase().endsWith(`/${fileName}`));
        if (suffixMatch instanceof TFile) {
          return suffixMatch;
        }
      }
    }

    return null;
  }

  getExcalidrawExportUrlsFromElement(element) {
    if (!(element instanceof Element)) {
      return [];
    }

    const urls = new Set();
    const pushUrl = (value) => {
      if (typeof value !== "string" || !value.includes(".excalidraw")) {
        return;
      }

      if (value.includes(",")) {
        for (const part of value.split(",")) {
          const candidate = part.trim().split(/\s+/)[0];
          if (candidate.includes(".excalidraw")) {
            urls.add(candidate);
          }
        }
        return;
      }

      urls.add(value.trim());
    };

    if (element instanceof HTMLImageElement) {
      pushUrl(element.currentSrc);
      pushUrl(element.getAttribute("src"));
      pushUrl(element.getAttribute("srcset"));
    }

    pushUrl(element.getAttribute("src"));
    pushUrl(element.getAttribute("srcset"));
    pushUrl(element.getAttribute("data"));
    pushUrl(element.getAttribute("href"));
    pushUrl(element.getAttribute("xlink:href"));

    return Array.from(urls);
  }

  getExcalidrawSourcePathCandidates(url) {
    if (typeof url !== "string" || !url.includes(".excalidraw")) {
      return [];
    }

    const basePathRaw = this.app?.vault?.adapter?.basePath;
    const basePath = typeof basePathRaw === "string" ? basePathRaw.replace(/\\/g, "/") : "";
    const rawCandidates = new Set();
    const addCandidate = (value) => {
      if (typeof value !== "string" || !value.includes(".excalidraw")) {
        return;
      }

      const stripped = value.split("#")[0].split("?")[0].trim();
      if (!stripped) {
        return;
      }

      const converted = stripped.replace(
        /\.excalidraw(?:\.(?:dark|light))?\.(svg|png)$/i,
        ".excalidraw.md"
      );
      if (!converted.endsWith(".excalidraw.md")) {
        return;
      }

      rawCandidates.add(converted);
    };

    addCandidate(url);
    try {
      const parsed = new URL(url, window.location.href);
      addCandidate(parsed.pathname);
    } catch (error) {}

    const normalized = [];
    for (const candidate of rawCandidates) {
      const decoded = decodeURIComponent(candidate).replace(/\\/g, "/");
      if (basePath && decoded.toLowerCase().startsWith(basePath.toLowerCase())) {
        const relative = decoded.slice(basePath.length).replace(/^\/+/, "");
        if (relative) {
          normalized.push(normalizePath(relative));
        }
      }

      const trimmedLeadingSlash = decoded.replace(/^\/+/, "");
      if (trimmedLeadingSlash) {
        normalized.push(normalizePath(trimmedLeadingSlash));
      }
    }

    return Array.from(new Set(normalized));
  }

  getObsidianThemeMode() {
    return document.body.classList.contains("theme-dark") ? "dark" : "light";
  }

  getObsidianBackgroundColor(contextEl = null) {
    const candidates = [
      contextEl instanceof HTMLElement ? contextEl : null,
      contextEl instanceof HTMLElement ? contextEl.closest(".workspace-leaf-content") : null,
      document.querySelector(".workspace-leaf.mod-active .workspace-leaf-content"),
      document.body,
      document.documentElement,
    ];

    for (const candidate of candidates) {
      if (!(candidate instanceof HTMLElement)) {
        continue;
      }

      const styles = window.getComputedStyle(candidate);
      const cssVar = styles.getPropertyValue("--background-primary").trim();
      if (cssVar) {
        return cssVar;
      }

      const backgroundColor = styles.backgroundColor?.trim();
      if (backgroundColor && backgroundColor !== "rgba(0, 0, 0, 0)" && backgroundColor !== "transparent") {
        return backgroundColor;
      }
    }

    return this.getObsidianThemeMode() === "dark" ? "#1e1e1e" : "#ffffff";
  }

  syncExcalidrawRuledPageOverlay(root) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    const host = this.findExcalidrawRuledPageHost(root);
    this.clearStaleExcalidrawRuledPageHosts(root, host);
    const existing = root.querySelector(".writing-bridge-excalidraw-ruled-page");
    if (
      !(host instanceof HTMLElement) ||
      !this.isFeatureEnabled("ruledPageFeatureEnabled") ||
      !this.settings.ruledPageEnabled
    ) {
      if (existing instanceof HTMLElement) {
        existing.remove();
      }
      return;
    }

    const overlay =
      existing instanceof HTMLDivElement ? existing : this.createExcalidrawRuledPageOverlay();
    host.classList.add("writing-bridge-excalidraw-ruled-host");
    if (!overlay.isConnected) {
      host.appendChild(overlay);
    } else if (overlay.parentElement !== host) {
      overlay.remove();
      host.appendChild(overlay);
    }

    const view = this.getExcalidrawViewForRoot(root);
    const state = this.getExcalidrawRuledPageState(view);
    if (!state) {
      overlay.remove();
      return;
    }

    this.applyRuledPageOverlayState(overlay, state);
  }

  createExcalidrawRuledPageOverlay() {
    const overlay = document.createElement("div");
    overlay.className = "writing-bridge-excalidraw-ruled-page";
    overlay.setAttribute("aria-hidden", "true");
    return overlay;
  }

  getExcalidrawRuledPageState(view) {
    const target = this.normalizeCanvasTarget(this.getCanvasTarget());
    const fallbackView = target?.kind === "excalidraw" ? target.view : null;
    const api = this.getExcalidrawApi(view) ?? this.getExcalidrawApi(fallbackView);
    if (!api) {
      return null;
    }

    const appState = api.getAppState?.();
    const zoom = Number(appState?.zoom?.value ?? 1) || 1;
    const safeZoom = zoom > 0 ? zoom : 1;
    const sceneGridSize = Number(appState?.gridStep ?? appState?.gridSize ?? 8) || 8;
    const rawStepPx = Math.max(1, sceneGridSize * RULED_PAGE_GRID_MULTIPLIER * safeZoom);
    const stepPx = Math.max(RULED_PAGE_MIN_SCREEN_STEP_PX, rawStepPx);
    const sceneOriginViewport = this.excalidrawSceneToViewportCoords(view, { x: 0, y: 0 });
    const offsetPx = this.mod(Number(sceneOriginViewport?.y ?? 0), stepPx);
    const lineThicknessPx = Math.max(1, Math.min(3, 1.5 * safeZoom));
    const opacity = this.getRuledPageOpacityForScreenStep(stepPx);

    return {
      stepPx,
      offsetPx,
      lineThicknessPx,
      opacity,
      visible: true,
    };
  }

  excalidrawSceneToViewportCoords(view, point) {
    const plugin = this.app?.plugins?.plugins?.[EXCALIDRAW_PLUGIN_ID];
    const excalidrawLibCandidates = [
      plugin?.excalidrawLib,
      plugin?.ea?.plugin?.excalidrawLib,
      typeof window !== "undefined" ? window.ExcalidrawLib : null,
      typeof window !== "undefined" ? window.excalidrawLib : null,
    ];
    const api = this.getExcalidrawApi(view);
    const appState = api?.getAppState?.();
    const scenePoint = {
      x: Number(point?.x ?? 0),
      y: Number(point?.y ?? 0),
    };

    for (const excalidrawLib of excalidrawLibCandidates) {
      if (!excalidrawLib || typeof excalidrawLib.sceneCoordsToViewportCoords !== "function") {
        continue;
      }

      try {
        const viewportPoint = excalidrawLib.sceneCoordsToViewportCoords(scenePoint, appState);
        if (
          viewportPoint &&
          Number.isFinite(Number(viewportPoint.x)) &&
          Number.isFinite(Number(viewportPoint.y))
        ) {
          return {
            x: Number(viewportPoint.x),
            y: Number(viewportPoint.y),
          };
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`Excalidraw scene-to-viewport transform failed: ${message}`);
      }
    }

    const zoom = Number(appState?.zoom?.value ?? 1) || 1;
    const viewportWidth = Number(appState?.width ?? view?.containerEl?.clientWidth ?? 0) || 0;
    const viewportHeight = Number(appState?.height ?? view?.containerEl?.clientHeight ?? 0) || 0;
    const scrollX = Number(appState?.scrollX ?? 0);
    const scrollY = Number(appState?.scrollY ?? 0);
    return {
      x: viewportWidth / 2 + (scenePoint.x + scrollX) * zoom,
      y: viewportHeight / 2 + (scenePoint.y + scrollY) * zoom,
    };
  }

  syncExcalidrawRuledPageToggles() {
    if (!this.isFeatureEnabled("ruledPageFeatureEnabled")) {
      this.destroyExcalidrawRuledPageToggles();
      return;
    }

    const activeTarget = this.getCanvasTarget();
    const activeRoot =
      activeTarget?.kind === "excalidraw" ? this.getExcalidrawLeafRootForView(activeTarget.view) : null;
    const activeRootId =
      activeRoot instanceof HTMLElement ? this.ensureStyleToggleRootId(activeRoot) : "";
    const roots = document.querySelectorAll('.workspace-leaf-content[data-type="excalidraw"]');
    for (const root of roots) {
      this.syncExcalidrawRuledPageToggle(root, activeRoot);
    }

    for (const button of document.querySelectorAll('.writing-bridge-ruled-toggle[data-target-kind="excalidraw"]')) {
      if (!(button instanceof HTMLElement)) {
        continue;
      }

      if (!activeRootId || button.dataset.rootId !== activeRootId) {
        button.remove();
      }
    }
  }

  destroyExcalidrawRuledPageToggles() {
    for (const button of document.querySelectorAll('.writing-bridge-ruled-toggle[data-target-kind="excalidraw"]')) {
      if (button instanceof HTMLElement) {
        button.remove();
      }
    }
  }

  syncExcalidrawRuledPageToggle(root, activeRoot = null) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    const wrapper = this.findExcalidrawContainer(root);
    const rootId = this.ensureStyleToggleRootId(root);
    const existing = root.querySelector(
      `.writing-bridge-ruled-toggle[data-root-id="${rootId}"][data-target-kind="excalidraw"]`
    );

    if (!(wrapper instanceof HTMLElement) || activeRoot !== root) {
      if (existing instanceof HTMLElement) {
        existing.remove();
      }
      return;
    }

    const button =
      existing instanceof HTMLButtonElement ? existing : this.createExcalidrawRuledPageToggleButton();
    if (!button.dataset.rootId) {
      button.dataset.rootId = rootId;
    }

    if (!button.isConnected) {
      wrapper.appendChild(button);
    } else if (button.parentElement !== wrapper) {
      button.remove();
      wrapper.appendChild(button);
    }

    const enabled = Boolean(this.settings.ruledPageEnabled);
    button.classList.toggle("is-on", enabled);
    button.classList.toggle("is-off", !enabled);
    button.textContent = "Lines";
    button.setAttribute("aria-label", enabled ? "Hide ruled lines" : "Show ruled lines");
    button.setAttribute("aria-pressed", String(enabled));
    button.title = enabled ? "Hide ruled lines" : "Show ruled lines";
    this.positionExcalidrawRuledPageToggle(button, root);
  }

  createExcalidrawRuledPageToggleButton() {
    const button = this.createRuledPageToggleButton();
    button.dataset.targetKind = "excalidraw";
    return button;
  }

  positionExcalidrawRuledPageToggle(button, root) {
    if (!(button instanceof HTMLElement) || !(root instanceof HTMLElement)) {
      return;
    }

    const wrapper = this.findExcalidrawContainer(root);
    const rootRect = root.getBoundingClientRect();
    const wrapperRect = wrapper instanceof HTMLElement ? wrapper.getBoundingClientRect() : null;
    if (rootRect.width <= 0 || rootRect.height <= 0 || !wrapperRect || wrapperRect.width <= 0) {
      return;
    }

    const rightUiAnchor = this.findExcalidrawRightUiAnchor(root);
    const anchorRect = rightUiAnchor instanceof HTMLElement ? rightUiAnchor.getBoundingClientRect() : null;
    const top = anchorRect
      ? Math.max(12, Math.round(anchorRect.top - wrapperRect.top + 6))
      : Math.max(18, Math.round(rootRect.top - wrapperRect.top + 18));
    const right = anchorRect
      ? Math.max(-10, Math.round(wrapperRect.right - anchorRect.right - 10))
      : -10;
    button.style.top = `${top}px`;
    button.style.right = `${right}px`;
    button.style.left = `auto`;
  }

  findExcalidrawRightUiAnchor(root) {
    if (!(root instanceof HTMLElement)) {
      return null;
    }

    const selectors = [
      ".selected-shape-actions",
      ".layer-ui__wrapper__top-right",
      ".excalidraw-ui-top-right",
      ".App-menu_top",
      ".FixedSideContainer.FixedSideContainer_side_top",
    ];

    for (const selector of selectors) {
      for (const candidate of root.querySelectorAll(selector)) {
        if (!(candidate instanceof HTMLElement)) {
          continue;
        }

        const rect = candidate.getBoundingClientRect();
        if (rect.width > 0 && rect.height > 0) {
          return candidate;
        }
      }
    }

    return null;
  }

  syncExcalidrawStylePanelToggles() {
    this.destroyExcalidrawStylePanelToggles();
  }

  destroyExcalidrawStylePanelToggles() {
    for (const button of document.querySelectorAll('.writing-bridge-style-toggle[data-target-kind="excalidraw"]')) {
      if (button instanceof HTMLElement) {
        button.remove();
      }
    }

    for (const root of document.querySelectorAll('.workspace-leaf-content[data-type="excalidraw"]')) {
      if (root instanceof HTMLElement) {
        root.classList.remove("writing-bridge-style-panel-collapsed");
      }
    }
  }

  syncExcalidrawStylePanelToggle(root, activeRoot = null) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    root.classList.toggle(
      "writing-bridge-style-panel-collapsed",
      Boolean(this.settings.stylePanelCollapsed)
    );

    const panel = root.querySelector(".selected-shape-actions");
    const rootId = this.ensureStyleToggleRootId(root);
    const existing = document.querySelector(
      `.writing-bridge-style-toggle[data-root-id="${rootId}"][data-target-kind="excalidraw"]`
    );

    if (!(panel instanceof HTMLElement) || activeRoot !== root) {
      if (existing instanceof HTMLElement) {
        existing.remove();
      }
      return;
    }

    const button =
      existing instanceof HTMLButtonElement ? existing : this.createExcalidrawStylePanelToggleButton();
    if (!button.dataset.rootId) {
      button.dataset.rootId = rootId;
    }

    if (!button.isConnected) {
      document.body.appendChild(button);
    }

    const panelIsOpen = !this.settings.stylePanelCollapsed;
    button.classList.toggle("is-open", panelIsOpen);
    button.textContent = panelIsOpen ? ">" : "<";
    button.setAttribute("aria-label", panelIsOpen ? "Hide Excalidraw styles" : "Show Excalidraw styles");
    button.setAttribute("aria-pressed", String(panelIsOpen));
    button.title = panelIsOpen ? "Hide Excalidraw styles" : "Show Excalidraw styles";
    this.positionExcalidrawStylePanelToggle(button, root, panel, panelIsOpen);
  }

  createExcalidrawStylePanelToggleButton() {
    const button = this.createStylePanelToggleButton();
    button.dataset.targetKind = "excalidraw";
    return button;
  }

  positionExcalidrawStylePanelToggle(button, root, panel, panelIsOpen) {
    if (!(button instanceof HTMLElement) || !(root instanceof HTMLElement)) {
      return;
    }

    const rootRect = root.getBoundingClientRect();
    const panelRect = panel instanceof HTMLElement ? panel.getBoundingClientRect() : null;
    if (rootRect.width <= 0 || rootRect.height <= 0) {
      return;
    }

    const margin = 8;
    const handleWidth = 22;
    const handleHeight = 42;
    const rootTop = Math.max(margin, Math.round(rootRect.top));
    const rootBottom = Math.min(window.innerHeight - margin, Math.round(rootRect.bottom));

    let top = rootTop + 58;
    let left = Math.round(rootRect.right) - handleWidth - margin;

    if (panelIsOpen && panelRect && panelRect.width > 0 && panelRect.height > 0) {
      top = Math.round(panelRect.top + 8);
      left = Math.round(panelRect.left - handleWidth - 6);
    }

    top = Math.max(rootTop + margin, Math.min(top, rootBottom - handleHeight - margin));
    left = Math.max(margin, Math.min(left, window.innerWidth - handleWidth - margin));

    button.style.top = `${top}px`;
    button.style.left = `${left}px`;
  }

  getTldrawTargetForRoot(root, preferredTarget = null) {
    const candidates = [
      this.normalizeTldrawTarget(preferredTarget),
      this.normalizeTldrawTarget(this.lastActiveTldrawTarget),
      this.normalizeTldrawTarget(this.getTldrawTarget()),
    ];

    for (const candidate of candidates) {
      if (!candidate) {
        continue;
      }

      const candidateRoot = this.getTldrawTargetRoot(candidate);
      if (candidateRoot === root) {
        return candidate;
      }
    }

    return null;
  }

  normalizeTldrawTarget(target) {
    if (!target) {
      return null;
    }

    const editor = target.editor ?? this.getCurrentTldrawEditor(target.container ?? null);
    const rawContainer =
      target.container instanceof HTMLElement
        ? target.container
        : typeof editor?.getContainer === "function"
          ? editor.getContainer()
          : null;
    const container = this.findTldrawContainer(rawContainer) ?? rawContainer;

    if (!editor && !(container instanceof HTMLElement)) {
      return null;
    }

    return { editor, container };
  }

  getTldrawTargetRoot(target) {
    const container =
      target?.container instanceof HTMLElement
        ? target.container
        : typeof target?.editor?.getContainer === "function"
          ? target.editor.getContainer()
          : null;
    return container instanceof HTMLElement ? container.closest(".tldraw-view-root") : null;
  }

  getRuledPageOverlayState(target) {
    const editor = target?.editor;
    if (
      !editor ||
      typeof editor.getCamera !== "function" ||
      typeof editor.getDocumentSettings !== "function"
    ) {
      return null;
    }

    const documentSettings = editor.getDocumentSettings();
    const gridSize = Number(documentSettings?.gridSize ?? 10);
    const camera = typeof editor.getCamera === "function" ? editor.getCamera() : null;
    const zoom = Number(camera?.z ?? camera?.zoom ?? 1) || 1;
    const safeZoom = zoom > 0 ? zoom : 1;
    const rawStepPx = Math.max(1, gridSize * RULED_PAGE_GRID_MULTIPLIER);
    const stepPx = Math.max(rawStepPx, RULED_PAGE_MIN_SCREEN_STEP_PX / safeZoom);
    const offsetPx = 0;
    const lineThicknessPx = Math.max(RULED_PAGE_LINE_THICKNESS_PX, RULED_PAGE_LINE_THICKNESS_PX / safeZoom);
    const opacity = this.getRuledPageOpacityForScreenStep(stepPx * safeZoom);

    return {
      stepPx,
      offsetPx,
      lineThicknessPx,
      opacity,
      visible: true,
    };
  }

  getConfiguredTldrawDefaultDrawSize() {
    if (!this.isFeatureEnabled("tldrawDefaultPenSizeEnabled")) {
      return null;
    }

    const size = this.settings?.tldrawDefaultDrawSize;
    return size && size !== "default" && TLDRAW_DRAW_SIZE_OPTIONS.includes(size) ? size : null;
  }

  getCurrentTldrawDrawSize(editor) {
    if (!editor || typeof editor.getStyleForNextShape !== "function") {
      return null;
    }

    const size = editor.getStyleForNextShape(TLDRAW_SIZE_STYLE);
    return typeof size === "string" && size ? size : null;
  }

  isTldrawDrawSizeResetValue(size) {
    return (
      !size ||
      size === "default" ||
      size === TLDRAW_SIZE_STYLE.defaultValue
    );
  }

  getTldrawDrawSizeApplyMarker(editor, root = null) {
    const pageId =
      typeof editor?.getCurrentPageId === "function"
        ? `${editor.getCurrentPageId() ?? ""}`
        : "";
    const storeId = typeof editor?.store?.id === "string" ? editor.store.id : "";
    const leafPath =
      root instanceof HTMLElement
        ? root.closest(".workspace-leaf")?.querySelector(".workspace-leaf-content")?.getAttribute("data-path") ?? ""
        : "";
    return `${storeId}::${pageId}::${leafPath}`;
  }

  applyPreferredTldrawDrawSize(root, target) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    const preferredSize = this.getConfiguredTldrawDefaultDrawSize();
    if (!preferredSize) {
      return;
    }

    const editor = target?.editor;
    if (!editor || typeof editor.setStyleForNextShapes !== "function") {
      return;
    }

    const marker = this.getTldrawDrawSizeApplyMarker(editor, root);
    const currentSize = this.getCurrentTldrawDrawSize(editor);
    const rootId = this.ensureStyleToggleRootId(root);
    const previous = this.appliedTldrawDrawSizeByRootId.get(rootId);
    const sameScope = previous?.editor === editor && previous?.marker === marker;

    if (currentSize === preferredSize) {
      if (!sameScope || previous?.size !== currentSize) {
        this.appliedTldrawDrawSizeByRootId.set(rootId, { editor, size: currentSize, marker });
      }
      return;
    }

    if (
      sameScope &&
      currentSize &&
      currentSize !== preferredSize &&
      !this.isTldrawDrawSizeResetValue(currentSize)
    ) {
      if (previous?.size !== currentSize) {
        this.appliedTldrawDrawSizeByRootId.set(rootId, { editor, size: currentSize, marker });
      }
      return;
    }

    try {
      if (currentSize !== preferredSize) {
        editor.setStyleForNextShapes(TLDRAW_SIZE_STYLE, preferredSize);
      }
      const appliedSize = this.getCurrentTldrawDrawSize(editor) ?? preferredSize;
      this.appliedTldrawDrawSizeByRootId.set(rootId, { editor, size: appliedSize, marker });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`failed to apply preferred tldraw draw size: ${message}`);
    }
  }

  syncStylePanelToggles() {
    if (!this.isFeatureEnabled("stylePanelToggleEnabled")) {
      this.destroyStylePanelToggles();
      return;
    }

    const liveRootIds = new Set();
    const roots = document.querySelectorAll(".tldraw-view-root");
    for (const root of roots) {
      if (!(root instanceof HTMLElement)) {
        continue;
      }

      liveRootIds.add(this.ensureStyleToggleRootId(root));
      this.syncStylePanelToggle(root);
    }

    for (const button of document.querySelectorAll(".writing-bridge-style-toggle")) {
      if (!(button instanceof HTMLElement)) {
        continue;
      }

      if (!liveRootIds.has(button.dataset.rootId ?? "")) {
        button.remove();
      }
    }

    this.destroyExcalidrawStylePanelToggles();
  }

  syncStylePanelToggle(root) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    root.classList.toggle(
      "writing-bridge-style-panel-collapsed",
      Boolean(this.settings.stylePanelCollapsed)
    );

    const panel = root.querySelector(".tlui-style-panel");
    const rootId = this.ensureStyleToggleRootId(root);
    const existing = document.querySelector(
      `.writing-bridge-style-toggle[data-root-id="${rootId}"]`
    );

    if (!(panel instanceof HTMLElement)) {
      if (existing instanceof HTMLElement) {
        existing.remove();
      }
      return;
    }

    const button =
      existing instanceof HTMLButtonElement ? existing : this.createStylePanelToggleButton();

    if (!button.dataset.rootId) {
      button.dataset.rootId = rootId;
    }

    if (!button.isConnected) {
      document.body.appendChild(button);
    }

    const panelIsOpen = !this.settings.stylePanelCollapsed;
    button.classList.toggle("is-open", panelIsOpen);
    button.textContent = panelIsOpen ? ">" : "<";
    button.setAttribute("aria-label", panelIsOpen ? "Hide tldraw styles" : "Show tldraw styles");
    button.setAttribute("aria-pressed", String(panelIsOpen));
    button.title = panelIsOpen ? "Hide tldraw styles" : "Show tldraw styles";
    this.positionStylePanelToggle(button, root, panel, panelIsOpen);
  }

  createStylePanelToggleButton() {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "writing-bridge-style-toggle";
    button.addEventListener("pointerup", (event) => {
      event.stopPropagation();
      event.preventDefault();
      void this.toggleStylePanelCollapsed();
    });
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
    return button;
  }

  ensureStyleToggleRootId(root) {
    if (typeof root.dataset.tldrawPenBridgeRootId === "string" && root.dataset.tldrawPenBridgeRootId) {
      return root.dataset.tldrawPenBridgeRootId;
    }

    this.stylePanelToggleId += 1;
    const rootId = `writing-bridge-root-${this.stylePanelToggleId}`;
    root.dataset.tldrawPenBridgeRootId = rootId;
    return rootId;
  }

  positionStylePanelToggle(button, root, panel, panelIsOpen) {
    if (!(button instanceof HTMLElement) || !(root instanceof HTMLElement) || !(panel instanceof HTMLElement)) {
      return;
    }

    const rootRect = root.getBoundingClientRect();
    const panelRect = panel.getBoundingClientRect();
    if (rootRect.width <= 0 || rootRect.height <= 0) {
      return;
    }

    const margin = 8;
    const handleWidth = 22;
    const handleHeight = 42;
    const rootTop = Math.max(margin, Math.round(rootRect.top));
    const rootBottom = Math.min(window.innerHeight - margin, Math.round(rootRect.bottom));

    let top = rootTop + 54;
    let left = Math.round(rootRect.right) - handleWidth - margin;

    if (panelIsOpen && panelRect.width > 0 && panelRect.height > 0) {
      top = Math.round(panelRect.top + 8);
      left = Math.round(panelRect.left - handleWidth - 6);
    }

    top = Math.max(rootTop + margin, Math.min(top, rootBottom - handleHeight - margin));
    left = Math.max(margin, Math.min(left, window.innerWidth - handleWidth - margin));

    button.style.top = `${top}px`;
    button.style.left = `${left}px`;
  }

  async toggleStylePanelCollapsed() {
    if (!this.isFeatureEnabled("stylePanelToggleEnabled")) {
      this.showFeatureDisabledNotice("Canvas style panel toggle");
      return;
    }

    this.settings.stylePanelCollapsed = !this.settings.stylePanelCollapsed;
    await this.saveSettings();
    this.syncStylePanelToggles();
  }

  destroyStylePanelToggles() {
    for (const button of document.querySelectorAll(".writing-bridge-style-toggle")) {
      if (button instanceof HTMLElement) {
        button.remove();
      }
    }

    for (const root of document.querySelectorAll(".tldraw-view-root")) {
      if (root instanceof HTMLElement) {
        root.classList.remove("writing-bridge-style-panel-collapsed");
      }
    }

    for (const root of document.querySelectorAll('.workspace-leaf-content[data-type="excalidraw"]')) {
      if (root instanceof HTMLElement) {
        root.classList.remove("writing-bridge-style-panel-collapsed");
      }
    }
  }

  refreshSelectionMenu() {
    if (!this.isFeatureEnabled("selectionMenuEnabled")) {
      this.lastActiveCanvasTarget = null;
      this.lastActiveTldrawTarget = null;
      this.selectionMenuState.lastSelectionKey = "";
      this.selectionMenuState.dismissedSelectionKey = "";
      this.hideSelectionMenu();
      return;
    }

    const target = this.getCanvasTarget();
    if (!target) {
      this.lastActiveCanvasTarget = null;
      this.lastActiveTldrawTarget = null;
      this.selectionMenuState.lastSelectionKey = "";
      this.selectionMenuState.dismissedSelectionKey = "";
      this.hideSelectionMenu();
      return;
    }

    if (target.kind === "tldraw") {
      this.lastActiveTldrawTarget = {
        kind: "tldraw",
        editor: target.editor,
        container: target?.container ?? this.lastActiveTldrawTarget?.container ?? null,
      };
    }
    this.lastActiveCanvasTarget = target;

    const selectionIds =
      target.kind === "tldraw"
        ? this.getTldrawSelectionIds(target.editor)
        : this.getExcalidrawSelectionIds(target.view);

    if (!selectionIds.length) {
      this.selectionMenuState.lastSelectionKey = "";
      this.selectionMenuState.dismissedSelectionKey = "";
      this.hideSelectionMenu();
      return;
    }

    const selectionKey = `${target.kind}:${selectionIds.join("|")}`;
    if (selectionKey !== this.selectionMenuState.lastSelectionKey) {
      this.selectionMenuState.lastSelectionKey = selectionKey;
      this.selectionMenuState.dismissedSelectionKey = "";
    }

    if (this.selectionMenuState.dismissedSelectionKey === selectionKey) {
      this.hideSelectionMenu();
      return;
    }

    const bounds = this.getSelectionMenuBoundsForTarget(target);
    if (!bounds) {
      this.hideSelectionMenu();
      return;
    }

    const menu = this.ensureSelectionMenu();
    this.syncSelectionMenuButtons(menu);
    const visibleButtons = Array.from(menu.querySelectorAll("button")).filter((button) => !button.hidden);
    if (visibleButtons.length <= 1) {
      this.hideSelectionMenu();
      return;
    }
    menu.setAttribute("data-selection-key", selectionKey);
    this.positionSelectionMenu(menu, bounds);
    menu.classList.add("is-visible");
  }

  ensureSelectionMenu() {
    if (this.selectionMenuEl instanceof HTMLElement) {
      return this.selectionMenuEl;
    }

    const menu = document.createElement("div");
    menu.className = "writing-bridge-selection-menu";

    const copyButton = this.createSelectionMenuButton(
      "Copy",
      async () => {
        this.dismissSelectionMenu(menu);
        await this.copySelectedAsText(this.lastActiveCanvasTarget ?? this.getCanvasTarget());
      },
      "writing-bridge-selection-menu__primary",
      "copySelectedTextEnabled"
    );
    const cutButton = this.createSelectionMenuButton(
      "Cut",
      async () => {
        this.dismissSelectionMenu(menu);
        await this.cutSelectedContent(this.lastActiveCanvasTarget ?? this.getCanvasTarget());
      },
      "",
      "copySelectedTextEnabled"
    );
    const screenshotButton = this.createSelectionMenuButton(
      "Screenshot",
      async () => {
        this.dismissSelectionMenu(menu);
        await this.copySelectedAsScreenshot(this.lastActiveCanvasTarget ?? this.getCanvasTarget());
      },
      "",
      "selectionScreenshotEnabled"
    );
    const drawButton = this.createSelectionMenuButton(
      "Draw",
      () => {
        this.dismissSelectionMenu(menu);
        this.runToolCommand("draw");
      },
      "",
      "toolCommandsEnabled"
    );
    const panButton = this.createSelectionMenuButton(
      "Pan",
      () => {
        this.dismissSelectionMenu(menu);
        this.runToolCommand("hand");
      },
      "",
      "toolCommandsEnabled"
    );
    const closeButton = this.createSelectionMenuButton("Close", () => {
      this.dismissSelectionMenu(menu);
    });

    menu.append(copyButton, cutButton, screenshotButton, drawButton, panButton, closeButton);
    this.syncSelectionMenuButtons(menu);
    document.body.appendChild(menu);
    this.selectionMenuEl = menu;
    return menu;
  }

  syncSelectionMenuButtons(menu) {
    if (!(menu instanceof HTMLElement)) {
      return;
    }

    for (const button of menu.querySelectorAll("button")) {
      if (!(button instanceof HTMLButtonElement)) {
        continue;
      }

      const featureKey = button.dataset.featureToggle;
      const enabled = !featureKey || this.isFeatureEnabled(featureKey);
      button.hidden = !enabled;
      button.disabled = !enabled;
    }
  }

  createSelectionMenuButton(label, onClick, extraClass = "", featureToggleKey = "") {
    const button = document.createElement("button");
    button.type = "button";
    button.className = extraClass;
    button.textContent = label;
    if (featureToggleKey) {
      button.dataset.featureToggle = featureToggleKey;
    }
    let handledPointerUp = false;

    button.addEventListener("pointerdown", (event) => {
      event.stopPropagation();
      if (event.cancelable) {
        event.preventDefault();
      }
    });
    button.addEventListener("pointerup", (event) => {
      handledPointerUp = true;
      event.stopPropagation();
      if (event.cancelable) {
        event.preventDefault();
      }
      void onClick();
      window.setTimeout(() => {
        handledPointerUp = false;
      }, 0);
    });
    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
      if (handledPointerUp) {
        return;
      }
      void onClick();
    });
    button.addEventListener("keydown", (event) => {
      if (event.key !== "Enter" && event.key !== " ") {
        return;
      }

      event.preventDefault();
      event.stopPropagation();
      void onClick();
    });
    return button;
  }

  runToolCommand(toolId) {
    if (!this.isFeatureEnabled("toolCommandsEnabled")) {
      this.showFeatureDisabledNotice("Canvas tool commands");
      return;
    }

    const command = TOOL_COMMANDS.find((entry) => entry.toolId === toolId);
    const target = this.lastActiveCanvasTarget ?? this.getCanvasTarget();
    if (!command || !target) {
      return;
    }

    this.activateTool(target, command);
  }

  async cutSelectedContent(target) {
    const resolvedTarget = this.normalizeCanvasTarget(target ?? this.getCanvasTarget());
    if (!resolvedTarget) {
      new Notice("No active canvas selection to cut");
      return;
    }

    let snapshot = null;
    try {
      snapshot = await this.prepareCutSnapshot(resolvedTarget);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`cut snapshot failed: ${message}`);
      new Notice(this.getFailureNoticeMessage("Cut text", error));
      return;
    }

    if (!snapshot) {
      this.log("cut skipped: no snapshot was prepared");
      return;
    }

    const didDelete = await this.deleteSelectedContentAfterCopy(resolvedTarget);

    let didWrite = false;
    try {
      didWrite = await this.writeCutSnapshotToClipboard(snapshot);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`cut clipboard write failed: ${message}`);
      new Notice(this.getFailureNoticeMessage("Cut text", error));
      return;
    }

    if (!didWrite) {
      new Notice("Cut text failed: no text could be recognized");
      return;
    }

    if (!didDelete) {
      this.log(`cut downgraded to copy for ${resolvedTarget.kind}: selection was not removed`);
      new Notice("Copied selected text");
      return;
    }

    this.log(`cut completed for ${resolvedTarget.kind}`);
    new Notice("Cut text");
  }

  async prepareCutSnapshot(target) {
    const liveTarget = this.normalizeCanvasTarget(target) ?? this.getCanvasTarget();
    if (!liveTarget) {
      return null;
    }

    if (liveTarget.kind === "excalidraw") {
      return this.prepareExcalidrawCutSnapshot(liveTarget);
    }

    return this.prepareTldrawCutSnapshot(liveTarget);
  }

  async prepareTldrawCutSnapshot(target) {
    const container = target?.container ?? this.getActiveTldrawContainer();
    const domText = this.getActiveDomTextSelection(container);
    if (domText) {
      return {
        type: "text",
        text: domText,
        source: "tldraw-dom",
      };
    }

    const editor = target?.editor ?? this.getCurrentTldrawEditor(container);
    if (!editor) {
      this.log("cut snapshot failed: no active tldraw editor");
      new Notice("No active tldraw canvas");
      return null;
    }

    const selectedShapes = this.getSelectedLeafShapes(editor);
    if (!selectedShapes.length) {
      this.log("cut snapshot skipped: no selected tldraw shapes");
      new Notice("No tldraw selection to cut");
      return null;
    }

    const directText = this.tryExtractDirectText(editor, selectedShapes);
    if (directText) {
      return {
        type: "text",
        text: directText,
        source: "tldraw-direct",
      };
    }

    const shapeIds = selectedShapes.map((shape) => shape.id);
    const strokePayload = this.buildStrokeRecognitionPayload(editor, selectedShapes);
    let blob = null;
    try {
      blob = await this.exportSelectionBlob(editor, shapeIds, {
        padding: OCR_EXPORT_PADDING,
        scale: OCR_EXPORT_SCALE,
        pixelRatio: OCR_EXPORT_PIXEL_RATIO,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`tldraw cut snapshot export failed: ${message}`);
    }

    if (!strokePayload && !(blob instanceof Blob)) {
      new Notice("No text found in selection");
      return null;
    }

    return {
      type: "canvas-selection",
      source: "tldraw",
      strokePayload,
      blob,
    };
  }

  async prepareExcalidrawCutSnapshot(target) {
    const normalizedTarget = this.normalizeExcalidrawTarget(target);
    const container = normalizedTarget?.container ?? this.getActiveExcalidrawContainer();
    const domText = this.getActiveDomTextSelection(container);
    if (domText) {
      return {
        type: "text",
        text: domText,
        source: "excalidraw-dom",
      };
    }

    const view = normalizedTarget?.view ?? this.getCurrentExcalidrawView(container);
    if (!view) {
      this.log("cut snapshot failed: no active Excalidraw view");
      new Notice("No active Excalidraw canvas");
      return null;
    }

    const selectedElements = this.getExcalidrawSelectedElements(view);
    if (!selectedElements.length) {
      this.log("cut snapshot skipped: no selected Excalidraw elements");
      new Notice("No Excalidraw selection to cut");
      return null;
    }

    const directText = this.tryExtractExcalidrawDirectText(selectedElements);
    if (directText) {
      return {
        type: "text",
        text: directText,
        source: "excalidraw-direct",
      };
    }

    const strokePayload = this.buildExcalidrawStrokeRecognitionPayload(view, selectedElements);
    let blob = null;
    try {
      blob = await this.captureExcalidrawSelectionBlob(view, {
        padding: OCR_EXPORT_PADDING,
        scale: OCR_EXPORT_SCALE,
        pixelRatio: OCR_EXPORT_PIXEL_RATIO,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`Excalidraw cut snapshot export failed: ${message}`);
    }

    if (!strokePayload && !(blob instanceof Blob)) {
      new Notice("No text found in selection");
      return null;
    }

    return {
      type: "canvas-selection",
      source: "excalidraw",
      strokePayload,
      blob,
    };
  }

  async writeCutSnapshotToClipboard(snapshot) {
    if (!snapshot) {
      return false;
    }

    if (snapshot.type === "text") {
      await this.writeClipboardText(snapshot.text);
      this.log(`cut clipboard text written from ${snapshot.source}`);
      return true;
    }

    let normalizedText = null;

    if (snapshot.strokePayload) {
      try {
        const result = await this.runInkStrokeRecognizer(snapshot.strokePayload);
        const strokeText = this.normalizeCopiedText(result?.text);
        this.log(
          `cut ink recognizer source=${snapshot.source} status=${result?.status ?? "unknown"} confidence=${result?.confidence ?? "unknown"} strokes=${snapshot.strokePayload?.strokes?.length ?? 0}`
        );

        if (strokeText && this.shouldTrustStrokeRecognition(result, strokeText, snapshot.strokePayload)) {
          normalizedText = this.mergeStrokeTextWithDirectEntries(strokeText, snapshot.strokePayload);
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`cut ink recognizer failed for ${snapshot.source}: ${message}`);
      }
    }

    if (!normalizedText && snapshot.blob instanceof Blob) {
      const extractedText = await this.extractBlobTextWithOcr(snapshot.blob);
      normalizedText = this.normalizeCopiedText(extractedText);
    }

    if (!normalizedText) {
      return false;
    }

    await this.writeClipboardText(normalizedText);
    this.log(`cut clipboard text written from ${snapshot.source}`);
    return true;
  }

  async deleteSelectedContentAfterCopy(target) {
    const liveTarget = this.normalizeCanvasTarget(target) ?? this.getCanvasTarget();
    if (!liveTarget) {
      return false;
    }

    const container =
      liveTarget.kind === "excalidraw"
        ? liveTarget.container ?? this.getActiveExcalidrawContainer()
        : liveTarget.container ?? this.getActiveTldrawContainer();
    if (this.deleteSelectedDomText(container)) {
      return true;
    }

    if (liveTarget.kind === "excalidraw") {
      return this.deleteSelectedExcalidrawContent(liveTarget);
    }

    return this.deleteSelectedTldrawContent(liveTarget);
  }

  deleteSelectedDomText(container) {
    const selection = document.getSelection();
    const rawText = selection?.toString();
    if (!rawText || !rawText.trim()) {
      return false;
    }

    if (container) {
      const anchorWithin = this.nodeIsWithinContainer(selection?.anchorNode ?? null, container);
      const focusWithin = this.nodeIsWithinContainer(selection?.focusNode ?? null, container);
      if (!anchorWithin && !focusWithin) {
        return false;
      }
    }

    try {
      selection.deleteFromDocument?.();
      this.log("deleted active DOM text selection after cut");
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`delete DOM text selection failed: ${message}`);
      return false;
    }
  }

  deleteSelectedTldrawContent(target) {
    const editor = target?.editor ?? this.getCurrentTldrawEditor(target?.container ?? null);
    if (!editor) {
      return false;
    }

    const shapeIds = this.getTldrawSelectionIds(editor);
    if (!shapeIds.length || typeof editor.deleteShapes !== "function") {
      return false;
    }

    try {
      if (typeof editor.markHistoryStoppingPoint === "function") {
        editor.markHistoryStoppingPoint("cut");
      }
      editor.deleteShapes(shapeIds);
      this.log(`deleted ${shapeIds.length} tldraw shapes after cut`);
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`delete tldraw selection failed: ${message}`);
      return false;
    }
  }

  deleteSelectedExcalidrawContent(target) {
    const normalizedTarget = this.normalizeExcalidrawTarget(target);
    const view = normalizedTarget?.view ?? this.getCurrentExcalidrawView(normalizedTarget?.container ?? null);
    if (!view) {
      return false;
    }

    const selectedElements = this.getExcalidrawSelectedElements(view);
    if (!selectedElements.length) {
      return false;
    }

    const api = normalizedTarget?.api ?? this.getExcalidrawApi(view);
    const selectedIds = new Set(selectedElements.map((element) => element?.id).filter(Boolean));
    if (api && typeof api.getSceneElements === "function" && typeof api.updateScene === "function") {
      try {
        const elements = Array.from(api.getSceneElements() ?? []);
        const updatedElements = elements.map((element) =>
          selectedIds.has(element?.id) ? { ...element, isDeleted: true } : element
        );
        api.updateScene({
          elements: updatedElements,
          appState: {
            selectedElementIds: {},
          },
          commitToHistory: true,
        });
        this.focusExcalidrawTarget(normalizedTarget ?? target);
        this.log(`deleted ${selectedIds.size} Excalidraw elements after cut via updateScene`);
        return true;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`delete Excalidraw selection via updateScene failed: ${message}`);
      }
    }

    const root = this.getExcalidrawRootForView(view) ?? normalizedTarget?.container;
    if (!(root instanceof HTMLElement)) {
      return false;
    }

    this.focusExcalidrawTarget(normalizedTarget ?? target);
    this.sendShortcutKey(root, "Delete", "Delete");
    this.log(`deleted ${selectedIds.size} Excalidraw elements after cut via Delete key fallback`);
    return true;
  }

  getSelectionMenuBoundsForTarget(target) {
    if (target?.kind === "excalidraw") {
      return this.getExcalidrawSelectionMenuBounds(target);
    }

    return this.getSelectionMenuBounds(target?.editor);
  }

  getSelectionMenuBounds(editor) {
    if (typeof editor.getSelectionScreenBounds === "function") {
      const screenBounds = editor.getSelectionScreenBounds();
      if (screenBounds) {
        return screenBounds;
      }
    }

    if (
      typeof editor.getSelectionPageBounds === "function" &&
      typeof editor.pageToScreen === "function"
    ) {
      const pageBounds = editor.getSelectionPageBounds();
      if (!pageBounds) {
        return null;
      }

      const point = editor.pageToScreen(pageBounds.point);
      return {
        x: point.x,
        y: point.y,
        w: pageBounds.w,
        h: pageBounds.h,
      };
    }

    if (
      typeof editor?.getShapePageBounds === "function" &&
      typeof editor?.pageToScreen === "function"
    ) {
      let mergedBounds = null;
      for (const shapeId of this.getTldrawSelectionIds(editor)) {
        const bounds = editor.getShapePageBounds(shapeId);
        if (!bounds) {
          continue;
        }
        mergedBounds = mergedBounds ? this.mergeBounds(mergedBounds, bounds) : this.cloneBounds(bounds);
      }

      if (mergedBounds) {
        const point = editor.pageToScreen({ x: mergedBounds.x, y: mergedBounds.y });
        return {
          x: point.x,
          y: point.y,
          w: mergedBounds.w,
          h: mergedBounds.h,
        };
      }
    }

    return null;
  }

  getTldrawSelectionIds(editor) {
    return typeof editor?.getSelectedShapeIds === "function" ? Array.from(editor.getSelectedShapeIds() ?? []) : [];
  }

  getExcalidrawSelectionIds(view) {
    return this.getExcalidrawSelectedElements(view).map((element) => element.id).filter(Boolean);
  }

  positionSelectionMenu(menu, bounds) {
    const menuWidth = menu.offsetWidth || 260;
    const menuHeight = menu.offsetHeight || 44;
    const margin = 12;
    const selectionCenterX = Number(bounds.x ?? 0) + Number(bounds.w ?? 0) / 2;
    const selectionTop = Number(bounds.y ?? 0);
    const selectionBottom = selectionTop + Number(bounds.h ?? 0);

    let left = selectionCenterX - menuWidth / 2;
    left = Math.max(margin, Math.min(left, window.innerWidth - menuWidth - margin));

    let top = selectionTop - menuHeight - margin;
    if (top < margin) {
      top = selectionBottom + margin;
    }

    menu.style.left = `${left}px`;
    menu.style.top = `${top}px`;
  }

  hideSelectionMenu() {
    if (this.selectionMenuEl instanceof HTMLElement) {
      this.selectionMenuEl.classList.remove("is-visible");
    }
  }

  dismissSelectionMenu(menu) {
    const selectionKey = menu.getAttribute("data-selection-key") ?? "";
    this.selectionMenuState.dismissedSelectionKey = selectionKey;
    this.hideSelectionMenu();
  }

  destroySelectionMenu() {
    if (this.selectionMenuEl instanceof HTMLElement) {
      this.selectionMenuEl.remove();
      this.selectionMenuEl = null;
    }
  }

  async copySelectedAsText(target) {
    if (!this.isFeatureEnabled("copySelectedTextEnabled")) {
      this.showFeatureDisabledNotice("Copy selected as text");
      return false;
    }

    const now = Date.now();
    if (this.copyInFlight || now - this.lastCopyStartedAt < COPY_DEBOUNCE_MS) {
      this.log("copy skipped: already running or recently triggered");
      return false;
    }

    this.copyInFlight = true;
    this.lastCopyStartedAt = now;

    try {
      const liveTarget = this.normalizeCanvasTarget(target) ?? this.getCanvasTarget();
      if (liveTarget?.kind === "excalidraw") {
        return await this.copySelectedExcalidrawAsText(liveTarget);
      }

      const container = liveTarget?.container ?? this.getActiveTldrawContainer();
      const domText = this.getActiveDomTextSelection(container);
      if (domText) {
        await this.writeClipboardText(domText);
        this.log("copied active DOM text selection");
        new Notice("Copied selected text");
        return true;
      }

      const editor = liveTarget?.editor ?? this.getCurrentTldrawEditor(container);
      if (!editor) {
        this.log("copy failed: no active tldraw editor");
        new Notice("No active tldraw canvas");
        return false;
      }

      const selectedShapes = this.getSelectedLeafShapes(editor);
      if (!selectedShapes.length) {
        this.log("copy skipped: no selected shapes");
        new Notice("No tldraw selection to copy");
        return false;
      }

      const strokeText = await this.tryExtractStrokeText(editor, selectedShapes);
      if (strokeText) {
        await this.writeClipboardText(strokeText);
        this.log(`copied stroke-recognized text from ${selectedShapes.length} shapes`);
        new Notice("Copied selection as text");
        return true;
      }

      const directText = this.tryExtractDirectText(editor, selectedShapes);
      if (directText) {
        await this.writeClipboardText(directText);
        this.log(`copied direct text from ${selectedShapes.length} shapes`);
        new Notice("Copied selection as text");
        return true;
      }

      const extractedText = await this.extractSelectionTextWithOcr(editor, selectedShapes.map((shape) => shape.id));
      const normalizedText = this.normalizeCopiedText(extractedText);
      if (!normalizedText) {
        this.log("ocr produced no text");
        new Notice("No text found in selection");
        return false;
      }

      await this.writeClipboardText(normalizedText);
      this.log(`copied OCR text from ${selectedShapes.length} shapes`);
      new Notice("Copied selection as text");
      return true;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`copy failed: ${message}`);
      new Notice(this.getFailureNoticeMessage("Copy selected as text", error));
      return false;
    } finally {
      this.copyInFlight = false;
    }
  }

  async copySelectedAsScreenshot(target) {
    if (!this.isFeatureEnabled("selectionScreenshotEnabled")) {
      this.showFeatureDisabledNotice("Copy selected screenshot");
      return;
    }

    if (this.screenshotInFlight) {
      this.log("screenshot skipped: already running");
      return;
    }

    this.screenshotInFlight = true;

    try {
      const liveTarget = this.normalizeCanvasTarget(target) ?? this.getCanvasTarget();
      if (liveTarget?.kind === "excalidraw") {
        await this.copySelectedExcalidrawAsScreenshot(liveTarget);
        return;
      }

      const container = liveTarget?.container ?? this.getActiveTldrawContainer();
      const editor = liveTarget?.editor ?? this.getCurrentTldrawEditor(container);
      if (!editor) {
        this.log("screenshot failed: no active tldraw editor");
        new Notice("No active tldraw canvas");
        return;
      }

      const selectedShapes = this.getSelectedLeafShapes(editor);
      if (!selectedShapes.length) {
        this.log("screenshot skipped: no selected shapes");
        new Notice("No tldraw selection to screenshot");
        return;
      }

      const blob = await this.exportSelectionBlob(editor, selectedShapes.map((shape) => shape.id), {
        padding: SCREENSHOT_EXPORT_PADDING,
        scale: SCREENSHOT_EXPORT_SCALE,
        pixelRatio: SCREENSHOT_EXPORT_PIXEL_RATIO,
      });
      await this.writeClipboardImage(blob);
      this.log(`copied selection screenshot from ${selectedShapes.length} shapes`);
      new Notice("Copied selection screenshot");
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`screenshot failed: ${message}`);
      new Notice(this.getFailureNoticeMessage("Copy selected screenshot", error));
    } finally {
      this.screenshotInFlight = false;
    }
  }

  getActiveDomTextSelection(container) {
    const selection = document.getSelection();
    const rawText = selection?.toString();
    if (!rawText || !rawText.trim()) {
      return null;
    }

    if (!container) {
      return this.normalizeCopiedText(rawText);
    }

    const anchorWithin = this.nodeIsWithinContainer(selection?.anchorNode ?? null, container);
    const focusWithin = this.nodeIsWithinContainer(selection?.focusNode ?? null, container);
    if (!anchorWithin && !focusWithin) {
      return null;
    }

    return this.normalizeCopiedText(rawText);
  }

  nodeIsWithinContainer(node, container) {
    if (!(container instanceof HTMLElement) || !node) {
      return false;
    }

    if (node instanceof HTMLElement) {
      return container.contains(node);
    }

    return node.parentElement instanceof HTMLElement && container.contains(node.parentElement);
  }

  normalizeExcalidrawTarget(target) {
    if (!target) {
      return null;
    }

    const view = this.isExcalidrawView(target.view)
      ? target.view
      : this.getCurrentExcalidrawView(target.container ?? null);
    const api = this.getExcalidrawApi(target.api ? { excalidrawAPI: target.api } : view);
    const rawContainer =
      target.container instanceof HTMLElement
        ? target.container
        : this.getExcalidrawRootForView(view) ?? null;
    const container = this.findExcalidrawContainer(rawContainer) ?? rawContainer;

    if (!view && !(container instanceof HTMLElement)) {
      return null;
    }

    return { view, api, container };
  }

  getExcalidrawSelectedElements(view) {
    if (typeof view?.getViewSelectedElements === "function") {
      try {
        return Array.from(view.getViewSelectedElements(true) ?? []);
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`excalidraw selected elements read failed: ${message}`);
      }
    }

    const api = this.getExcalidrawApi(view);
    const selectedElementIds = api?.getAppState?.()?.selectedElementIds;
    if (!selectedElementIds || typeof api?.getSceneElements !== "function") {
      return [];
    }

    const selectedIds = new Set(Object.keys(selectedElementIds));
    return Array.from(api.getSceneElements() ?? []).filter((element) => selectedIds.has(element.id));
  }

  getExcalidrawSelectionMenuBounds(target) {
    const view = target?.view;
    const selectedElements = this.getExcalidrawSelectedElements(view);
    if (!selectedElements.length) {
      return null;
    }

    return this.getExcalidrawSelectionScreenBounds(view, selectedElements);
  }

  getExcalidrawSelectionScreenBounds(view, elements, padding = 0) {
    const root = this.getExcalidrawRootForView(view);
    const api = this.getExcalidrawApi(view);
    if (!(root instanceof HTMLElement) || !api || !elements.length) {
      return null;
    }

    const wrapperRect = root.getBoundingClientRect();
    if (wrapperRect.width <= 0 || wrapperRect.height <= 0) {
      return null;
    }

    const sceneBounds = this.getExcalidrawSceneBounds(view, elements);
    if (!sceneBounds) {
      return null;
    }

    const appState = api.getAppState?.();
    const zoom = Number(appState?.zoom?.value ?? 1) || 1;
    const scrollX = Number(appState?.scrollX ?? 0);
    const scrollY = Number(appState?.scrollY ?? 0);

    const x = wrapperRect.left + (sceneBounds.x + scrollX) * zoom - padding;
    const y = wrapperRect.top + (sceneBounds.y + scrollY) * zoom - padding;
    const w = sceneBounds.w * zoom + padding * 2;
    const h = sceneBounds.h * zoom + padding * 2;

    if (!Number.isFinite(x) || !Number.isFinite(y) || w <= 0 || h <= 0) {
      return null;
    }

    return { x, y, w, h };
  }

  getExcalidrawSceneBounds(view, elements) {
    const plugin = this.app?.plugins?.plugins?.[EXCALIDRAW_PLUGIN_ID];
    const ea = plugin?.ea;
    if (ea && typeof ea.setView === "function" && typeof ea.getBoundingBox === "function") {
      try {
        ea.setView(view);
        const bounds = ea.getBoundingBox(elements);
        if (bounds && Number.isFinite(bounds.width) && Number.isFinite(bounds.height)) {
          return {
            x: Number(bounds.topX ?? 0),
            y: Number(bounds.topY ?? 0),
            w: Number(bounds.width ?? 0),
            h: Number(bounds.height ?? 0),
          };
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`excalidraw bounding box via EA failed: ${message}`);
      }
    }

    let minX = Infinity;
    let minY = Infinity;
    let maxX = -Infinity;
    let maxY = -Infinity;

    for (const element of elements) {
      const bounds = this.getExcalidrawElementBounds(element);
      if (!bounds) {
        continue;
      }

      minX = Math.min(minX, bounds.x);
      minY = Math.min(minY, bounds.y);
      maxX = Math.max(maxX, bounds.x + bounds.w);
      maxY = Math.max(maxY, bounds.y + bounds.h);
    }

    if (!Number.isFinite(minX) || !Number.isFinite(minY) || !Number.isFinite(maxX) || !Number.isFinite(maxY)) {
      return null;
    }

    return {
      x: minX,
      y: minY,
      w: Math.max(0, maxX - minX),
      h: Math.max(0, maxY - minY),
    };
  }

  getExcalidrawElementBounds(element) {
    if (!element) {
      return null;
    }

    const x1 = Number(element.x ?? 0);
    const y1 = Number(element.y ?? 0);
    const x2 = x1 + Number(element.width ?? 0);
    const y2 = y1 + Number(element.height ?? 0);
    const left = Math.min(x1, x2);
    const top = Math.min(y1, y2);
    const width = Math.abs(x2 - x1);
    const height = Math.abs(y2 - y1);
    const angle = Number(element.angle ?? 0);

    if (!Number.isFinite(angle) || Math.abs(angle) < 0.0001 || width === 0 || height === 0) {
      return { x: left, y: top, w: width, h: height };
    }

    const cx = left + width / 2;
    const cy = top + height / 2;
    const corners = [
      { x: left, y: top },
      { x: left + width, y: top },
      { x: left + width, y: top + height },
      { x: left, y: top + height },
    ].map((corner) => this.rotatePoint(corner, { x: cx, y: cy }, angle));

    const minX = Math.min(...corners.map((corner) => corner.x));
    const minY = Math.min(...corners.map((corner) => corner.y));
    const maxX = Math.max(...corners.map((corner) => corner.x));
    const maxY = Math.max(...corners.map((corner) => corner.y));
    return {
      x: minX,
      y: minY,
      w: maxX - minX,
      h: maxY - minY,
    };
  }

  rotatePoint(point, center, angle) {
    const dx = Number(point.x ?? 0) - Number(center.x ?? 0);
    const dy = Number(point.y ?? 0) - Number(center.y ?? 0);
    const cos = Math.cos(angle);
    const sin = Math.sin(angle);
    return {
      x: Number(center.x ?? 0) + dx * cos - dy * sin,
      y: Number(center.y ?? 0) + dx * sin + dy * cos,
    };
  }

  async copySelectedExcalidrawAsText(target) {
    const container = target?.container ?? this.getActiveExcalidrawContainer();
    const domText = this.getActiveDomTextSelection(container);
    if (domText) {
      await this.writeClipboardText(domText);
      this.log("copied active Excalidraw DOM text selection");
      new Notice("Copied selected text");
      return true;
    }

    const view = target?.view ?? this.getCurrentExcalidrawView(container);
    if (!view) {
      this.log("copy failed: no active Excalidraw view");
      new Notice("No active Excalidraw canvas");
      return false;
    }

    const selectedElements = this.getExcalidrawSelectedElements(view);
    if (!selectedElements.length) {
      this.log("copy skipped: no selected Excalidraw elements");
      new Notice("No Excalidraw selection to copy");
      return false;
    }

    const strokeText = await this.tryExtractExcalidrawStrokeText(view, selectedElements);
    if (strokeText) {
      await this.writeClipboardText(strokeText);
      this.log(`copied Excalidraw stroke text from ${selectedElements.length} elements`);
      new Notice("Copied selection as text");
      return true;
    }

    const directText = this.tryExtractExcalidrawDirectText(selectedElements);
    if (directText) {
      await this.writeClipboardText(directText);
      this.log(`copied Excalidraw direct text from ${selectedElements.length} elements`);
      new Notice("Copied selection as text");
      return true;
    }

    const blob = await this.captureExcalidrawSelectionBlob(view, {
      padding: OCR_EXPORT_PADDING,
      scale: OCR_EXPORT_SCALE,
      pixelRatio: OCR_EXPORT_PIXEL_RATIO,
    });
    const extractedText = await this.extractBlobTextWithOcr(blob);
    const normalizedText = this.normalizeCopiedText(extractedText);
    if (!normalizedText) {
      this.log("Excalidraw OCR produced no text");
      new Notice("No text found in selection");
      return false;
    }

    await this.writeClipboardText(normalizedText);
    this.log(`copied Excalidraw OCR text from ${selectedElements.length} elements`);
    new Notice("Copied selection as text");
    return true;
  }

  async copySelectedExcalidrawAsScreenshot(target) {
    const view = target?.view ?? this.getCurrentExcalidrawView(target?.container ?? null);
    if (!view) {
      this.log("screenshot failed: no active Excalidraw view");
      new Notice("No active Excalidraw canvas");
      return;
    }

    const selectedElements = this.getExcalidrawSelectedElements(view);
    if (!selectedElements.length) {
      this.log("screenshot skipped: no selected Excalidraw elements");
      new Notice("No Excalidraw selection to screenshot");
      return;
    }

    const blob = await this.captureExcalidrawSelectionBlob(view, {
      padding: SCREENSHOT_EXPORT_PADDING,
      scale: SCREENSHOT_EXPORT_SCALE,
      pixelRatio: SCREENSHOT_EXPORT_PIXEL_RATIO,
    });
    await this.writeClipboardImage(blob);
    this.log(`copied Excalidraw selection screenshot from ${selectedElements.length} elements`);
    new Notice("Copied selection screenshot");
  }

  tryExtractExcalidrawDirectText(elements) {
    const entries = [];

    for (const element of elements) {
      if (!this.isExcalidrawDirectTextElement(element)) {
        return null;
      }

      const text = this.getExcalidrawElementPlainText(element);
      if (!text) {
        continue;
      }

      const bounds = this.getExcalidrawElementBounds(element);
      if (!bounds) {
        return null;
      }

      entries.push({ text, bounds });
    }

    if (!entries.length) {
      return null;
    }

    return this.layoutEntriesAsText(entries);
  }

  isExcalidrawDirectTextElement(element) {
    return Boolean(element && EXCALIDRAW_DIRECT_TEXT_TYPES.has(element.type));
  }

  getExcalidrawElementPlainText(element) {
    return this.normalizeCopiedText(typeof element?.text === "string" ? element.text : "") ?? "";
  }

  async tryExtractExcalidrawStrokeText(view, elements) {
    const payload = this.buildExcalidrawStrokeRecognitionPayload(view, elements);
    if (!payload) {
      return null;
    }

    try {
      const result = await this.runInkStrokeRecognizer(payload);
      const strokeText = this.normalizeCopiedText(result?.text);
      this.log(
        `excalidraw ink recognizer status=${result?.status ?? "unknown"} confidence=${result?.confidence ?? "unknown"} strokes=${payload.strokes.length}`
      );

      if (!strokeText) {
        return null;
      }

      if (!this.shouldTrustStrokeRecognition(result, strokeText, payload)) {
        this.log("excalidraw ink recognizer deferred to image transcription");
        return null;
      }

      return this.mergeStrokeTextWithDirectEntries(strokeText, payload);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`excalidraw ink stroke recognizer failed: ${message}`);
      return null;
    }
  }

  buildExcalidrawStrokeRecognitionPayload(view, elements) {
    const directEntries = [];
    const drawStrokes = [];
    let drawBounds = null;

    const sortedElements = [...elements].sort((left, right) => {
      const leftBounds = this.getExcalidrawElementBounds(left);
      const rightBounds = this.getExcalidrawElementBounds(right);
      const leftTop = Number(leftBounds?.y ?? 0);
      const rightTop = Number(rightBounds?.y ?? 0);
      if (Math.abs(leftTop - rightTop) > 4) {
        return leftTop - rightTop;
      }

      return Number(leftBounds?.x ?? 0) - Number(rightBounds?.x ?? 0);
    });

    for (const element of sortedElements) {
      const bounds = this.getExcalidrawElementBounds(element);
      if (!bounds) {
        continue;
      }

      if (element.type === "freedraw") {
        const strokeSegments = this.extractExcalidrawFreedrawStrokes(element);
        if (!strokeSegments.length) {
          continue;
        }

        drawStrokes.push(...strokeSegments);
        drawBounds = drawBounds ? this.mergeBounds(drawBounds, bounds) : this.cloneBounds(bounds);
        continue;
      }

      if (this.isExcalidrawDirectTextElement(element)) {
        const text = this.getExcalidrawElementPlainText(element);
        if (!text) {
          continue;
        }

        directEntries.push({ text, bounds });
      }
    }

    if (!drawStrokes.length) {
      return null;
    }

    return {
      mode: "text",
      strokes: drawStrokes,
      alternateCount: 5,
      directEntries,
      drawBounds,
    };
  }

  extractExcalidrawFreedrawStrokes(element) {
    const points = Array.isArray(element?.points) ? element.points : [];
    if (!points.length) {
      return [];
    }

    const pressures = Array.isArray(element?.pressures) ? element.pressures : [];
    const absolutePoints = points.map((point, index) => ({
      x: Number(element.x ?? 0) + Number(point?.[0] ?? 0),
      y: Number(element.y ?? 0) + Number(point?.[1] ?? 0),
      z: Number(pressures[index] ?? 0.5),
    }));

    return absolutePoints.length > 0 ? [{ points: absolutePoints }] : [];
  }

  async captureExcalidrawSelectionBlob(view, options = {}) {
    const selectedElements = this.getExcalidrawSelectedElements(view);
    if (!selectedElements.length) {
      throw new Error("No Excalidraw selection to capture");
    }

    const plugin = this.app?.plugins?.plugins?.[EXCALIDRAW_PLUGIN_ID];
    const ea = plugin?.ea;
    if (ea && typeof ea.setView === "function" && typeof ea.createViewSVG === "function") {
      try {
        ea.setView(view);
        const api = this.getExcalidrawApi(view);
        const theme = api?.getAppState?.()?.theme === "dark" ? "dark" : "light";
        const svg = await ea.createViewSVG({
          withBackground: true,
          theme,
          padding: Number(options.padding ?? 0),
          selectedOnly: true,
          skipInliningFonts: false,
          embedScene: false,
        });
        return await this.renderSvgElementToPngBlob(svg, {
          scale: Number(options.scale ?? 1),
          pixelRatio: Number(options.pixelRatio ?? window.devicePixelRatio ?? 1),
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`Excalidraw SVG export failed, falling back to screen capture: ${message}`);
      }
    }

    if (typeof window?.require !== "function") {
      throw new Error("Excalidraw export is not available");
    }

    const wrapper = this.getExcalidrawRootForView(view);
    if (!(wrapper instanceof HTMLElement)) {
      throw new Error("Excalidraw wrapper was not found");
    }

    const bounds = this.getExcalidrawSelectionScreenBounds(
      view,
      selectedElements,
      Number(options.padding ?? 0)
    );
    if (!bounds) {
      throw new Error("Could not determine Excalidraw selection bounds");
    }

    const layerUi = wrapper.querySelector(".layer-ui__wrapper");
    const bottomBar = wrapper.querySelector(".App-bottom-bar");
    const restoreLayerUi = layerUi instanceof HTMLElement ? layerUi.style.display : null;
    const restoreBottomBar = bottomBar instanceof HTMLElement ? bottomBar.style.display : null;

    try {
      if (layerUi instanceof HTMLElement) {
        layerUi.style.display = "none";
      }
      if (bottomBar instanceof HTMLElement) {
        bottomBar.style.display = "none";
      }

      await new Promise((resolve) => window.setTimeout(resolve, 40));
      const { remote } = window.require("electron");
      const webContents = remote?.getCurrentWebContents?.();
      if (!webContents || typeof webContents.capturePage !== "function") {
        throw new Error("Electron capture API is not available");
      }

      const dpr = Number(window.devicePixelRatio ?? 1) || 1;
      const maxWidth = Math.max(1, Math.round(window.innerWidth * dpr));
      const maxHeight = Math.max(1, Math.round(window.innerHeight * dpr));
      const captureRect = {
        x: Math.max(0, Math.round(bounds.x * dpr)),
        y: Math.max(0, Math.round(bounds.y * dpr)),
        width: Math.max(1, Math.round(bounds.w * dpr)),
        height: Math.max(1, Math.round(bounds.h * dpr)),
      };
      captureRect.width = Math.min(captureRect.width, Math.max(1, maxWidth - captureRect.x));
      captureRect.height = Math.min(captureRect.height, Math.max(1, maxHeight - captureRect.y));

      const image = await webContents.capturePage(captureRect);
      const png = image?.toPNG?.();
      if (!(png instanceof Buffer) || png.length === 0) {
        throw new Error("Excalidraw selection capture returned an empty image");
      }

      return new Blob([png], { type: "image/png" });
    } finally {
      if (layerUi instanceof HTMLElement && restoreLayerUi !== null) {
        layerUi.style.display = restoreLayerUi;
      }
      if (bottomBar instanceof HTMLElement && restoreBottomBar !== null) {
        bottomBar.style.display = restoreBottomBar;
      }
    }
  }

  async renderSvgElementToPngBlob(svgElement, options = {}) {
    if (!(svgElement instanceof SVGElement)) {
      throw new Error("Excalidraw SVG export returned an invalid element");
    }

    const size = this.getSvgViewportSize(svgElement);
    if (!size) {
      throw new Error("Excalidraw SVG export did not include a valid size");
    }

    const scale = Math.max(0.25, Number(options.scale ?? 1) || 1);
    const pixelRatio = Math.max(1, Number(options.pixelRatio ?? 1) || 1);
    const renderScale = scale * pixelRatio;
    const svgClone = svgElement.cloneNode(true);
    if (!(svgClone instanceof SVGElement)) {
      throw new Error("Excalidraw SVG export could not be cloned");
    }

    if (!svgClone.getAttribute("xmlns")) {
      svgClone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
    }
    if (!svgClone.getAttribute("xmlns:xlink")) {
      svgClone.setAttribute("xmlns:xlink", "http://www.w3.org/1999/xlink");
    }
    svgClone.setAttribute("width", String(size.width));
    svgClone.setAttribute("height", String(size.height));

    const serializer = new XMLSerializer();
    const svgMarkup = serializer.serializeToString(svgClone);
    const svgBlob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
    const svgUrl = URL.createObjectURL(svgBlob);

    try {
      const image = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error("Failed to load Excalidraw SVG export"));
        img.decoding = "async";
        img.src = svgUrl;
      });

      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.ceil(size.width * renderScale));
      canvas.height = Math.max(1, Math.ceil(size.height * renderScale));
      const context = canvas.getContext("2d");
      if (!context) {
        throw new Error("Canvas 2D context is not available");
      }

      context.setTransform(renderScale, 0, 0, renderScale, 0, 0);
      context.imageSmoothingEnabled = true;
      context.drawImage(image, 0, 0, size.width, size.height);

      const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!(pngBlob instanceof Blob) || pngBlob.size === 0) {
        throw new Error("Excalidraw SVG export rasterized to an empty image");
      }

      return pngBlob;
    } finally {
      URL.revokeObjectURL(svgUrl);
    }
  }

  async renderSvgStringToPngBlob(svgMarkup, options = {}) {
    if (typeof svgMarkup !== "string" || svgMarkup.trim() === "") {
      throw new Error("SVG export did not include any markup");
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(svgMarkup, "image/svg+xml");
    const svgElement = doc.documentElement instanceof SVGElement ? doc.documentElement : null;
    if (!(svgElement instanceof SVGElement)) {
      throw new Error("SVG export could not be parsed");
    }

    const width = Number(options.width ?? 0);
    const height = Number(options.height ?? 0);
    const size =
      width > 0 && height > 0
        ? { width, height }
        : this.getSvgViewportSize(svgElement);
    if (!size) {
      throw new Error("SVG export did not include a valid size");
    }

    const scale = Math.max(0.25, Number(options.scale ?? 1) || 1);
    const pixelRatio = Math.max(1, Number(options.pixelRatio ?? 1) || 1);
    const renderScale = scale * pixelRatio;
    const svgBlob = new Blob([svgMarkup], { type: "image/svg+xml;charset=utf-8" });
    const svgUrl = URL.createObjectURL(svgBlob);

    try {
      const image = await new Promise((resolve, reject) => {
        const img = new Image();
        img.onload = () => resolve(img);
        img.onerror = () => reject(new Error("Failed to load SVG export"));
        img.decoding = "async";
        img.src = svgUrl;
      });

      const canvas = document.createElement("canvas");
      canvas.width = Math.max(1, Math.ceil(size.width * renderScale));
      canvas.height = Math.max(1, Math.ceil(size.height * renderScale));
      const context = canvas.getContext("2d");
      if (!context) {
        throw new Error("Canvas 2D context is not available");
      }

      context.setTransform(renderScale, 0, 0, renderScale, 0, 0);
      context.imageSmoothingEnabled = true;
      context.drawImage(image, 0, 0, size.width, size.height);

      const pngBlob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
      if (!(pngBlob instanceof Blob) || pngBlob.size === 0) {
        throw new Error("SVG export rasterized to an empty image");
      }

      return pngBlob;
    } finally {
      URL.revokeObjectURL(svgUrl);
    }
  }

  getSvgViewportSize(svgElement) {
    const viewBox = svgElement?.viewBox?.baseVal;
    if (viewBox && Number.isFinite(viewBox.width) && Number.isFinite(viewBox.height) && viewBox.width > 0 && viewBox.height > 0) {
      return {
        width: Number(viewBox.width),
        height: Number(viewBox.height),
      };
    }

    const width = Number.parseFloat(svgElement?.getAttribute?.("width") ?? "");
    const height = Number.parseFloat(svgElement?.getAttribute?.("height") ?? "");
    if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
      return { width, height };
    }

    try {
      const bbox = typeof svgElement?.getBBox === "function" ? svgElement.getBBox() : null;
      if (bbox && Number.isFinite(bbox.width) && Number.isFinite(bbox.height) && bbox.width > 0 && bbox.height > 0) {
        return {
          width: Number(bbox.width),
          height: Number(bbox.height),
        };
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`Excalidraw SVG size fallback failed: ${message}`);
    }

    return null;
  }

  getSelectedLeafShapes(editor) {
    if (typeof editor.getSelectedShapeIds !== "function" || typeof editor.getShape !== "function") {
      return [];
    }

    const selectedShapeIds = Array.from(editor.getSelectedShapeIds() ?? []);
    const seen = new Set();
    const leafShapes = [];
    const queue = [...selectedShapeIds];

    while (queue.length > 0) {
      const shapeId = queue.shift();
      if (!shapeId || seen.has(shapeId)) {
        continue;
      }

      seen.add(shapeId);
      const shape = editor.getShape(shapeId);
      if (!shape) {
        continue;
      }

      if (shape.type === "group" && typeof editor.getSortedChildIdsForParent === "function") {
        const childIds = editor.getSortedChildIdsForParent(shape.id) ?? [];
        queue.unshift(...childIds);
        continue;
      }

      leafShapes.push(shape);
    }

    return leafShapes;
  }

  tryExtractDirectText(editor, shapes) {
    const entries = [];

    for (const shape of shapes) {
      if (!this.isDirectTextShape(shape)) {
        return null;
      }

      const text = this.getShapePlainText(editor, shape);
      if (!text) {
        continue;
      }

      const bounds = typeof editor.getShapePageBounds === "function" ? editor.getShapePageBounds(shape.id) : null;
      if (!bounds) {
        return null;
      }

      entries.push({ text, bounds });
    }

    if (!entries.length) {
      return null;
    }

    return this.layoutEntriesAsText(entries);
  }

  isDirectTextShape(shape) {
    return Boolean(shape && DIRECT_TEXT_SHAPE_TYPES.has(shape.type));
  }

  getShapePlainText(editor, shape) {
    try {
      const util = editor.getShapeUtil(shape);
      if (!util || typeof util.getText !== "function") {
        return "";
      }

      return this.normalizeCopiedText(util.getText(shape)) ?? "";
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`shape text read failed for ${shape?.id ?? "unknown"}: ${message}`);
      return "";
    }
  }

  layoutEntriesAsText(entries) {
    const shapedEntries = entries
      .map((entry) => ({
        text: entry.text,
        top: Number(entry.bounds.y ?? 0),
        left: Number(entry.bounds.x ?? 0),
        bottom: Number((entry.bounds.y ?? 0) + (entry.bounds.h ?? 0)),
      }))
      .sort((a, b) => {
        if (Math.abs(a.top - b.top) > 4) {
          return a.top - b.top;
        }
        return a.left - b.left;
      });

    const rows = [];

    for (const entry of shapedEntries) {
      const currentRow = rows[rows.length - 1];
      if (!currentRow || entry.top > currentRow.bottom - 6) {
        rows.push({
          bottom: entry.bottom,
          entries: [entry],
        });
        continue;
      }

      currentRow.bottom = Math.max(currentRow.bottom, entry.bottom);
      currentRow.entries.push(entry);
    }

    const text = rows
      .map((row) => row.entries.sort((a, b) => a.left - b.left).map((entry) => entry.text).join(" "))
      .join("\n");

    return this.normalizeCopiedText(text);
  }

  async tryExtractStrokeText(editor, shapes) {
    const payload = this.buildStrokeRecognitionPayload(editor, shapes);
    if (!payload) {
      return null;
    }

    try {
      const result = await this.runInkStrokeRecognizer(payload);
      const strokeText = this.normalizeCopiedText(result?.text);
      this.log(
        `ink stroke recognizer status=${result?.status ?? "unknown"} confidence=${result?.confidence ?? "unknown"} strokes=${payload.strokes.length}`
      );

      if (!strokeText) {
        return null;
      }

      if (!this.shouldTrustStrokeRecognition(result, strokeText, payload)) {
        this.log("ink stroke recognizer deferred to image transcription");
        return null;
      }

      return this.mergeStrokeTextWithDirectEntries(strokeText, payload);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`ink stroke recognizer failed: ${message}`);
      return null;
    }
  }

  buildStrokeRecognitionPayload(editor, shapes) {
    const directEntries = [];
    const drawStrokes = [];
    let drawBounds = null;

    const sortedShapes = [...shapes].sort((left, right) => {
      const leftBounds = typeof editor.getShapePageBounds === "function" ? editor.getShapePageBounds(left.id) : null;
      const rightBounds = typeof editor.getShapePageBounds === "function" ? editor.getShapePageBounds(right.id) : null;
      const leftTop = Number(leftBounds?.y ?? 0);
      const rightTop = Number(rightBounds?.y ?? 0);
      if (Math.abs(leftTop - rightTop) > 4) {
        return leftTop - rightTop;
      }

      return Number(leftBounds?.x ?? 0) - Number(rightBounds?.x ?? 0);
    });

    for (const shape of sortedShapes) {
      const bounds = typeof editor.getShapePageBounds === "function" ? editor.getShapePageBounds(shape.id) : null;
      if (!bounds) {
        continue;
      }

      if (shape.type === "draw") {
        const strokeSegments = this.extractDrawShapeStrokes(editor, shape);
        if (!strokeSegments.length) {
          continue;
        }

        drawStrokes.push(...strokeSegments);
        drawBounds = drawBounds ? this.mergeBounds(drawBounds, bounds) : this.cloneBounds(bounds);
        continue;
      }

      if (this.isDirectTextShape(shape)) {
        const text = this.getShapePlainText(editor, shape);
        if (!text) {
          continue;
        }

        directEntries.push({ text, bounds });
      }
    }

    if (!drawStrokes.length) {
      return null;
    }

    return {
      mode: "text",
      strokes: drawStrokes,
      alternateCount: 5,
      directEntries,
      drawBounds,
    };
  }

  mergeStrokeTextWithDirectEntries(strokeText, payload) {
    if (!payload || payload.directEntries.length === 0) {
      return strokeText;
    }

    return this.layoutEntriesAsText([
      ...payload.directEntries,
      {
        text: strokeText,
        bounds: payload.drawBounds,
      },
    ]);
  }

  shouldTrustStrokeRecognition(result, strokeText, payload) {
    if (!strokeText) {
      return false;
    }

    if (this.isSuspiciousRecognitionText(strokeText, payload?.strokes?.length ?? 0)) {
      return false;
    }

    const confidence = `${result?.confidence ?? ""}`.trim().toLowerCase();
    if (confidence === "strong") {
      return true;
    }

    if (confidence === "intermediate") {
      return true;
    }

    const strokeCount = Number(payload?.strokes?.length ?? 0);
    return strokeCount > 0 && strokeCount <= 2;
  }

  isSuspiciousRecognitionText(text, strokeCount) {
    const normalized = this.normalizeCopiedText(text);
    if (!normalized) {
      return true;
    }

    const compact = normalized.replace(/\s+/g, "");
    if (!compact) {
      return true;
    }

    const letterDigitCount = (compact.match(/[A-Za-z0-9]/g) ?? []).length;
    const mathCount = (compact.match(/[+\-*/=()[\]{}<>^_%]/g) ?? []).length;
    const suspiciousCount = (compact.match(/[@\\`~]/g) ?? []).length;
    const unsupportedCount = (compact.match(/[^A-Za-z0-9\s.,;:!?'"()[\]{}<>+\-*/=^_%&$#|]/g) ?? []).length;
    const compactLength = compact.length;
    const weirdRatio = (unsupportedCount + suspiciousCount) / compactLength;

    if (strokeCount >= 4 && compactLength <= 1) {
      return true;
    }

    if (letterDigitCount === 0 && mathCount === 0) {
      return true;
    }

    if (strokeCount >= 6 && letterDigitCount + mathCount <= 2) {
      return true;
    }

    if (weirdRatio >= 0.2) {
      return true;
    }

    return suspiciousCount >= 2 && letterDigitCount <= Math.max(1, Math.floor(compactLength / 3));
  }

  extractDrawShapeStrokes(editor, shape) {
    if (!shape?.props?.segments || typeof editor.getShapePageTransform !== "function") {
      return [];
    }

    const transform = editor.getShapePageTransform(shape.id);
    if (!transform || typeof transform.applyToPoint !== "function") {
      return [];
    }

    const strokes = [];
    for (const segment of shape.props.segments) {
      const localPoints = decodeDrawSegmentPath(segment?.path);
      if (localPoints.length === 0) {
        continue;
      }

      const pagePoints = localPoints.map((point) => {
        const transformed = transform.applyToPoint(point);
        return {
          x: Number(transformed?.x ?? point.x),
          y: Number(transformed?.y ?? point.y),
          z: Number(point.z ?? 0.5),
        };
      });

      if (pagePoints.length > 0) {
        strokes.push({ points: pagePoints });
      }
    }

    return strokes;
  }

  cloneBounds(bounds) {
    return {
      x: Number(bounds.x ?? 0),
      y: Number(bounds.y ?? 0),
      w: Number(bounds.w ?? 0),
      h: Number(bounds.h ?? 0),
    };
  }

  mergeBounds(left, right) {
    const minX = Math.min(Number(left.x ?? 0), Number(right.x ?? 0));
    const minY = Math.min(Number(left.y ?? 0), Number(right.y ?? 0));
    const maxX = Math.max(Number(left.x ?? 0) + Number(left.w ?? 0), Number(right.x ?? 0) + Number(right.w ?? 0));
    const maxY = Math.max(Number(left.y ?? 0) + Number(left.h ?? 0), Number(right.y ?? 0) + Number(right.h ?? 0));
    return {
      x: minX,
      y: minY,
      w: maxX - minX,
      h: maxY - minY,
    };
  }

  async runInkStrokeRecognizer(payload) {
    await this.ensureTempExportFolder();

    const requestPath = this.getVaultAbsolutePath(STROKE_REQUEST_FILE);
    const helperPath = this.getVaultAbsolutePath(
      normalizePath(`${this.app.vault.configDir}/plugins/writing-bridge/runtime/${INK_HELPER_EXE}`)
    );

    if (!fs.existsSync(helperPath)) {
      throw new Error("Ink stroke recognizer helper is missing");
    }

    fs.writeFileSync(requestPath, JSON.stringify(payload), "utf8");

    try {
      const stdout = await new Promise((resolve, reject) => {
        execFile(
          helperPath,
          [requestPath],
          { windowsHide: true, maxBuffer: 1024 * 1024, encoding: "utf8" },
          (error, stdout2, stderr2) => {
            const stdoutText = typeof stdout2 === "string" ? stdout2 : `${stdout2 ?? ""}`;
            if (error && stdoutText.trim() !== "") {
              resolve(stdoutText);
              return;
            }

            if (error) {
              const stderrText = typeof stderr2 === "string" && stderr2.trim() !== "" ? stderr2.trim() : error.message;
              reject(new Error(stderrText));
              return;
            }

            resolve(stdoutText);
          }
        );
      });

      const parsed = JSON.parse(stdout);
      if (!parsed?.ok && parsed?.error) {
        throw new Error(parsed.error);
      }

      return parsed;
    } finally {
      if (fs.existsSync(requestPath)) {
        fs.unlinkSync(requestPath);
      }
    }
  }

  async extractSelectionTextWithOcr(editor, shapeIds) {
    const blob = await this.exportSelectionBlob(editor, shapeIds, {
      padding: OCR_EXPORT_PADDING,
      scale: OCR_EXPORT_SCALE,
      pixelRatio: OCR_EXPORT_PIXEL_RATIO,
    });

    return this.extractBlobTextWithOcr(blob);
  }

  async extractBlobTextWithOcr(blob) {
    const failures = [];

    for (const provider of this.getPreferredVisionProviders()) {
      const providerLabel = this.getVisionProviderLabel(provider);
      try {
        const text = await this.extractBlobTextWithPreferredVisionProvider(blob, provider);
        if (this.normalizeCopiedText(text)) {
          this.log(`ocr backend=${provider}`);
          return text;
        }

        this.pushFailureDetail(failures, providerLabel, "returned an empty transcription");
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`${provider} vision failed: ${message}`);
        this.pushFailureDetail(failures, providerLabel, error);
      }
    }

    const mistralApiKey = this.getConfiguredMistralApiKey();
    if (mistralApiKey) {
      try {
        const text = await this.extractSelectionTextWithMistralVision(blob, mistralApiKey);
        if (this.normalizeCopiedText(text)) {
          this.log("ocr backend=mistral-vision");
          return text;
        }

        this.pushFailureDetail(failures, "Mistral vision", "returned an empty transcription");
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`mistral vision failed: ${message}`);
        this.pushFailureDetail(failures, "Mistral vision", error);
      }

      try {
        const text = await this.extractSelectionTextWithMistral(blob, mistralApiKey);
        if (this.normalizeCopiedText(text)) {
          this.log("ocr backend=mistral");
          return text;
        }

        this.pushFailureDetail(failures, "Mistral OCR", "returned an empty transcription");
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        this.log(`mistral ocr failed: ${message}`);
        this.pushFailureDetail(failures, "Mistral OCR", error);
      }
    }

    const textExtractor = this.app?.plugins?.plugins?.[TEXT_EXTRACTOR_PLUGIN_ID];
    const extractText = textExtractor?.api?.extractText;
    if (typeof extractText !== "function") {
      this.pushFailureDetail(failures, "Text Extractor", "plugin is not available");
      throw new Error(`OCR transcription failed: ${failures.join("; ")}`);
    }

    const file = await this.writeTempExport(blob);
    await this.clearTextExtractorCache(file.path);

    try {
      const text = await extractText.call(textExtractor.api, file);
      if (this.normalizeCopiedText(text)) {
        this.log("ocr backend=text-extractor");
        return text;
      }

      this.pushFailureDetail(failures, "Text Extractor", "returned no text");
    } catch (error) {
      this.pushFailureDetail(failures, "Text Extractor", error);
    } finally {
      await this.clearTextExtractorCache(file.path);
    }

    throw new Error(
      `OCR transcription failed: ${failures.join("; ") || "no transcription backend returned any text"}`
    );
  }

  async extractBlobTextWithPreferredVisionProvider(blob, provider) {
    if (provider === VISION_PROVIDER_GEMINI) {
      const apiKey = this.getConfiguredGeminiApiKey();
      if (!apiKey) {
        throw new Error("Gemini API key is not configured");
      }

      return this.extractSelectionTextWithGemini(blob, apiKey, this.getConfiguredGeminiModel());
    }

    if (provider === VISION_PROVIDER_OPENROUTER) {
      const apiKey = this.getConfiguredOpenRouterApiKey();
      if (!apiKey) {
        throw new Error("OpenRouter API key is not configured");
      }

      return this.extractSelectionTextWithOpenRouter(blob, apiKey);
    }

    if (provider === VISION_PROVIDER_CUSTOM) {
      const baseUrl = this.getConfiguredCustomOpenAiBaseUrl();
      if (!baseUrl) {
        throw new Error("Custom OpenAI base URL is not configured");
      }

      const model = this.getConfiguredCustomOpenAiModel();
      if (!model) {
        throw new Error("Custom OpenAI model is not configured");
      }

      return this.extractSelectionTextWithCustomOpenAi(
        blob,
        baseUrl,
        this.getConfiguredCustomOpenAiApiKey(),
        model
      );
    }

    throw new Error(`Unsupported vision provider: ${provider}`);
  }

  getPreferredVisionProviders() {
    const preferred = this.getConfiguredVisionProvider();
    const ordered = [preferred, ...VISION_PROVIDERS.filter((p) => p !== preferred)];
    const seen = new Set();
    const result = [];
    for (const provider of ordered) {
      if (seen.has(provider)) {
        continue;
      }
      seen.add(provider);
      if (this.isVisionProviderConfigured(provider)) {
        result.push(provider);
      }
    }
    return result;
  }

  isVisionProviderConfigured(provider) {
    if (provider === VISION_PROVIDER_OPENROUTER) {
      return Boolean(this.getConfiguredOpenRouterApiKey());
    }

    if (provider === VISION_PROVIDER_GEMINI) {
      return Boolean(this.getConfiguredGeminiApiKey());
    }

    if (provider === VISION_PROVIDER_CUSTOM) {
      return Boolean(this.getConfiguredCustomOpenAiBaseUrl()) && Boolean(this.getConfiguredCustomOpenAiModel());
    }

    return false;
  }

  getConfiguredVisionProvider() {
    const value = this.settings?.visionLlmProvider;
    return VISION_PROVIDERS.includes(value) ? value : VISION_PROVIDER_OPENROUTER;
  }

  getVisionProviderLabel(provider) {
    if (provider === VISION_PROVIDER_GEMINI) {
      return "Gemini";
    }

    if (provider === VISION_PROVIDER_OPENROUTER) {
      return "OpenRouter";
    }

    if (provider === VISION_PROVIDER_CUSTOM) {
      return "Custom OpenAI";
    }

    return String(provider ?? "Unknown provider");
  }

  getConfiguredOpenRouterApiKey() {
    const fromSettings =
      typeof this.settings?.openRouterApiKey === "string" ? this.settings.openRouterApiKey.trim() : "";
    if (fromSettings !== "") {
      return fromSettings;
    }

    const envValue =
      typeof process?.env?.OPENROUTER_API_KEY === "string" ? process.env.OPENROUTER_API_KEY.trim() : "";
    return envValue !== "" ? envValue : null;
  }

  getConfiguredGeminiApiKey() {
    const fromSettings =
      typeof this.settings?.geminiApiKey === "string" ? this.settings.geminiApiKey.trim() : "";
    if (fromSettings !== "") {
      return fromSettings;
    }

    const envValue =
      typeof process?.env?.GEMINI_API_KEY === "string" ? process.env.GEMINI_API_KEY.trim() : "";
    return envValue !== "" ? envValue : null;
  }

  getConfiguredOpenRouterModels() {
    const raw =
      typeof this.settings?.openRouterModels === "string" && this.settings.openRouterModels.trim() !== ""
        ? this.settings.openRouterModels
        : OPENROUTER_DEFAULT_MODELS.join("\n");

    const parsed = raw
      .split(/[\r\n,]+/)
      .map((value) => value.trim())
      .filter(Boolean);

    return parsed.length > 0 ? parsed : [...OPENROUTER_DEFAULT_MODELS];
  }

  getConfiguredGeminiModel() {
    const value =
      typeof this.settings?.geminiModel === "string" ? this.settings.geminiModel.trim() : "";
    return value !== "" ? value : GEMINI_DEFAULT_MODEL;
  }

  getConfiguredCustomOpenAiBaseUrl() {
    const fromSettings =
      typeof this.settings?.customOpenAiBaseUrl === "string" ? this.settings.customOpenAiBaseUrl.trim() : "";
    if (fromSettings !== "") {
      return fromSettings;
    }

    const envValue =
      typeof process?.env?.CUSTOM_OPENAI_BASE_URL === "string"
        ? process.env.CUSTOM_OPENAI_BASE_URL.trim()
        : "";
    return envValue !== "" ? envValue : null;
  }

  getConfiguredCustomOpenAiApiKey() {
    const fromSettings =
      typeof this.settings?.customOpenAiApiKey === "string" ? this.settings.customOpenAiApiKey.trim() : "";
    if (fromSettings !== "") {
      return fromSettings;
    }

    const envValue =
      typeof process?.env?.CUSTOM_OPENAI_API_KEY === "string"
        ? process.env.CUSTOM_OPENAI_API_KEY.trim()
        : "";
    return envValue !== "" ? envValue : null;
  }

  getConfiguredCustomOpenAiModel() {
    const value =
      typeof this.settings?.customOpenAiModel === "string" ? this.settings.customOpenAiModel.trim() : "";
    return value !== "" ? value : null;
  }

  resolveCustomOpenAiChatCompletionsUrl(baseUrl) {
    const trimmed = String(baseUrl ?? "").trim().replace(/\/+$/, "");
    if (trimmed === "") {
      throw new Error("Custom OpenAI base URL is not configured");
    }

    if (/\/chat\/completions$/i.test(trimmed)) {
      return trimmed;
    }

    if (/\/v\d+(?:beta\d*)?$/i.test(trimmed)) {
      return `${trimmed}/chat/completions`;
    }

    return `${trimmed}/v1/chat/completions`;
  }

  getConfiguredMistralApiKey() {
    const liveSettings = this.app?.plugins?.plugins?.[MARKER_PLUGIN_ID]?.settings?.mistralaiApiKey;
    if (typeof liveSettings === "string" && liveSettings.trim() !== "") {
      return liveSettings.trim();
    }

    try {
      const basePath = this.app?.vault?.adapter?.basePath;
      if (typeof basePath !== "string" || basePath === "") {
        return null;
      }

      const settingsPath = path.join(
        basePath,
        this.app.vault.configDir,
        "plugins",
        MARKER_PLUGIN_ID,
        "data.json"
      );
      if (!fs.existsSync(settingsPath)) {
        return null;
      }

      const raw = fs.readFileSync(settingsPath, "utf8");
      const parsed = JSON.parse(raw);
      return typeof parsed?.mistralaiApiKey === "string" && parsed.mistralaiApiKey.trim() !== ""
        ? parsed.mistralaiApiKey.trim()
        : null;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`failed to load marker-api settings: ${message}`);
      return null;
    }
  }

  async extractSelectionTextWithMistralVision(blob, apiKey) {
    const arrayBuffer = await blob.arrayBuffer();
    const base64Image = Buffer.from(arrayBuffer).toString("base64");
    const dataUrl = `data:image/png;base64,${base64Image}`;
    const response = await this.requestJson("https://api.mistral.ai/v1/chat/completions", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        model: "mistral-small-latest",
        temperature: 0,
        messages: [
          {
            role: "user",
            content: [
              {
                type: "text",
                text: this.getVisionTranscriptionPrompt(),
              },
              {
                type: "image_url",
                image_url: dataUrl,
              },
            ],
          },
        ],
      }),
    });

    const firstChoice = Array.isArray(response?.choices) ? response.choices[0] : null;
    const content = firstChoice?.message?.content;
    if (typeof content === "string") {
      return content;
    }

    if (Array.isArray(content)) {
      return content
        .map((item) => (typeof item?.text === "string" ? item.text : ""))
        .join("\n");
    }

    throw new Error("Mistral vision response did not contain text");
  }

  async extractSelectionTextWithGemini(blob, apiKey, model) {
    const arrayBuffer = await blob.arrayBuffer();
    const base64Image = Buffer.from(arrayBuffer).toString("base64");
    const response = await this.requestJson(
      `${GEMINI_API_URL_BASE}/${encodeURIComponent(model)}:generateContent?key=${encodeURIComponent(apiKey)}`,
      {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          generationConfig: {
            temperature: 0,
            maxOutputTokens: 600,
          },
          contents: [
            {
              role: "user",
              parts: [
                {
                  text: this.getVisionTranscriptionPrompt(),
                },
                {
                  inline_data: {
                    mime_type: "image/png",
                    data: base64Image,
                  },
                },
              ],
            },
          ],
        }),
      }
    );

    const candidate = Array.isArray(response?.candidates) ? response.candidates[0] : null;
    const parts = Array.isArray(candidate?.content?.parts) ? candidate.content.parts : [];
    const text = parts
      .map((part) => (typeof part?.text === "string" ? part.text : ""))
      .join("\n");

    if (text.trim() !== "") {
      return text;
    }

    throw new Error("Gemini response did not contain text");
  }

  async extractSelectionTextWithOpenRouter(blob, apiKey) {
    const arrayBuffer = await blob.arrayBuffer();
    const base64Image = Buffer.from(arrayBuffer).toString("base64");
    const models = this.getConfiguredOpenRouterModels();
    const requestBody = {
      temperature: 0,
      max_tokens: 600,
      provider: {
        sort: "throughput",
        allow_fallbacks: true,
      },
      messages: [
        {
          role: "user",
          content: [
            {
              type: "text",
              text: this.getVisionTranscriptionPrompt(),
            },
            {
              type: "image_url",
              image_url: {
                url: `data:image/png;base64,${base64Image}`,
              },
            },
          ],
        },
      ],
    };

    if (models.length === 1) {
      requestBody.model = models[0];
    } else {
      requestBody.models = models;
    }

    const response = await this.requestJson(OPENROUTER_API_URL, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
        "Content-Type": "application/json",
        "HTTP-Referer": "obsidian://writing-bridge",
        "X-Title": "writing bridge",
      },
      body: JSON.stringify(requestBody),
    });

    const firstChoice = Array.isArray(response?.choices) ? response.choices[0] : null;
    const content = firstChoice?.message?.content;
    if (typeof content === "string") {
      return content;
    }

    if (Array.isArray(content)) {
      return content
        .map((item) => {
          if (typeof item === "string") {
            return item;
          }

          return typeof item?.text === "string" ? item.text : "";
        })
        .join("\n");
    }

    throw new Error("OpenRouter response did not contain text");
  }

  async extractSelectionTextWithCustomOpenAi(blob, baseUrl, apiKey, model) {
    const arrayBuffer = await blob.arrayBuffer();
    const base64Image = Buffer.from(arrayBuffer).toString("base64");
    const dataUrl = `data:image/png;base64,${base64Image}`;
    const url = this.resolveCustomOpenAiChatCompletionsUrl(baseUrl);

    const headers = { "Content-Type": "application/json" };
    if (apiKey) {
      headers.Authorization = `Bearer ${apiKey}`;
    }

    const response = await this.requestJson(url, {
      method: "POST",
      headers,
      body: JSON.stringify({
        model,
        temperature: 0,
        max_tokens: 600,
        messages: [
          {
            role: "user",
            content: [
              {
                type: "text",
                text: this.getVisionTranscriptionPrompt(),
              },
              {
                type: "image_url",
                image_url: { url: dataUrl },
              },
            ],
          },
        ],
      }),
    });

    const firstChoice = Array.isArray(response?.choices) ? response.choices[0] : null;
    const content = firstChoice?.message?.content;
    if (typeof content === "string") {
      return content;
    }

    if (Array.isArray(content)) {
      return content
        .map((item) => {
          if (typeof item === "string") {
            return item;
          }

          return typeof item?.text === "string" ? item.text : "";
        })
        .join("\n");
    }

    throw new Error("Custom OpenAI response did not contain text");
  }

  async extractSelectionTextWithMistral(blob, apiKey) {
    const fileId = await this.uploadMistralOcrFile(blob, apiKey);

    try {
      const signedUrl = await this.getMistralSignedUrl(fileId, apiKey);
      const response = await this.requestJson("https://api.mistral.ai/v1/ocr", {
        method: "POST",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
        body: JSON.stringify({
          model: "mistral-ocr-latest",
          document: {
            type: "image_url",
            image_url: signedUrl,
          },
          include_image_base64: false,
        }),
      });

      const pages = Array.isArray(response?.pages) ? response.pages : [];
      const text = pages
        .map((page) => this.normalizeCopiedText(page?.markdown))
        .filter(Boolean)
        .join("\n\n");

      return text;
    } finally {
      await this.deleteMistralFile(fileId, apiKey);
    }
  }

  getVisionTranscriptionPrompt() {
    const configuredPrompt =
      typeof this.settings?.visionTranscriptionPrompt === "string"
        ? this.settings.visionTranscriptionPrompt.trim()
        : "";
    return configuredPrompt !== "" ? configuredPrompt : DEFAULT_VISION_TRANSCRIPTION_PROMPT;
  }

  async uploadMistralOcrFile(blob, apiKey) {
    const formData = new FormData();
    formData.append("purpose", "ocr");
    formData.append("file", blob, "tldraw-selection.png");

    const response = await this.requestJson("https://api.mistral.ai/v1/files", {
      method: "POST",
      headers: {
        Authorization: `Bearer ${apiKey}`,
      },
      body: formData,
    });

    if (typeof response?.id !== "string" || response.id.trim() === "") {
      throw new Error("Mistral upload did not return a file id");
    }

    return response.id;
  }

  async getMistralSignedUrl(fileId, apiKey) {
    const response = await this.requestJson(`https://api.mistral.ai/v1/files/${fileId}/url`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${apiKey}`,
      },
    });

    if (typeof response?.url !== "string" || response.url.trim() === "") {
      throw new Error("Mistral signed URL request did not return a URL");
    }

    return response.url;
  }

  async deleteMistralFile(fileId, apiKey) {
    if (typeof fileId !== "string" || fileId.trim() === "") {
      return;
    }

    try {
      await fetch(`https://api.mistral.ai/v1/files/${fileId}`, {
        method: "DELETE",
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`mistral file cleanup failed: ${message}`);
    }
  }

  async requestJson(url, options) {
    const response = await fetch(url, options);
    const bodyText = await response.text();
    let parsedBody = null;

    if (bodyText) {
      try {
        parsedBody = JSON.parse(bodyText);
      } catch (_error) {
        parsedBody = null;
      }
    }

    if (!response.ok) {
      const message =
        typeof parsedBody?.message === "string"
          ? parsedBody.message
          : typeof parsedBody?.error === "string"
            ? parsedBody.error
            : typeof parsedBody?.error?.message === "string"
              ? parsedBody.error.message
              : typeof parsedBody?.detail === "string"
                ? parsedBody.detail
                : typeof parsedBody?.details === "string"
                  ? parsedBody.details
                  : "";
      const statusMessage = `${response.status} ${response.statusText}`.trim();
      const detail = message || bodyText;
      throw new Error(detail ? `HTTP ${statusMessage}: ${detail}` : `HTTP ${statusMessage}`);
    }

    return parsedBody;
  }

  async writeTempExport(blob) {
    await this.ensureTempExportFolder();

    const arrayBuffer = await blob.arrayBuffer();
    const existing = this.app.vault.getAbstractFileByPath(TEMP_EXPORT_FILE);
    if (existing instanceof TFile) {
      await this.app.vault.modifyBinary(existing, arrayBuffer);
      return existing;
    }

    await this.app.vault.createBinary(TEMP_EXPORT_FILE, arrayBuffer);

    const created = this.app.vault.getAbstractFileByPath(TEMP_EXPORT_FILE);
    if (created instanceof TFile) {
      return created;
    }

    throw new Error("Temp OCR export was created but not indexed by Obsidian");
  }

  async ensureTempExportFolder() {
    const folderPath = normalizePath(TEMP_EXPORT_FOLDER);
    if (this.app.vault.getAbstractFileByPath(folderPath)) {
      return;
    }

    try {
      await this.app.vault.createFolder(folderPath);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      if (!message.toLowerCase().includes("exists")) {
        throw error;
      }
    }
  }

  async clearTextExtractorCache(filePath) {
    const md5 = crypto.createHash("md5").update(filePath, "utf8").digest("hex");
    const cachePath = normalizePath(
      `${this.app.vault.configDir}/plugins/${TEXT_EXTRACTOR_PLUGIN_ID}/cache/${md5}.json`
    );

    if (await this.app.vault.adapter.exists(cachePath)) {
      await this.app.vault.adapter.remove(cachePath);
    }
  }

  async exportSelectionBlob(editor, shapeIds, options = {}) {
    const padding = Number(options.padding ?? 24);
    const scale = Number(options.scale ?? 3);
    const pixelRatio = Number(options.pixelRatio ?? OCR_EXPORT_PIXEL_RATIO);

    if (typeof editor?.toImage === "function") {
      const imageResult = await editor.toImage(shapeIds, {
        background: true,
        format: "png",
        padding,
        scale,
        pixelRatio,
      });

      const blob = imageResult?.blob;
      if (!(blob instanceof Blob)) {
        throw new Error("tldraw export did not return an image");
      }

      return blob;
    }

    if (typeof editor?.getSvgString === "function") {
      const svgResult = await editor.getSvgString(shapeIds, {
        background: true,
        scale: 1,
        padding,
        darkMode: document.body.classList.contains("theme-dark"),
        preserveAspectRatio: "xMidYMid meet",
      });
      const svgMarkup = typeof svgResult === "string" ? svgResult : svgResult?.svg;
      if (typeof svgMarkup !== "string" || svgMarkup.trim() === "") {
        throw new Error("tldraw SVG export did not return markup");
      }

      return this.renderSvgStringToPngBlob(svgMarkup, {
        scale,
        pixelRatio,
        width: Number(svgResult?.width ?? 0),
        height: Number(svgResult?.height ?? 0),
      });
    }

    throw new Error("tldraw editor cannot export selections");
  }

  installTldrawEmbedPreviewFormatBridge() {
    this.syncTldrawEmbedPreviewFormats();

    const observer = new MutationObserver((mutations) => {
      for (const mutation of mutations) {
        if (mutation.type === "attributes" && mutation.attributeName === "src") {
          this.scheduleTldrawEmbedPreviewFormatSync();
          return;
        }

        if (mutation.type === "childList" && (mutation.addedNodes.length || mutation.removedNodes.length)) {
          this.scheduleTldrawEmbedPreviewFormatSync();
          return;
        }
      }
    });

    observer.observe(document.body, {
      attributes: true,
      attributeFilter: ["src"],
      childList: true,
      subtree: true,
    });
    this.register(() => observer.disconnect());

    const syncAfterTldrawFileChange = (file) => {
      if (file instanceof TFile && this.isTldrawSourceFile(file)) {
        this.scheduleTldrawEmbedPreviewFormatSync(250);
      }
    };
    this.registerEvent(this.app.vault.on("modify", syncAfterTldrawFileChange));
    this.registerEvent(this.app.vault.on("rename", syncAfterTldrawFileChange));
    this.registerEvent(this.app.vault.on("delete", syncAfterTldrawFileChange));
  }

  scheduleTldrawEmbedPreviewFormatSync(delayMs = 80) {
    if (this.pendingTldrawEmbedPreviewSyncTimeout) {
      window.clearTimeout(this.pendingTldrawEmbedPreviewSyncTimeout);
    }

    this.pendingTldrawEmbedPreviewSyncTimeout = window.setTimeout(() => {
      this.pendingTldrawEmbedPreviewSyncTimeout = 0;
      this.syncTldrawEmbedPreviewFormats();
    }, delayMs);
  }

  syncTldrawEmbedPreviewFormats() {
    if (!this.isFeatureEnabled("tldrawEmbedPreviewBridgeEnabled")) {
      this.restoreTldrawEmbedPreviewSourceUrls();
      return;
    }

    const format = this.settings?.tldrawEmbedPreviewFormat ?? DEFAULT_SETTINGS.tldrawEmbedPreviewFormat;
    if (format !== "png") {
      this.restoreTldrawEmbedPreviewSourceUrls();
      return;
    }

    for (const image of document.querySelectorAll(".ptl-markdown-embed .ptl-tldraw-image img[src]")) {
      if (image instanceof HTMLImageElement && this.isTldrawEmbedPreviewImage(image)) {
        void this.ensureTldrawEmbedPreviewPng(image);
      }
    }
  }

  restoreTldrawEmbedPreviewSourceUrls() {
    for (const image of document.querySelectorAll(".ptl-markdown-embed .ptl-tldraw-image img[src]")) {
      if (!(image instanceof HTMLImageElement)) {
        continue;
      }

      const bridgeUrl = this.tldrawEmbedPreviewObjectUrlByImage.get(image);
      const sourceUrl = this.tldrawEmbedPreviewSourceUrlByImage.get(image);
      if (!bridgeUrl || !sourceUrl || (image.currentSrc || image.src) !== bridgeUrl) {
        continue;
      }

      image.src = sourceUrl;
      image.removeAttribute("data-tldraw-pen-bridge-preview-format");
      this.tldrawEmbedPreviewObjectUrlByImage.delete(image);
      this.tldrawEmbedPreviewSourceUrlByImage.delete(image);
      URL.revokeObjectURL(bridgeUrl);
      this.tldrawEmbedPreviewObjectUrls.delete(bridgeUrl);
    }
  }

  isTldrawEmbedPreviewImage(image) {
    if (!(image instanceof HTMLImageElement)) {
      return false;
    }

    if (image.closest(".tldraw-view-root")) {
      return false;
    }

    const previewRoot = image.closest(".ptl-tldraw-image");
    return previewRoot instanceof HTMLElement && Boolean(previewRoot.closest(".ptl-markdown-embed"));
  }

  async ensureTldrawEmbedPreviewPng(image) {
    const currentSrc = image.currentSrc || image.src;
    if (typeof currentSrc !== "string" || currentSrc === "") {
      return;
    }

    if (this.tldrawEmbedPreviewObjectUrlByImage.get(image) === currentSrc) {
      return;
    }

    if (this.tldrawEmbedPreviewConversionByImage.get(image) === currentSrc) {
      return;
    }

    this.tldrawEmbedPreviewConversionByImage.set(image, currentSrc);
    try {
      const svgMarkup = await this.readSvgMarkupFromImageUrl(currentSrc);
      if (!svgMarkup) {
        return;
      }

      const pngBlob = await this.renderSvgStringToPngBlob(svgMarkup, {
        scale: 1,
        pixelRatio: Math.max(1, Math.min(2, window.devicePixelRatio || 1)),
      });
      if (!(pngBlob instanceof Blob) || pngBlob.size === 0) {
        return;
      }

      const previousBridgeUrl = this.tldrawEmbedPreviewObjectUrlByImage.get(image);
      if (previousBridgeUrl) {
        URL.revokeObjectURL(previousBridgeUrl);
        this.tldrawEmbedPreviewObjectUrls.delete(previousBridgeUrl);
      }

      const pngUrl = URL.createObjectURL(pngBlob);
      this.tldrawEmbedPreviewObjectUrls.add(pngUrl);
      this.tldrawEmbedPreviewObjectUrlByImage.set(image, pngUrl);
      this.tldrawEmbedPreviewSourceUrlByImage.set(image, currentSrc);
      image.dataset.tldrawPenBridgePreviewFormat = "png";
      image.src = pngUrl;
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`tldraw embed PNG preview conversion failed: ${message}`);
    } finally {
      if (this.tldrawEmbedPreviewConversionByImage.get(image) === currentSrc) {
        this.tldrawEmbedPreviewConversionByImage.delete(image);
      }
    }
  }

  async readSvgMarkupFromImageUrl(url) {
    if (typeof url !== "string" || url === "") {
      return null;
    }

    const response = await fetch(url);
    if (!response.ok) {
      return null;
    }

    const contentType = response.headers.get("content-type") ?? "";
    if (contentType.includes("image/png")) {
      return null;
    }

    const text = await response.text();
    const trimmed = text.trimStart();
    if (!contentType.includes("svg") && !trimmed.startsWith("<svg") && !trimmed.startsWith("<?xml")) {
      return null;
    }

    return text;
  }

  isTldrawSourceFile(file) {
    if (!(file instanceof TFile)) {
      return false;
    }

    if (file.extension === "tldr") {
      return true;
    }

    if (file.extension !== "md") {
      return false;
    }

    const cache = this.app.metadataCache.getFileCache(file);
    return Boolean(cache?.frontmatter?.["tldraw-file"]);
  }

  normalizeCopiedText(value) {
    if (typeof value !== "string") {
      return null;
    }

    const normalized = value.replace(/\r\n?/g, "\n").trim();
    return normalized === "" ? null : normalized;
  }

  getErrorMessage(error) {
    if (error instanceof Error && typeof error.message === "string" && error.message.trim() !== "") {
      return error.message.trim();
    }

    if (typeof error === "string" && error.trim() !== "") {
      return error.trim();
    }

    return "Unknown error";
  }

  wrapErrorWithContext(context, error) {
    return new Error(`${context}: ${this.getErrorMessage(error)}`);
  }

  getFailureNoticeMessage(operation, error, maxLength = 220) {
    const detail = this.getErrorMessage(error).replace(/\s+/g, " ").trim();
    const message = `${operation} failed${detail ? `: ${detail}` : ""}`;
    return message.length <= maxLength ? message : `${message.slice(0, maxLength - 1)}…`;
  }

  pushFailureDetail(failures, label, errorOrMessage) {
    if (!Array.isArray(failures)) {
      return;
    }

    const detail =
      typeof errorOrMessage === "string" ? errorOrMessage.trim() : this.getErrorMessage(errorOrMessage);
    failures.push(detail ? `${label}: ${detail}` : label);
  }

  async writeClipboardText(text) {
    try {
      const { clipboard } = require("electron");
      await clipboard.writeText(text);
    } catch (error) {
      throw this.wrapErrorWithContext("Clipboard text write failed", error);
    }
  }

  async writeClipboardImage(blob) {
    try {
      const arrayBuffer = await blob.arrayBuffer();
      const { clipboard, nativeImage } = require("electron");
      const image = nativeImage.createFromBuffer(Buffer.from(arrayBuffer));
      if (typeof image?.isEmpty === "function" && image.isEmpty()) {
        throw new Error("selection screenshot export produced an empty image");
      }

      clipboard.writeImage(image);
    } catch (error) {
      throw this.wrapErrorWithContext("Clipboard image write failed", error);
    }
  }

  getVaultAbsolutePath(vaultRelativePath) {
    const basePath = this.app?.vault?.adapter?.basePath;
    if (typeof basePath !== "string" || basePath === "") {
      throw new Error("Vault base path is not available");
    }

    return path.join(basePath, vaultRelativePath);
  }

  findTldrawContainer(root) {
    if (!(root instanceof HTMLElement)) {
      return null;
    }

    const selectors = [
      ".tl-container.tl-container__focused",
      ".tldraw-view-root .tl-container",
      ".ptl-tldraw-image .tl-container",
      `.workspace-leaf-content[data-type="${INK_WRITING_VIEW_TYPE}"] .tl-container`,
      `.workspace-leaf-content[data-type="${INK_DRAWING_VIEW_TYPE}"] .tl-container`,
      ".ddc_ink_writing-editor .tl-container",
      ".ddc_ink_drawing-editor .tl-container",
      ".ddc_ink_embed .tl-container",
    ];

    for (const selector of selectors) {
      const candidate = root.matches?.(selector) ? root : root.querySelector(selector);
      if (candidate instanceof HTMLElement && this.isTldrawContainer(candidate)) {
        return candidate;
      }
    }

    return null;
  }

  getOwningTldrawContainer(element) {
    if (!(element instanceof HTMLElement)) {
      return null;
    }

    const candidate = element.closest(".tl-container");
    if (candidate instanceof HTMLElement && this.isTldrawContainer(candidate)) {
      return candidate;
    }

    return null;
  }

  isTldrawContainer(element) {
    return (
      element instanceof HTMLElement &&
      element.classList.contains("tl-container") &&
      Boolean(element.closest(".tldraw-view-root, .ptl-tldraw-image") || this.isInkTldrawContainer(element))
    );
  }

  findExcalidrawContainer(root) {
    if (!(root instanceof HTMLElement)) {
      return null;
    }

    const selectors = [
      '.workspace-leaf-content[data-type="excalidraw"] .excalidraw-wrapper',
      ".excalidraw-view .excalidraw-wrapper",
      ".excalidraw-wrapper",
    ];

    for (const selector of selectors) {
      const candidate = root.matches?.(selector) ? root : root.querySelector(selector);
      if (candidate instanceof HTMLElement && this.isExcalidrawContainer(candidate)) {
        return candidate;
      }
    }

    return null;
  }

  findExcalidrawRuledPageHost(root) {
    if (!(root instanceof HTMLElement)) {
      return null;
    }

    const selectors = [
      '.workspace-leaf-content[data-type="excalidraw"] .excalidraw__canvas-wrapper',
      ".excalidraw-view .excalidraw__canvas-wrapper",
      ".excalidraw__canvas-wrapper",
      '.workspace-leaf-content[data-type="excalidraw"] .excalidraw-wrapper',
      ".excalidraw-view .excalidraw-wrapper",
      ".excalidraw-wrapper",
    ];

    for (const selector of selectors) {
      const candidate = root.matches?.(selector) ? root : root.querySelector(selector);
      if (candidate instanceof HTMLElement) {
        return candidate;
      }
    }

    return this.findExcalidrawContainer(root);
  }

  clearStaleExcalidrawRuledPageHosts(root, activeHost = null) {
    if (!(root instanceof HTMLElement)) {
      return;
    }

    for (const host of root.querySelectorAll(".writing-bridge-excalidraw-ruled-host")) {
      if (!(host instanceof HTMLElement) || host === activeHost) {
        continue;
      }

      host.classList.remove("writing-bridge-excalidraw-ruled-host");
    }
  }

  getOwningExcalidrawContainer(element) {
    if (!(element instanceof HTMLElement)) {
      return null;
    }

    const candidate = element.closest(".excalidraw-wrapper");
    if (candidate instanceof HTMLElement && this.isExcalidrawContainer(candidate)) {
      return candidate;
    }

    return null;
  }

  isExcalidrawContainer(element) {
    return (
      element instanceof HTMLElement &&
      element.classList.contains("excalidraw-wrapper") &&
      Boolean(element.closest('.workspace-leaf-content[data-type="excalidraw"], .excalidraw-view'))
    );
  }

  getExcalidrawRootForView(view) {
    const explicitRoot = view?.excalidrawWrapperRef?.current;
    if (explicitRoot instanceof HTMLElement) {
      return explicitRoot;
    }

    const containerEl = view?.containerEl;
    if (containerEl instanceof HTMLElement) {
      const container = this.findExcalidrawContainer(containerEl);
      if (container) {
        return container;
      }
    }

    const activeContainer = this.getActiveExcalidrawContainer();
    if (activeContainer instanceof HTMLElement) {
      return activeContainer;
    }

    return null;
  }

  getExcalidrawLeafRootForView(view) {
    const containerEl = view?.containerEl;
    if (containerEl instanceof HTMLElement) {
      const root = containerEl.closest('.workspace-leaf-content[data-type="excalidraw"]');
      if (root instanceof HTMLElement) {
        return root;
      }
    }

    const wrapper = this.getExcalidrawRootForView(view);
    return wrapper instanceof HTMLElement
      ? wrapper.closest('.workspace-leaf-content[data-type="excalidraw"]')
      : null;
  }

  getExcalidrawViewForRoot(root) {
    if (!(root instanceof HTMLElement)) {
      return null;
    }

    const wrapper = this.findExcalidrawContainer(root);
    const leaves = this.app.workspace.getLeavesOfType(EXCALIDRAW_VIEW_TYPE) ?? [];
    for (const leaf of leaves) {
      const view = leaf?.view;
      if (!this.isExcalidrawView(view)) {
        continue;
      }

      const leafRoot = this.getExcalidrawLeafRootForView(view);
      const viewWrapper = this.getExcalidrawRootForView(view);
      if (
        leafRoot === root ||
        (leafRoot instanceof HTMLElement && (leafRoot.contains(root) || root.contains(leafRoot))) ||
        (wrapper instanceof HTMLElement &&
          viewWrapper instanceof HTMLElement &&
          (wrapper === viewWrapper || wrapper.contains(viewWrapper) || viewWrapper.contains(wrapper)))
      ) {
        return view;
      }
    }

    return null;
  }

  focusExcalidrawTarget(target) {
    const root = this.getExcalidrawRootForView(target?.view) ?? target?.container;
    if (root instanceof HTMLElement) {
      root.focus?.({ preventScroll: true });
    }
  }

  sendToolKey(target, key, code) {
    this.focusTldrawTarget(target);

    window.requestAnimationFrame(() => {
      const liveTarget = target.isConnected ? target : this.getTldrawTarget()?.container;
      if (!(liveTarget instanceof HTMLElement)) {
        return;
      }

      this.sendShortcutKey(liveTarget, key, code);
    });
  }

  sendShortcutKey(target, key, code, extraInit = {}) {
    if (!(target instanceof HTMLElement)) {
      return;
    }

    this.focusTldrawTarget(target);
    this.log(`sending synthetic key ${key}${extraInit.ctrlKey ? " with Ctrl" : ""}`);

    for (const type of ["keydown", "keyup"]) {
      const event = new KeyboardEvent(type, {
        key,
        code,
        bubbles: true,
        cancelable: true,
        composed: true,
        ...extraInit,
      });

      target.dispatchEvent(event);
    }
  }

  dispatchClipboardCommand(target, type) {
    if (!(target instanceof HTMLElement) || (type !== "cut" && type !== "copy" && type !== "paste")) {
      return false;
    }

    const ownerDocument = target.ownerDocument;
    if (!ownerDocument) {
      return false;
    }

    this.focusTldrawTarget(target);
    this.log(`dispatching clipboard event ${type}`);

    try {
      const event = new ClipboardEvent(type, {
        bubbles: true,
        cancelable: true,
        composed: true,
      });

      return ownerDocument.dispatchEvent(event);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      this.log(`clipboard event ${type} failed: ${message}`);
      return false;
    }
  }

  focusTldrawTarget(target) {
    if (!(target instanceof HTMLElement)) {
      return;
    }

    target.focus({ preventScroll: true });
  }

  log(message) {
    try {
      const basePath = this.app?.vault?.adapter?.basePath;
      const rootPath =
        typeof basePath === "string" && basePath !== ""
          ? basePath
          : this.app.vault.configDir;
      const logPath = path.join(
        rootPath,
        this.app.vault.configDir,
        "plugins",
        "writing-bridge",
        "bridge.log"
      );
      fs.mkdirSync(path.dirname(logPath), { recursive: true });
      fs.appendFileSync(logPath, `${new Date().toISOString()} ${message}\n`, "utf8");
    } catch (_error) {
      // Logging should never be the thing that breaks the tool.
    }
  }

  mod(value, divisor) {
    if (!Number.isFinite(value) || !Number.isFinite(divisor) || divisor === 0) {
      return 0;
    }

    return ((value % divisor) + divisor) % divisor;
  }
};

function decodeDrawSegmentPath(base64) {
  if (typeof base64 !== "string" || base64.length === 0) {
    return [];
  }

  const bytes = Buffer.from(base64, "base64");
  if (bytes.length < 12) {
    return [];
  }

  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const points = [];
  let x = view.getFloat32(0, true);
  let y = view.getFloat32(4, true);
  let z = view.getFloat32(8, true);
  points.push({ x, y, z });

  for (let offset = 12; offset < bytes.length; offset += 6) {
    x += float16BitsToNumber(view.getUint16(offset, true));
    y += float16BitsToNumber(view.getUint16(offset + 2, true));
    z += float16BitsToNumber(view.getUint16(offset + 4, true));
    points.push({ x, y, z });
  }

  return points;
}

function float16BitsToNumber(bits) {
  const sign = bits >> 15;
  const exp = (bits >> 10) & 31;
  const frac = bits & 1023;
  if (exp === 0) {
    const value = frac * FLOAT16_POW2_SUBNORMAL;
    return sign ? -value : value;
  }

  if (exp === 31) {
    if (frac) {
      return Number.NaN;
    }

    return sign ? -Infinity : Infinity;
  }

  const magnitude = FLOAT16_POW2[exp] * FLOAT16_MANTISSA[frac];
  return sign ? -magnitude : magnitude;
}

class WritingBridgeSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();
    const currentGeminiModel = this.plugin.getConfiguredGeminiModel();
    const geminiModelOptions = {
      "gemini-2.5-flash": "gemini-2.5-flash",
      "gemini-2.5-flash-lite": "gemini-2.5-flash-lite",
      "gemini-2.5-pro": "gemini-2.5-pro",
    };

    if (!geminiModelOptions[currentGeminiModel]) {
      geminiModelOptions[currentGeminiModel] = `${currentGeminiModel} (current)`;
    }

    containerEl.createEl("h3", { text: "Feature toggles" });

    new Setting(containerEl)
      .setName("Enable tool hotkeys and commands")
      .setDesc("Draw, Select, and Hand tool commands and the matching quick-action menu buttons.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("toolCommandsEnabled")).onChange(async (value) => {
          this.plugin.settings.toolCommandsEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.destroySelectionMenu();
          this.plugin.refreshSelectionMenu();
        })
      );

    new Setting(containerEl)
      .setName("Enable copy selected as text")
      .setDesc("Allow OCR or direct text copy from canvas selections, plus the Copy and Cut buttons in the floating selection menu.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("copySelectedTextEnabled")).onChange(async (value) => {
          this.plugin.settings.copySelectedTextEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.destroySelectionMenu();
          this.plugin.refreshSelectionMenu();
        })
      );

    new Setting(containerEl)
      .setName("Enable selection screenshot copy")
      .setDesc("Allow copying the current canvas selection as an image and show the Screenshot button in the floating selection menu.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("selectionScreenshotEnabled")).onChange(async (value) => {
          this.plugin.settings.selectionScreenshotEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.destroySelectionMenu();
          this.plugin.refreshSelectionMenu();
        })
      );

    new Setting(containerEl)
      .setName("Enable region screenshots")
      .setDesc("Allow region screenshot commands, including the scrolling-region capture helpers.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("regionScreenshotEnabled")).onChange(async (value) => {
          this.plugin.settings.regionScreenshotEnabled = value;
          await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("Enable floating selection menu")
      .setDesc("Show the bubble menu over active tldraw and Excalidraw selections.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("selectionMenuEnabled")).onChange(async (value) => {
          this.plugin.settings.selectionMenuEnabled = value;
          await this.plugin.saveSettings();
          if (!value) {
            this.plugin.hideSelectionMenu();
          }
          this.plugin.destroySelectionMenu();
          this.plugin.refreshSelectionMenu();
        })
      );

    new Setting(containerEl)
      .setName("Enable canvas style panel toggle")
      .setDesc("Show the pull-tab that collapses or opens the floating style and action panels.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("stylePanelToggleEnabled")).onChange(async (value) => {
          this.plugin.settings.stylePanelToggleEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.syncStylePanelToggles();
        })
      );

    new Setting(containerEl)
      .setName("Enable ruled page overlay")
      .setDesc("Show notebook lines, line toggle buttons, and the ruled-page command on supported canvases.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("ruledPageFeatureEnabled")).onChange(async (value) => {
          this.plugin.settings.ruledPageFeatureEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.syncRuledPageOverlays();
          this.plugin.syncRuledPageToggles();
        })
      );

    new Setting(containerEl)
      .setName("Enable sidebar pen assist")
      .setDesc("Replay missed pen taps on sidebars, ribbons, and tab chrome.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("sidebarPenAssistEnabled")).onChange(async (value) => {
          this.plugin.settings.sidebarPenAssistEnabled = value;
          this.plugin.sidebarPenTapState = null;
          await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("Enable default tldraw pen size")
      .setDesc("Auto-apply the configured tldraw draw size when a canvas first becomes active.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("tldrawDefaultPenSizeEnabled")).onChange(async (value) => {
          this.plugin.settings.tldrawDefaultPenSizeEnabled = value;
          this.plugin.appliedTldrawDrawSizeByRootId.clear();
          await this.plugin.saveSettings();
          this.plugin.syncRuledPageOverlays();
        })
      );


    new Setting(containerEl)
      .setName("Enable default Excalidraw zoom")
      .setDesc("Auto-apply the configured Excalidraw zoom when a leaf first opens in this session.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("excalidrawDefaultZoomEnabled")).onChange(async (value) => {
          this.plugin.settings.excalidrawDefaultZoomEnabled = value;
          this.plugin.appliedExcalidrawZoomLeaves = new WeakSet();
          await this.plugin.saveSettings();
          this.plugin.syncExcalidrawDefaultZoom();
        })
      );

    new Setting(containerEl)
      .setName("Enable tldraw embed preview conversion")
      .setDesc("Convert tldraw markdown embeds to bridge-managed PNG previews when that format is selected.")
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.isFeatureEnabled("tldrawEmbedPreviewBridgeEnabled")).onChange(async (value) => {
          this.plugin.settings.tldrawEmbedPreviewBridgeEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.syncTldrawEmbedPreviewFormats();
        })
      );

    containerEl.createEl("h3", { text: "Behavior settings" });

    new Setting(containerEl)
      .setName("Default tldraw pen size")
      .setDesc("Apply this draw size when a tldraw canvas first becomes active. You can still change it manually after that.")
      .addDropdown((dropdown) => {
        dropdown
          .addOption("default", "Use tldraw default")
          .addOption("s", "Small")
          .addOption("m", "Medium")
          .addOption("l", "Large")
          .addOption("xl", "Extra large")
          .setValue(this.plugin.settings.tldrawDefaultDrawSize ?? "s")
          .onChange(async (value) => {
            this.plugin.settings.tldrawDefaultDrawSize = value;
            this.plugin.appliedTldrawDrawSizeByRootId.clear();
            await this.plugin.saveSettings();
            this.plugin.syncRuledPageOverlays();
          });
      });

    new Setting(containerEl)
      .setName("Tldraw embed preview format")
      .setDesc("Convert tldraw page embeds to PNG previews by default. SVG leaves tldraw's normal vector preview alone.")
      .addDropdown((dropdown) => {
        dropdown
          .addOption("png", "PNG")
          .addOption("svg", "SVG")
          .addOption("default", "Tldraw default")
          .setValue(this.plugin.settings.tldrawEmbedPreviewFormat ?? "png")
          .onChange(async (value) => {
            this.plugin.settings.tldrawEmbedPreviewFormat = value;
            await this.plugin.saveSettings();
            this.plugin.syncTldrawEmbedPreviewFormats();
          });
      });

    new Setting(containerEl)
      .setName("Show ruled page background")
      .setDesc("Overlay notebook-style horizontal lines on supported drawing canvases.")
      .addToggle((toggle) =>
        toggle.setValue(Boolean(this.plugin.settings.ruledPageEnabled)).onChange(async (value) => {
          this.plugin.settings.ruledPageEnabled = value;
          await this.plugin.saveSettings();
          this.plugin.syncRuledPageOverlays();
        })
      );


    new Setting(containerEl)
      .setName("Default Excalidraw zoom")
      .setDesc("Apply this zoom the first time each Excalidraw leaf opens this session. Accepts a ratio (e.g. 1.5) or a percent (e.g. 150%). Leave blank to use Excalidraw's own zoom.")
      .addText((text) =>
        text
          .setPlaceholder("e.g. 1.5 or 150%")
          .setValue(this.plugin.settings.excalidrawDefaultZoom ?? "")
          .onChange(async (value) => {
            this.plugin.settings.excalidrawDefaultZoom = value;
            await this.plugin.saveSettings();
            this.plugin.appliedExcalidrawZoomLeaves = new WeakSet();
            this.plugin.syncExcalidrawDefaultZoom();
          })
      );

    new Setting(containerEl)
      .setName("Collapse canvas style panel by default")
      .setDesc("Hide the floating style/actions panel until you click the pull-tab toggle.")
      .addToggle((toggle) =>
        toggle.setValue(Boolean(this.plugin.settings.stylePanelCollapsed)).onChange(async (value) => {
          this.plugin.settings.stylePanelCollapsed = value;
          await this.plugin.saveSettings();
          this.plugin.syncStylePanelToggles();
        })
      );

    new Setting(containerEl)
      .setName("AI transcription prompt")
      .setDesc("Custom instructions for handwriting and math copy-to-text. Leave blank to use the built-in prompt.")
      .addTextArea((text) => {
        text
          .setPlaceholder(DEFAULT_VISION_TRANSCRIPTION_PROMPT)
          .setValue(this.plugin.settings.visionTranscriptionPrompt ?? "")
          .onChange(async (value) => {
            this.plugin.settings.visionTranscriptionPrompt = value;
            await this.plugin.saveSettings();
          });
        text.inputEl.rows = 6;
        text.inputEl.cols = 36;
      });

    new Setting(containerEl)
      .setName("Primary image LLM provider")
      .setDesc("Use this provider first for handwriting and math image transcription. The other providers run as fallbacks in declared order.")
      .addDropdown((dropdown) =>
        dropdown
          .addOption(VISION_PROVIDER_OPENROUTER, "OpenRouter")
          .addOption(VISION_PROVIDER_GEMINI, "Gemini")
          .addOption(VISION_PROVIDER_CUSTOM, "Custom OpenAI-compatible")
          .setValue(this.plugin.getConfiguredVisionProvider())
          .onChange(async (value) => {
            this.plugin.settings.visionLlmProvider = value;
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Gemini API key")
      .setDesc("Used when Gemini is the selected image recognition provider or LLM fallback.")
      .addText((text) =>
        text
          .setPlaceholder("AIza...")
          .setValue(this.plugin.settings.geminiApiKey ?? "")
          .onChange(async (value) => {
            this.plugin.settings.geminiApiKey = value.trim();
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Gemini model")
      .setDesc("Choose the Gemini vision model used for copy recognition.")
      .addDropdown((dropdown) => {
        for (const [value, label] of Object.entries(geminiModelOptions)) {
          dropdown.addOption(value, label);
        }

        dropdown.setValue(currentGeminiModel).onChange(async (value) => {
          this.plugin.settings.geminiModel = value;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName("OpenRouter API key")
      .setDesc("Used when OpenRouter is the selected image recognition provider or LLM fallback.")
      .addText((text) =>
        text
          .setPlaceholder("sk-or-...")
          .setValue(this.plugin.settings.openRouterApiKey ?? "")
          .onChange(async (value) => {
            this.plugin.settings.openRouterApiKey = value.trim();
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("OpenRouter models")
      .setDesc("One model per line. The bridge routes across them for the fastest acceptable vision result.")
      .addTextArea((text) => {
        text
          .setPlaceholder(OPENROUTER_DEFAULT_MODELS.join("\n"))
          .setValue(this.plugin.settings.openRouterModels ?? "")
          .onChange(async (value) => {
            this.plugin.settings.openRouterModels = value;
            await this.plugin.saveSettings();
          });
        text.inputEl.rows = 4;
        text.inputEl.cols = 36;
      });

    new Setting(containerEl)
      .setName("Custom OpenAI-compatible base URL")
      .setDesc("Full /chat/completions URL for any OpenAI-compatible vision endpoint. The bridge will append /v1/chat/completions if you only paste the host.")
      .addText((text) =>
        text
          .setPlaceholder("https://api.openai.com/v1/chat/completions")
          .setValue(this.plugin.settings.customOpenAiBaseUrl ?? "")
          .onChange(async (value) => {
            this.plugin.settings.customOpenAiBaseUrl = value.trim();
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Custom OpenAI-compatible API key")
      .setDesc("Sent as a Bearer token. Leave blank for endpoints that do not require authentication.")
      .addText((text) =>
        text
          .setPlaceholder("sk-... (optional)")
          .setValue(this.plugin.settings.customOpenAiApiKey ?? "")
          .onChange(async (value) => {
            this.plugin.settings.customOpenAiApiKey = value.trim();
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Custom OpenAI-compatible model")
      .setDesc("Model name to send in the chat completion request, e.g. gpt-4o-mini or your local model id.")
      .addText((text) =>
        text
          .setPlaceholder("gpt-4o-mini")
          .setValue(this.plugin.settings.customOpenAiModel ?? "")
          .onChange(async (value) => {
            this.plugin.settings.customOpenAiModel = value.trim();
            await this.plugin.saveSettings();
          })
      );
  }
}
