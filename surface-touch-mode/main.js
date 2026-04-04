const { Notice, Plugin, PluginSettingTab, Setting } = require("obsidian");

const DEFAULT_SETTINGS = {
  touchShimEnabled: true,
  modalDragShimEnabled: true,
  tabDragShimEnabled: true,
  edgeZonesEnabled: true,
  edgeZoneWidth: 18,
};

const TOUCH_POINTER_TYPES = new Set(["touch", "pen"]);
const SIDEBAR_COMMAND_IDS = {
  left: ["app:toggle-left-sidebar", "app:toggle-left-dock", "workspace:toggle-left-sidebar"],
  right: ["app:toggle-right-sidebar", "app:toggle-right-dock", "workspace:toggle-right-sidebar"],
};
const SIDEBAR_OPEN_CLASSES = {
  left: "is-left-sidedock-open",
  right: "is-right-sidedock-open",
};
const SHIM_TARGET_SELECTOR = [
  ".workspace-ribbon .side-dock-ribbon-action",
  ".workspace-ribbon .sidebar-toggle-button",
  ".sidebar-toggle-button",
  ".mod-sidedock .nav-action-button",
  ".mod-sidedock .clickable-icon",
].join(", ");
const SHIM_SCOPE_SELECTOR = [
  ".workspace-ribbon",
  ".mod-sidedock",
  ".sidebar-toggle-button",
].join(", ");
const TLDRAW_MENU_TARGET_SELECTOR = [
  ".tlui-menu .tlui-button__menu",
  ".tlui-menu .tlui-menu__submenu__trigger",
  "[data-radix-menu-content] .tlui-button__menu",
  "[data-radix-menu-content] .tlui-menu__submenu__trigger",
  ".tlui-popover__content .tlui-button__menu",
  ".tlui-popover__content .tlui-menu__submenu__trigger",
].join(", ");
const FLOATING_UI_SELECTOR = [
  ".workspace-split.mod-floating",
  ".popover",
  ".hover-popover",
  ".menu",
  ".suggestion-container",
  ".prompt",
  ".modal-container",
  ".mod-root-popout",
].join(", ");
const TAP_MAX_DISTANCE = 26;
const TAP_MAX_DURATION_MS = 900;
const TOUCH_CLASS_HOLD_MS = 1400;
const GHOST_CLICK_MAX_DISTANCE = 44;
const GHOST_CLICK_HOLD_MS = 900;
const MODAL_DRAG_HANDLE_HEIGHT = 56;
const MODAL_DRAG_MARGIN = 12;
const SIDEBAR_HOTSPOT_PADDING = 12;
const MODAL_HANDLE_LEFT_INSET = 12;
const MODAL_HANDLE_RIGHT_INSET = 88;
const TAB_DRAG_HANDLE_HEIGHT = 32;
const TAB_DRAG_HANDLE_SIDE_INSET = 6;
const ELIGIBLE_TAB_HEADER_SELECTOR = ".workspace-tabs.mod-top .workspace-tab-header";
const TAB_DRAG_START_DISTANCE = 4;
const TAB_DRAG_HANDLE_MIN_WIDTH = 48;

module.exports = class SurfaceTouchModePlugin extends Plugin {
  async onload() {
    await this.loadSettings();
    this.activeTap = null;
    this.activeModalDrag = null;
    this.activeTabDrag = null;
    this.touchClassTimer = 0;
    this.suppressedClick = null;
    this.rootEl = null;
    this.styleEl = null;
    this.overlayRefreshFrame = 0;
    this.overlayObserver = null;

    this.addCommand({
      id: "toggle-left-sidebar",
      name: "Surface Touch: Toggle Left Sidebar",
      callback: () => {
        this.toggleSidebar("left");
      },
    });

    this.addCommand({
      id: "toggle-right-sidebar",
      name: "Surface Touch: Toggle Right Sidebar",
      callback: () => {
        this.toggleSidebar("right");
      },
    });

    this.addSettingTab(new SurfaceTouchModeSettingTab(this.app, this));

    this.installPointerShim();
    this.applySettings();
    this.startOverlayObserver();
    this.registerEvent(this.app.workspace.on("layout-change", () => this.scheduleOverlayRefresh()));
    this.registerEvent(this.app.workspace.on("active-leaf-change", () => this.scheduleOverlayRefresh()));
    this.register(() => this.teardown());
  }

  onunload() {
    this.teardown();
  }

  async loadSettings() {
    this.settings = {
      ...DEFAULT_SETTINGS,
      ...((await this.loadData()) ?? {}),
    };
  }

  async saveSettings() {
    await this.saveData(this.settings);
    this.applySettings();
  }

  applySettings() {
    this.installStyle();
    this.scheduleOverlayRefresh();
  }

  installPointerShim() {
    this.boundPointerDown = this.handlePointerDown.bind(this);
    this.boundPointerMove = this.handlePointerMove.bind(this);
    this.boundPointerUp = this.handlePointerUp.bind(this);
    this.boundPointerCancel = this.handlePointerCancel.bind(this);
    this.boundClickCapture = this.handleClickCapture.bind(this);
    this.boundResize = this.scheduleOverlayRefresh.bind(this);
    this.boundScroll = this.scheduleOverlayRefresh.bind(this);

    window.addEventListener("pointerdown", this.boundPointerDown, true);
    window.addEventListener("pointermove", this.boundPointerMove, true);
    window.addEventListener("pointerup", this.boundPointerUp, true);
    window.addEventListener("pointercancel", this.boundPointerCancel, true);
    window.addEventListener("click", this.boundClickCapture, true);
    window.addEventListener("resize", this.boundResize, true);
    window.addEventListener("scroll", this.boundScroll, true);

    this.register(() => {
      window.removeEventListener("pointerdown", this.boundPointerDown, true);
      window.removeEventListener("pointermove", this.boundPointerMove, true);
      window.removeEventListener("pointerup", this.boundPointerUp, true);
      window.removeEventListener("pointercancel", this.boundPointerCancel, true);
      window.removeEventListener("click", this.boundClickCapture, true);
      window.removeEventListener("resize", this.boundResize, true);
      window.removeEventListener("scroll", this.boundScroll, true);
    });
  }

  handlePointerDown(event) {
    if (!this.isPrimaryTouchLikePointer(event)) {
      return;
    }

    if (this.settings.tabDragShimEnabled) {
      const tabHeader = this.getTabDragHeaderFromEventTarget(event.target);
      if (tabHeader) {
        this.beginTabDragCandidate(event, tabHeader);
        return;
      }
    }

    if (!this.settings.touchShimEnabled) {
      return;
    }

    if (this.hasBlockingModalOpen()) {
      return;
    }

    const target = this.getShimTarget(event.target);
    if (!target || this.isInsideSurfaceTouchOverlay(target) || !this.isAllowedShimTarget(target)) {
      return;
    }

    this.markTouchInput();
    this.activeTap = {
      pointerId: event.pointerId,
      target,
      x: event.clientX,
      y: event.clientY,
      startedAt: Date.now(),
    };
  }

  handlePointerMove(event) {
    const tabDrag = this.activeTabDrag;
    if (tabDrag && tabDrag.pointerId === event.pointerId) {
      this.handleTabDragPointerMove(event, tabDrag);
      return;
    }

    const drag = this.activeModalDrag;
    if (!drag || drag.pointerId !== event.pointerId) {
      return;
    }

    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }

    this.markTouchInput();

    const dx = Number(event.clientX) - drag.startX;
    const dy = Number(event.clientY) - drag.startY;
    if (!drag.moved && Math.hypot(dx, dy) >= 3) {
      drag.moved = true;
    }

    const width = Number(drag.rect.width);
    const height = Number(drag.rect.height);
    const maxLeft = Math.max(MODAL_DRAG_MARGIN, window.innerWidth - width - MODAL_DRAG_MARGIN);
    const maxTop = Math.max(MODAL_DRAG_MARGIN, window.innerHeight - height - MODAL_DRAG_MARGIN);
    const nextLeft = Math.max(
      MODAL_DRAG_MARGIN,
      Math.min(drag.rect.left + dx, maxLeft)
    );
    const nextTop = Math.max(
      MODAL_DRAG_MARGIN,
      Math.min(drag.rect.top + dy, maxTop)
    );
    const nextOffsetX = drag.startOffsetX + (nextLeft - drag.rect.left);
    const nextOffsetY = drag.startOffsetY + (nextTop - drag.rect.top);

    this.applyModalDragOffset(drag.modal, nextOffsetX, nextOffsetY);
    this.scheduleOverlayRefresh();
  }

  handlePointerUp(event) {
    const tabDrag = this.activeTabDrag;
    if (tabDrag && tabDrag.pointerId === event.pointerId) {
      this.finishTabDrag(event, tabDrag);
      return;
    }

    const drag = this.activeModalDrag;
    if (drag && drag.pointerId === event.pointerId) {
      this.finishModalDrag(event, drag);
      return;
    }

    if (!this.isTouchLikeCompletionPointer(event)) {
      return;
    }

    if (!this.settings.touchShimEnabled) {
      return;
    }

    const activeTap = this.activeTap;
    this.activeTap = null;
    if (!activeTap || activeTap.pointerId !== event.pointerId) {
      return;
    }

    const elapsed = Date.now() - activeTap.startedAt;
    const distance = Math.hypot(event.clientX - activeTap.x, event.clientY - activeTap.y);
    if (elapsed > TAP_MAX_DURATION_MS || distance > TAP_MAX_DISTANCE) {
      return;
    }

    const target = this.getShimTarget(event.target);
    if (!target || !this.isAllowedShimTarget(target)) {
      return;
    }

    if (!this.targetsMatch(activeTap.target, target)) {
      return;
    }

    const clickTarget = this.getSyntheticClickTarget(target);
    if (!clickTarget) {
      return;
    }

    this.queueSuppressedClick(event);
    this.markTouchInput();
    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }

    if (this.getSidebarSide(clickTarget)) {
      this.toggleSidebar(this.getSidebarSide(clickTarget));
      return;
    }

    this.dispatchSyntheticClick(clickTarget);
  }

  handleClickCapture(event) {
    const suppressed = this.suppressedClick;
    if (!suppressed) {
      return;
    }

    if (Date.now() > suppressed.expiresAt) {
      this.suppressedClick = null;
      return;
    }

    if (!this.shouldSuppressGhostClick(event, suppressed)) {
      return;
    }

    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }
    this.suppressedClick = null;
  }

  isPrimaryTouchLikePointer(event) {
    return Boolean(
      event &&
      event.isPrimary &&
      event.button === 0 &&
      TOUCH_POINTER_TYPES.has(`${event.pointerType ?? ""}`.toLowerCase())
    );
  }

  isTouchLikeCompletionPointer(event) {
    return Boolean(
      event &&
      event.isPrimary &&
      TOUCH_POINTER_TYPES.has(`${event.pointerType ?? ""}`.toLowerCase())
    );
  }

  getShimTarget(node) {
    if (!(node instanceof Element)) {
      return null;
    }

    return node.closest([SHIM_TARGET_SELECTOR, TLDRAW_MENU_TARGET_SELECTOR].join(", "));
  }

  isAllowedShimTarget(node) {
    if (!(node instanceof Element)) {
      return false;
    }

    if (this.isTldrawMenuTarget(node)) {
      return true;
    }

    return !this.isInsideFloatingUi(node) && Boolean(node.closest(SHIM_SCOPE_SELECTOR));
  }

  getSyntheticClickTarget(node) {
    if (!(node instanceof Element)) {
      return null;
    }

    const tldrawMenuTarget = node.closest(".tlui-button__menu, .tlui-menu__submenu__trigger");
    if (tldrawMenuTarget instanceof HTMLElement && this.isTldrawMenuTarget(tldrawMenuTarget)) {
      return tldrawMenuTarget;
    }

    const target = node.closest(
      [
        ".workspace-ribbon .side-dock-ribbon-action",
        ".sidebar-toggle-button .clickable-icon",
        ".sidebar-toggle-button",
        ".mod-sidedock .nav-action-button",
        ".clickable-icon",
      ].join(", ")
    );

    return target instanceof HTMLElement ? target : null;
  }

  isTldrawMenuTarget(node) {
    return Boolean(node instanceof Element && node.closest(TLDRAW_MENU_TARGET_SELECTOR));
  }

  startModalDrag(event, descriptor) {
    const modal = descriptor.modal;
    const rect = descriptor.rect;
    this.resetLegacyModalInlineLayout(modal);
    const existingOffset = this.getModalDragOffset(modal);
    const captureTarget = event.target instanceof HTMLElement ? event.target : null;

    this.activeModalDrag = {
      pointerId: event.pointerId,
      modal,
      rect,
      startX: Number(event.clientX),
      startY: Number(event.clientY),
      startOffsetX: existingOffset.x,
      startOffsetY: existingOffset.y,
      captureTarget,
      moved: false,
    };

    this.markTouchInput();
    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }

    if (typeof captureTarget?.setPointerCapture === "function") {
      try {
        captureTarget.setPointerCapture(event.pointerId);
      } catch (_error) {
        // Pointer capture is flaky enough that this should not kill the drag.
      }
    }
  }

  startTabDrag(event, descriptor) {
    const header = descriptor.header;
    const rect = descriptor.rect;
    const captureTarget = descriptor.captureTarget ?? (event.target instanceof HTMLElement ? event.target : null);

    this.activeTabDrag = {
      pointerId: event.pointerId,
      header,
      rect,
      startX: Number(event.clientX),
      startY: Number(event.clientY),
      screenStartX: Number(event.screenX),
      screenStartY: Number(event.screenY),
      captureTarget,
      started: false,
      moved: false,
      startModifiers: {
        altKey: Boolean(event.altKey),
        ctrlKey: Boolean(event.ctrlKey),
        metaKey: Boolean(event.metaKey),
        shiftKey: Boolean(event.shiftKey),
      },
    };
  }

  beginTabDragCandidate(event, header) {
    if (!(header instanceof HTMLElement)) {
      return;
    }

    const baseTarget = header.querySelector(".workspace-tab-header-inner");
    const dragTarget = baseTarget instanceof HTMLElement ? baseTarget : header;
    const baseRect = dragTarget.getBoundingClientRect();
    if (!this.isUsableRect(baseRect) || baseRect.width < TAB_DRAG_HANDLE_MIN_WIDTH) {
      return;
    }

    const rect = {
      left: baseRect.left + TAB_DRAG_HANDLE_SIDE_INSET,
      top: baseRect.top,
      width: Math.max(0, baseRect.width - TAB_DRAG_HANDLE_SIDE_INSET * 2),
      height: Math.min(TAB_DRAG_HANDLE_HEIGHT, baseRect.height),
    };
    if (!this.isUsableRect(rect)) {
      return;
    }

    const clientX = Number(event.clientX);
    const clientY = Number(event.clientY);
    const withinHandle =
      clientX >= rect.left &&
      clientX <= rect.left + rect.width &&
      clientY >= rect.top &&
      clientY <= rect.top + rect.height;
    if (!withinHandle) {
      return;
    }

    this.startTabDrag(event, {
      header,
      rect,
      captureTarget: header,
    });
  }

  handleTabDragPointerMove(event, drag) {
    this.markTouchInput();

    const dx = Number(event.clientX) - drag.startX;
    const dy = Number(event.clientY) - drag.startY;
    const distance = Math.hypot(dx, dy);
    if (!drag.started && distance >= TAB_DRAG_START_DISTANCE) {
      drag.started = true;
      drag.moved = true;
      if (typeof drag.captureTarget?.setPointerCapture === "function") {
        try {
          drag.captureTarget.setPointerCapture(event.pointerId);
        } catch (_error) {
          // Pointer capture is flaky enough that this should not kill the drag.
        }
      }
      this.dispatchSyntheticPointerEvent(drag.header, "pointerdown", {
        clientX: drag.startX,
        clientY: drag.startY,
        screenX: drag.screenStartX,
        screenY: drag.screenStartY,
        buttons: 1,
        button: 0,
        ...drag.startModifiers,
      });
      this.dispatchSyntheticMouseEvent(drag.header, "mousedown", {
        clientX: drag.startX,
        clientY: drag.startY,
        screenX: drag.screenStartX,
        screenY: drag.screenStartY,
        buttons: 1,
        button: 0,
        ...drag.startModifiers,
      });
      if (event.cancelable) {
        event.preventDefault();
      }
      event.stopPropagation();
      if (typeof event.stopImmediatePropagation === "function") {
        event.stopImmediatePropagation();
      }
    }

    if (!drag.started) {
      return;
    }

    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }

    const moveTarget = this.getUnderlyingElementAtPoint(event.clientX, event.clientY) ?? drag.header;
    this.dispatchSyntheticPointerEvent(moveTarget, "pointermove", {
      clientX: Number(event.clientX),
      clientY: Number(event.clientY),
      screenX: Number(event.screenX),
      screenY: Number(event.screenY),
      buttons: 1,
      button: 0,
      altKey: Boolean(event.altKey),
      ctrlKey: Boolean(event.ctrlKey),
      metaKey: Boolean(event.metaKey),
      shiftKey: Boolean(event.shiftKey),
    });
    this.dispatchSyntheticMouseEvent(moveTarget, "mousemove", {
      clientX: Number(event.clientX),
      clientY: Number(event.clientY),
      screenX: Number(event.screenX),
      screenY: Number(event.screenY),
      buttons: 1,
      button: 0,
      altKey: Boolean(event.altKey),
      ctrlKey: Boolean(event.ctrlKey),
      metaKey: Boolean(event.metaKey),
      shiftKey: Boolean(event.shiftKey),
    });
    this.dispatchSyntheticMouseEvent(document, "mousemove", {
      clientX: Number(event.clientX),
      clientY: Number(event.clientY),
      screenX: Number(event.screenX),
      screenY: Number(event.screenY),
      buttons: 1,
      button: 0,
      altKey: Boolean(event.altKey),
      ctrlKey: Boolean(event.ctrlKey),
      metaKey: Boolean(event.metaKey),
      shiftKey: Boolean(event.shiftKey),
    });
  }

  finishTabDrag(event, drag) {
    this.activeTabDrag = null;
    if (drag.started && typeof drag.captureTarget?.releasePointerCapture === "function") {
      try {
        drag.captureTarget.releasePointerCapture(event.pointerId);
      } catch (_error) {
        // Pointer capture is flaky enough that this should not kill the drag.
      }
    }

    if (drag.started) {
      if (event.cancelable) {
        event.preventDefault();
      }
      event.stopPropagation();
      if (typeof event.stopImmediatePropagation === "function") {
        event.stopImmediatePropagation();
      }

      const upTarget = this.getUnderlyingElementAtPoint(event.clientX, event.clientY) ?? drag.header;
      this.dispatchSyntheticPointerEvent(upTarget, "pointerup", {
        clientX: Number(event.clientX),
        clientY: Number(event.clientY),
        screenX: Number(event.screenX),
        screenY: Number(event.screenY),
        buttons: 0,
        button: 0,
        altKey: Boolean(event.altKey),
        ctrlKey: Boolean(event.ctrlKey),
        metaKey: Boolean(event.metaKey),
        shiftKey: Boolean(event.shiftKey),
      });
      this.dispatchSyntheticMouseEvent(upTarget, "mouseup", {
        clientX: Number(event.clientX),
        clientY: Number(event.clientY),
        screenX: Number(event.screenX),
        screenY: Number(event.screenY),
        buttons: 0,
        button: 0,
        altKey: Boolean(event.altKey),
        ctrlKey: Boolean(event.ctrlKey),
        metaKey: Boolean(event.metaKey),
        shiftKey: Boolean(event.shiftKey),
      });
      this.dispatchSyntheticMouseEvent(document, "mouseup", {
        clientX: Number(event.clientX),
        clientY: Number(event.clientY),
        screenX: Number(event.screenX),
        screenY: Number(event.screenY),
        buttons: 0,
        button: 0,
        altKey: Boolean(event.altKey),
        ctrlKey: Boolean(event.ctrlKey),
        metaKey: Boolean(event.metaKey),
        shiftKey: Boolean(event.shiftKey),
      });
      this.queueSuppressedClick(event);
    }

    this.scheduleOverlayRefresh();
  }

  finishModalDrag(event, drag) {
    this.activeModalDrag = null;
    if (typeof drag.captureTarget?.releasePointerCapture === "function") {
      try {
        drag.captureTarget.releasePointerCapture(event.pointerId);
      } catch (_error) {
        // Pointer capture is flaky enough that this should not kill the drag.
      }
    }

    if (drag.moved) {
      this.queueSuppressedClick(event);
    }

    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }

    this.scheduleOverlayRefresh();
  }

  queueSuppressedClick(event) {
    this.suppressedClick = {
      x: Number(event?.clientX) || 0,
      y: Number(event?.clientY) || 0,
      expiresAt: Date.now() + GHOST_CLICK_HOLD_MS,
    };
  }

  shouldSuppressGhostClick(event, suppressed) {
    if (!(event instanceof MouseEvent)) {
      return false;
    }

    if (event.target instanceof Element && event.target.closest(".modal-container")) {
      return true;
    }

    if (this.isInsideFloatingUi(event.target) || !this.isAllowedShimTarget(event.target)) {
      return false;
    }

    const distance = Math.hypot(
      Number(event.clientX) - Number(suppressed.x),
      Number(event.clientY) - Number(suppressed.y)
    );
    return distance <= GHOST_CLICK_MAX_DISTANCE;
  }

  hasBlockingModalOpen() {
    return Boolean(document.querySelector(".modal-container"));
  }

  isInsideFloatingUi(node) {
    return Boolean(node instanceof Element && node.closest(FLOATING_UI_SELECTOR));
  }

  getSidebarSide(target) {
    if (!(target instanceof Element)) {
      return null;
    }

    if (target.closest(".sidebar-toggle-button.mod-left")) {
      return "left";
    }

    if (target.closest(".sidebar-toggle-button.mod-right")) {
      return "right";
    }

    return null;
  }

  targetsMatch(left, right) {
    return (
      left === right ||
      left?.contains?.(right) ||
      right?.contains?.(left)
    );
  }

  dispatchSyntheticClick(target) {
    if (!(target instanceof HTMLElement)) {
      return;
    }

    target.focus({ preventScroll: true });
    if (typeof target.click === "function") {
      target.click();
      return;
    }

    for (const type of ["mousedown", "mouseup", "click"]) {
      target.dispatchEvent(
        new MouseEvent(type, {
          bubbles: true,
          cancelable: true,
          composed: true,
          view: window,
        })
      );
    }
  }

  toggleSidebar(side) {
    if (side !== "left" && side !== "right") {
      return false;
    }

    if (this.executeKnownSidebarCommand(side)) {
      return true;
    }

    if (this.toggleSidebarViaWorkspace(side)) {
      return true;
    }

    const targetButton = this.findSidebarButton(side);
    if (targetButton) {
      this.dispatchSyntheticClick(targetButton);
      return true;
    }

    new Notice(`Could not toggle ${side} sidebar`);
    return false;
  }

  executeKnownSidebarCommand(side) {
    const commandIds = SIDEBAR_COMMAND_IDS[side] ?? [];
    for (const commandId of commandIds) {
      if (this.app.commands?.commands?.[commandId]) {
        this.app.commands.executeCommandById(commandId);
        return true;
      }
    }

    return false;
  }

  toggleSidebarViaWorkspace(side) {
    const workspace = this.app.workspace;
    const split = side === "left" ? workspace?.leftSplit : workspace?.rightSplit;
    if (!split) {
      return false;
    }

    if (typeof split.toggle === "function") {
      split.toggle();
      return true;
    }

    const isOpen = document.querySelector(".workspace")?.classList.contains(SIDEBAR_OPEN_CLASSES[side]);
    if (typeof split.expand === "function" && typeof split.collapse === "function") {
      if (isOpen) {
        split.collapse();
      } else {
        split.expand();
      }
      return true;
    }

    return false;
  }

  findSidebarButton(side) {
    const selectors = [
      `.sidebar-toggle-button.mod-${side} .clickable-icon`,
      `.sidebar-toggle-button.mod-${side}`,
      `.workspace-ribbon .sidebar-toggle-button.mod-${side} .clickable-icon`,
    ];

    for (const selector of selectors) {
      const found = Array.from(document.querySelectorAll(selector)).find((node) =>
        node instanceof HTMLElement && this.isInteractableElement(node)
      );
      if (found instanceof HTMLElement) {
        return found;
      }
    }

    return null;
  }

  markTouchInput() {
    document.body.classList.add("surface-touch-mode-last-input");
    window.clearTimeout(this.touchClassTimer);
    this.touchClassTimer = window.setTimeout(() => {
      document.body.classList.remove("surface-touch-mode-last-input");
    }, TOUCH_CLASS_HOLD_MS);
  }

  clearActiveTap() {
    this.activeTap = null;
    this.activeModalDrag = null;
    this.activeTabDrag = null;
  }

  handlePointerCancel(event) {
    const tabDrag = this.activeTabDrag;
    if (tabDrag && tabDrag.pointerId === event?.pointerId) {
      this.activeTabDrag = null;
      if (typeof tabDrag.captureTarget?.releasePointerCapture === "function") {
        try {
          tabDrag.captureTarget.releasePointerCapture(event.pointerId);
        } catch (_error) {
          // Pointer capture is flaky enough that this should not kill the drag.
        }
      }
      this.scheduleOverlayRefresh();
    }

    const drag = this.activeModalDrag;
    if (drag && drag.pointerId === event?.pointerId) {
      this.activeModalDrag = null;
      if (typeof drag.captureTarget?.releasePointerCapture === "function") {
        try {
          drag.captureTarget.releasePointerCapture(event.pointerId);
        } catch (_error) {
          // Pointer capture is flaky enough that this should not kill the drag.
        }
      }
      this.scheduleOverlayRefresh();
    }

    this.activeTap = null;
  }

  installStyle() {
    if (!(this.styleEl instanceof HTMLStyleElement)) {
      this.styleEl = document.createElement("style");
      this.styleEl.setAttribute("data-surface-touch-mode", "true");
      document.head.appendChild(this.styleEl);
    }

    this.styleEl.textContent = `
      body.surface-touch-mode-enabled {
        --surface-touch-edge-zone-width: ${Number(this.settings.edgeZoneWidth) || DEFAULT_SETTINGS.edgeZoneWidth}px;
      }

      body.surface-touch-mode-enabled .workspace-ribbon .side-dock-ribbon-action,
      body.surface-touch-mode-enabled .workspace-ribbon .sidebar-toggle-button,
      body.surface-touch-mode-enabled .sidebar-toggle-button,
      body.surface-touch-mode-enabled .mod-sidedock .nav-action-button {
        touch-action: manipulation;
        -webkit-tap-highlight-color: transparent;
      }

      body.surface-touch-mode-last-input .workspace-ribbon .side-dock-ribbon-action,
      body.surface-touch-mode-last-input .workspace-ribbon .sidebar-toggle-button,
      body.surface-touch-mode-last-input .sidebar-toggle-button,
      body.surface-touch-mode-last-input .mod-sidedock .nav-action-button,
      body.surface-touch-mode-last-input .mod-sidedock .clickable-icon {
        transition: none !important;
        animation: none !important;
      }

      body.surface-touch-mode-last-input .workspace-ribbon .side-dock-ribbon-action:active,
      body.surface-touch-mode-last-input .workspace-ribbon .sidebar-toggle-button:active,
      body.surface-touch-mode-last-input .sidebar-toggle-button:active,
      body.surface-touch-mode-last-input .mod-sidedock .nav-action-button:active,
      body.surface-touch-mode-last-input .mod-sidedock .clickable-icon:active,
      body.surface-touch-mode-last-input .workspace-ribbon .side-dock-ribbon-action:active svg,
      body.surface-touch-mode-last-input .workspace-ribbon .sidebar-toggle-button:active svg,
      body.surface-touch-mode-last-input .sidebar-toggle-button:active svg,
      body.surface-touch-mode-last-input .mod-sidedock .nav-action-button:active svg,
      body.surface-touch-mode-last-input .mod-sidedock .clickable-icon:active svg {
        transform: none !important;
        filter: none !important;
        opacity: 1 !important;
        scale: 1 !important;
      }

      .surface-touch-mode-root {
        position: fixed;
        inset: 0;
        pointer-events: none;
        z-index: 9999;
      }

      .surface-touch-mode-edge {
        pointer-events: auto;
        touch-action: manipulation;
        -webkit-tap-highlight-color: transparent;
      }

      .surface-touch-mode-hotspot,
      .surface-touch-mode-modal-handle {
        position: fixed;
        pointer-events: auto;
        background: transparent;
        -webkit-tap-highlight-color: transparent;
      }

      .surface-touch-mode-edge {
        position: fixed;
        top: 0;
        bottom: 0;
        width: var(--surface-touch-edge-zone-width);
        background: transparent;
        opacity: 0;
      }

      .surface-touch-mode-edge.mod-left {
        left: 0;
      }

      .surface-touch-mode-edge.mod-right {
        right: 0;
      }

      .surface-touch-mode-hotspot {
        touch-action: manipulation;
      }

      .surface-touch-mode-modal-handle {
        touch-action: none;
      }
      body.surface-touch-mode-enabled .modal.surface-touch-mode-modal-dragged {
        transform: translate3d(
          var(--surface-touch-modal-drag-x, 0px),
          var(--surface-touch-modal-drag-y, 0px),
          0
        ) !important;
      }
    `;

    document.body.classList.add("surface-touch-mode-enabled");
  }

  ensureOverlay() {
    if (!document.body) {
      return;
    }

    if (!(this.rootEl instanceof HTMLElement)) {
      this.rootEl = document.createElement("div");
      this.rootEl.className = "surface-touch-mode-root";
      document.body.appendChild(this.rootEl);
    }

    this.rootEl.replaceChildren();

    const suppressRightAssist = this.shouldSuppressSidebarAssist("right");

    if (this.settings.edgeZonesEnabled) {
      this.rootEl.appendChild(this.createEdgeZone("left"));
      if (!suppressRightAssist) {
        this.rootEl.appendChild(this.createEdgeZone("right"));
      }
    }

    if (this.settings.touchShimEnabled) {
      const leftHotspot = this.createSidebarHotspot("left");
      const rightHotspot = suppressRightAssist ? null : this.createSidebarHotspot("right");
      if (leftHotspot) {
        this.rootEl.appendChild(leftHotspot);
      }
      if (rightHotspot) {
        this.rootEl.appendChild(rightHotspot);
      }
    }

    if (this.settings.modalDragShimEnabled) {
      for (const handle of this.createModalDragHandles()) {
        this.rootEl.appendChild(handle);
      }
    }

  }

  shouldSuppressSidebarAssist(side) {
    if (side !== "right") {
      return false;
    }

    const drawingSelector = [
      '.workspace-leaf-content[data-type="excalidraw"]',
      '.workspace-leaf-content[data-type="tldraw-view"]',
    ].join(", ");
    const threshold = Math.max(56, Number(this.settings.edgeZoneWidth ?? 0) + 56);

    for (const node of document.querySelectorAll(drawingSelector)) {
      if (!(node instanceof HTMLElement) || !this.isInteractableElement(node)) {
        continue;
      }

      const rect = node.getBoundingClientRect();
      if (rect.right >= window.innerWidth - threshold) {
        return true;
      }
    }

    return false;
  }

  createEdgeZone(side) {
    const zone = document.createElement("div");
    zone.className = `surface-touch-mode-edge mod-${side}`;
    zone.setAttribute("aria-hidden", "true");
    zone.addEventListener("pointerup", (event) => this.handleOverlayActivator(side, event));
    return zone;
  }

  handleOverlayActivator(side, event) {
    if (!this.isPrimaryTouchLikePointer(event) && event.pointerType !== "mouse") {
      return;
    }

    if (this.hasBlockingModalOpen()) {
      return;
    }

    if (event.cancelable) {
      event.preventDefault();
    }
    event.stopPropagation();
    if (typeof event.stopImmediatePropagation === "function") {
      event.stopImmediatePropagation();
    }

    this.queueSuppressedClick(event);
    this.markTouchInput();
    this.toggleSidebar(side);
  }

  createSidebarHotspot(side) {
    const button = this.findSidebarButton(side);
    if (!(button instanceof HTMLElement)) {
      return null;
    }

    const rect = button.getBoundingClientRect();
    if (!this.isUsableRect(rect)) {
      return null;
    }

    const hotspot = document.createElement("div");
    hotspot.className = `surface-touch-mode-hotspot mod-${side}`;
    hotspot.setAttribute("aria-hidden", "true");
    this.positionOverlayNode(hotspot, {
      left: rect.left - SIDEBAR_HOTSPOT_PADDING,
      top: rect.top - SIDEBAR_HOTSPOT_PADDING,
      width: rect.width + SIDEBAR_HOTSPOT_PADDING * 2,
      height: rect.height + SIDEBAR_HOTSPOT_PADDING * 2,
    });
    hotspot.addEventListener("pointerup", (event) => this.handleOverlayActivator(side, event));
    hotspot.addEventListener("contextmenu", (event) => {
      event.preventDefault();
    });
    return hotspot;
  }

  createModalDragHandles() {
    const handles = [];
    const modals = Array.from(document.querySelectorAll(".modal"));
    for (const modal of modals) {
      if (!(modal instanceof HTMLElement) || !this.isInteractableElement(modal)) {
        continue;
      }

      this.resetLegacyModalInlineLayout(modal);

      const rect = modal.getBoundingClientRect();
      if (!this.isUsableRect(rect)) {
        continue;
      }

      const width = rect.width - MODAL_HANDLE_LEFT_INSET - MODAL_HANDLE_RIGHT_INSET;
      if (width < 96) {
        continue;
      }

      const handle = document.createElement("div");
      handle.className = "surface-touch-mode-modal-handle";
      handle.setAttribute("aria-hidden", "true");
      this.positionOverlayNode(handle, {
        left: rect.left + MODAL_HANDLE_LEFT_INSET,
        top: rect.top,
        width,
        height: Math.min(MODAL_DRAG_HANDLE_HEIGHT, rect.height),
      });
      handle.addEventListener("pointerdown", (event) => this.handleModalOverlayPointerDown(modal, event));
      handle.addEventListener("contextmenu", (event) => {
        event.preventDefault();
      });
      handles.push(handle);
    }

    return handles;
  }

  isEligibleTabDragHeader(header) {
    if (!(header instanceof HTMLElement)) {
      return false;
    }

    if (!header.closest(".workspace-tabs.mod-top")) {
      return false;
    }

    if (header.closest(".mod-left-split, .mod-right-split, .workspace-drawer-tab-options, .mod-sidedock")) {
      return false;
    }

    return true;
  }

  handleModalOverlayPointerDown(modal, event) {
    if (!this.isPrimaryTouchLikePointer(event)) {
      return;
    }

    this.resetLegacyModalInlineLayout(modal);
    const rect = modal.getBoundingClientRect();
    if (!this.isUsableRect(rect)) {
      return;
    }

    this.startModalDrag(event, { modal, rect });
  }

  getModalDragOffset(modal) {
    if (!(modal instanceof HTMLElement)) {
      return { x: 0, y: 0 };
    }

    const x = Number.parseFloat(modal.dataset.surfaceTouchDragX ?? "0");
    const y = Number.parseFloat(modal.dataset.surfaceTouchDragY ?? "0");
    return {
      x: Number.isFinite(x) ? x : 0,
      y: Number.isFinite(y) ? y : 0,
    };
  }

  applyModalDragOffset(modal, x, y) {
    if (!(modal instanceof HTMLElement)) {
      return;
    }

    const roundedX = Math.round(Number.isFinite(x) ? x : 0);
    const roundedY = Math.round(Number.isFinite(y) ? y : 0);
    modal.dataset.surfaceTouchDragX = String(roundedX);
    modal.dataset.surfaceTouchDragY = String(roundedY);
    modal.style.setProperty("--surface-touch-modal-drag-x", `${roundedX}px`);
    modal.style.setProperty("--surface-touch-modal-drag-y", `${roundedY}px`);
    modal.classList.add("surface-touch-mode-modal-dragged");
  }

  resetLegacyModalInlineLayout(modal) {
    if (!(modal instanceof HTMLElement)) {
      return;
    }

    const style = modal.style;
    const looksLikeLegacyShim =
      style.position === "fixed" ||
      style.margin === "0" ||
      style.width !== "" ||
      style.maxWidth !== "" ||
      style.left !== "" ||
      style.top !== "";

    if (!looksLikeLegacyShim) {
      return;
    }

    style.removeProperty("position");
    style.removeProperty("left");
    style.removeProperty("top");
    style.removeProperty("margin");
    style.removeProperty("width");
    style.removeProperty("max-width");
    if (!modal.classList.contains("surface-touch-mode-modal-dragged")) {
      style.removeProperty("transform");
    }
  }

  getTabDragHeaderFromEventTarget(node) {
    if (!(node instanceof Element)) {
      return null;
    }

    if (node.closest(".workspace-tab-header-inner-close-button, .workspace-tab-header-status-icon, .clickable-icon")) {
      return null;
    }

    const header = node.closest(ELIGIBLE_TAB_HEADER_SELECTOR);
    if (!(header instanceof HTMLElement) || !this.isEligibleTabDragHeader(header)) {
      return null;
    }

    return header;
  }

  getUnderlyingElementAtPoint(x, y) {
    if (!(this.rootEl instanceof HTMLElement)) {
      return document.elementFromPoint(Number(x) || 0, Number(y) || 0);
    }

    const previousPointerEvents = this.rootEl.style.pointerEvents;
    this.rootEl.style.pointerEvents = "none";
    try {
      return document.elementFromPoint(Number(x) || 0, Number(y) || 0);
    } finally {
      this.rootEl.style.pointerEvents = previousPointerEvents;
    }
  }

  dispatchSyntheticPointerEvent(target, type, init) {
    if (!(target instanceof EventTarget) || typeof window.PointerEvent !== "function") {
      return;
    }

    target.dispatchEvent(
      new PointerEvent(type, {
        bubbles: true,
        cancelable: true,
        composed: true,
        view: window,
        pointerId: 1,
        isPrimary: true,
        pointerType: "mouse",
        ...init,
      })
    );
  }

  dispatchSyntheticMouseEvent(target, type, init) {
    if (!(target instanceof EventTarget)) {
      return;
    }

    target.dispatchEvent(
      new MouseEvent(type, {
        bubbles: true,
        cancelable: true,
        composed: true,
        view: window,
        ...init,
      })
    );
  }

  positionOverlayNode(node, rect) {
    if (!(node instanceof HTMLElement) || !rect) {
      return;
    }

    node.style.left = `${Math.round(Number(rect.left) || 0)}px`;
    node.style.top = `${Math.round(Number(rect.top) || 0)}px`;
    node.style.width = `${Math.max(0, Math.round(Number(rect.width) || 0))}px`;
    node.style.height = `${Math.max(0, Math.round(Number(rect.height) || 0))}px`;
  }

  isUsableRect(rect) {
    return Boolean(
      rect &&
      Number.isFinite(rect.width) &&
      Number.isFinite(rect.height) &&
      rect.width > 0 &&
      rect.height > 0
    );
  }

  isInteractableElement(node) {
    if (!(node instanceof HTMLElement)) {
      return false;
    }

    const rect = node.getBoundingClientRect();
    if (!this.isUsableRect(rect)) {
      return false;
    }

    const style = window.getComputedStyle(node);
    return style.display !== "none" && style.visibility !== "hidden";
  }

  startOverlayObserver() {
    if (this.overlayObserver || !(document.body instanceof HTMLElement)) {
      return;
    }

    this.overlayObserver = new MutationObserver((mutations) => {
      const shouldRefresh = mutations.some((mutation) => !this.isManagedOverlayNode(mutation.target));
      if (shouldRefresh) {
        this.scheduleOverlayRefresh();
      }
    });
    this.overlayObserver.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["class", "style"],
    });

    this.register(() => {
      if (this.overlayObserver) {
        this.overlayObserver.disconnect();
        this.overlayObserver = null;
      }
    });
  }

  scheduleOverlayRefresh() {
    if (this.overlayRefreshFrame) {
      return;
    }

    this.overlayRefreshFrame = window.requestAnimationFrame(() => {
      this.overlayRefreshFrame = 0;
      this.ensureOverlay();
    });
  }

  isManagedOverlayNode(node) {
    if (!(node instanceof Node)) {
      return false;
    }

    if (this.rootEl instanceof HTMLElement && (node === this.rootEl || this.rootEl.contains(node))) {
      return true;
    }

    return node instanceof HTMLStyleElement && node === this.styleEl;
  }

  isInsideSurfaceTouchOverlay(node) {
    return node instanceof Element && node.closest(".surface-touch-mode-root");
  }

  teardown() {
    window.clearTimeout(this.touchClassTimer);
    if (this.overlayRefreshFrame) {
      window.cancelAnimationFrame(this.overlayRefreshFrame);
      this.overlayRefreshFrame = 0;
    }
    document.body?.classList.remove("surface-touch-mode-enabled", "surface-touch-mode-last-input");
    this.activeTap = null;
    this.activeModalDrag = null;
    this.activeTabDrag = null;
    this.suppressedClick = null;
    if (this.rootEl instanceof HTMLElement) {
      this.rootEl.remove();
      this.rootEl = null;
    }
    if (this.styleEl instanceof HTMLStyleElement) {
      this.styleEl.remove();
      this.styleEl = null;
    }
  }
};

class SurfaceTouchModeSettingTab extends PluginSettingTab {
  constructor(app, plugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();

    new Setting(containerEl)
      .setName("Touch click shim")
      .setDesc("Turn touch and pen taps on sidebar and ribbon controls into reliable clicks.")
      .addToggle((toggle) =>
        toggle.setValue(Boolean(this.plugin.settings.touchShimEnabled)).onChange(async (value) => {
          this.plugin.settings.touchShimEnabled = value;
          await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("Modal drag shim")
      .setDesc("Let pen and touch drag Obsidian modal title bars immediately instead of waiting on long-press behavior.")
      .addToggle((toggle) =>
        toggle.setValue(Boolean(this.plugin.settings.modalDragShimEnabled)).onChange(async (value) => {
          this.plugin.settings.modalDragShimEnabled = value;
          await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("Tab drag shim")
      .setDesc("Let pen and touch drag note tabs from the top tab bar without needing a keyboard or trackpad.")
      .addToggle((toggle) =>
        toggle.setValue(Boolean(this.plugin.settings.tabDragShimEnabled)).onChange(async (value) => {
          this.plugin.settings.tabDragShimEnabled = value;
          await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("Edge sidebar zones")
      .setDesc("Tap the left or right screen edge to open the corresponding sidebar.")
      .addToggle((toggle) =>
        toggle.setValue(Boolean(this.plugin.settings.edgeZonesEnabled)).onChange(async (value) => {
          this.plugin.settings.edgeZonesEnabled = value;
          await this.plugin.saveSettings();
        })
      );

    new Setting(containerEl)
      .setName("Edge zone width")
      .setDesc("Width of each invisible sidebar edge zone in pixels.")
      .addText((text) =>
        text.setValue(String(this.plugin.settings.edgeZoneWidth)).onChange(async (value) => {
          const parsed = Number(value);
          if (!Number.isFinite(parsed) || parsed < 8) {
            return;
          }

          this.plugin.settings.edgeZoneWidth = Math.round(parsed);
          await this.plugin.saveSettings();
        })
      );
  }
}
