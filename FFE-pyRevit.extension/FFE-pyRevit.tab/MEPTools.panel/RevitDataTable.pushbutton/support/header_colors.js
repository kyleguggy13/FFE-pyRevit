/* Column fills belong to the working layout; only explicit Save writes them to Revit. */
(function () {
  "use strict";
  const validColor = value => /^#[0-9a-f]{6}$/i.test(value);

  // Choose the higher-contrast text color for arbitrary custom fills.
  function textColor(color) {
    const rgb = [1, 3, 5].map(offset => parseInt(color.slice(offset, offset + 2), 16) / 255)
      .map(value => value <= 0.04045 ? value / 12.92 : ((value + 0.055) / 1.055) ** 2.4);
    return rgb[0] * 0.2126 + rgb[1] * 0.7152 + rgb[2] * 0.0722 > 0.179 ? "#000000" : "#ffffff";
  }

  /** Own the header context menu, draft color picker, and scoped column styling. */
  class HeaderColorControls {
    constructor(options) {
      this.options = options;
      this.listeners = [];
      this.grid = document.getElementById("grid");
      this.sheet = document.createElement("style");
      document.head.append(this.sheet);
      this.menu = document.createElement("div");
      this.menu.id = "header-color-menu";
      this.menu.setAttribute("role", "menu");
      this.menu.hidden = true;
      this.change = this.button("Change color…", () => this.openPicker());
      this.reset = this.button("Reset color", () => {
        if (options.isReady()) options.setColor(this.target, null);
        this.closeMenu(true);
      });
      this.menu.append(this.change, this.reset);
      this.dialog = document.createElement("dialog");
      this.dialog.id = "header-color-dialog";
      this.dialog.setAttribute("aria-labelledby", "header-color-title");
      this.dialog.innerHTML = '<form method="dialog"><h2 id="header-color-title"></h2>' +
        '<p id="header-color-caption"></p><label>Fill color<input id="header-color-input" type="color"></label>' +
        '<div class="dialog-actions"><button value="cancel">Cancel</button>' +
        '<button class="primary" value="apply">Apply</button></div></form>';
      document.body.append(this.menu, this.dialog);
      this.listen(this.dialog, "close", () => {
        const color = this.dialog.querySelector("input").value;
        if (this.dialog.returnValue === "apply" && options.isReady() && validColor(color)) {
          options.setColor(this.target, color);
        }
        this.restoreFocus();
      });
      this.listen(this.grid, "contextmenu", event => this.openMenu(event));
      this.listen(document, "pointerdown", event => {
        if (!this.menu.contains(event.target)) this.closeMenu();
      });
      this.listen(document, "scroll", () => this.closeMenu(), true);
      this.listen(window, "resize", () => this.closeMenu());
      this.listen(this.menu, "keydown", event => {
        if (event.key === "Escape") { event.preventDefault(); this.closeMenu(true); }
        else if (event.key === "Tab") this.closeMenu();
        else if (["ArrowUp", "ArrowDown", "Home", "End"].includes(event.key)) {
          event.preventDefault();
          const buttons = [...this.menu.querySelectorAll("button:enabled")];
          const index = buttons.indexOf(document.activeElement);
          const next = event.key === "Home" ? 0 : event.key === "End" ? buttons.length - 1
            : (index + (event.key === "ArrowDown" ? 1 : -1) + buttons.length) % buttons.length;
          buttons[next].focus();
        }
      });
    }

    listen(target, type, handler, capture = false) {
      target.addEventListener(type, handler, capture);
      this.listeners.push({ target, type, handler, capture });
    }

    button(text, handler) {
      const button = document.createElement("button");
      button.type = "button";
      button.setAttribute("role", "menuitem");
      button.textContent = text;
      this.listen(button, "click", handler);
      return button;
    }

    groupFor(field) { return this.options.getGroups().find(group => group.fields.includes(field)); }

    colorFor(target) {
      return target.group ? this.groupFor(target.field)?.color : this.options.getColumnColors().get(target.field);
    }

    openMenu(event) {
      const header = event.target.closest(".tabulator-col");
      if (!header || !this.grid.querySelector(".tabulator-header").contains(header)) return;
      event.preventDefault();
      if (!this.options.isReady()) return;
      const group = header.classList.contains("tabulator-col-group");
      const field = group ? header.querySelector("[tabulator-field]")?.getAttribute("tabulator-field")
        : header.getAttribute("tabulator-field");
      if (!field) return;
      this.target = { field, group };
      this.menu.setAttribute("aria-label", group ? "Column group color" : "Column color");
      this.reset.disabled = !this.colorFor(this.target);
      this.menu.hidden = false;
      const bounds = header.getBoundingClientRect();
      this.menu.style.left = `${Math.max(4, Math.min(event.clientX || bounds.left, innerWidth - this.menu.offsetWidth - 4))}px`;
      this.menu.style.top = `${Math.max(4, Math.min(event.clientY || bounds.bottom, innerHeight - this.menu.offsetHeight - 4))}px`;
      this.change.focus();
    }

    closeMenu(focus = false) {
      const wasOpen = !this.menu.hidden;
      this.menu.hidden = true;
      if (wasOpen && focus) this.restoreFocus();
    }

    restoreFocus() {
      // Headers are not normally keyboard focus targets; keep grid navigation usable.
      this.grid.focus({ preventScroll: true });
    }

    openPicker() {
      if (!this.options.isReady()) return;
      this.closeMenu();
      const target = this.target;
      const group = this.groupFor(target.field);
      this.dialog.querySelector("h2").textContent = target.group ? "Column group color" : "Column color";
      this.dialog.querySelector("p").textContent = target.group
        ? `${group.name}: applies to the band and its columns. Individual column colors take priority.`
        : this.options.table.getColumn(target.field).getDefinition().title;
      this.dialog.querySelector("input").value = this.colorFor(target) || group?.color || "#d9efed";
      this.dialog.returnValue = "cancel";
      this.dialog.showModal();
    }

    /** CSS selectors color virtualized/new rows without rewriting cell formatters. */
    update() {
      if (!this.options.isReady()) {
        this.closeMenu();
        if (this.dialog.open) { this.dialog.returnValue = "cancel"; this.dialog.close(); }
      }
      const rules = [];
      for (const column of this.options.table.getColumns()) {
        const field = column.getField();
        const color = this.options.getColumnColors().get(field) || this.groupFor(field)?.color;
        if (!validColor(color)) continue;
        const selector = `[tabulator-field="${CSS.escape(field)}"]`;
        const ink = textColor(color);
        rules.push(`#grid .tabulator-header .tabulator-col${selector},
          #grid .tabulator-header .tabulator-col${selector} > .tabulator-col-content,
          #grid .tabulator-row:not(.tabulator-selected) .tabulator-cell${selector}
          { background-color: ${color}; color: ${ink}; }
          #grid .tabulator-col${selector} .sort-priority { color: ${ink}; }
          #grid .tabulator-col${selector}[aria-sort="ascending"] .tabulator-arrow { border-bottom-color: ${ink}; }
          #grid .tabulator-col${selector}[aria-sort="descending"] .tabulator-arrow { border-top-color: ${ink}; }`);
      }
      const css = rules.join("\n");
      if (this.sheet.textContent !== css) this.sheet.textContent = css;
      for (const column of this.options.table.getColumns(true)) {
        const children = column.getSubColumns();
        if (!children.length) continue;
        const color = this.groupFor(children[0].getField())?.color;
        for (const element of [column.getElement(), column.getElement().querySelector(":scope > .tabulator-col-content")]) {
          element.style.backgroundColor = validColor(color) ? color : "";
          element.style.color = validColor(color) ? textColor(color) : "";
        }
      }
    }

    destroy() {
      for (const { target, type, handler, capture } of this.listeners) target.removeEventListener(type, handler, capture);
      this.sheet.remove();
      this.menu.remove();
      this.dialog.remove();
    }
  }
  window.HeaderColorControls = HeaderColorControls;
})();
