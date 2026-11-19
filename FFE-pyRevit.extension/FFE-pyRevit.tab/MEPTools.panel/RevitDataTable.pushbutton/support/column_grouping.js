/* Native header dragging uses Tabulator's public column API for both drop targets. */
(function () {
  "use strict";
  /** Connect native drags to ordered grouping levels and public column moves. */
  class ColumnGroupingDrag {
    /** Wire drop targets; callbacks keep the working/saved layout in the existing state. */
    constructor(options) {
      this.options = options;
      this.grid = document.getElementById("grid");
      this.zone = document.getElementById("group-drop-zone");
      this.label = document.getElementById("group-drop-label");
      this.chips = document.getElementById("group-chips");
      this.chipSignature = "";
      this.clear = document.getElementById("group-clear");
      this.session = null;
      this.overZone = false;
      this.resizeStart = false;
      this.marker = null;
      this.listeners = [];
      this.listen(this.grid, "mousedown", event => {
        this.resizeStart = Boolean(event.target.closest(".tabulator-col-resize-handle"));
      }, true);
      this.listen(this.grid, "dragstart", event => this.start(event));
      this.listen(this.grid, "dragover", event => this.overHeader(event));
      this.listen(this.grid, "dragleave", event => {
        if (!this.grid.contains(event.relatedTarget)) this.clearMarker();
      });
      this.listen(this.grid, "drop", event => this.dropHeader(event));
      this.listen(this.zone, "dragover", event => this.overGrouping(event));
      this.listen(this.zone, "dragleave", event => {
        if (!this.zone.contains(event.relatedTarget)) {
          this.overZone = false;
          this.clearMarker();
          this.render();
        }
      });
      this.listen(this.zone, "drop", event => this.dropGrouping(event));
      this.listen(this.chips, "dragstart", event => {
        const chip = event.target.closest(".group-chip");
        if (chip && !event.target.closest("button")) this.begin(event, chip.dataset.field, "chip", chip);
        else event.preventDefault();
      });
      this.listen(this.chips, "click", event => {
        const button = event.target.closest(".group-remove");
        if (button && this.options.isReady()) {
          const field = button.closest(".group-chip").dataset.field;
          this.options.setGroups(this.options.getGroups().filter(item => item !== field));
          this.zone.focus();
        }
      });
      this.listen(this.chips, "keydown", event => {
        const chip = event.target.closest(".group-chip");
        if (!chip || event.target !== chip || !event.altKey || !["ArrowLeft", "ArrowRight"].includes(event.key) || !this.options.isReady()) return;
        event.preventDefault();
        const fields = [...this.options.getGroups()];
        const index = fields.indexOf(chip.dataset.field);
        const next = index + (event.key === "ArrowLeft" ? -1 : 1);
        if (next < 0 || next >= fields.length) return;
        [fields[index], fields[next]] = [fields[next], fields[index]];
        this.options.setGroups(fields);
        Array.from(this.chips.querySelectorAll(".group-chip")).find(item => item.dataset.field === fields[next])?.focus();
      });
      this.listen(document, "dragend", () => this.cancel());
      this.listen(document, "keydown", event => { if (event.key === "Escape") this.cancel(); });
      this.listen(window, "blur", () => this.cancel());
      this.listen(this.clear, "click", () => {
        if (this.options.isReady()) this.options.setGroups([]);
      });
    }

    listen(target, type, handler, capture = false) {
      target.addEventListener(type, handler, capture);
      this.listeners.push({ target, type, handler, capture });
    }

    column(field) {
      return this.options.table.getColumns().find(column => column.isVisible() && column.getField() === field);
    }

    band(field) {
      return this.options.table.getColumns(true).find(column =>
        column.getSubColumns()[0]?.getField() === field && column.isVisible());
    }

    header(event) {
      const header = event.target.closest?.(".tabulator-col");
      return header && this.grid.contains(header) &&
        (header.hasAttribute("tabulator-field") || header.dataset.headerGroup) ? header : null;
    }

    // A whole band lands beside the target band, including when hovering its children.
    targetHeader(event) {
      const header = this.header(event);
      return this.session?.source === "band" && header
        ? header.closest(".tabulator-col-group") || header : header;
    }

    valid() {
      return Boolean(this.session && this.options.isReady() &&
        (this.session.source === "band" ? this.band(this.session.field) : this.column(this.session.field)));
    }

    // Only drags originating in this grid can change its grouping or order.
    start(event) {
      const header = this.header(event);
      if (!header) return;
      if (!this.options.isReady() || this.resizeStart || !event.dataTransfer) {
        event.preventDefault();
        return;
      }
      const field = header.dataset.headerGroup || header.getAttribute("tabulator-field");
      this.begin(event, field, header.dataset.headerGroup ? "band" : "header", header);
    }

    // Keep the source node mounted throughout native dragging, including chip drags.
    begin(event, field, source, element) {
      if (!this.options.isReady() || !event.dataTransfer) {
        event.preventDefault();
        return;
      }
      if (!(source === "band" ? this.band(field) : this.column(field))) {
        event.preventDefault();
        return;
      }
      this.session = { field, source };
      event.dataTransfer.effectAllowed = "copyMove";
      event.dataTransfer.setData("application/x-ffe-table-column", field);
      element.classList.add("column-drag-source");
      this.render();
    }

    overGrouping(event) {
      if (!this.valid() || this.session.source === "band") return;
      event.preventDefault();
      event.dataTransfer.dropEffect = this.session.source === "chip" ? "move" : "copy";
      this.overZone = true;
      this.clearMarker();
      const chip = event.target.closest(".group-chip");
      if (chip && chip.dataset.field !== this.session.field) {
        const bounds = chip.getBoundingClientRect();
        chip.classList.add(event.clientX > bounds.left + bounds.width / 2 ? "group-drop-after" : "group-drop-before");
        this.marker = chip;
      }
      this.render();
    }

    dropGrouping(event) {
      event.preventDefault();
      if (!this.valid() || this.session.source === "band") {
        this.cancel();
        return;
      }
      const field = this.session.field;
      const fields = [...this.options.getGroups()];
      const chip = event.target.closest(".group-chip");
      // A repeated header drop on empty space is a no-op, never a duplicate level.
      if (chip?.dataset.field === field || (this.session.source === "header" && fields.includes(field) && !chip)) {
        this.cancel();
        return;
      }
      const ordered = fields.filter(item => item !== field);
      let index = ordered.length;
      if (chip) {
        const bounds = chip.getBoundingClientRect();
        index = ordered.indexOf(chip.dataset.field) + (event.clientX > bounds.left + bounds.width / 2 ? 1 : 0);
      }
      ordered.splice(index, 0, field);
      this.cancel();
      this.options.setGroups(ordered);
    }

    overHeader(event) {
      if (!this.valid() || this.session.source === "chip") return;
      const header = this.targetHeader(event);
      if (!header) return;
      event.preventDefault();
      event.dataTransfer.dropEffect = "move";
      this.overZone = false;
      this.clearMarker();
      if ((header.dataset.headerGroup || header.getAttribute("tabulator-field")) !== this.session.field) {
        const bounds = header.getBoundingClientRect();
        header.classList.add(event.clientX > bounds.left + bounds.width / 2 ? "column-drop-after" : "column-drop-before");
        this.marker = header;
      }
      this.render();
    }

    // Delay all layout changes until drop so cancellation needs no restoration.
    dropHeader(event) {
      if (!this.valid() || this.session.source === "chip") return;
      const header = this.targetHeader(event);
      if (!header) return;
      event.preventDefault();
      const source = this.session.field;
      const sourceKind = this.session.source;
      const bounds = header.getBoundingClientRect();
      const after = event.clientX > bounds.left + bounds.width / 2;
      const band = header.dataset.headerGroup && this.band(header.dataset.headerGroup);
      const members = band?.getSubColumns().filter(column => column.isVisible());
      const target = band && sourceKind !== "band"
        ? members[after ? members.length - 1 : 0]?.getField()
        : header.dataset.headerGroup || header.getAttribute("tabulator-field");
      this.cancel();
      if (source === target) return;
      if (sourceKind === "band") this.options.moveBand(source, target, after);
      else if (this.column(target)) this.options.moveColumn(source, target, after);
    }

    clearMarker() {
      if (this.marker) this.marker.classList.remove("column-drop-before", "column-drop-after", "group-drop-before", "group-drop-after");
      this.marker = null;
    }

    /** Clear transient drag feedback without changing the working layout. */
    cancel() {
      this.session = null;
      this.overZone = false;
      this.clearMarker();
      for (const root of [this.grid, this.chips]) {
        for (const element of root.querySelectorAll(".column-drag-source")) element.classList.remove("column-drag-source");
      }
      this.render();
    }

    render() {
      const fields = this.options.getGroups();
      const ready = this.options.isReady();
      this.zone.classList.toggle("group-drop-ready", Boolean(this.session && this.session.source !== "band"));
      this.zone.classList.toggle("group-drop-over", this.overZone);
      this.label.textContent = this.overZone && this.session
        ? `Drop to group by ${this.options.caption(this.session.field)}`
        : fields.length ? "Drag another column here" : "Drag column headers here to group";
      const signature = JSON.stringify([fields.map(field => [field, this.options.caption(field)]), ready]);
      if (signature !== this.chipSignature) {
        this.chipSignature = signature;
        const steps = fields.map((field, index) => {
          const step = document.createElement("li");
          step.className = "group-step";
          if (index) {
            const arrow = document.createElement("span");
            arrow.className = "group-arrow";
            arrow.setAttribute("aria-hidden", "true");
            step.append(arrow);
          }
          const chip = document.createElement("div");
          chip.className = "group-chip";
          chip.dataset.field = field;
          chip.draggable = ready;
          chip.tabIndex = ready ? 0 : -1;
          chip.title = "Drag to reorder grouping; Alt+Left/Right also moves this level";
          chip.setAttribute("aria-label", `Grouping level ${index + 1}: ${this.options.caption(field)}`);
          const grip = document.createElement("span");
          grip.className = "group-grip";
          grip.textContent = "⠿";
          grip.setAttribute("aria-hidden", "true");
          const name = document.createElement("span");
          name.className = "group-chip-name";
          name.textContent = this.options.caption(field);
          const remove = document.createElement("button");
          remove.type = "button";
          remove.className = "group-remove";
          remove.textContent = "×";
          remove.disabled = !ready;
          remove.setAttribute("aria-label", `Remove grouping by ${this.options.caption(field)}`);
          chip.append(grip, name, remove);
          step.append(chip);
          return step;
        });
        this.chips.replaceChildren(...steps);
      }
      this.clear.hidden = !fields.length;
      this.clear.disabled = !ready;
    }

    /** Refresh drag eligibility and the chip after data/columns/layout changes. */
    update() {
      const ready = this.options.isReady();
      this.zone.setAttribute("aria-disabled", String(!ready));
      for (const column of this.options.table.getColumns()) column.getElement().draggable = ready && column.isVisible();
      for (const column of this.options.table.getColumns(true)) {
        const children = column.getSubColumns();
        if (!children.length) continue;
        const header = column.getElement();
        header.dataset.headerGroup = children[0].getField();
        header.draggable = ready && column.isVisible();
        header.title = column.getDefinition().title + " — drag to move this column group";
      }
      if (!ready && this.session) this.cancel();
      else this.render();
    }

    /** Release document/window listeners when the grid is destroyed. */
    destroy() {
      this.cancel();
      for (const { target, type, handler, capture } of this.listeners) target.removeEventListener(type, handler, capture);
      this.listeners = [];
    }
  }
  window.ColumnGroupingDrag = ColumnGroupingDrag;
})();
