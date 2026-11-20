/* Shared definitions are saved explicitly; this controller owns only UI working state. */
(function () {
  "use strict";
  const element = id => document.getElementById(id);
  const clone = value => JSON.parse(JSON.stringify(value));

  class SavedTableControls {
    constructor(options) {
      this.options = options;
      this.current = null;
      this.baseline = null;
      this.starter = null;
      this.pending = null;
      this.serial = 0;
      this.warnings = [];
      this.needsRepair = false;
      this.dialogResolve = null;
      element("saved-table").addEventListener("change", () => {
        const id = element("saved-table").value;
        element("saved-table").value = this.current?.id || "";
        if (id !== (this.current?.id || "")) this.guard(() => this.open(id));
      });
      element("table-save").addEventListener("click", () => this.save());
      element("table-reopen").addEventListener("click", () => {
        if (this.current) this.guard(() => this.open(this.current.id));
      });
      element("table-save-as").addEventListener("click", () => this.save(true));
      element("table-rename").addEventListener("click", () => this.rename());
      element("table-delete").addEventListener("click", () => this.remove());
      element("table-dialog-form").addEventListener("submit", event => {
        event.preventDefault();
        const input = element("table-name");
        if (!input.disabled && (!input.value.trim() || input.value.trim().length > 100)) {
          input.setCustomValidity("Enter a table name of 1 to 100 characters.");
          input.reportValidity();
          return;
        }
        this.resolveDialog({ choice: "confirm", name: input.value.trim() });
      });
      element("table-name").addEventListener("input", () => element("table-name").setCustomValidity(""));
      element("table-dialog-cancel").addEventListener("click", () => this.resolveDialog(null));
      element("table-dialog-discard").addEventListener("click", () => this.resolveDialog({ choice: "discard" }));
      element("table-dialog").addEventListener("close", () => {
        if (!element("table-dialog").open) this.resolveDialog(null);
      });
    }

    dirty() {
      return this.needsRepair || (this.baseline !== null && JSON.stringify(this.options.getLayout()) !== this.baseline);
    }

    onSnapshot() {
      if (this.baseline === null) {
        this.starter = clone(this.options.getLayout());
        this.baseline = JSON.stringify(this.starter);
      }
    }

    update() {
      const snapshot = this.options.getSnapshot();
      const loading = this.options.isLoading();
      const ready = Boolean(snapshot) && !loading;
      const canRequest = ready && this.options.isActive() && !this.options.isBusy() && Boolean(this.options.bridge);
      const canWrite = canRequest && snapshot.canSaveTables === true;
      element("saved-table").disabled = !canRequest;
      element("table-reopen").disabled = !canRequest || !this.current;
      element("table-save").disabled = !canWrite || (Boolean(this.current) && !this.dirty());
      element("table-save-as").disabled = !canWrite;
      element("table-rename").disabled = !canWrite || !this.current;
      element("table-delete").disabled = !canWrite || !this.current;
      if (!ready) return;
      element("table-dirty").textContent = this.dirty() ? "Unsaved changes" : "";
      const catalog = snapshot.savedTables || [];
      const names = new Map();
      for (const table of catalog) names.set(table.name, (names.get(table.name) || 0) + 1);
      const choices = [{ id: "", name: "Unsaved table" }, ...catalog];
      if (this.current && !catalog.some(table => table.id === this.current.id)) {
        choices.push({ ...this.current, name: `${this.current.name} (unavailable)` });
      }
      element("saved-table").replaceChildren(...choices.map(table => {
        const option = document.createElement("option");
        option.value = table.id;
        option.textContent = table.id && (names.get(table.name) > 1 || table.name === "Unsaved table")
          ? `${table.name} — ${table.modifiedBy || "Team"} (${table.id.slice(0, 8)})` : table.name;
        option.title = table.modifiedAt ? `Modified by ${table.modifiedBy || "unknown"} · ${table.modifiedAt}` : "";
        return option;
      }));
      element("saved-table").value = this.current?.id || "";
      const warnings = [...(snapshot.tableIssues || []), ...this.warnings];
      const latest = this.current && catalog.find(table => table.id === this.current.id);
      if (latest && latest.revision !== this.current.revision) {
        warnings.push("The shared definition changed. Reopen it to load current settings, or use Save As to retain your working view.");
      }
      element("table-note").textContent = warnings.join("\n");
      element("table-note").hidden = !warnings.length;
    }

    ask({ title, message, name = null, confirm = "Save", discard = false, writable = true }) {
      if (element("table-dialog").open) return Promise.resolve(null);
      element("table-dialog-title").textContent = title;
      element("table-dialog-message").textContent = message;
      element("table-name-label").hidden = name === null;
      element("table-name").required = name !== null;
      element("table-name").disabled = name === null;
      element("table-name").value = name || "";
      element("table-name").setCustomValidity("");
      element("table-dialog-confirm").textContent = confirm;
      element("table-dialog-confirm").disabled = !writable;
      element("table-dialog-discard").hidden = !discard;
      const answer = new Promise(resolve => { this.dialogResolve = resolve; });
      element("table-dialog").showModal();
      if (name !== null) {
        element("table-name").focus();
        element("table-name").select();
      } else element("table-dialog-cancel").focus();
      return answer;
    }

    resolveDialog(answer) {
      const resolve = this.dialogResolve;
      if (!resolve) return;
      this.dialogResolve = null;
      element("table-dialog").close();
      resolve(answer);
    }

    async guard(continuation) {
      if (this.options.isBusy()) {
        this.options.status("Wait for the current Revit request before closing or switching tables.", "warning");
        return;
      }
      if (!this.dirty()) return continuation();
      const answer = await this.ask({ title: "Unsaved table changes", discard: true,
        message: "Save your working view before continuing? Saving updates the shared definition in this model.",
        writable: this.options.isActive() && this.options.getSnapshot().canSaveTables === true });
      if (answer?.choice === "confirm") await this.save(false, continuation);
      else if (answer?.choice === "discard") continuation();
    }

    send(action, payload, continuation = null) {
      if (this.options.isBusy() || !this.options.isActive() || !this.options.bridge) return;
      const requestId = `table-${++this.serial}`;
      this.pending = { action, requestId, layout: payload.layout && clone(payload.layout), continuation };
      this.options.setBusy(true);
      this.options.status("Waiting for Revit…");
      try {
        this.options.bridge.postMessage({ ...payload, action, requestId,
          documentKey: this.options.getSnapshot().documentKey });
      } catch (error) {
        this.pending = null;
        this.options.setBusy(false);
        this.options.status(`Could not contact Revit: ${error.message}`, "error");
      }
    }

    open(id) {
      this.send("loadTable", { tableId: id });
    }

    async save(asNew = false, continuation = null) {
      let name = this.current?.name;
      if (!this.current || asNew) {
        const answer = await this.ask({ title: "Save table as", name: name ? `${name} copy` : "Spaces",
          message: "This table definition is stored in the model. Save/Sync Revit to persist and share it." });
        if (!answer || !answer.name) return;
        name = answer.name;
      }
      this.send("saveTable", { tableId: asNew ? "" : this.current?.id || "",
        revision: asNew ? "" : this.current?.revision || "", name, layout: this.options.getLayout() }, continuation);
    }

    async rename() {
      if (!this.current) return;
      const answer = await this.ask({ title: "Rename shared table", name: this.current.name,
        message: "Rename this definition for the project team. Unsaved layout changes remain in your working view.", confirm: "Rename" });
      if (answer?.name) this.send("renameTable", { tableId: this.current.id, revision: this.current.revision, name: answer.name });
    }

    async remove() {
      if (!this.current) return;
      const answer = await this.ask({ title: "Delete shared table", confirm: "Delete",
        message: `Delete “${this.current.name}” from this model? Its definition is removed for the project team. Your current table remains open; Spaces are unchanged.` });
      if (answer) this.send("deleteTable", { tableId: this.current.id, revision: this.current.revision });
    }

    defaultLayout(snapshot) {
      const layout = clone(this.starter);
      layout.filters = { phase: snapshot.defaultPhaseId, level: "all", query: "" };
      return layout;
    }

    async handleResult(message) {
      if (!this.pending || message.requestId !== this.pending.requestId) return;
      const pending = this.pending;
      this.pending = null;
      try {
        if (!message.success) {
          this.options.status(message.message || "The saved-table operation failed. Your working view has been retained.", "error");
          return;
        }
        if (message.action === "loadTable") {
          await this.options.loadSnapshot(message.snapshot, message.layout || this.defaultLayout(message.snapshot));
          this.current = message.table;
          this.baseline = JSON.stringify(this.options.getLayout());
          this.warnings = message.warnings || [];
          this.needsRepair = Boolean(this.current && this.warnings.length);
          this.options.status(this.warnings.length ? "Loaded the available settings. Review the saved-table warnings."
            : `Opened ${this.current?.name || "an unsaved table"}.`, this.warnings.length ? "warning" : "success");
        } else {
          if (message.catalog) Object.assign(this.options.getSnapshot(), message.catalog);
          this.current = message.table;
          if (message.action === "saveTable") {
            this.baseline = JSON.stringify(message.layout || pending.layout);
            this.warnings = [];
            this.needsRepair = false;
          } else if (message.action === "deleteTable") {
            this.baseline = JSON.stringify(this.starter);
            this.warnings = [];
            this.needsRepair = false;
          }
          this.options.status("Table definition updated in the model. Save/Sync Revit to persist and share this change.", "success");
        }
      } catch (error) {
        this.options.status(`Could not restore the saved table: ${error.message}`, "error");
        return;
      } finally {
        this.options.setBusy(false);
        this.update();
      }
      if (pending.continuation) pending.continuation();
    }

    close() {
      const colorDialog = element("header-color-dialog");
      if (colorDialog?.open) { colorDialog.returnValue = "cancel"; colorDialog.close(); }
      const publishDialog = element("publish-dialog");
      if (publishDialog?.open && !this.options.isBusy()) publishDialog.close();
      if (element("parameter-dialog").open) element("parameter-dialog").close();
      if (element("header-groups-dialog").open) element("header-groups-dialog").close();
      this.guard(() => this.options.bridge?.postMessage({ action: "closeWindow",
        documentKey: this.options.getSnapshot().documentKey }));
    }
  }

  window.SavedTableControls = SavedTableControls;
})();
