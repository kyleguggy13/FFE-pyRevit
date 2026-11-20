/* Explicit cloud publishing uses the host bridge; the page retains its offline CSP. */
(function () {
  "use strict";
  const byId = id => document.getElementById(id);

  /** Review a destination and correlate results without changing saved-table state. */
  class CloudPublishControls {
    constructor(options) {
      this.options = options;
      this.pending = null;
      this.catalog = null;
      this.serial = 0;
      byId("publish").addEventListener("click", () => this.open());
      byId("publish-reload").addEventListener("click", () => this.reload());
      byId("publish-cancel").addEventListener("click", () => byId("publish-dialog").close());
      byId("publish-dialog").addEventListener("cancel", event => { if (this.pending) event.preventDefault(); });
      byId("publish-project").addEventListener("change", () => this.update());
      byId("publish-form").addEventListener("submit", event => {
        event.preventDefault();
        if (byId("publish-confirm").disabled) return;
        this.send("publishSnapshot", { projectId: byId("publish-project").value });
      });
    }

    ready() {
      return Boolean(this.options.bridge && this.options.getSnapshot()) && this.options.isActive() &&
        !this.options.isLoading() && !this.options.isBusy();
    }

    update() {
      const ready = this.ready();
      byId("publish").disabled = !ready;
      byId("publish-cancel").disabled = Boolean(this.pending);
      byId("publish-reload").disabled = !ready;
      byId("publish-project").disabled = !ready || Boolean(this.catalog?.model);
      byId("publish-confirm").disabled = !ready || !this.catalog || !byId("publish-project").value;
    }

    open() {
      if (!this.ready()) return;
      const snapshot = this.options.getSnapshot();
      byId("publish-scope").textContent = `Publish all ${snapshot.rows.length} Spaces from ${snapshot.documentTitle}, across all phases and levels, ` +
        `with the ${snapshot.selectedParameterIds?.length || 0} loaded parameter column(s). ` +
        "Current data is collected when you publish. Table filters and row selections do not limit the upload.";
      byId("publish-dialog").showModal();
      this.reload();
    }

    reload() {
      if (!this.ready()) return;
      this.catalog = null;
      byId("publish-project").replaceChildren();
      byId("publish-destination").textContent = "Loading cloud projects…";
      this.send("publishCatalog");
    }

    send(action, payload = {}) {
      if (!this.ready()) return;
      this.pending = { requestId: `publish-${++this.serial}`, action };
      byId("publish-error").textContent = "";
      this.options.setBusy(true);
      this.options.status(action === "publishCatalog" ? "Loading cloud projects…" : "Publishing Space data…");
      try {
        this.options.bridge.postMessage({ ...payload, ...this.pending,
          documentKey: this.options.getSnapshot().documentKey });
      } catch (error) {
        this.handleResult({ requestId: this.pending.requestId, success: false, message: `Could not contact Revit: ${error.message}` });
      }
    }

    handleResult(message) {
      if (!this.pending || this.pending.requestId !== message.requestId) return;
      const action = this.pending.action;
      this.pending = null;
      if (!message.success) {
        this.catalog = null;
        byId("publish-error").textContent = `${message.message || "Publishing failed."} Reload projects before retrying.`;
        byId("publish-destination").textContent = "The publish request did not complete. Your working table is retained.";
        this.options.status(message.message || "Publishing failed.", "error");
      } else if (action === "publishCatalog") {
        this.catalog = message.data;
        const choices = this.catalog.projects.map(project => {
          const option = document.createElement("option");
          option.value = project.id;
          option.textContent = project.name;
          return option;
        });
        byId("publish-project").replaceChildren(...choices);
        if (this.catalog.model) byId("publish-project").value = this.catalog.model.projectId;
        byId("publish-destination").textContent = this.catalog.model
          ? `Replace the published Space dataset for “${this.catalog.model.name}”. Saved cloud layouts are retained.`
          : "Create a cloud model for this Revit model in the selected project. Save/Sync Revit afterward to share its cloud link.";
        if (!choices.length) byId("publish-error").textContent = "No cloud projects are available. Set up a project in FFE Mechanical, then reload.";
        this.options.status(choices.length ? "Choose a destination and click Publish." : "No cloud projects are available.");
      } else {
        byId("publish-dialog").close();
        this.catalog = null;
        this.options.status(`Published ${message.data.rowCount} Space(s) and ${message.data.parameterCount} parameter column(s) to FFE Mechanical.`, "success");
      }
      this.options.setBusy(false);
      this.update();
    }
  }
  window.CloudPublishControls = CloudPublishControls;
})();
