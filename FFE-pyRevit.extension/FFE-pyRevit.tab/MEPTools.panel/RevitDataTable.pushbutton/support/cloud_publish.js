/* Loadwise v3 sync uses the host bridge; preview stays local until confirmed. */
(function () {
  "use strict";
  const byId = id => document.getElementById(id);
  const pageSize = 50;
  const actionStatus = {
    publishCatalog: "Loading Loadwise projects…",
    publishReview: "Collecting current Spaces…",
    publishCommit: "Syncing Spaces to Loadwise…",
  };

  class CloudPublishControls {
    constructor(options) {
      this.options = options;
      this.pending = null;
      this.catalog = null;
      this.serial = 0;
      this.resetReview();
      byId("publish").addEventListener("click", () => this.open());
      byId("publish-reload").addEventListener("click", () => this.reload());
      byId("publish-cancel").addEventListener("click", () => byId("publish-dialog").close());
      byId("publish-dialog").addEventListener("cancel", event => { if (this.pending) event.preventDefault(); });
      byId("publish-project").addEventListener("change", () => this.update());
      byId("publish-previous").addEventListener("click", () => { this.page--; this.renderPage(); });
      byId("publish-next").addEventListener("click", () => { this.page++; this.renderPage(); });
      byId("publish-form").addEventListener("submit", event => {
        event.preventDefault();
        if (byId("publish-confirm").disabled) return;
        this.send(this.review ? "publishCommit" : "publishReview", this.review
          ? { reviewId: this.review.reviewId } : { projectId: byId("publish-project").value });
      });
    }

    ready() {
      return Boolean(this.options.bridge && this.options.getSnapshot()) && this.options.isActive() &&
        !this.options.isLoading() && !this.options.isBusy() && !this.pending;
    }

    resetReview() {
      this.review = null;
      this.retry = false;
      this.page = 0;
      byId("publish-review").hidden = true;
      byId("publish-changes").replaceChildren();
      byId("publish-warnings").textContent = "";
    }

    update() {
      const ready = this.ready();
      byId("publish").disabled = !ready;
      byId("publish-cancel").disabled = Boolean(this.pending);
      byId("publish-reload").disabled = !ready;
      byId("publish-project").disabled = !ready || Boolean(this.catalog?.projectId || this.review);
      byId("publish-previous").disabled = !ready || this.page === 0;
      byId("publish-next").disabled = !ready || !this.review ||
        (this.page + 1) * pageSize >= this.review.request.spaces.length;
      byId("publish-confirm").textContent = this.retry ? "Retry sync" : this.review ? "Sync Spaces" : "Collect Spaces";
      byId("publish-confirm").disabled = !ready || !this.catalog || !byId("publish-project").value;
    }

    renderPage() {
      const spaces = this.review.request.spaces;
      const rows = spaces.slice(this.page * pageSize, (this.page + 1) * pageSize).map(space => {
        const row = document.createElement("tr");
        for (const key of ["number", "name", "level", "area", "ceilingHeight"]) {
          row.insertCell().textContent = space[key] ?? "Unavailable";
        }
        return row;
      });
      byId("publish-changes").replaceChildren(...rows);
      byId("publish-page").textContent = spaces.length
        ? "Page " + (this.page + 1) + " of " + Math.ceil(spaces.length / pageSize) : "";
      byId("publish-summary").textContent = spaces.length + " Space(s) ready to sync. " +
        this.review.skippedCount + " Space(s) skipped. Missing rooms are reported and retained.";
      byId("publish-warnings").textContent = this.review.warnings.join("\n");
      this.update();
    }

    open() {
      if (!this.ready()) return;
      byId("publish-scope").textContent = "Collect all host-model Spaces across every phase and level. " +
        "Sync updates room number, name, level, area and height, including unplaced and zero-area Spaces. " +
        "Table filters and selections do not limit sync. Review the collected values before confirming.";
      byId("publish-dialog").showModal();
      this.reload();
    }

    reload() {
      if (!this.ready()) return;
      this.catalog = null;
      this.resetReview();
      byId("publish-project").replaceChildren();
      byId("publish-destination").textContent = actionStatus.publishCatalog;
      this.send("publishCatalog");
    }

    send(action, payload = {}) {
      if (!this.ready()) return;
      this.pending = { requestId: "publish-" + (++this.serial), action };
      byId("publish-error").textContent = "";
      this.options.setBusy(true);
      this.options.status(actionStatus[action]);
      this.update();
      try {
        this.options.bridge.postMessage({ ...payload, ...this.pending, documentKey: this.options.getSnapshot().documentKey });
      } catch (error) {
        this.handleResult({ requestId: this.pending.requestId, success: false, message: "Could not contact Revit: " + error.message });
      }
    }

    handleResult(message) {
      if (!this.pending || this.pending.requestId !== message.requestId) return;
      const action = this.pending.action;
      this.pending = null;
      if (!message.success) {
        const error = message.message || "Space sync failed.";
        this.retry = Boolean(message.retry);
        if (!this.retry) { this.catalog = null; this.resetReview(); }
        byId("publish-error").textContent = error + (this.retry
          ? " The result is unconfirmed. Retry sync to resend the same collected values."
          : " Reload projects and collect Spaces again.");
        this.options.status(error, "error");
      } else if (action === "publishCatalog") {
        this.catalog = message.data;
        const choices = this.catalog.projects.map(project => {
          const option = document.createElement("option");
          option.value = project.id;
          option.textContent = project.name;
          return option;
        });
        byId("publish-project").replaceChildren(...choices);
        if (this.catalog.projectId) byId("publish-project").value = this.catalog.projectId;
        byId("publish-destination").textContent = this.catalog.projectId
          ? "Connected to this model's Loadwise project."
          : "Choose a project. The first sync saves this model's project connection; Save/Sync Revit to share it.";
        if (!choices.length) byId("publish-error").textContent = "No projects are available. Create one in Loadwise, then reload.";
        this.options.status(choices.length ? "Choose a project and collect Spaces." : "No Loadwise projects are available.");
      } else if (action === "publishReview") {
        this.review = message.data;
        byId("publish-review").hidden = false;
        byId("publish-destination").textContent = "Collected locally. Confirm Sync Spaces to send these values to Loadwise.";
        this.renderPage();
        this.options.status("Review the collected Spaces, then confirm sync.");
      } else {
        const result = message.data;
        const missing = result.notInRevit.length;
        const report = "Synced to Loadwise: " + result.created + " created, " + result.updated + " updated, " +
          result.unchanged + " unchanged, " + result.linked + " linked by number. " + missing +
          " room(s) absent from this collection and retained. " + this.review.skippedCount + " Space(s) skipped.";
        byId("publish-dialog").close();
        const kind = missing || this.review.skippedCount ? "warning" : "success";
        this.catalog = null;
        this.resetReview();
        this.options.status(report, kind);
      }
      this.options.setBusy(false);
      this.update();
    }
  }
  window.CloudPublishControls = CloudPublishControls;
})();
