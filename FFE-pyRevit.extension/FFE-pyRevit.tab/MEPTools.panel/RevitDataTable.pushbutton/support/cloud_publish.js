/* Repository publishing uses the host bridge and the packaged page's offline policy. */
(function () {
  "use strict";
  const byId = id => document.getElementById(id);
  const pageSize = 50;
  const actionStatus = {
    publishCatalog: "Loading cloud projects…",
    publishReview: "Collecting and comparing Space values…",
    publishCommit: "Committing selected Space changes…",
  };

  function stateText(state, displayText) {
    if (state.state === "missing") return "Missing";
    if (state.state === "unset") return "Unset";
    const value = state.value;
    if (value && typeof value === "object") {
      return value.uniqueId || value.sentinel || (value.elementId ? "Space #" + value.elementId : JSON.stringify(value));
    }
    if (value === "") return "(empty text)";
    const raw = String(value);
    return displayText && displayText !== raw ? displayText + " (" + raw + ")" : raw;
  }

  /** Review repository differences and stage only the changes chosen in this dialog. */
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
      byId("publish-message").addEventListener("input", () => this.update());
      byId("publish-previous").addEventListener("click", () => { this.page--; this.renderPage(); });
      byId("publish-next").addEventListener("click", () => { this.page++; this.renderPage(); });
      byId("publish-select-local").addEventListener("click", () => {
        for (const [key, item] of this.candidates) {
          if (item.classification === "local_change" && !(item.field === "@space" && item.proposed.state === "missing")) {
            this.select(key, true);
          }
        }
        this.update();
      });
      byId("publish-clear").addEventListener("click", () => { this.selection.clear(); this.update(); });
      byId("publish-form").addEventListener("submit", event => {
        event.preventDefault();
        if (byId("publish-confirm").disabled) return;
        if (this.review) {
          this.send("publishCommit", { reviewId: this.review.collectionId, selection: [...this.selection],
            message: byId("publish-message").value.trim() });
        } else {
          this.send("publishReview", { projectId: byId("publish-project").value });
        }
      });
    }

    ready() {
      return Boolean(this.options.bridge && this.options.getSnapshot()) && this.options.isActive() &&
        !this.options.isLoading() && !this.options.isBusy();
    }

    resetReview() {
      this.review = null;
      this.retry = false;
      this.candidates = [];
      this.selection = new Set();
      this.additions = new Map();
      this.page = 0;
      byId("publish-review").hidden = true;
      byId("publish-message").value = "";
      byId("publish-changes").replaceChildren();
    }

    update() {
      const ready = this.ready();
      const canEdit = ready && !this.retry;
      byId("publish").disabled = !ready;
      byId("publish-cancel").disabled = Boolean(this.pending);
      byId("publish-reload").disabled = !ready;
      byId("publish-project").disabled = !ready || Boolean(this.catalog?.model || this.catalog?.projectId || this.review);
      byId("publish-message").disabled = !canEdit;
      byId("publish-select-local").disabled = !canEdit || !this.candidates.length;
      byId("publish-clear").disabled = !canEdit || !this.selection.size;
      byId("publish-previous").disabled = Boolean(this.pending) || this.page === 0;
      byId("publish-next").disabled = Boolean(this.pending) || (this.page + 1) * pageSize >= this.candidates.length;
      byId("publish-confirm").textContent = this.retry ? "Retry commit" : this.review ? "Commit selected" : "Collect changes";
      byId("publish-confirm").disabled = !ready || (this.review
        ? !this.selection.size || !byId("publish-message").value.trim()
        : !this.catalog || !byId("publish-project").value);
      if (this.review) {
        const { incomingCount, skippedCount, warningCount } = this.review;
        let summary = this.candidates.length + " proposal(s). " + this.selection.size + " selected.";
        if (incomingCount) summary += " " + incomingCount + " incoming value(s) retained.";
        if (skippedCount) summary += " " + skippedCount + " Space(s) unavailable. Deletions are excluded.";
        if (warningCount) summary += " " + warningCount + " collection warning(s).";
        byId("publish-summary").textContent = summary;
        byId("publish-changes").querySelectorAll("input").forEach(input => {
          input.checked = this.selection.has(input.value);
          input.disabled = !canEdit;
        });
      }
    }

    select(key, checked) {
      const item = this.review.candidates[key];
      if (checked) {
        this.selection.add(key);
        if (this.additions.has(item.spaceId)) this.selection.add(this.additions.get(item.spaceId));
      } else {
        this.selection.delete(key);
        if (this.additions.get(item.spaceId) === key) {
          for (const selected of this.selection) {
            if (this.review.candidates[selected].spaceId === item.spaceId) this.selection.delete(selected);
          }
        }
      }
    }

    renderPage() {
      const rows = this.candidates.slice(this.page * pageSize, (this.page + 1) * pageSize).map(([key, item]) => {
        const row = document.createElement("tr");
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.value = key;
        const field = item.field === "@space" ? (item.proposed.state === "missing" ? "Remove Space"
          : item.current.state === "missing" ? "Add Space" : "Space identity")
          : item.metadata.definition?.name || item.field;
        checkbox.setAttribute("aria-label", "Stage " + field + " for " + item.spaceLabel + " (" + item.field + ")");
        checkbox.addEventListener("change", () => { this.select(key, checkbox.checked); this.update(); });
        row.insertCell().append(checkbox);
        row.insertCell().textContent = item.spaceLabel;
        const fieldCell = row.insertCell();
        fieldCell.textContent = field;
        if (item.classification !== "local_change") {
          const note = document.createElement("small");
          note.textContent = item.classification === "conflict" ? "Conflict · select to keep Revit value" :
            "No sync baseline · select to keep Revit value";
          fieldCell.append(note);
        }
        row.insertCell().textContent = stateText(item.current);
        row.insertCell().textContent = stateText(item.proposed, item.metadata.displayText);
        return row;
      });
      byId("publish-changes").replaceChildren(...rows);
      byId("publish-page").textContent = this.candidates.length
        ? "Page " + (this.page + 1) + " of " + Math.ceil(this.candidates.length / pageSize) : "";
      this.update();
    }

    open() {
      if (!this.ready()) return;
      const snapshot = this.options.getSnapshot();
      byId("publish-scope").textContent = "Collect current host-model Spaces across all phases and levels, including " +
        (snapshot.selectedParameterIds?.length || 0) + " loaded parameter column(s). Select changes before committing.";
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
        const error = message.message || "Publishing failed.";
        if (message.retry) {
          this.retry = true;
          byId("publish-error").textContent = error + " " +
            (message.retryHint || "Retry commit to check the same request.");
        } else {
          this.catalog = null;
          this.resetReview();
          byId("publish-error").textContent = error + " Reload projects to start a new review.";
          byId("publish-destination").textContent = "";
        }
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
        if (this.catalog.model) byId("publish-project").value = this.catalog.model.projectId;
        else if (this.catalog.projectId) byId("publish-project").value = this.catalog.projectId;
        byId("publish-destination").textContent = this.catalog.model
          ? "Linked model: " + this.catalog.model.name
          : "Collect changes to register this model in the selected project. Save/Sync Revit to persist the link.";
        if (!choices.length) byId("publish-error").textContent = "No cloud projects are available. Set up a project in FFE Mechanical, then reload.";
        this.options.status(choices.length ? "Choose a destination and collect changes." : "No cloud projects are available.");
      } else if (action === "publishReview") {
        this.review = message.data;
        this.candidates = Object.entries(this.review.candidates);
        for (const [key, item] of this.candidates) {
          if (item.field === "@space" && item.current.state === "missing" && item.proposed.state === "value") {
            this.additions.set(item.spaceId, key);
          }
        }
        byId("publish-review").hidden = false;
        byId("publish-destination").textContent = "Collection submitted. Review Space changes before committing.";
        this.renderPage();
        this.options.status(this.candidates.length ? "Review and stage Space changes." : "No Space changes to commit.");
      } else {
        byId("publish-dialog").close();
        this.catalog = null;
        this.resetReview();
        this.options.status(message.data.noop ? "The repository already contains those values." :
          "Committed " + message.data.changeCount + " change(s) to FFE Mechanical.", "success");
      }
      this.options.setBusy(false);
      this.update();
    }
  }
  window.CloudPublishControls = CloudPublishControls;
})();
