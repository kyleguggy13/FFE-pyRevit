/* Named Space membership is saved explicitly as model metadata through Revit. */
(function () {
  "use strict";
  const byId = id => document.getElementById(id);

  class SpaceGroupControls {
    constructor(options) {
      this.options = options;
      this.current = null;
      this.members = new Set();
      this.pending = null;
      this.serial = 0;
      this.confirmDelete = false;
      byId("space-groups").addEventListener("click", () => this.open());
      byId("space-group-choice").addEventListener("change", () => this.load(byId("space-group-choice").value));
      byId("space-group-search").addEventListener("input", () => this.renderMembers());
      byId("space-group-name").addEventListener("input", () => { this.confirmDelete = false; this.update(); });
      byId("space-group-use-selection").addEventListener("click", () => {
        this.members = new Set(this.options.getSelectedUniqueIds());
        this.confirmDelete = false;
        this.renderMembers();
      });
      for (const [id, checked] of [["space-group-select-listed", true], ["space-group-clear-listed", false]]) {
        byId(id).addEventListener("click", () => {
          for (const row of this.filteredRows()) {
            if (checked) this.members.add(row.uniqueId);
            else this.members.delete(row.uniqueId);
          }
          this.confirmDelete = false;
          this.renderMembers();
        });
      }
      byId("space-group-cancel").addEventListener("click", () => byId("space-group-dialog").close());
      byId("space-group-dialog").addEventListener("cancel", event => { if (this.pending) event.preventDefault(); });
      byId("space-group-form").addEventListener("submit", event => {
        event.preventDefault();
        if (byId("space-group-save").disabled) return;
        this.send("saveSpaceGroup", { name: byId("space-group-name").value.trim(), members: [...this.members] });
      });
      byId("space-group-delete").addEventListener("click", () => {
        if (!this.current || byId("space-group-delete").disabled) return;
        if (!this.confirmDelete) {
          this.confirmDelete = true;
          byId("space-group-error").textContent = "Delete this group? Its Spaces will be kept. Click Confirm delete to continue.";
          this.update();
        } else this.send("deleteSpaceGroup");
      });
    }

    ready() {
      return Boolean(this.options.bridge && this.options.getSnapshot()) && this.options.isActive() &&
        !this.options.isLoading() && !this.options.isBusy() && !this.pending;
    }

    update() {
      const ready = this.ready();
      const canWrite = ready && this.options.getSnapshot().canEditSpaceGroups === true;
      byId("space-groups").disabled = !ready;
      byId("space-group-choice").disabled = Boolean(this.pending);
      byId("space-group-cancel").disabled = Boolean(this.pending);
      byId("space-group-name").disabled = !canWrite;
      byId("space-group-save").disabled = !canWrite || !this.members.size || !byId("space-group-name").value.trim();
      byId("space-group-delete").disabled = !canWrite || !this.current;
      byId("space-group-delete").textContent = this.confirmDelete ? "Confirm delete" : "Delete group";
      byId("space-group-use-selection").disabled = !canWrite || !this.options.getSelectedUniqueIds().length;
      for (const id of ["space-group-select-listed", "space-group-clear-listed"]) byId(id).disabled = !canWrite;
      byId("space-group-members").querySelectorAll("input").forEach(input => { input.disabled = !canWrite; });
      byId("space-group-count").textContent = this.members.size + " Space(s) in this group. " +
        (this.options.getSnapshot()?.canEditSpaceGroups ? "" : "Group editing is unavailable in this model.");
    }

    renderChoices() {
      const groups = this.options.getSnapshot().spaceGroups || [];
      const names = new Map();
      for (const group of groups) names.set(group.name, (names.get(group.name) || 0) + 1);
      byId("space-group-choice").replaceChildren(...[{ id: "", name: "New group" }, ...groups].map(group => {
        const option = document.createElement("option");
        option.value = group.id;
        option.textContent = group.name + (names.get(group.name) > 1 ? " (" + group.id.slice(0, 8) + ")" : "");
        return option;
      }));
    }

    load(id, seedSelection = true) {
      const group = (this.options.getSnapshot().spaceGroups || []).find(item => item.id === id);
      this.current = group ? JSON.parse(JSON.stringify(group)) : null;
      this.members = new Set(group ? group.members : seedSelection ? this.options.getSelectedUniqueIds() : []);
      this.confirmDelete = false;
      byId("space-group-choice").value = group?.id || "";
      byId("space-group-name").value = group?.name || "";
      byId("space-group-search").value = "";
      byId("space-group-error").textContent = (this.options.getSnapshot().groupIssues || []).join("\n");
      this.renderMembers();
    }

    filteredRows() {
      const rows = [...(this.options.getSnapshot()?.rows || [])];
      const known = new Set(rows.map(row => row.uniqueId));
      for (const member of this.members) {
        if (!known.has(member)) rows.push({ uniqueId: member, number: "", name: member + " (unavailable)", level: "", phase: "" });
      }
      const query = byId("space-group-search").value.trim().toLowerCase();
      return rows.filter(row => [row.number, row.name, row.level, row.phase].join(" ").toLowerCase().includes(query));
    }

    renderMembers() {
      byId("space-group-members").replaceChildren(...this.filteredRows().map(row => {
        const label = document.createElement("label");
        label.className = "space-check";
        const input = document.createElement("input");
        input.type = "checkbox";
        input.value = row.uniqueId;
        input.checked = this.members.has(row.uniqueId);
        input.addEventListener("change", () => {
          if (input.checked) this.members.add(row.uniqueId);
          else this.members.delete(row.uniqueId);
          this.confirmDelete = false;
          this.update();
        });
        const text = document.createElement("span");
        text.textContent = [row.number, row.name, row.level, row.phase].filter(Boolean).join(" · ");
        label.append(input, text);
        return label;
      }));
      this.update();
    }

    open() {
      if (!this.ready()) return;
      this.renderChoices();
      this.load("");
      byId("space-group-dialog").showModal();
      byId("space-group-name").focus();
    }

    send(action, payload = {}) {
      if (!this.ready()) return;
      this.pending = { action, requestId: "space-group-" + (++this.serial) };
      byId("space-group-error").textContent = "";
      this.options.setBusy(true);
      this.update();
      try {
        this.options.bridge.postMessage({ ...payload, ...this.pending, groupId: this.current?.id || "",
          revision: this.current?.revision || "", documentKey: this.options.getSnapshot().documentKey });
      } catch (error) {
        this.handleResult({ requestId: this.pending.requestId, success: false, message: "Could not contact Revit: " + error.message });
      }
    }

    handleResult(message) {
      if (!this.pending || message.requestId !== this.pending.requestId) return;
      this.pending = null;
      this.options.setBusy(false);
      if (!message.success) {
        byId("space-group-error").textContent = message.message || "Could not save the Space group.";
        this.update();
        return;
      }
      Object.assign(this.options.getSnapshot(), message.catalog);
      this.renderChoices();
      this.load(message.group?.id || "", false);
      this.options.status(message.group ? "Saved Space group. Save/Sync Revit to share it." : "Deleted Space group. Spaces were kept.", "success");
      this.options.setBusy(false);
    }
  }
  window.SpaceGroupControls = SpaceGroupControls;
})();
