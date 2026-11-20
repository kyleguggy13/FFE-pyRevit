/* Named column bands are local layout settings until the table is explicitly saved. */
(function () {
  "use strict";
  const byId = id => document.getElementById(id);

  /** Edit a draft of column memberships without changing the live grid on Cancel. */
  class HeaderGroupControls {
    constructor(options) {
      this.options = options;
      this.draft = [];
      byId("column-groups").addEventListener("click", () => this.open());
      byId("header-group-add").addEventListener("click", () => {
        this.draft.push({ name: "", fields: [] });
        this.render();
        byId("header-group-list").lastElementChild.querySelector("input[type=text]").focus();
      });
      byId("header-groups-cancel").addEventListener("click", () => byId("header-groups-dialog").close());
      byId("header-group-list").addEventListener("input", event => {
        const index = Number(event.target.closest("[data-index]")?.dataset.index);
        if (event.target.matches("input[type=text]")) this.draft[index].name = event.target.value;
        byId("header-groups-error").textContent = "";
      });
      byId("header-group-list").addEventListener("change", event => {
        if (!event.target.matches("input[type=checkbox]")) return;
        const index = Number(event.target.closest("[data-index]").dataset.index);
        const field = event.target.dataset.field;
        this.draft[index].fields = this.draft[index].fields.filter(item => item !== field);
        if (event.target.checked) this.draft[index].fields.push(field);
        this.render();
        Array.from(byId("header-group-list").children[index].querySelectorAll("input[type=checkbox]"))
          .find(input => input.dataset.field === field)?.focus();
      });
      byId("header-group-list").addEventListener("click", event => {
        if (!event.target.closest(".header-group-remove")) return;
        this.draft.splice(Number(event.target.closest("[data-index]").dataset.index), 1);
        this.render();
      });
      byId("header-groups-form").addEventListener("submit", event => {
        event.preventDefault();
        if (!this.options.isReady()) return;
        if (this.draft.some(group => !group.name.trim() || group.name.trim().length > 100 || !group.fields.length)) {
          byId("header-groups-error").textContent = "Give every group a name and select at least one column.";
          return;
        }
        this.options.setGroups(this.draft.map(group => ({ ...group, name: group.name.trim(), fields: [...group.fields] })));
        byId("header-groups-dialog").close();
      });
    }

    /** Start each editing session from the currently applied header bands. */
    open() {
      if (!this.options.isReady()) return;
      this.draft = this.options.getGroups().map(group => ({ ...group, fields: [...group.fields] }));
      byId("header-groups-error").textContent = "";
      this.render();
      byId("header-groups-dialog").showModal();
    }

    /** Build names and checkboxes with DOM text so custom names remain literal text. */
    render() {
      const columns = this.options.getColumns();
      const editors = this.draft.map((group, index) => {
        const editor = document.createElement("fieldset");
        editor.dataset.index = index;
        const legend = document.createElement("legend");
        legend.textContent = `Header group ${index + 1}`;
        const heading = document.createElement("div");
        heading.className = "header-group-heading";
        const label = document.createElement("label");
        label.textContent = "Group name";
        const input = document.createElement("input");
        input.type = "text";
        input.required = true;
        input.maxLength = 100;
        input.placeholder = "e.g. Dimensions";
        input.value = group.name;
        label.append(input);
        const remove = document.createElement("button");
        remove.type = "button";
        remove.className = "header-group-remove";
        remove.textContent = "Remove group";
        heading.append(label, remove);
        const choices = document.createElement("div");
        choices.className = "header-group-columns";
        for (const column of columns) {
          const choice = document.createElement("label");
          const checkbox = document.createElement("input");
          checkbox.type = "checkbox";
          checkbox.dataset.field = column.field;
          checkbox.checked = group.fields.includes(column.field);
          checkbox.disabled = this.draft.some((other, otherIndex) => otherIndex !== index && other.fields.includes(column.field));
          const text = document.createElement("span");
          text.textContent = `${column.title}${column.visible ? "" : " (hidden)"}`;
          choice.append(checkbox, text);
          if (checkbox.disabled) choice.title = "This column belongs to another header group.";
          choices.append(choice);
        }
        editor.append(legend, heading, choices);
        return editor;
      });
      byId("header-group-list").replaceChildren(...editors);
      byId("header-groups-empty").hidden = Boolean(editors.length);
      byId("header-group-add").disabled = this.draft.length >= columns.length ||
        columns.every(column => this.draft.some(group => group.fields.includes(column.field)));
    }
  }
  window.HeaderGroupControls = HeaderGroupControls;
})();
