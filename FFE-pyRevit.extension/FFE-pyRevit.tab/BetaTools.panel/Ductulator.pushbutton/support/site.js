(function initSiteNavigation(globalScope) {
  "use strict";

  function createOverflowControl() {
    var wrapper = document.createElement("div");
    wrapper.className = "tool-ribbon__select-wrap";
    wrapper.hidden = true;

    var select = document.createElement("select");
    select.className = "tool-ribbon__select";
    select.setAttribute("aria-label", "More navigation options");
    wrapper.appendChild(select);

    return { wrapper: wrapper, select: select };
  }

  function buildSelectOptions(select, links) {
    select.innerHTML = "";

    var placeholder = document.createElement("option");
    placeholder.value = "";
    placeholder.textContent = "More";
    placeholder.selected = true;
    select.appendChild(placeholder);

    links.forEach(function (link) {
      var option = document.createElement("option");
      option.value = link.getAttribute("href") || "";
      option.textContent = link.textContent.trim();
      select.appendChild(option);
    });
  }

  function setupRibbon(ribbon) {
    var links = Array.from(ribbon.querySelectorAll(":scope > .tool-ribbon__link"));
    if (!links.length) {
      return;
    }

    var overflow = createOverflowControl();
    ribbon.appendChild(overflow.wrapper);

    function render() {
      var isMobile = globalScope.matchMedia("(max-width: 720px)").matches;
      var activeLink = links.find(function (link) {
        return link.classList.contains("is-active") || link.getAttribute("aria-current") === "page";
      }) || links[0];

      ribbon.classList.toggle("tool-ribbon--mobile", isMobile);

      if (!isMobile) {
        links.forEach(function (link) {
          link.classList.remove("tool-ribbon__link--hidden");
        });
        overflow.wrapper.hidden = true;
        overflow.select.selectedIndex = 0;
        return;
      }

      var hiddenLinks = [];
      links.forEach(function (link) {
        var shouldHide = link !== activeLink;
        link.classList.toggle("tool-ribbon__link--hidden", shouldHide);
        if (shouldHide) {
          hiddenLinks.push(link);
        }
      });

      buildSelectOptions(overflow.select, hiddenLinks);
      overflow.wrapper.hidden = hiddenLinks.length === 0;
      overflow.select.selectedIndex = 0;
    }

    overflow.select.addEventListener("change", function handleNavigation() {
      if (!overflow.select.value) {
        return;
      }
      globalScope.location.assign(overflow.select.value);
    });

    if (typeof globalScope.addEventListener === "function") {
      globalScope.addEventListener("resize", render);
    }

    render();
  }

  function setupContactForm(form) {
    var formId = form.getAttribute("data-formspree-form-id");
    if (!formId || !form.id) {
      return;
    }

    if (typeof globalScope.formspree === "function") {
      globalScope.formspree("initForm", {
        formElement: "#" + form.id,
        formId: formId
      });
    }
  }

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", function onReady() {
      document.querySelectorAll(".tool-ribbon").forEach(setupRibbon);
      document.querySelectorAll(".contact-form[data-formspree-form-id]").forEach(setupContactForm);
    });
  }
}(typeof window !== "undefined" ? window : globalThis));
