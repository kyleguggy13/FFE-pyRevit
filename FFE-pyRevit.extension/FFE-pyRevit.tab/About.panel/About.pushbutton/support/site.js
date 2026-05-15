(function attachAboutApp(globalScope) {
  "use strict";

  function hasWebViewBridge() {
    return Boolean(
      globalScope.chrome &&
      globalScope.chrome.webview &&
      typeof globalScope.chrome.webview.postMessage === "function"
    );
  }

  function postWebViewMessage(message) {
    if (!hasWebViewBridge()) {
      return false;
    }

    globalScope.chrome.webview.postMessage(JSON.stringify(message));
    return true;
  }

  function isExternalUrl(url) {
    var value = String(url || "").toLowerCase();
    return (
      value.indexOf("http://") === 0 ||
      value.indexOf("https://") === 0 ||
      value.indexOf("mailto:") === 0
    );
  }

  function openExternal(url) {
    if (!isExternalUrl(url)) {
      return;
    }

    if (postWebViewMessage({ type: "openExternal", url: url })) {
      return;
    }

    globalScope.location.href = url;
  }

  function closeWindow() {
    if (postWebViewMessage({ type: "closeWindow" })) {
      return;
    }

    try {
      globalScope.close();
    } catch (ignore) {
      // Browser fallback only.
    }
  }

  function setText(selector, value) {
    Array.prototype.forEach.call(document.querySelectorAll(selector), function (element) {
      element.textContent = value;
    });
  }

  function loadData(data) {
    data = data || {};

    if (data.version) {
      setText("[data-about-version]", data.version);
    }

    var changelog = document.querySelector("#changelog-text");
    if (changelog) {
      changelog.textContent = data.changelog || "No changelog available.";
    }
  }

  function setupExternalLinks() {
    Array.prototype.forEach.call(document.querySelectorAll("a[href]"), function (link) {
      var href = link.getAttribute("href");
      if (!isExternalUrl(href)) {
        return;
      }

      link.addEventListener("click", function (event) {
        event.preventDefault();
        openExternal(href);
      });
    });
  }

  function init() {
    setupExternalLinks();

    var closeButton = document.querySelector("#close-window");
    if (closeButton) {
      closeButton.addEventListener("click", closeWindow);
    }

    postWebViewMessage({ type: "appReady" });
  }

  globalScope.ffeAbout = {
    loadData: loadData
  };

  if (typeof document !== "undefined") {
    document.addEventListener("DOMContentLoaded", init);
  }
}(typeof window !== "undefined" ? window : globalThis));
