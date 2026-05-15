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

  function clearElement(element) {
    while (element.firstChild) {
      element.removeChild(element.firstChild);
    }
  }

  function appendInlineMarkdown(parent, text) {
    var source = String(text || "");
    var pattern = /(\*\*[^*]+\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))/g;
    var lastIndex = 0;
    var match;

    function appendText(value) {
      if (value) {
        parent.appendChild(document.createTextNode(value));
      }
    }

    while ((match = pattern.exec(source)) !== null) {
      appendText(source.slice(lastIndex, match.index));

      var token = match[0];
      var node = null;

      if (token.indexOf("**") === 0) {
        node = document.createElement("strong");
        node.textContent = token.slice(2, -2);
      } else if (token.indexOf("`") === 0) {
        node = document.createElement("code");
        node.textContent = token.slice(1, -1);
      } else {
        var linkMatch = token.match(/^\[([^\]]+)\]\(([^)]+)\)$/);
        if (linkMatch) {
          node = document.createElement("a");
          node.textContent = linkMatch[1];
          node.href = linkMatch[2];
          if (isExternalUrl(linkMatch[2])) {
            node.addEventListener("click", function (event) {
              event.preventDefault();
              openExternal(this.href);
            });
          }
        }
      }

      if (node) {
        parent.appendChild(node);
      } else {
        appendText(token);
      }

      lastIndex = pattern.lastIndex;
    }

    appendText(source.slice(lastIndex));
  }

  function renderMarkdown(target, markdownText) {
    var markdown = String(markdownText || "No changelog available.").replace(/<!--[\s\S]*?-->/g, "");
    var lines = markdown.split(/\r?\n/);
    var currentList = null;
    var currentParagraph = null;

    clearElement(target);

    function closeList() {
      currentList = null;
    }

    function closeParagraph() {
      currentParagraph = null;
    }

    lines.forEach(function (line) {
      var trimmed = line.replace(/\s+$/, "");
      var content = trimmed.trim();
      var headingMatch;
      var listMatch;
      var paragraphLine;

      if (!content) {
        closeList();
        closeParagraph();
        return;
      }

      headingMatch = content.match(/^(#{1,6})\s+(.+)$/);
      if (headingMatch) {
        closeList();
        closeParagraph();
        var level = Math.min(headingMatch[1].length, 4);
        var heading = document.createElement("h" + level);
        appendInlineMarkdown(heading, headingMatch[2]);
        target.appendChild(heading);
        return;
      }

      listMatch = content.match(/^[-*]\s+(.+)$/);
      if (listMatch) {
        closeParagraph();
        if (!currentList) {
          currentList = document.createElement("ul");
          target.appendChild(currentList);
        }
        var item = document.createElement("li");
        appendInlineMarkdown(item, listMatch[1]);
        currentList.appendChild(item);
        return;
      }

      closeList();
      paragraphLine = content;
      if (!currentParagraph) {
        currentParagraph = document.createElement("p");
        target.appendChild(currentParagraph);
      } else {
        currentParagraph.appendChild(document.createTextNode(" "));
      }
      appendInlineMarkdown(currentParagraph, paragraphLine);
    });
  }

  function loadData(data) {
    data = data || {};

    if (data.version) {
      setText("[data-about-version]", data.version);
    }

    var changelog = document.querySelector("#changelog-content");
    if (changelog) {
      renderMarkdown(changelog, data.changelog);
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
