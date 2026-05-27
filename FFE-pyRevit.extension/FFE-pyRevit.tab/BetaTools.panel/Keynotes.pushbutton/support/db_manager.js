(function attachKeynoteDb(globalScope) {
  "use strict";

  var supabaseClient = null;
  var activeChannel = null;
  var currentLibraryId = "";
  var activeClaimsChannel = null;
  var currentClaimsLibraryId = "";

  function text(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value);
  }

  function trim(value) {
    return text(value).replace(/^\s+|\s+$/g, "");
  }

  function getSupabaseGlobal() {
    return globalScope.supabase || null;
  }

  function makeError(error, fallbackMessage) {
    var message = fallbackMessage || "Supabase request failed.";
    if (error) {
      message = error.message || error.error_description || error.details || message;
    }
    return new Error(message);
  }

  function normalizeSnapshot(snapshot) {
    snapshot = snapshot || {};
    snapshot.entries = snapshot.entries || [];
    snapshot.datasetVersion = Number(snapshot.datasetVersion || snapshot.dataset_version || 0);
    snapshot.entryCount = Number(snapshot.entryCount || snapshot.entries.length || 0);
    snapshot.libraryId = text(snapshot.libraryId || snapshot.library_id);
    snapshot.libraryKey = text(snapshot.libraryKey || snapshot.library_key);
    snapshot.displayPath = text(snapshot.displayPath || snapshot.display_path);
    snapshot.encoding = text(snapshot.encoding || "utf-8");
    snapshot.lineEnding = snapshot.lineEnding || snapshot.line_ending || "\r\n";
    return snapshot;
  }

  function normalizeClaims(data) {
    data = data || {};
    data.libraryId = text(data.libraryId || data.library_id);
    data.libraryKey = text(data.libraryKey || data.library_key);
    data.claims = (data.claims || []).map(function (claim) {
      return {
        claimKey: text(claim.claimKey || claim.claim_key),
        dbId: text(claim.dbId || claim.db_id),
        key: text(claim.key || claim.keynote_key),
        clientId: text(claim.clientId || claim.client_id),
        clientName: text(claim.clientName || claim.client_name),
        updatedAt: text(claim.updatedAt || claim.updated_at)
      };
    });
    return data;
  }

  function rpc(functionName, args, options) {
    options = options || {};

    if (!supabaseClient) {
      return Promise.reject(new Error("Supabase is not configured."));
    }

    return supabaseClient.rpc(functionName, args || {}).then(function (response) {
      var data = response.data || {};
      if (response.error) {
        throw makeError(response.error);
      }
      if (data.status === "error") {
        throw new Error(data.message || "Supabase request failed.");
      }
      if (options.raw) {
        return data;
      }
      return normalizeSnapshot(data);
    });
  }

  function configure(settings) {
    var supabaseApi = getSupabaseGlobal();
    var url = trim(settings && settings.url);
    var anonKey = trim(settings && (settings.anonKey || settings.publishableKey));

    if (!supabaseApi || typeof supabaseApi.createClient !== "function") {
      throw new Error("The Supabase JavaScript client did not load.");
    }
    if (!url || !anonKey) {
      throw new Error("Supabase URL and publishable key are required.");
    }

    unsubscribe();
    supabaseClient = supabaseApi.createClient(url, anonKey, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
        detectSessionInUrl: false
      }
    });
    return supabaseClient;
  }

  function ensureLibrary(payload) {
    payload = payload || {};
    return rpc("ensure_keynote_library", {
      p_library_key: payload.libraryKey,
      p_display_path: payload.displayPath || payload.keynotePath || "",
      p_encoding: payload.encoding || "utf-8",
      p_line_ending: payload.lineEnding || "\r\n",
      p_seed_entries: payload.entries || [],
      p_client_id: payload.clientId || "",
      p_client_name: payload.clientName || ""
    });
  }

  function getSnapshot(libraryKey) {
    return rpc("get_keynote_snapshot", {
      p_library_key: libraryKey
    });
  }

  function syncFileSnapshot(payload) {
    payload = payload || {};
    return rpc("sync_keynote_file_snapshot", {
      p_library_key: payload.libraryKey,
      p_display_path: payload.displayPath || payload.keynotePath || "",
      p_encoding: payload.encoding || "utf-8",
      p_line_ending: payload.lineEnding || "\r\n",
      p_file_hash: payload.fileHash || "",
      p_last_write_utc: payload.lastWriteUtc || null,
      p_entries: payload.entries || [],
      p_client_id: payload.clientId || "",
      p_client_name: payload.clientName || ""
    });
  }

  function saveChanges(payload) {
    payload = payload || {};
    return rpc("save_keynote_changes", {
      p_library_key: payload.libraryKey,
      p_client_id: payload.clientId || "",
      p_client_name: payload.clientName || "",
      p_base_dataset_version: payload.baseDatasetVersion || 0,
      p_changes: payload.changes || {}
    }, { raw: true }).then(function (data) {
      if (data && data.snapshot) {
        data.snapshot = normalizeSnapshot(data.snapshot);
      }
      if (data && data.status === "ready") {
        return normalizeSnapshot(data);
      }
      return data || {};
    });
  }

  function getEditClaims(libraryKey) {
    return rpc("get_keynote_edit_claims", {
      p_library_key: libraryKey
    }, { raw: true }).then(normalizeClaims);
  }

  function setEditClaims(payload) {
    payload = payload || {};
    return rpc("replace_keynote_edit_claims", {
      p_library_key: payload.libraryKey,
      p_client_id: payload.clientId || "",
      p_client_name: payload.clientName || "",
      p_claims: payload.claims || []
    }, { raw: true }).then(normalizeClaims);
  }

  function removeChannel(channel) {
    if (channel && supabaseClient && typeof supabaseClient.removeChannel === "function") {
      supabaseClient.removeChannel(channel);
    }
  }

  function unsubscribe() {
    var channel = activeChannel;
    var claimsChannel = activeClaimsChannel;
    activeChannel = null;
    activeClaimsChannel = null;
    currentLibraryId = "";
    currentClaimsLibraryId = "";

    removeChannel(channel);
    removeChannel(claimsChannel);
  }

  function subscribeLibrary(libraryId, clientId, handlers) {
    handlers = handlers || {};
    libraryId = text(libraryId);
    clientId = text(clientId);

    if (!supabaseClient || !libraryId) {
      return null;
    }
    if (currentLibraryId === libraryId && activeChannel) {
      return activeChannel;
    }

    removeChannel(activeChannel);
    activeChannel = null;
    currentLibraryId = libraryId;
    activeChannel = supabaseClient.channel("keynote-library-" + libraryId)
      .on(
        "postgres_changes",
        {
          event: "UPDATE",
          schema: "public",
          table: "keynote_libraries",
          filter: "id=eq." + libraryId
        },
        function (payload) {
          var nextClientId = text(payload && payload.new && payload.new.last_saved_by_client_id);
          if (clientId && nextClientId === clientId) {
            return;
          }
          if (typeof handlers.onRemoteChange === "function") {
            handlers.onRemoteChange(payload || {});
          }
        }
      )
      .subscribe(function (status) {
        if (typeof handlers.onStatus === "function") {
          handlers.onStatus(status);
        }
      });

    return activeChannel;
  }

  function subscribeEditClaims(libraryId, clientId, handlers) {
    handlers = handlers || {};
    libraryId = text(libraryId);

    if (!supabaseClient || !libraryId) {
      return null;
    }
    if (currentClaimsLibraryId === libraryId && activeClaimsChannel) {
      return activeClaimsChannel;
    }

    removeChannel(activeClaimsChannel);
    activeClaimsChannel = null;
    currentClaimsLibraryId = libraryId;
    activeClaimsChannel = supabaseClient.channel("keynote-edit-claims-" + libraryId)
      .on(
        "postgres_changes",
        {
          event: "*",
          schema: "public",
          table: "keynote_edit_claims",
          filter: "library_id=eq." + libraryId
        },
        function (payload) {
          if (typeof handlers.onClaimsChanged === "function") {
            handlers.onClaimsChanged(payload || {});
          }
        }
      )
      .subscribe(function (status) {
        if (typeof handlers.onStatus === "function") {
          handlers.onStatus(status);
        }
      });

    return activeClaimsChannel;
  }

  globalScope.ffeKeynoteDb = {
    configure: configure,
    ensureLibrary: ensureLibrary,
    getSnapshot: getSnapshot,
    syncFileSnapshot: syncFileSnapshot,
    saveChanges: saveChanges,
    getEditClaims: getEditClaims,
    setEditClaims: setEditClaims,
    subscribeLibrary: subscribeLibrary,
    subscribeEditClaims: subscribeEditClaims,
    unsubscribe: unsubscribe
  };
}(typeof window !== "undefined" ? window : this));
