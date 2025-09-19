# -*- coding: UTF-8 -*-
# from pyrevit import forms
# from pyrevit import EXEC_PARAMS

# forms.alert(EXEC_PARAMS.event_args.CurrentActiveView.Name)


# from __future__ import annotations
import os, json, time, tempfile

from pyrevit import forms, revit

doc = revit.doc
doc_path = doc.PathName or "<Untitled>"
safe = doc_path.replace(":", "_").replace("\\", "_").replace("/", "_")
stamp_path = os.path.join(tempfile.gettempdir(), "pyrevit_syncstamp_" + safe + ".json")

elapsed_s = None
if os.path.exists(stamp_path):
    try:
        with open(stamp_path, "r") as f:
            payload = json.load(f)
        t0 = float(payload.get("t0", time.time()))
        elapsed_s = max(0.0, time.time() - t0)
    finally:
        try:
            os.remove(stamp_path)
        except Exception:
            pass

# Keep your hook lightweight—toast a quick, non-blocking message.
if elapsed_s is not None:
    forms.toast("Sync completed in {:.1f}s".format(elapsed_s), title="pyRevit", appid="pyRevit")
else:
    forms.toast("Sync completed", title="pyRevit", appid="pyRevit")