# -*- coding: utf-8 -*-
"""Purpose: select compatible WebView2 WPF/Core assemblies without mixing SDKs.
Behaviors: inspect identities, verify CLR binding, and report recovery details.
Revit API: targets Revit 2025/2026; no document access or transactions.
Design: adapt the tested Keynotes loader without changing existing commands.
"""
import os
import clr
from System import Activator, AppDomain
from System.Reflection import Assembly, AssemblyName
from pyrevit import HOST_APP, BIN_DIR

revit_app = HOST_APP.app

def get_revit_install_dir():
    """
    Get the Revit installation directory, which is needed to load the bundled WebView2 assemblies.
    This function tries multiple methods to find the path in case of different Revit versions or configurations.
    """
    try:
        app_path = revit_app.ApplicationPath
        if app_path and os.path.exists(app_path):
            return os.path.dirname(app_path)
    except:
        pass

    try:
        version = str(revit_app.VersionNumber) or "2026"
    except:
        version = "2026"

    return os.path.join(
        os.environ.get("ProgramFiles", r"C:\Program Files"),
        "Autodesk",
        "Revit {0}".format(version)
    )


def get_loaded_webview2_assemblies():
    """Inspect managed WebView2 assemblies without loading another SDK version."""
    return [assembly for assembly in AppDomain.CurrentDomain.GetAssemblies()
            if assembly.GetName().Name.startswith("Microsoft.Web.WebView2.")]


def describe_webview2_assembly(assembly):
    """Include assembly identity and origin even for assemblies with no file location."""
    try:
        location = assembly.Location or "[location unavailable]"
    except Exception:
        location = "[location unavailable]"
    return "{0} | {1}".format(assembly.GetName().FullName, location)


def format_webview2_error(error, details=None):
    """Preserve failure details and explain recovery from persistent DLL bindings."""
    lines = list(details or [])
    try:
        loaded = get_loaded_webview2_assemblies()
        lines.extend("Loaded: " + describe_webview2_assembly(assembly) for assembly in loaded)
        if not loaded:
            lines.append("No managed WebView2 assemblies are loaded.")
    except Exception as inspection_error:
        lines.append("Could not inspect loaded assemblies: {0}".format(inspection_error))
    return ("{0}\n\nWebView2 assembly details:\n{1}\n\n"
            "Fully close and restart Revit after installing this fix or changing add-ins. "
            "A pyRevit reload cannot replace loaded DLLs. If the failure persists after "
            "restart, share these versions and paths to identify the conflicting or missing DLLs."
            ).format(error, "\n".join(lines))


def get_webview2_core_reference(wpf_assembly):
    """Read the WPF control's declared Core dependency rather than assuming compatibility."""
    references = [name for name in wpf_assembly.GetReferencedAssemblies()
                  if name.Name == "Microsoft.Web.WebView2.Core"]
    if len(references) != 1:
        raise Exception("The WebView2 WPF assembly must reference exactly one Core assembly.")
    return references[0].FullName


def webview2_core_identity(assembly_name):
    """Compare SDK identities including version, culture, and public key token."""
    return "Microsoft.Web.WebView2.Core," + assembly_name.FullName.split(",", 1)[1]


def load_webview2_types():
    """Select one compatible Core/WPF pair and verify the actual CLR dependency binding."""
    core_name = "Microsoft.Web.WebView2.Core"
    wpf_name = "Microsoft.Web.WebView2.Wpf"
    details = []
    try:
        loaded = get_loaded_webview2_assemblies()
        cores = [assembly for assembly in loaded if assembly.GetName().Name == core_name]
        wpfs = [assembly for assembly in loaded if assembly.GetName().Name == wpf_name]
        # Multiple load contexts can expose incompatible types even with identical names.
        if len(cores) > 1 or len(wpfs) > 1:
            raise Exception("Multiple WebView2 Core or WPF assemblies are already loaded.")
        core = cores[0] if cores else None
        wpf = wpfs[0] if wpfs else None
        expected_core = core.GetName().FullName if core is not None else None
        expected_wpf = wpf.GetName().FullName if wpf is not None else None
        if wpf is not None:
            reference = get_webview2_core_reference(wpf)
            if expected_core and reference != expected_core:
                raise Exception("Loaded WebView2 WPF requires {0}, but Core is {1}.".format(
                    reference, expected_core))
            expected_core = reference

        directories = []
        for assembly in (core, wpf):
            if assembly is not None:
                try:
                    if assembly.Location:
                        directories.append(os.path.dirname(assembly.Location))
                except Exception:
                    pass  # A locationless assembly can still match an installed SDK.
        for directory in (BIN_DIR, get_revit_install_dir()):
            if directory not in directories:
                directories.append(directory)

        def inspect_file(directory, name):
            path = os.path.join(directory, name + ".dll")
            if not os.path.isfile(path):
                details.append("Not found: " + path)
                return None
            identity = AssemblyName.GetAssemblyName(path)
            details.append("Candidate: {0} | {1}".format(identity.FullName, path))
            if identity.Name != name:
                raise Exception("Unexpected assembly identity in {0}.".format(path))
            return path, identity

        core_path = wpf_path = None
        if expected_core is None:
            # Inspect both files before loading either; a partial pyRevit SDK can fall
            # back to Revit, but an installed mismatched pair must be reported.
            for directory in directories:
                core_file = inspect_file(directory, core_name)
                wpf_file = inspect_file(directory, wpf_name)
                if core_file and wpf_file:
                    if core_file[1].FullName != webview2_core_identity(wpf_file[1]):
                        raise Exception("WebView2 Core and WPF SDK identities do not match.")
                    core_path, wpf_path = core_file[0], wpf_file[0]
                    expected_core = core_file[1].FullName
                    expected_wpf = wpf_file[1].FullName
                    break
        else:
            for directory in directories:
                if core is None and core_path is None:
                    candidate = inspect_file(directory, core_name)
                    if candidate and candidate[1].FullName == expected_core:
                        core_path = candidate[0]
                if wpf is None and wpf_path is None:
                    candidate = inspect_file(directory, wpf_name)
                    if candidate and webview2_core_identity(candidate[1]) == expected_core:
                        wpf_path = candidate[0]
                        expected_wpf = candidate[1].FullName

        if (core is None and core_path is None) or (wpf is None and wpf_path is None):
            raise Exception("No compatible WebView2 Core/WPF pair was found{0}.".format(
                " for " + expected_core if expected_core else ""))
        if core is None:
            core = Assembly.LoadFrom(core_path)
        if core.GetName().FullName != expected_core:
            raise Exception("The CLR loaded a different WebView2 Core assembly than selected.")
        if wpf is None:
            wpf = Assembly.LoadFrom(wpf_path)
        if wpf.GetName().FullName != expected_wpf:
            raise Exception("The CLR loaded a different WebView2 WPF assembly than selected.")
        if get_webview2_core_reference(wpf) != expected_core:
            raise Exception("The selected WebView2 WPF assembly references a different Core SDK.")

        webview_type = wpf.GetType("Microsoft.Web.WebView2.Wpf.WebView2", True)
        creation_type = wpf.GetType("Microsoft.Web.WebView2.Wpf.CoreWebView2CreationProperties", True)
        # Reading the property's type resolves the dependency through the CLR. File
        # names alone cannot detect a binding redirected to another add-in's Core.
        bound_core = webview_type.GetProperty("CoreWebView2").PropertyType.Assembly
        if bound_core != core:
            raise Exception("WebView2 WPF bound to an unexpected Core assembly: {0}".format(
                describe_webview2_assembly(bound_core)))

        try:
            native_dir = os.path.dirname(core.Location) if core.Location else None
        except Exception:
            native_dir = None
        if native_dir and native_dir not in os.environ.get("PATH", "").split(os.pathsep):
            os.environ["PATH"] = native_dir + os.pathsep + os.environ.get("PATH", "")
        if hasattr(clr, "GetPythonType"):
            return clr.GetPythonType(webview_type), clr.GetPythonType(creation_type)
        # pythonnet does not expose GetPythonType; preserve the two callable constructors
        # without importing an ambiguous Microsoft.Web.WebView2 namespace.
        return (lambda: Activator.CreateInstance(webview_type),
                lambda: Activator.CreateInstance(creation_type))
    except Exception as error:
        raise Exception(format_webview2_error(error, details))



