"""Exercise Keynotes' real loader with fake CLR/file boundaries, without starting Revit.

Verify selection before loading, dependency binding after loading, and diagnostics.
The live Revit 2024/2025/2026 checks are in keynote_annotation_acceptance.md.
"""
import ast
import ntpath
import pathlib
import types
import unittest


SCRIPT = (pathlib.Path(__file__).resolve().parents[1]
          / 'FFE-pyRevit.extension/FFE-pyRevit.tab/DesignTools.panel/Keynotes.pushbutton/script.py')
CORE = 'Microsoft.Web.WebView2.Core'
WPF = 'Microsoft.Web.WebView2.Wpf'
OLD = '1.0.1343.22'
NEW = '1.0.4129.50'
PYREVIT = r'C:\pyRevit\bin\netfx'
REVIT = r'C:\Autodesk\Revit 2024'


class Identity:
    def __init__(self, name, version, token='2a8ab48044d2601e'):
        self.Name = name
        self.FullName = '{}, Version={}, Culture=neutral, PublicKeyToken={}'.format(name, version, token)


class FakeAssembly:
    """Model file metadata separately from the Core actually bound by the CLR."""
    def __init__(self, name, version, directory, host, reference=None, token='2a8ab48044d2601e'):
        self.identity = Identity(name, version, token)
        self.Location = ntpath.join(directory, name + '.dll')
        self.reference = reference or Identity(CORE, version, token)
        self.host = host
        self.bound_core = None
        self.webview_python_type = type('WebView2', (), {})
        self.creation_python_type = type('CoreWebView2CreationProperties', (), {})

    def GetName(self):
        return self.identity

    def GetReferencedAssemblies(self):
        return [self.reference]

    def GetType(self, name, throw_on_error):
        if name.endswith('.WebView2'):
            def get_property(property_name):
                assert property_name == 'CoreWebView2'
                core = self.bound_core or next(a for a in self.host.loaded if a.GetName().Name == CORE)
                return types.SimpleNamespace(PropertyType=types.SimpleNamespace(Assembly=core))
            return types.SimpleNamespace(python_type=self.webview_python_type, GetProperty=get_property)
        assert name.endswith('.CoreWebView2CreationProperties')
        return types.SimpleNamespace(python_type=self.creation_python_type)


class LoaderHost:
    """Record disk inspections and loads so rejected pairs cannot pollute the process."""
    def __init__(self):
        self.files = {}
        self.loaded = []
        self.loads = []
        self.inspections = []
        self.redirects = {}

    def add_pair(self, directory, version):
        return (self.add(CORE, directory, version), self.add(WPF, directory, version))

    def add(self, name, directory, version, **kwargs):
        assembly = FakeAssembly(name, version, directory, self, **kwargs)
        self.files[assembly.Location] = assembly
        return assembly

    def inspect(self, path):
        self.inspections.append(path)
        return self.files[path].GetName()

    def load(self, path):
        self.loads.append(path)
        assembly = self.redirects.get(path, self.files[path])
        self.loaded.append(assembly)
        return assembly


class KeynoteWebView2Tests(unittest.TestCase):
    def setUp(self):
        self.host = LoaderHost()
        self.environ = {'PATH': r'C:\Windows'}
        tree = ast.parse(SCRIPT.read_text(encoding='utf-8'))
        definitions = [node for node in tree.body if isinstance(node, ast.FunctionDef)]
        self.api = {
            'os': types.SimpleNamespace(
                path=types.SimpleNamespace(dirname=ntpath.dirname, join=ntpath.join,
                                           isfile=lambda path: path in self.host.files),
                environ=self.environ, pathsep=';'),
            'Assembly': types.SimpleNamespace(LoadFrom=self.host.load),
            'AssemblyName': types.SimpleNamespace(GetAssemblyName=self.host.inspect),
            'AppDomain': types.SimpleNamespace(CurrentDomain=types.SimpleNamespace(
                GetAssemblies=lambda: self.host.loaded)),
            'clr': types.SimpleNamespace(GetPythonType=lambda reflected: reflected.python_type),
            'BIN_DIR': PYREVIT,
        }
        exec(compile(ast.Module(body=definitions, type_ignores=[]), str(SCRIPT), 'exec'), self.api)
        self.api['get_revit_install_dir'] = lambda: REVIT

    def load(self):
        return self.api['load_webview2_types']()

    def test_prefers_pyrevit_pair_and_preserves_constructor_interface(self):
        core, wpf = self.host.add_pair(PYREVIT, NEW)
        self.host.add_pair(REVIT, OLD)
        result = self.load()
        self.assertEqual((wpf.webview_python_type, wpf.creation_python_type), result)
        self.assertEqual([core.Location, wpf.Location], self.host.loads)
        self.assertEqual([core.Location, wpf.Location], self.host.inspections)
        self.assertTrue(self.environ['PATH'].startswith(PYREVIT + ';'))
        path = self.environ['PATH']
        self.assertEqual(result, self.load())
        self.assertEqual(path, self.environ['PATH'])
        self.assertEqual(2, len(self.host.loads))

    def test_reuses_loaded_pair_without_inspecting_or_loading_files(self):
        core, wpf = self.host.add_pair(REVIT, OLD)
        self.host.loaded.extend([core, wpf])
        self.host.add_pair(PYREVIT, NEW)
        self.assertEqual(wpf.webview_python_type, self.load()[0])
        self.assertEqual([], self.host.loads)
        self.assertEqual([], self.host.inspections)

    def test_loaded_core_selects_matching_wpf_from_its_directory(self):
        core, wpf = self.host.add_pair(REVIT, OLD)
        self.host.loaded.append(core)
        self.host.add_pair(PYREVIT, NEW)
        self.assertEqual(wpf.webview_python_type, self.load()[0])
        self.assertEqual([wpf.Location], self.host.loads)

    def test_loaded_core_can_find_matching_wpf_in_another_directory(self):
        core = self.host.add(CORE, r'C:\OtherAddin', NEW)
        self.host.loaded.append(core)
        _, wpf = self.host.add_pair(PYREVIT, NEW)
        self.host.add_pair(REVIT, OLD)
        self.assertEqual(wpf.webview_python_type, self.load()[0])
        self.assertEqual([wpf.Location], self.host.loads)

    def test_loaded_wpf_selects_its_declared_core_dependency(self):
        core, wpf = self.host.add_pair(REVIT, OLD)
        self.host.loaded.append(wpf)
        self.host.add_pair(PYREVIT, NEW)
        self.assertEqual(wpf.webview_python_type, self.load()[0])
        self.assertEqual([core.Location], self.host.loads)

    def test_unavailable_pyrevit_pair_falls_back_to_revit_without_mixing(self):
        for partial in (False, True):
            with self.subTest(partial=partial):
                self.setUp()
                if partial:
                    self.host.add(CORE, PYREVIT, NEW)
                core, wpf = self.host.add_pair(REVIT, OLD)
                self.assertEqual(wpf.webview_python_type, self.load()[0])
                self.assertEqual([core.Location, wpf.Location], self.host.loads)

    def test_missing_files_report_paths_and_restart_guidance(self):
        with self.assertRaisesRegex(Exception, 'No compatible WebView2') as failure:
            self.load()
        self.assertIn(ntpath.join(PYREVIT, CORE + '.dll'), str(failure.exception))
        self.assertIn(ntpath.join(REVIT, WPF + '.dll'), str(failure.exception))
        self.assertIn('Fully close and restart Revit', str(failure.exception))
        self.assertEqual([], self.host.loads)

    def test_mismatched_installed_pair_is_rejected_before_loading(self):
        core = self.host.add(CORE, PYREVIT, OLD)
        wpf = self.host.add(WPF, PYREVIT, NEW)
        self.host.add_pair(REVIT, OLD)
        with self.assertRaisesRegex(Exception, 'SDK identities do not match') as failure:
            self.load()
        for value in (OLD, NEW, core.Location, wpf.Location):
            self.assertIn(value, str(failure.exception))
        self.assertEqual([], self.host.loads)

    def test_loaded_incompatible_pair_is_rejected_without_new_loads(self):
        core = self.host.add(CORE, REVIT, OLD)
        wpf = self.host.add(WPF, PYREVIT, NEW)
        self.host.loaded.extend([core, wpf])
        with self.assertRaisesRegex(Exception, 'Loaded WebView2 WPF requires') as failure:
            self.load()
        for value in (OLD, NEW, core.Location, wpf.Location):
            self.assertIn(value, str(failure.exception))
        self.assertEqual([], self.host.loads)

    def test_loaded_core_cannot_fall_back_to_an_incompatible_wpf(self):
        core = self.host.add(CORE, r'C:\OtherAddin', NEW)
        self.host.loaded.append(core)
        self.host.add_pair(REVIT, OLD)
        with self.assertRaisesRegex(Exception, 'No compatible WebView2'):
            self.load()
        self.assertEqual([], self.host.loads)

    def test_public_key_identity_must_match_before_loading(self):
        self.host.add(CORE, PYREVIT, NEW)
        self.host.add(WPF, PYREVIT, NEW, token='different')
        with self.assertRaisesRegex(Exception, 'SDK identities do not match'):
            self.load()
        self.assertEqual([], self.host.loads)

    def test_declared_reference_is_checked_after_loading(self):
        core, wpf = self.host.add_pair(PYREVIT, NEW)
        wpf.reference = Identity(CORE, OLD)
        with self.assertRaisesRegex(Exception, 'references a different Core SDK'):
            self.load()

    def test_actual_core_binding_is_checked_beyond_declared_reference(self):
        core, wpf = self.host.add_pair(PYREVIT, NEW)
        wrong_core = self.host.add(CORE, REVIT, OLD)
        wpf.bound_core = wrong_core
        with self.assertRaisesRegex(Exception, 'bound to an unexpected Core') as failure:
            self.load()
        self.assertIn(wrong_core.Location, str(failure.exception))

    def test_core_load_redirect_is_rejected_before_loading_wpf(self):
        core, _ = self.host.add_pair(PYREVIT, NEW)
        self.host.redirects[core.Location] = self.host.add(CORE, REVIT, OLD)
        with self.assertRaisesRegex(Exception, 'different WebView2 Core assembly than selected'):
            self.load()
        self.assertEqual([core.Location], self.host.loads)

    def test_wpf_load_redirect_is_rejected_even_if_it_references_selected_core(self):
        _, wpf = self.host.add_pair(PYREVIT, NEW)
        redirected = self.host.add(WPF, REVIT, OLD, reference=Identity(CORE, NEW))
        self.host.redirects[wpf.Location] = redirected
        with self.assertRaisesRegex(Exception, 'different WebView2 WPF assembly than selected'):
            self.load()

    def test_multiple_loaded_core_contexts_are_rejected(self):
        self.host.loaded.extend(self.host.add(CORE, directory, NEW)
                                for directory in (PYREVIT, REVIT))
        with self.assertRaisesRegex(Exception, 'Multiple WebView2'):
            self.load()
        self.assertEqual([], self.host.loads)

    def test_pythonnet_returns_callable_constructors_from_selected_types(self):
        _, wpf = self.host.add_pair(PYREVIT, NEW)
        self.api['clr'] = types.SimpleNamespace()
        self.api['Activator'] = types.SimpleNamespace(CreateInstance=lambda reflected: reflected.python_type())
        webview, creation = self.load()
        self.assertIsInstance(webview(), wpf.webview_python_type)
        self.assertIsInstance(creation(), wpf.creation_python_type)

    def test_async_failure_includes_versions_paths_and_original_exception(self):
        core, wpf = self.host.add_pair(PYREVIT, NEW)
        self.host.loaded.extend([core, wpf])
        tree = ast.parse(SCRIPT.read_text(encoding='utf-8'))
        window = next(node for node in tree.body if isinstance(node, ast.ClassDef)
                      and node.name == 'KeynoteManagerWindow')
        handler = next(node for node in window.body if isinstance(node, ast.FunctionDef)
                       and node.name == 'on_core_webview2_initialized')
        exec(compile(ast.Module(body=[handler], type_ignores=[]), str(SCRIPT), 'exec'), self.api)
        status = types.SimpleNamespace(Text='')
        args = types.SimpleNamespace(IsSuccess=False, InitializationException='Method not found: constructor')
        self.api['on_core_webview2_initialized'](types.SimpleNamespace(status_text=status), None, args)
        for value in ('Method not found: constructor', NEW, core.Location, wpf.Location,
                      'Fully close and restart Revit', 'pyRevit reload cannot replace loaded DLLs'):
            self.assertIn(value, status.Text)


if __name__ == '__main__':
    unittest.main()
