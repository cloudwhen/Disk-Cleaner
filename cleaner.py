import os
import shutil
import glob
import winreg
from pathlib import Path
import ctypes


class DiskCleaner:
    def __init__(self):
        self.results = {}
        self.temp_dirs = [
            os.environ.get("TEMP", ""),
            os.environ.get("TMP", ""),
            os.path.join(os.environ.get("SYSTEMROOT", ""), "Temp"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Temp"),
        ]

        self.browser_cache_dirs = self._get_browser_cache_dirs()

        self.edge_dirs = self._get_edge_dirs()

        self.jianying_dirs = self._get_jianying_dirs()

        self.log_dirs = [
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "INetCache"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "INetCookies"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "History"),
            os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Recent"),
        ]

        self.installed_software = self._get_installed_software()

        self.common_residual_patterns = [
            'anyviewer', 'teamviewer', 'anydesk', 'remotedesktop',
            'splashtop', 'logmein', 'supremo',
            'vnc', 'radmin', 'dameware',
            'microsoft office', 'office'
        ]

    def _get_installed_software(self):
        installed = set()

        try:
            registry_paths = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            ]

            for root_key, sub_key in registry_paths:
                try:
                    with winreg.OpenKey(root_key, sub_key) as key:
                        i = 0
                        while True:
                            try:
                                software_key = winreg.EnumKey(key, i)
                                with winreg.OpenKey(key, software_key) as subkey:
                                    try:
                                        display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                        if display_name:
                                            installed.add(display_name.lower())
                                    except WindowsError:
                                        pass
                                i += 1
                            except WindowsError:
                                break
                except WindowsError:
                    pass
        except Exception:
            pass

        return installed

    def _get_browser_cache_dirs(self):
        cache_dirs = []

        try:
            chrome_path = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data", "Default", "Cache")
            if os.path.exists(chrome_path):
                cache_dirs.append(("Chrome缓存", chrome_path))

            chrome_code_cache = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Google", "Chrome", "User Data", "Default", "Code Cache")
            if os.path.exists(chrome_code_cache):
                cache_dirs.append(("Chrome代码缓存", chrome_code_cache))

            edge_path = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Edge", "User Data", "Default", "Cache")
            if os.path.exists(edge_path):
                cache_dirs.append(("Edge缓存", edge_path))

            firefox_path = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Mozilla", "Firefox", "Profiles")
            if os.path.exists(firefox_path):
                for profile in os.listdir(firefox_path):
                    cache_path = os.path.join(firefox_path, profile, "cache2")
                    if os.path.exists(cache_path):
                        cache_dirs.append(("Firefox缓存", cache_path))
        except Exception:
            pass

        return cache_dirs

    def _get_edge_dirs(self):
        edge_dirs = []

        try:
            localappdata = os.environ.get("LOCALAPPDATA", "")
            programdata = os.environ.get("PROGRAMDATA", "")

            edge_webview = os.path.join(programdata, "Microsoft", "EdgeUpdate")
            if os.path.exists(edge_webview):
                for item in os.listdir(edge_webview):
                    item_path = os.path.join(edge_webview, item)
                    if os.path.isdir(item_path) and "WebView" in item:
                        edge_dirs.append(("Edge WebView旧版本", item_path))

            edge_core = os.path.join(programdata, "Microsoft", "EdgeCore")
            if os.path.exists(edge_core):
                for item in os.listdir(edge_core):
                    item_path = os.path.join(edge_core, item)
                    if os.path.isdir(item_path):
                        edge_dirs.append(("Edge Core旧版本", item_path))

            edge_cache = os.path.join(localappdata, "Microsoft", "Edge", "User Data")
            if os.path.exists(edge_cache):
                for profile in os.listdir(edge_cache):
                    profile_path = os.path.join(edge_cache, profile)
                    if os.path.isdir(profile_path):
                        cache_path = os.path.join(profile_path, "Cache")
                        if os.path.exists(cache_path):
                            edge_dirs.append(("Edge浏览器缓存", cache_path))

                        code_cache = os.path.join(profile_path, "Code Cache")
                        if os.path.exists(code_cache):
                            edge_dirs.append(("Edge代码缓存", code_cache))

                        gpucache = os.path.join(profile_path, "GPUCache")
                        if os.path.exists(gpucache):
                            edge_dirs.append(("Edge GPU缓存", gpucache))
        except Exception:
            pass

        return edge_dirs

    def _get_jianying_dirs(self):
        jianying_dirs = []

        try:
            localappdata = os.environ.get("LOCALAPPDATA", "")

            capcut_base = os.path.join(localappdata, "CapCut")
            if os.path.exists(capcut_base):
                material_cache = os.path.join(capcut_base, "MaterialCache")
                if os.path.exists(material_cache):
                    jianying_dirs.append(("剪映如下缓存内容", material_cache))

                audio_cache = os.path.join(capcut_base, "AudioCache")
                if os.path.exists(audio_cache):
                    jianying_dirs.append(("剪映音频缓存", audio_cache))

                workspace_cache = os.path.join(capcut_base, "WorkspaceCache")
                if os.path.exists(workspace_cache):
                    jianying_dirs.append(("剪映工作平台缓存", workspace_cache))

                effect_cache = os.path.join(capcut_base, "EffectCache")
                if os.path.exists(effect_cache):
                    jianying_dirs.append(("剪映艺术特效缓存", effect_cache))

                temp_cache = os.path.join(capcut_base, "Temp")
                if os.path.exists(temp_cache):
                    jianying_dirs.append(("剪映临时文件", temp_cache))

                preview_cache = os.path.join(capcut_base, "PreviewCache")
                if os.path.exists(preview_cache):
                    jianying_dirs.append(("剪映预览缓存", preview_cache))

                export_cache = os.path.join(capcut_base, "ExportCache")
                if os.path.exists(export_cache):
                    jianying_dirs.append(("剪映导出缓存", export_cache))
        except Exception:
            pass

        return jianying_dirs

    def scan_c_drive(self):
        results = {}

        results.update(self._scan_temp_files())
        results.update(self._scan_browser_cache())
        results.update(self._scan_edge_dirs())
        results.update(self._scan_jianying_dirs())
        results.update(self._scan_system_logs())
        results.update(self._scan_recycle_bin())
        results.update(self._scan_program_files_residuals())
        results.update(self._scan_appdata_residuals())
        results.update(self._scan_invalid_shortcuts())

        return results

    def _scan_program_files_residuals(self):
        results = {}
        program_dirs = []

        program_files = os.path.join("C:\\", "Program Files")
        program_files_x86 = os.path.join("C:\\", "Program Files (x86)")

        if os.path.exists(program_files):
            program_dirs.append(program_files)
        if os.path.exists(program_files_x86):
            program_dirs.append(program_files_x86)

        for program_dir in program_dirs:
            try:
                for item in os.listdir(program_dir):
                    item_path = os.path.join(program_dir, item)
                    if os.path.isdir(item_path):
                        residual_info = self._is_residual_directory(item, item_path)
                        if residual_info:
                            size = self._get_dir_size(item_path)
                            if size > 0:
                                file_id = f"program_residual_{len(results)}"
                                results[file_id] = {
                                    "category": f"软件残留 ({residual_info})",
                                    "path": item_path,
                                    "size": size
                                }
            except Exception:
                continue

        return results

    def _is_residual_directory(self, dir_name, dir_path):
        dir_name_lower = dir_name.lower()

        for installed in self.installed_software:
            if self._is_name_match(dir_name_lower, installed):
                return None

        if self._has_uninstall_exe(dir_path):
            return None

        for pattern in self.common_residual_patterns:
            if pattern in dir_name_lower:
                if not self._is_name_match_any(dir_name_lower):
                    return f"常见残留 ({pattern})"

        return None

    def _is_name_match(self, dir_name, installed_name):
        dir_name_clean = dir_name.replace(" ", "").replace("-", "").replace("_", "")
        installed_clean = installed_name.replace(" ", "").replace("-", "").replace("_", "")

        if dir_name_clean in installed_clean or installed_clean in dir_name_clean:
            return True

        return False

    def _is_name_match_any(self, dir_name):
        for installed in self.installed_software:
            if self._is_name_match(dir_name, installed):
                return True
        return False

    def _has_uninstall_exe(self, dir_path):
        try:
            for root, dirs, files in os.walk(dir_path):
                for file in files:
                    file_lower = file.lower()
                    if "uninstall" in file_lower and file_lower.endswith('.exe'):
                        return True
        except Exception:
            pass
        return False

    def _scan_appdata_residuals(self):
        results = {}

        appdata_dirs = [
            os.environ.get("APPDATA", ""),
            os.environ.get("LOCALAPPDATA", ""),
            os.environ.get("PROGRAMDATA", "")
        ]

        for appdata_dir in appdata_dirs:
            if not appdata_dir or not os.path.exists(appdata_dir):
                continue

            try:
                for item in os.listdir(appdata_dir):
                    item_path = os.path.join(appdata_dir, item)
                    if os.path.isdir(item_path):
                        residual_info = self._is_residual_appdata_directory(item, item_path)
                        if residual_info:
                            size = self._get_dir_size(item_path)
                            if size > 0:
                                file_id = f"appdata_residual_{len(results)}"
                                results[file_id] = {
                                    "category": f"软件残留 ({residual_info})",
                                    "path": item_path,
                                    "size": size
                                }
            except Exception:
                continue

        return results

    def _is_residual_appdata_directory(self, dir_name, dir_path):
        dir_name_lower = dir_name.lower()

        for installed in self.installed_software:
            if self._is_name_match(dir_name_lower, installed):
                return None

        for pattern in self.common_residual_patterns:
            if pattern in dir_name_lower:
                if not self._is_name_match_any(dir_name_lower):
                    return f"常见残留 ({pattern})"

        return None

    def _scan_invalid_shortcuts(self):
        results = {}

        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        start_menu_path = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs")
        public_start_menu = os.path.join(os.environ.get("PROGRAMDATA", ""), "Microsoft", "Windows", "Start Menu", "Programs")

        shortcut_paths = [desktop_path, start_menu_path, public_start_menu]

        for shortcut_dir in shortcut_paths:
            if not shortcut_dir or not os.path.exists(shortcut_dir):
                continue

            try:
                for root, dirs, files in os.walk(shortcut_dir):
                    for file in files:
                        if file.lower().endswith('.lnk'):
                            shortcut_path = os.path.join(root, file)
                            if self._is_invalid_shortcut(shortcut_path):
                                try:
                                    size = os.path.getsize(shortcut_path)
                                    file_id = f"invalid_shortcut_{len(results)}"
                                    results[file_id] = {
                                        "category": "无效快捷方式",
                                        "path": shortcut_path,
                                        "size": size
                                    }
                                except Exception:
                                    continue
            except Exception:
                continue

        return results

    def _is_invalid_shortcut(self, shortcut_path):
        try:
            import pythoncom
            from win32com.shell import shell, shellcon

            shell_link = pythoncom.CoCreateInstance(
                shell.CLSID_ShellLink,
                None,
                pythoncom.CLSCTX_INPROC_SERVER,
                shell.IID_IShellLink
            )

            persist_file = shell_link.QueryInterface(pythoncom.IID_IPersistFile)
            persist_file.Load(shortcut_path)

            target = shell_link.GetPath(shell.SLGP_SHORTPATH)[0]

            if not target:
                return True

            if os.path.exists(target):
                return False

            return True
        except Exception:
            return False

    def _scan_temp_files(self):
        results = {}

        for temp_dir in self.temp_dirs:
            if not temp_dir or not os.path.exists(temp_dir):
                continue

            try:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size = os.path.getsize(file_path)
                            file_id = f"temp_{len(results)}"
                            results[file_id] = {
                                "category": "临时文件",
                                "path": file_path,
                                "size": size
                            }
                        except Exception:
                            continue
            except Exception:
                continue

        return results

    def _scan_browser_cache(self):
        results = {}

        for category, cache_dir in self.browser_cache_dirs:
            if not os.path.exists(cache_dir):
                continue

            try:
                for root, dirs, files in os.walk(cache_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size = os.path.getsize(file_path)
                            file_id = f"browser_{len(results)}"
                            results[file_id] = {
                                "category": category,
                                "path": file_path,
                                "size": size
                            }
                        except Exception:
                            continue
            except Exception:
                continue

        return results

    def _scan_edge_dirs(self):
        results = {}

        for category, edge_dir in self.edge_dirs:
            if not os.path.exists(edge_dir):
                continue

            try:
                if os.path.isdir(edge_dir):
                    size = self._get_dir_size(edge_dir)
                    if size > 0:
                        file_id = f"edge_{len(results)}"
                        results[file_id] = {
                            "category": category,
                            "path": edge_dir,
                            "size": size
                        }
                else:
                    for root, dirs, files in os.walk(edge_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                size = os.path.getsize(file_path)
                                file_id = f"edge_{len(results)}"
                                results[file_id] = {
                                    "category": category,
                                    "path": file_path,
                                    "size": size
                                }
                            except Exception:
                                continue
            except Exception:
                continue

        return results

    def _scan_jianying_dirs(self):
        results = {}

        for category, jianying_dir in self.jianying_dirs:
            if not os.path.exists(jianying_dir):
                continue

            try:
                if os.path.isdir(jianying_dir):
                    size = self._get_dir_size(jianying_dir)
                    if size > 0:
                        file_id = f"jianying_{len(results)}"
                        results[file_id] = {
                            "category": category,
                            "path": jianying_dir,
                            "size": size
                        }
                else:
                    for root, dirs, files in os.walk(jianying_dir):
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                size = os.path.getsize(file_path)
                                file_id = f"jianying_{len(results)}"
                                results[file_id] = {
                                    "category": category,
                                    "path": file_path,
                                    "size": size
                                }
                            except Exception:
                                continue
            except Exception:
                continue

        return results

    def _scan_system_logs(self):
        results = {}

        for log_dir in self.log_dirs:
            if not os.path.exists(log_dir):
                continue

            try:
                for root, dirs, files in os.walk(log_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size = os.path.getsize(file_path)
                            file_id = f"log_{len(results)}"
                            results[file_id] = {
                                "category": "系统日志",
                                "path": file_path,
                                "size": size
                            }
                        except Exception:
                            continue
            except Exception:
                continue

        prefetch_dir = os.path.join(os.environ.get("SYSTEMROOT", ""), "Prefetch")
        if os.path.exists(prefetch_dir):
            try:
                for file in os.listdir(prefetch_dir):
                    if file.endswith(".pf"):
                        file_path = os.path.join(prefetch_dir, file)
                        try:
                            size = os.path.getsize(file_path)
                            file_id = f"prefetch_{len(results)}"
                            results[file_id] = {
                                "category": "预读取文件",
                                "path": file_path,
                                "size": size
                            }
                        except Exception:
                            continue
            except Exception:
                pass

        return results

    def _scan_recycle_bin(self):
        results = {}

        drives = ["C:\\", "D:\\", "E:\\", "F:\\"]

        for drive in drives:
            if not os.path.exists(drive):
                continue

            recycle_path = os.path.join(drive, "$Recycle.Bin")
            if os.path.exists(recycle_path):
                try:
                    for root, dirs, files in os.walk(recycle_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                size = os.path.getsize(file_path)
                                file_id = f"recycle_{len(results)}"
                                results[file_id] = {
                                    "category": "回收站",
                                    "path": file_path,
                                    "size": size
                                }
                            except Exception:
                                continue
                except Exception:
                    continue

        return results

    def clean_items(self, selected_items):
        success_count = 0
        total_size = 0

        for item_path in selected_items:
            try:
                if os.path.isfile(item_path):
                    size = os.path.getsize(item_path)
                    os.remove(item_path)
                    success_count += 1
                    total_size += size
                elif os.path.isdir(item_path):
                    size = self._get_dir_size(item_path)
                    shutil.rmtree(item_path)
                    success_count += 1
                    total_size += size
            except Exception as e:
                continue

        return success_count, total_size

    def _get_dir_size(self, path):
        total_size = 0
        try:
            for dirpath, dirnames, filenames in os.walk(path):
                for filename in filenames:
                    file_path = os.path.join(dirpath, filename)
                    try:
                        total_size += os.path.getsize(file_path)
                    except Exception:
                        continue
        except Exception:
            pass
        return total_size

    def clean_temp_files(self):
        cleaned_count = 0
        total_size = 0

        for temp_dir in self.temp_dirs:
            if not temp_dir or not os.path.exists(temp_dir):
                continue

            try:
                for root, dirs, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size = os.path.getsize(file_path)
                            os.remove(file_path)
                            cleaned_count += 1
                            total_size += size
                        except Exception:
                            continue
            except Exception:
                continue

        return cleaned_count, total_size

    def clean_browser_cache(self):
        cleaned_count = 0
        total_size = 0

        for category, cache_dir in self.browser_cache_dirs:
            if not os.path.exists(cache_dir):
                continue

            try:
                for root, dirs, files in os.walk(cache_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size = os.path.getsize(file_path)
                            os.remove(file_path)
                            cleaned_count += 1
                            total_size += size
                        except Exception:
                            continue
            except Exception:
                continue

        return cleaned_count, total_size

    def clean_system_logs(self):
        cleaned_count = 0
        total_size = 0

        for log_dir in self.log_dirs:
            if not os.path.exists(log_dir):
                continue

            try:
                for root, dirs, files in os.walk(log_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        try:
                            size = os.path.getsize(file_path)
                            os.remove(file_path)
                            cleaned_count += 1
                            total_size += size
                        except Exception:
                            continue
            except Exception:
                continue

        prefetch_dir = os.path.join(os.environ.get("SYSTEMROOT", ""), "Prefetch")
        if os.path.exists(prefetch_dir):
            try:
                for file in os.listdir(prefetch_dir):
                    if file.endswith(".pf"):
                        file_path = os.path.join(prefetch_dir, file)
                        try:
                            size = os.path.getsize(file_path)
                            os.remove(file_path)
                            cleaned_count += 1
                            total_size += size
                        except Exception:
                            continue
            except Exception:
                pass

        return cleaned_count, total_size

    def clean_recycle_bin(self):
        cleaned_count = 0
        total_size = 0

        drives = ["C:\\", "D:\\", "E:\\", "F:\\"]

        for drive in drives:
            if not os.path.exists(drive):
                continue

            recycle_path = os.path.join(drive, "$Recycle.Bin")
            if os.path.exists(recycle_path):
                try:
                    for root, dirs, files in os.walk(recycle_path):
                        for file in files:
                            file_path = os.path.join(root, file)
                            try:
                                size = os.path.getsize(file_path)
                                os.remove(file_path)
                                cleaned_count += 1
                                total_size += size
                            except Exception:
                                continue
                except Exception:
                    continue

        return cleaned_count, total_size

    def clean_all(self):
        results = {
            "temp_files": self.clean_temp_files(),
            "browser_cache": self.clean_browser_cache(),
            "system_logs": self.clean_system_logs(),
            "recycle_bin": self.clean_recycle_bin()
        }

        total_files = sum(r[0] for r in results.values())
        total_size = sum(r[1] for r in results.values())

        return total_files, total_size, results
