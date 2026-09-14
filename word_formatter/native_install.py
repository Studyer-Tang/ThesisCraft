"""Current-user native Ribbon registration, with no administrator or script policy changes."""

import hashlib
import os
from pathlib import Path
import struct
import sys

PROG_ID = 'StudyTang.ThesisCraft'
CLASS_ID = '{F93F581A-74C6-4728-9406-392B6873CA1C}'
SETTINGS = r'Software\Study-Tang\ThesisCraft'
ADDINS = (r'Software\Microsoft\Office\Word\Addins', r'Software\Kingsoft\Office\WPS\Addins')


def validate_binary(data, machine):
    if len(data) < 64 or data[:2] != b'MZ':
        raise ValueError('原生插件 DLL 文件损坏。')
    offset = struct.unpack_from('<I', data, 60)[0]
    if offset + 6 > len(data) or data[offset:offset + 4] != b'PE\0\0':
        raise ValueError('原生插件 DLL 格式不正确。')
    if struct.unpack_from('<H', data, offset + 4)[0] != machine:
        raise ValueError('原生插件 DLL 位数不匹配。')


def _delete_tree(root, path, view):
    import winreg
    try:
        with winreg.OpenKey(root, path, 0, winreg.KEY_READ | winreg.KEY_WRITE | view) as key:
            children = [winreg.EnumKey(key, i) for i in range(winreg.QueryInfoKey(key)[0])]
        for child in children:
            _delete_tree(root, path + '\\' + child, view)
        winreg.DeleteKeyEx(root, path, view, 0)
    except FileNotFoundError:
        pass


def install(binary_dir=None, executable=None, project=None, python=None):
    if os.name != 'nt':
        raise RuntimeError('原生插件仅支持 Windows。')
    import winreg
    frozen = getattr(sys, 'frozen', False)
    project = Path(project or (Path(sys.executable).parent if frozen else Path(__file__).resolve().parents[1])).resolve()
    binary_dir = Path(binary_dir or (project/'Native-Addin/bin' if frozen else project/'addins/native/bin'))
    executable = Path(executable or sys.executable).resolve()
    python = not frozen if python is None else python
    if python and executable.with_name('pythonw.exe').is_file():
        executable = executable.with_name('pythonw.exe')
    if not executable.is_file() or not project.is_dir():
        raise ValueError('排版程序路径不存在。')
    payloads = {}
    for arch, machine in [('x86_64', 0x8664), ('i686', 0x14c)]:
        data = (binary_dir/arch/'ThesisCraft.Addin.dll').read_bytes()
        validate_binary(data, machine)
        payloads[arch] = data
    installed = []
    for arch, view in [('x86_64', winreg.KEY_WOW64_64KEY), ('i686', winreg.KEY_WOW64_32KEY)]:
        data = payloads[arch]
        digest = hashlib.sha256(data).hexdigest()
        folder = Path(os.environ['LOCALAPPDATA'])/'Study-Tang/ThesisCraft/NativeAddin'/digest[:16]/arch
        folder.mkdir(parents=True, exist_ok=True)
        dll = folder/'ThesisCraft.Addin.dll'
        if not dll.exists():
            with dll.open('xb') as stream:
                stream.write(data)
        if hashlib.sha256(dll.read_bytes()).hexdigest() != digest:
            raise ValueError('已安装 DLL 校验不匹配，请重新解压安装包。')

        def write(path, values):
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, path, 0, winreg.KEY_WRITE | view) as key:
                for name, value in values.items():
                    winreg.SetValueEx(key, name, 0, winreg.REG_DWORD if isinstance(value, int) else winreg.REG_SZ, value)

        write(fr'Software\Classes\CLSID\{CLASS_ID}\InprocServer32', {'': str(dll), 'ThreadingModel': 'Apartment'})
        write(fr'Software\Classes\{PROG_ID}\CLSID', {'': CLASS_ID})
        write(fr'Software\Classes\CLSID\{CLASS_ID}\ProgID', {'': PROG_ID})
        write(SETTINGS, {'NativeExecutable': str(executable), 'NativePython': '1' if python else '0', 'ProjectPath': str(project)})
        for location in ADDINS:
            write(location + '\\' + PROG_ID, {'FriendlyName': '学研排版', 'Description': 'ThesisCraft · Study-Tang', 'LoadBehavior': 3})
            _delete_tree(winreg.HKEY_CURRENT_USER, location + r'\StudyTang.WordFormatter', view)
        write(r'Software\Kingsoft\Office\WPS\AddinsWL', {PROG_ID: ''})
        installed.append(str(dll))
    return installed


def uninstall():
    import winreg
    for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
        paths = [fr'Software\Classes\CLSID\{CLASS_ID}', fr'Software\Classes\{PROG_ID}']
        paths += [location + '\\' + PROG_ID for location in ADDINS]
        for path in paths:
            _delete_tree(winreg.HKEY_CURRENT_USER, path, view)
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Kingsoft\Office\WPS\AddinsWL', 0, winreg.KEY_WRITE | view) as key:
                winreg.DeleteValue(key, PROG_ID)
        except FileNotFoundError:
            pass


def main(remove=False):
    from tkinter import messagebox
    from .gui import _create_root
    root, _ = _create_root()
    root.withdraw()
    try:
        if remove:
            uninstall()
            message = '原生入口已卸载，重启 Word/WPS 后生效。个人模板与排版程序已保留。'
        else:
            install()
            message = '安装完成！正常重启 Word/WPS 后，点击“学研排版 → 打开排版面板”。\n无需开启宏；原有文档不会被修改。'
        messagebox.showinfo('学研排版', message, parent=root)
        return 0
    except Exception as error:
        messagebox.showerror('安装未完成', str(error), parent=root)
        return 1
    finally:
        root.destroy()
