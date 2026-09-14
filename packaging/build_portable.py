"""Bundle the verified Windows executable with simple plugin launchers."""
from pathlib import Path
from importlib.metadata import distribution
import shutil
import subprocess
import sys
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from word_formatter.version import APP_NAME, APP_BINARY_BASENAME, __version__  # noqa: E402


def copy_dependency_notices(output):
    """Ship upstream redistribution notices beside the bundled runtime."""
    notices = output / 'THIRD-PARTY-LICENSES'
    notices.mkdir(exist_ok=True)
    packages = ('python-docx', 'docxcompose', 'bibtexparser', 'pypdf',
                'citeproc-py', 'babel', 'lxml', 'six', 'pyparsing',
                'typing_extensions', 'tkinterdnd2', 'pywin32', 'pillow', 'pyinstaller')
    versions = []
    for name in packages:
        package = distribution(name)
        versions.append(f'{name} {package.version}')
        for file in package.files or []:
            if not file.name.lower().startswith(('license', 'licence', 'copying', 'notice', 'authors')):
                continue
            source = Path(package.locate_file(file))
            if source.is_file():
                destination = notices / name / str(file).replace('..', '_').replace('\\', '/')
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(source, destination)
    base = Path(sys.base_prefix)
    for source in [base / 'LICENSE.txt', *base.glob('tcl/*/license*')]:
        if source.is_file():
            destination = notices / 'python-runtime' / source.relative_to(base)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination)
    (notices / 'VERSIONS.txt').write_text('\n'.join(versions) + '\n', encoding='utf-8')
    rust_root = Path(subprocess.check_output(['rustc', '--print', 'sysroot'], text=True).strip())
    rust_notices = rust_root/'share/doc/rust'
    destination = notices/'rust-standard-library'
    destination.mkdir(exist_ok=True)
    shutil.copy2(rust_notices/'COPYRIGHT-library.html', destination/'COPYRIGHT-library.html')
    shutil.copytree(rust_notices/'licenses', destination/'licenses', dirs_exist_ok=True)
    (destination/'VERSION.txt').write_text(subprocess.check_output(['rustc', '--version'], text=True), encoding='utf-8')


def main():
    name = f'{APP_BINARY_BASENAME}.v{__version__}'
    executable = ROOT/'dist'/f'{name}.exe'
    if not executable.exists():
        executable = ROOT/'release'/f'{name}.exe'
    if not executable.exists():
        raise FileNotFoundError('请先构建 Windows 可执行程序。')
    output = ROOT/'release'/name
    output.mkdir(parents=True,exist_ok=True)
    shutil.copy2(executable,output/f'{name}.exe')
    native = output/'Native-Addin'
    native.mkdir(exist_ok=True)
    for script in ('install.ps1', 'uninstall.ps1'):
        shutil.copy2(ROOT/'addins/native'/script, native/script)
    for arch in ('x86_64','i686'):
        target = native/'bin'/arch
        target.mkdir(parents=True,exist_ok=True)
        shutil.copy2(ROOT/'addins/native/bin'/arch/'ThesisCraft.Addin.dll', target/'ThesisCraft.Addin.dll')
    (output/'Install-Native-Addin.cmd').write_text(
        f'@echo off\ncd /d "%~dp0"\nstart "" "%~dp0{name}.exe" --install-native\n', encoding='ascii')
    (output/'Uninstall-Native-Addin.cmd').write_text(
        f'@echo off\ncd /d "%~dp0"\nstart "" "%~dp0{name}.exe" --uninstall-native\n',encoding='ascii')
    for filename, arguments in [('Paper-Studio.cmd','--academic'),('Desktop.cmd',''),('Word-Plugin.cmd','--office word'),('WPS-Plugin.cmd','--office wps')]:
        (output/filename).write_text(f'@echo off\ncd /d "%~dp0"\nstart "" "{name}.exe" {arguments}\n',encoding='ascii')
    shutil.copy2(ROOT/'LICENSE',output/'LICENSE')
    shutil.copy2(ROOT/'docs/THESIS-GUIDE.md',output/'论文工作台使用说明.md')
    shutil.copy2(ROOT/'word_formatter/academic/styles/NOTICE.md',output/'CSL-NOTICE.md')
    copy_dependency_notices(output)
    (output/'使用说明.txt').write_text(
        APP_NAME+'（学研排版） '+__version__+'\n作者：Study-Tang\n\n'
        'Paper-Studio.cmd：论文工作台\nDesktop.cmd：通用桌面版\nWord-Plugin.cmd：打开 Word 学研排版操作窗口\nWPS-Plugin.cmd：打开 WPS 学研排版操作窗口\n\n'
        '先打开并保存文档，再启动插件。在操作窗口核对文档名，点击“论文设置 / 检查”或“通用文档排版”。\n'
        '保存位置默认与原文件相同；可选择文件夹或恢复默认。已有同名文件自动增加序号；原件不覆盖。\n'
        '在桌面版调整参数并“保存为默认”，插件下次排版使用同一配置。\n'
        'Windows 本机 Word 16.0 和 WPS 12.1.0.28043 已通过按钮、排版和打开结果测试。\n'
        'Install-Native-Addin.cmd：安装原生功能区入口，之后正常启动 Word/WPS 自动出现“学研排版”。\n'
        'Uninstall-Native-Addin.cmd：卸载原生入口。只注册当前用户，无需管理员或开启宏。\n'
        '也可继续使用免注册的 Word-Plugin.cmd / WPS-Plugin.cmd 外部操作面板。\n'
        '可执行程序自带运行环境，插件仍需电脑已安装桌面版 Word/WPS。\n',encoding='utf-8-sig')
    archive = ROOT/'release'/f'{name}.Windows-portable.zip'
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as bundle:
        for path in output.rglob('*'):
            if path.is_file():
                bundle.write(path,path.relative_to(output.parent))
    print(archive)


if __name__ == '__main__':
    main()
