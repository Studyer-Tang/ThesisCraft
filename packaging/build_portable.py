"""Bundle the verified Windows executable with simple plugin launchers."""
from pathlib import Path
from importlib.metadata import distribution
import shutil
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
    for filename, arguments in [('Paper-Studio.cmd','--academic'),('Desktop.cmd',''),('Word-Plugin.cmd','--office word'),('WPS-Plugin.cmd','--office wps')]:
        (output/filename).write_text(f'@echo off\ncd /d "%~dp0"\nstart "" "{name}.exe" {arguments}\n',encoding='ascii')
    shutil.copy2(ROOT/'LICENSE',output/'LICENSE')
    shutil.copy2(ROOT/'docs/THESIS-GUIDE.md',output/'论文工作台使用说明.md')
    shutil.copy2(ROOT/'word_formatter/academic/styles/NOTICE.md',output/'CSL-NOTICE.md')
    copy_dependency_notices(output)
    (output/'使用说明.txt').write_text(
        APP_NAME+'（学研排版） '+__version__+'\n作者：Study-Tang\n\n'
        'Paper-Studio.cmd：论文工作台\nDesktop.cmd：通用桌面版\nWord-Plugin.cmd：启动 Word 工具栏插件\nWPS-Plugin.cmd：启动 WPS 工具栏插件\n\n'
        '插件启动后需保持后台进程运行。打开并保存文档，点击“排版为新副本”。\n'
        '输出存放在原文件旁，已有同名文件自动增加序号；原件不覆盖。\n'
        '在桌面版调整参数并“保存为默认”，插件下次排版使用同一配置。\n'
        'Windows 本机 Word 16.0 和 WPS 12.1.0.28043 已通过按钮、排版和打开结果测试。\n'
        '当前是外部工具栏插件，不随 Office 自动启动。无需开启宏或注册 DLL。\n'
        '可执行程序自带运行环境，插件仍需电脑已安装桌面版 Word/WPS。\n',encoding='utf-8-sig')
    archive = ROOT/'release'/f'{name}.Windows-portable.zip'
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as bundle:
        for path in output.rglob('*'):
            if path.is_file():
                bundle.write(path,path.relative_to(output.parent))
    print(archive)


if __name__ == '__main__':
    main()
