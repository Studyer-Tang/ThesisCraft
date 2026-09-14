"""Connect to an existing visible document, including secondary Office instances."""

import logging
from pathlib import Path

log = logging.getLogger(__name__)


def application_host(application):
    """WPS can report 'Microsoft Word'; prefer the actual installation path."""
    try:
        folder = Path(str(application.Path))
        for executable, host in [('wps.exe', 'wps'), ('WINWORD.EXE', 'word')]:
            if (folder / executable).is_file():
                return host
    except Exception:
        pass
    name = str(application.Name).casefold()
    if any(token in name for token in ('wps', 'kingsoft', '金山')):
        return 'wps'
    if 'microsoft word' in name:
        return 'word'
    return None


def choose_application(candidates, host, window_order, document_path=None, window_root=None):
    """Rank actual visible document windows ahead of the global active object."""
    requested = str(Path(document_path).resolve()).casefold() if document_path else None
    matches = []
    for application in candidates:
        try:
            if application_host(application) != host:
                continue
            if not application.Visible or not application.Documents.Count:
                continue
            windows = [int(window.Hwnd) & 0xffffffff for window in application.Windows]
            if window_root:
                windows = [window_root(hwnd) or hwnd for hwnd in windows]
            ranks = [window_order.index(hwnd) for hwnd in windows if hwnd in window_order]
            if not ranks:
                continue
            if requested and not any(
                str(Path(str(doc.FullName)).resolve()).casefold() == requested
                for doc in application.Documents
            ):
                continue
            matches.append((min(ranks), application))
        except Exception:
            log.debug('Skipping unavailable Office instance', exc_info=True)
    if not matches:
        label = 'Word' if host == 'word' else 'WPS'
        raise RuntimeError(f'请先在桌面版 {label} 中打开文档，再点击“连接当前文档”。')
    return min(matches, key=lambda item: item[0])[1]


def connect_application(host, document_path=None):
    import pythoncom
    import win32com.client
    import win32gui

    windows = []
    win32gui.EnumWindows(
        lambda hwnd, _: windows.append(hwnd) if win32gui.IsWindowVisible(hwnd) else None,
        None,
    )
    candidates = []
    rot = pythoncom.GetRunningObjectTable()
    context = pythoncom.CreateBindCtx(0)
    for moniker in rot.EnumRunning():
        try:
            name = moniker.GetDisplayName(context, None)
            if not name.casefold().endswith(('.docx', '.docm', '.doc', '.wps', '.wpt', '.dotx', '.dotm')):
                continue
            # ROT exposes IUnknown; Dispatch needs its IDispatch interface.
            dispatch = rot.GetObject(moniker).QueryInterface(pythoncom.IID_IDispatch)
            document = win32com.client.Dispatch(dispatch)
            candidates.append(document.Application)
        except Exception:
            log.debug('Skipping unavailable running document', exc_info=True)
    try:
        candidates.append(win32com.client.GetActiveObject(
            'Word.Application' if host == 'word' else 'KWPS.Application'
        ))
    except pythoncom.com_error:
        pass
    return choose_application(candidates, host, windows, document_path,
                              window_root=lambda hwnd: win32gui.GetAncestor(hwnd, 2))


def is_busy_error(error):
    code = getattr(error, 'hresult', error.args[0] if error.args else None)
    return code in (-2147418111, -2147417846)  # Rejected call / retry later.
