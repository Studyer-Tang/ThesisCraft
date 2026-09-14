//! Native COM bridge: no CLR, macros, document mutation, or work on the Office UI thread.
use std::os::windows::process::CommandExt;
use std::{
    ffi::c_void,
    ptr,
    sync::atomic::{AtomicU32, Ordering},
};

type Hr = i32;
const OK: Hr = 0;
const NO_INTERFACE: Hr = 0x80004002u32 as i32;
const POINTER: Hr = 0x80004003u32 as i32;
const NOT_IMPL: Hr = 0x80004001u32 as i32;
const UNKNOWN_NAME: Hr = 0x80020006u32 as i32;
const MEMBER_NOT_FOUND: Hr = 0x80020003u32 as i32;

#[repr(C)]
#[derive(PartialEq, Eq)]
pub struct Guid {
    a: u32,
    b: u16,
    c: u16,
    d: [u8; 8],
}
const UNKNOWN: Guid = Guid {
    a: 0,
    b: 0,
    c: 0,
    d: [0xc0, 0, 0, 0, 0, 0, 0, 0x46],
};
const DISPATCH: Guid = Guid {
    a: 0x20400,
    ..UNKNOWN
};
const FACTORY: Guid = Guid { a: 1, ..UNKNOWN };
const EXT: Guid = Guid {
    a: 0xb65ad801,
    b: 0xabaf,
    c: 0x11d0,
    d: [0xbb, 0x8b, 0, 0xa0, 0xc9, 0x0f, 0x27, 0x44],
};
const RIBBON: Guid = Guid {
    a: 0xc0396,
    ..UNKNOWN
};
const CLASS: Guid = Guid {
    a: 0xf93f581a,
    b: 0x74c6,
    c: 0x4728,
    d: [0x94, 0x06, 0x39, 0x2b, 0x68, 0x73, 0xca, 0x1c],
};
static OBJECTS: AtomicU32 = AtomicU32::new(0);
static LOCKS: AtomicU32 = AtomicU32::new(0);

#[link(name = "oleaut32")]
extern "system" {
    fn SysAllocString(value: *const u16) -> *mut u16;
}
#[link(name = "advapi32")]
extern "system" {
    fn RegGetValueW(
        key: isize,
        subkey: *const u16,
        value: *const u16,
        flags: u32,
        kind: *mut u32,
        data: *mut c_void,
        size: *mut u32,
    ) -> i32;
}
#[link(name = "kernel32")]
extern "system" {
    fn GetModuleFileNameW(module: isize, path: *mut u16, size: u32) -> u32;
}
#[link(name = "user32")]
extern "system" {
    fn MessageBoxW(hwnd: isize, text: *const u16, title: *const u16, flags: u32) -> i32;
}

fn wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(Some(0)).collect()
}
unsafe fn bstr(s: &str) -> *mut u16 {
    SysAllocString(wide(s).as_ptr())
}
unsafe fn read_name(s: *const u16) -> String {
    if s.is_null() {
        return String::new();
    }
    let mut n = 0;
    while n < 32768 && *s.add(n) != 0 {
        n += 1;
    }
    String::from_utf16_lossy(std::slice::from_raw_parts(s, n))
}
fn setting(name: &str) -> Option<String> {
    let mut buffer = vec![0u16; 32768];
    let mut size = (buffer.len() * 2) as u32;
    let status = unsafe {
        RegGetValueW(
            0x80000001u32 as i32 as isize,
            wide("Software\\Study-Tang\\ThesisCraft").as_ptr(),
            wide(name).as_ptr(),
            2,
            ptr::null_mut(),
            buffer.as_mut_ptr().cast(),
            &mut size,
        )
    };
    if status != 0 || size < 2 {
        return None;
    }
    Some(String::from_utf16_lossy(
        &buffer[..(size as usize / 2).saturating_sub(1)],
    ))
}
fn launch(settings: bool) {
    let mut module = vec![0u16; 32768];
    let n = unsafe { GetModuleFileNameW(0, module.as_mut_ptr(), module.len() as u32) };
    let host = if String::from_utf16_lossy(&module[..n as usize])
        .to_lowercase()
        .ends_with("wps.exe")
    {
        "wps"
    } else {
        "word"
    };
    let result = (|| -> Result<(), String> {
        let exe = setting("NativeExecutable").ok_or("请重新安装 ThesisCraft 原生插件。")?;
        if !std::path::Path::new(&exe).is_file() {
            return Err("ThesisCraft 程序已移动，请重新安装插件。".into());
        }
        let mut command = std::process::Command::new(exe);
        if setting("NativePython").as_deref() == Some("1") {
            command.args(["-m", "word_formatter"]);
        }
        if settings {
            command.arg("--academic");
        } else {
            command.args(["--office", host]);
        }
        if let Some(project) = setting("ProjectPath") {
            command.current_dir(project);
        }
        command
            .creation_flags(0x08000000)
            .spawn()
            .map_err(|e| e.to_string())?;
        Ok(())
    })();
    if let Err(error) = result {
        unsafe {
            MessageBoxW(0, wide(&error).as_ptr(), wide("学研排版").as_ptr(), 0x40);
        }
    }
}
const XML: &str = r#"<customUI xmlns="http://schemas.microsoft.com/office/2006/01/customui"><ribbon><tabs><tab id="ThesisCraftTab" label="学研排版"><group id="ThesisCraftTools" label="ThesisCraft"><button id="ThesisCraftPanel" label="打开排版面板" size="large" imageMso="FormatPainter" onAction="OpenPanel" screentip="排版当前文档" supertip="打开操作面板，核对文档后快速排版、插入题注、更新目录。只生成新副本。"/><button id="ThesisCraftSettings" label="模板与图片设置" imageMso="FilePageSetup" onAction="OpenSettings"/><button id="ThesisCraftAbout" label="关于" imageMso="Help" onAction="About"/></group></tab></tabs></ribbon></customUI>"#;

#[repr(C)]
struct Variant {
    vt: u16,
    reserved: [u16; 3],
    data: [usize; 2],
}
#[repr(C)]
struct DispatchVTable {
    query: unsafe extern "system" fn(*mut Node, *const Guid, *mut *mut c_void) -> Hr,
    add_ref: unsafe extern "system" fn(*mut Node) -> u32,
    release: unsafe extern "system" fn(*mut Node) -> u32,
    type_count: unsafe extern "system" fn(*mut Node, *mut u32) -> Hr,
    type_info: unsafe extern "system" fn(*mut Node, u32, u32, *mut *mut c_void) -> Hr,
    ids: unsafe extern "system" fn(
        *mut Node,
        *const Guid,
        *const *const u16,
        u32,
        u32,
        *mut i32,
    ) -> Hr,
    invoke: unsafe extern "system" fn(
        *mut Node,
        i32,
        *const Guid,
        u32,
        u16,
        *mut c_void,
        *mut Variant,
        *mut c_void,
        *mut u32,
    ) -> Hr,
}
#[repr(C)]
struct ExtVTable {
    dispatch: DispatchVTable,
    connection:
        unsafe extern "system" fn(*mut Node, *mut c_void, i32, *mut c_void, *mut *mut c_void) -> Hr,
    disconnection: unsafe extern "system" fn(*mut Node, i32, *mut *mut c_void) -> Hr,
    update: unsafe extern "system" fn(*mut Node, *mut *mut c_void) -> Hr,
    startup: unsafe extern "system" fn(*mut Node, *mut *mut c_void) -> Hr,
    shutdown: unsafe extern "system" fn(*mut Node, *mut *mut c_void) -> Hr,
}
#[repr(C)]
struct RibbonVTable {
    dispatch: DispatchVTable,
    ui: unsafe extern "system" fn(*mut Node, *const u16, *mut *mut u16) -> Hr,
}
#[repr(C)]
struct Node {
    table: *const c_void,
    owner: *mut Addin,
}
#[repr(C)]
struct Addin {
    ext: Node,
    ribbon: Node,
    refs: AtomicU32,
}

unsafe extern "system" fn query(this: *mut Node, iid: *const Guid, out: *mut *mut c_void) -> Hr {
    if out.is_null() || iid.is_null() {
        return POINTER;
    }
    *out = ptr::null_mut();
    let owner = (*this).owner;
    if *iid == UNKNOWN || *iid == DISPATCH || *iid == EXT {
        *out = (&mut (*owner).ext as *mut Node).cast();
    } else if *iid == RIBBON {
        *out = (&mut (*owner).ribbon as *mut Node).cast();
    } else {
        return NO_INTERFACE;
    }
    add_ref(this);
    OK
}
unsafe extern "system" fn add_ref(this: *mut Node) -> u32 {
    (*(*this).owner).refs.fetch_add(1, Ordering::Relaxed) + 1
}
unsafe extern "system" fn release(this: *mut Node) -> u32 {
    let owner = (*this).owner;
    let count = (*owner).refs.fetch_sub(1, Ordering::AcqRel) - 1;
    if count == 0 {
        drop(Box::from_raw(owner));
        OBJECTS.fetch_sub(1, Ordering::AcqRel);
    }
    count
}
unsafe extern "system" fn type_count(_: *mut Node, count: *mut u32) -> Hr {
    if count.is_null() {
        return POINTER;
    }
    *count = 0;
    OK
}
unsafe extern "system" fn type_info(_: *mut Node, _: u32, _: u32, out: *mut *mut c_void) -> Hr {
    if !out.is_null() {
        *out = ptr::null_mut();
    }
    NOT_IMPL
}
unsafe extern "system" fn ids(
    _: *mut Node,
    _: *const Guid,
    names: *const *const u16,
    count: u32,
    _: u32,
    out: *mut i32,
) -> Hr {
    if names.is_null() || out.is_null() {
        return POINTER;
    }
    for i in 0..count as usize {
        let id = match read_name(*names.add(i)).to_lowercase().as_str() {
            "getcustomui" | "onconnection" => 1,
            "ondisconnection" => 2,
            "onaddinsupdate" => 3,
            "onstartupcomplete" => 4,
            "onbeginshutdown" => 5,
            "openpanel" => 101,
            "opensettings" => 102,
            "about" => 103,
            _ => -1,
        };
        *out.add(i) = id;
        if id == -1 {
            return UNKNOWN_NAME;
        }
    }
    OK
}
unsafe extern "system" fn invoke(
    this: *mut Node,
    id: i32,
    _: *const Guid,
    _: u32,
    _: u16,
    _: *mut c_void,
    result: *mut Variant,
    _: *mut c_void,
    _: *mut u32,
) -> Hr {
    if !result.is_null() {
        ptr::write_bytes(result, 0, 1);
    }
    if (*this).table == (&RIBBON_TABLE as *const RibbonVTable).cast() && id == 1 {
        if result.is_null() {
            return POINTER;
        }
        (*result).vt = 8;
        (*result).data[0] = bstr(XML) as usize;
        return OK;
    }
    match id {
        1..=5 => OK,
        101 => {
            launch(false);
            OK
        }
        102 => {
            launch(true);
            OK
        }
        103 => {
            MessageBoxW(0,wide("ThesisCraft\n作者 Study-Tang · GPT 辅助开发\n原生功能区入口，排版核心独立运行。原件不覆盖。").as_ptr(),wide("学研排版").as_ptr(),0x40);
            OK
        }
        _ => MEMBER_NOT_FOUND,
    }
}
unsafe extern "system" fn connection(
    _: *mut Node,
    _: *mut c_void,
    _: i32,
    _: *mut c_void,
    _: *mut *mut c_void,
) -> Hr {
    OK
}
unsafe extern "system" fn disconnection(_: *mut Node, _: i32, _: *mut *mut c_void) -> Hr {
    OK
}
unsafe extern "system" fn event(_: *mut Node, _: *mut *mut c_void) -> Hr {
    OK
}
unsafe extern "system" fn ui(_: *mut Node, _: *const u16, out: *mut *mut u16) -> Hr {
    if out.is_null() {
        return POINTER;
    }
    *out = bstr(XML);
    if (*out).is_null() {
        0x8007000eu32 as i32
    } else {
        OK
    }
}
const fn dispatch_table() -> DispatchVTable {
    DispatchVTable {
        query,
        add_ref,
        release,
        type_count,
        type_info,
        ids,
        invoke,
    }
}
static EXT_TABLE: ExtVTable = ExtVTable {
    dispatch: dispatch_table(),
    connection,
    disconnection,
    update: event,
    startup: event,
    shutdown: event,
};
static RIBBON_TABLE: RibbonVTable = RibbonVTable {
    dispatch: dispatch_table(),
    ui,
};

#[repr(C)]
struct FactoryVTable {
    query: unsafe extern "system" fn(*mut Factory, *const Guid, *mut *mut c_void) -> Hr,
    add_ref: unsafe extern "system" fn(*mut Factory) -> u32,
    release: unsafe extern "system" fn(*mut Factory) -> u32,
    create:
        unsafe extern "system" fn(*mut Factory, *mut c_void, *const Guid, *mut *mut c_void) -> Hr,
    lock: unsafe extern "system" fn(*mut Factory, i32) -> Hr,
}
#[repr(C)]
struct Factory {
    table: *const FactoryVTable,
    refs: AtomicU32,
}
unsafe extern "system" fn factory_query(
    this: *mut Factory,
    iid: *const Guid,
    out: *mut *mut c_void,
) -> Hr {
    if iid.is_null() || out.is_null() {
        return POINTER;
    }
    *out = ptr::null_mut();
    if *iid != UNKNOWN && *iid != FACTORY {
        return NO_INTERFACE;
    }
    *out = this.cast();
    factory_add_ref(this);
    OK
}
unsafe extern "system" fn factory_add_ref(this: *mut Factory) -> u32 {
    (*this).refs.fetch_add(1, Ordering::Relaxed) + 1
}
unsafe extern "system" fn factory_release(this: *mut Factory) -> u32 {
    let n = (*this).refs.fetch_sub(1, Ordering::AcqRel) - 1;
    if n == 0 {
        drop(Box::from_raw(this));
        OBJECTS.fetch_sub(1, Ordering::AcqRel);
    }
    n
}
unsafe extern "system" fn create(
    _: *mut Factory,
    outer: *mut c_void,
    iid: *const Guid,
    out: *mut *mut c_void,
) -> Hr {
    if out.is_null() {
        return POINTER;
    }
    *out = ptr::null_mut();
    if !outer.is_null() {
        return 0x80040110u32 as i32;
    }
    let object = Box::into_raw(Box::new(Addin {
        ext: Node {
            table: (&EXT_TABLE as *const ExtVTable).cast(),
            owner: ptr::null_mut(),
        },
        ribbon: Node {
            table: (&RIBBON_TABLE as *const RibbonVTable).cast(),
            owner: ptr::null_mut(),
        },
        refs: AtomicU32::new(1),
    }));
    (*object).ext.owner = object;
    (*object).ribbon.owner = object;
    OBJECTS.fetch_add(1, Ordering::Relaxed);
    let node = &mut (*object).ext as *mut Node;
    let result = query(node, iid, out);
    release(node);
    result
}
unsafe extern "system" fn lock(_: *mut Factory, value: i32) -> Hr {
    if value != 0 {
        LOCKS.fetch_add(1, Ordering::Relaxed);
    } else {
        let _ = LOCKS.fetch_update(Ordering::AcqRel, Ordering::Relaxed, |n| n.checked_sub(1));
    }
    OK
}
static FACTORY_TABLE: FactoryVTable = FactoryVTable {
    query: factory_query,
    add_ref: factory_add_ref,
    release: factory_release,
    create,
    lock,
};

#[no_mangle]
pub unsafe extern "system" fn DllGetClassObject(
    clsid: *const Guid,
    iid: *const Guid,
    out: *mut *mut c_void,
) -> Hr {
    if out.is_null() || clsid.is_null() {
        return POINTER;
    }
    *out = ptr::null_mut();
    if *clsid != CLASS {
        return 0x80040111u32 as i32;
    }
    let factory = Box::into_raw(Box::new(Factory {
        table: &FACTORY_TABLE,
        refs: AtomicU32::new(1),
    }));
    OBJECTS.fetch_add(1, Ordering::Relaxed);
    let result = factory_query(factory, iid, out);
    factory_release(factory);
    result
}
#[no_mangle]
pub extern "system" fn DllCanUnloadNow() -> Hr {
    if OBJECTS.load(Ordering::Acquire) == 0 && LOCKS.load(Ordering::Acquire) == 0 {
        OK
    } else {
        1
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[link(name = "oleaut32")]
    extern "system" {
        fn SysFreeString(value: *mut u16);
    }

    #[test]
    fn native_com_identity_callbacks_and_lifetime() {
        unsafe {
            assert_eq!(DllCanUnloadNow(), OK);
            let mut factory = ptr::null_mut();
            assert_eq!(DllGetClassObject(&CLASS, &FACTORY, &mut factory), OK);
            assert_eq!(DllCanUnloadNow(), 1);
            let mut ext = ptr::null_mut();
            assert_eq!(create(factory.cast(), ptr::null_mut(), &EXT, &mut ext), OK);
            let mut ribbon = ptr::null_mut();
            assert_eq!(query(ext.cast(), &RIBBON, &mut ribbon), OK);
            let mut identity = ptr::null_mut();
            assert_eq!(query(ribbon.cast(), &UNKNOWN, &mut identity), OK);
            assert_eq!(identity, ext);
            let mut xml = ptr::null_mut();
            assert_eq!(ui(ribbon.cast(), ptr::null(), &mut xml), OK);
            assert!(read_name(xml).contains("onAction=\"OpenPanel\""));
            SysFreeString(xml);
            let name = wide("OpenPanel");
            let mut id = 0;
            assert_eq!(ids(ext.cast(), &UNKNOWN, &name.as_ptr(), 1, 0, &mut id), OK);
            assert_eq!(id, 101);
            let missing = wide("NotACommand");
            assert_eq!(
                ids(ext.cast(), &UNKNOWN, &missing.as_ptr(), 1, 0, &mut id),
                UNKNOWN_NAME
            );
            let mut count = 99;
            assert_eq!(type_count(ext.cast(), &mut count), OK);
            assert_eq!(count, 0);
            assert_eq!(
                connection(
                    ext.cast(),
                    ptr::null_mut(),
                    0,
                    ptr::null_mut(),
                    ptr::null_mut()
                ),
                OK
            );
            assert_eq!(disconnection(ext.cast(), 0, ptr::null_mut()), OK);
            release(identity.cast());
            release(ribbon.cast());
            release(ext.cast());
            factory_release(factory.cast());
            assert_eq!(DllCanUnloadNow(), OK);
            let mut out = 1usize as *mut c_void;
            assert_ne!(DllGetClassObject(&UNKNOWN, &FACTORY, &mut out), OK);
            assert!(out.is_null());
            assert_eq!(DllCanUnloadNow(), OK);
        }
    }
}
