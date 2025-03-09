#[doc(hidden)]
#[allow(unused)]
pub mod __for_macros {
    pub use ::com::{sys::LSTATUS, CLSID, IID};
    use std::ffi::CString;
    pub use windows::{
        core::HRESULT,
        Win32::{
            Foundation::{E_FAIL, HMODULE},
            System::Registry::{HKEY, HKEY_USERS},
        },
    };
    use windows::{
        core::PCSTR,
        Win32::{
            Foundation::{ERROR_SUCCESS, S_OK},
            Security::SECURITY_ATTRIBUTES,
            System::{
                LibraryLoader::GetModuleFileNameA,
                Ole::SELFREG_E_CLASS,
                Registry::{
                    RegCloseKey, RegCreateKeyExA, RegDeleteKeyA, RegSetValueExA,
                    KEY_ALL_ACCESS, REG_OPTION_NON_VOLATILE, REG_SZ,
                },
            },
        },
    };

    use super::*;

    pub struct RegistryKeyInfo {
        pub root: HKEY,
        pub key_path: CString,
        pub key_value_name: CString,
        pub key_value_data: String,
    }

    fn create_class_key(key_info: &RegistryKeyInfo) -> Result<HKEY, HRESULT> {
        let mut hk_result = HKEY::default();
        let lp_class = PCSTR::null();
        match unsafe {
            RegCreateKeyExA(
                key_info.root,
                PCSTR(key_info.key_path.as_ptr() as *const u8),
                None,
                lp_class,
                REG_OPTION_NON_VOLATILE,
                KEY_ALL_ACCESS,
                None,
                std::ptr::addr_of_mut!(hk_result),
                None,
            )
        } {
            result if result.is_ok() => Ok(hk_result),
            result => Err(result.to_hresult()),
        }
    }

    fn set_class_key(
        key_handle: HKEY,
        key_info: &RegistryKeyInfo,
    ) -> Result<(), HRESULT> {
        match unsafe {
            RegSetValueExA(
                key_handle,
                PCSTR(key_info.key_value_name.as_ptr() as *const u8),
                None,
                REG_SZ,
                Some(key_info.key_value_data.as_bytes()),
            )
        } {
            result if result.is_ok() => Ok(()),
            result => Err(result.to_hresult()),
        }
    }

    fn add_class_key(key_info: &RegistryKeyInfo) -> Result<(), HRESULT> {
        match create_class_key(key_info) {
            Err(e) => return Err(e),
            Ok(key_handle) => {
                let result = set_class_key(key_handle, key_info);
                unsafe { RegCloseKey(key_handle) };
                result
            }
        }
    }

    fn remove_class_key(key_info: &RegistryKeyInfo) -> windows::core::Result<()> {
        unsafe {
            RegDeleteKeyA(key_info.root, PCSTR(key_info.key_path.as_ptr() as *const u8))
                .ok()
        }
    }

    pub unsafe fn get_dll_file_path(hmodule: HMODULE) -> String {
        const MAX_FILE_PATH_LENGTH: usize = 260;

        let mut path = [0u8; MAX_FILE_PATH_LENGTH];

        let len = GetModuleFileNameA(Some(hmodule), &mut path);

        String::from_utf8(path[..len as usize].to_vec()).unwrap()
    }

    pub fn server_name_key_path(class_name: &str) -> String {
        format!("{}\\CLSID", class_name)
    }

    pub fn class_key_path(clsid: ::com::CLSID) -> String {
        format!("CLSID\\{{{}}}", clsid)
    }

    pub fn class_inproc_key_path(clsid: ::com::CLSID) -> String {
        format!("CLSID\\{{{}}}\\InprocServer32", clsid)
    }

    #[doc(hidden)]
    pub fn register_keys(registry_keys_to_add: &[RegistryKeyInfo]) -> HRESULT {
        for key_info in registry_keys_to_add.iter() {
            match add_class_key(key_info) {
                Ok(()) => {}
                Err(hr) => return SELFREG_E_CLASS,
            }
        }

        S_OK
    }

    #[doc(hidden)]
    pub fn unregister_keys(registry_keys_to_remove: &[RegistryKeyInfo]) -> HRESULT {
        let mut hr = S_OK;
        for key_info in registry_keys_to_remove.iter() {
            match remove_class_key(key_info) {
                Ok(()) => {}
                Err(_) => hr = SELFREG_E_CLASS,
            }
        }

        hr
    }

    #[inline]
    pub fn dll_register_server(relevant_keys: &mut Vec<RegistryKeyInfo>) -> HRESULT {
        let hr = register_keys(relevant_keys);
        if hr.is_err() {
            relevant_keys.reverse();
            unregister_keys(relevant_keys);
        }

        hr
    }

    #[inline]
    pub fn dll_unregister_server(relevant_keys: &mut Vec<RegistryKeyInfo>) -> HRESULT {
        relevant_keys.reverse();
        unregister_keys(relevant_keys)
    }

    pub fn get_current_user_sid() -> anyhow::Result<String> {
        use ::windows::{
            core::Owned,
            Win32::{
                Foundation::{CloseHandle, LocalFree, HANDLE, HLOCAL},
                Security::{
                    Authorization::ConvertSidToStringSidA, GetTokenInformation,
                    TokenUser, TOKEN_QUERY, TOKEN_USER,
                },
                System::{
                    Memory::{LocalAlloc, LPTR},
                    Threading::{
                        GetCurrentProcess, GetCurrentThread, OpenProcessToken,
                        OpenThreadToken,
                    },
                },
            },
        };

        let mut access_token: HANDLE = HANDLE(core::ptr::null_mut());
        unsafe {
            match OpenThreadToken(
                GetCurrentThread(),
                TOKEN_QUERY,
                true,
                &mut access_token as *mut HANDLE,
            ) {
                Ok(()) => {}
                Err(_) => OpenProcessToken(
                    GetCurrentProcess(),
                    TOKEN_QUERY,
                    &mut access_token as *mut HANDLE,
                )?,
            }
        }

        let mut bytes_required = 0;
        let _ = unsafe {
            GetTokenInformation(access_token, TokenUser, None, 0, &mut bytes_required)
        };
        let retval = match unsafe { LocalAlloc(LPTR, bytes_required as usize) } {
            Ok(buf) => {
                // Wrap the allocation in [Owned] so it will be automatically deallocated when it leaves scope
                let buf = unsafe { Owned::new(buf) };
                match unsafe {
                    GetTokenInformation(
                        access_token,
                        TokenUser,
                        Some(buf.0 as *mut _),
                        bytes_required,
                        &mut bytes_required,
                    )
                } {
                    Ok(()) => {
                        let user_token = unsafe { &*(buf.0 as *const TOKEN_USER) };
                        let mut pstr: *mut core::ffi::c_char = core::ptr::null_mut();
                        match unsafe {
                            ConvertSidToStringSidA(
                                user_token.User.Sid,
                                &mut pstr as *mut _ as *mut _,
                            )
                        } {
                            Ok(()) => {
                                let cstr = unsafe { core::ffi::CStr::from_ptr(pstr) };
                                let str = cstr
                                    .to_str()
                                    .map(String::from)
                                    .map_err(|e| anyhow::anyhow!(e));
                                unsafe { LocalFree(Some(HLOCAL(pstr as *mut _))) };
                                str
                            }
                            Err(e) => Err(anyhow::anyhow!(e)),
                        }
                    }
                    Err(e) => Err(anyhow::anyhow!(e)),
                }
            }
            Err(e) => Err(anyhow::anyhow!(e)),
        };

        unsafe { CloseHandle(access_token) };
        retval
    }
}

#[macro_export]
macro_rules! register_xll_module {
    (($server_name_one:literal, $class_id_one:ident, $class_type_one:ty), $(($server_name:literal, $class_id:ident, $class_type:ty)),*) => {
        static mut _HMODULE: $crate::xll_utils::__for_macros::HMODULE = $crate::xll_utils::__for_macros::HMODULE(std::ptr::null_mut());
        #[no_mangle]
        unsafe extern "system" fn DllMain(hmodule: $crate::xll_utils::__for_macros::HMODULE, fdw_reason: u32, _reserved: *mut ::core::ffi::c_void) -> i32 {
            const DLL_PROCESS_ATTACH: u32 = 1;
            if fdw_reason == DLL_PROCESS_ATTACH {
                unsafe { _HMODULE = hmodule; }
            }
            1
        }

        #[no_mangle]
        unsafe extern "system" fn DllGetClassObject(class_id: *const $crate::xll_utils::__for_macros::CLSID, iid: *const $crate::xll_utils::__for_macros::IID, result: *mut *mut ::core::ffi::c_void) -> $crate::xll_utils::__for_macros::HRESULT {
            assert!(!class_id.is_null(), "class id passed to DllGetClassObject should never be null");

            let class_id = unsafe { &*class_id };
            if class_id == &$class_id_one {
                let instance = <$class_type_one as ::com::production::Class>::Factory::allocate();
                $crate::xll_utils::__for_macros::HRESULT(instance.QueryInterface(&*iid, result))
            } $(else if class_id == &$class_id {
                let instance = <$class_type_one as ::com::production::Class>::Factory::allocate();
                $crate::xll_utils::__for_macros::HRESULT(instance.QueryInterface(&*iid, result))
            })* else {
                $crate::xll_utils::__for_macros::HRESULT(::com::sys::CLASS_E_CLASSNOTAVAILABLE)
            }
        }

        #[no_mangle]
        extern "system" fn DllRegisterServer() -> $crate::xll_utils::__for_macros::HRESULT {
            match $crate::xll_utils::__for_macros::get_current_user_sid() {
                Err(_) => $crate::xll_utils::__for_macros::E_FAIL,
                Ok(sid) => {
                    let root = $crate::xll_utils::__for_macros::HKEY_USERS;
                    let prefix = format!("{sid}_Classes\\");
                    let mut relevant_keys = get_relevant_registry_keys(root, prefix);
                    $crate::xll_utils::__for_macros::dll_register_server(&mut relevant_keys)
                }
            }
        }

        #[no_mangle]
        extern "system" fn DllUnregisterServer() -> $crate::xll_utils::__for_macros::HRESULT {
            match $crate::xll_utils::__for_macros::get_current_user_sid() {
                Err(_) => $crate::xll_utils::__for_macros::E_FAIL,
                Ok(sid) => {
                    let root = $crate::xll_utils::__for_macros::HKEY_USERS;
                    let prefix = format!("{sid}_Classes\\");
                    let mut relevant_keys = get_relevant_registry_keys(root, prefix);
                    $crate::xll_utils::__for_macros::dll_unregister_server(&mut relevant_keys)
                }
            }
        }

        fn get_relevant_registry_keys(root: $crate::xll_utils::__for_macros::HKEY, prefix: String) -> Vec<$crate::xll_utils::__for_macros::RegistryKeyInfo> {
            use $crate::xll_utils::__for_macros::RegistryKeyInfo;
            let file_path = unsafe { $crate::xll_utils::__for_macros::get_dll_file_path(_HMODULE) };
            vec![
                RegistryKeyInfo {
                    root,
                    key_path: ::std::ffi::CString::new(format!("{prefix}{}", $crate::xll_utils::__for_macros::server_name_key_path($server_name_one))).unwrap(),
                    key_value_name: c"".into(),
                    key_value_data: format!("{{{}}}", $class_id_one)
                },
                RegistryKeyInfo {
                    root,
                    key_path: ::std::ffi::CString::new(format!("{prefix}{}", $crate::xll_utils::__for_macros::class_key_path($class_id_one))).unwrap(),
                    key_value_name: c"".into(),
                    key_value_data: String::from($server_name_one),
                },
                RegistryKeyInfo {
                    root,
                    key_path: ::std::ffi::CString::new(format!("{prefix}{}", $crate::xll_utils::__for_macros::class_inproc_key_path($class_id_one))).unwrap(),
                    key_value_name: c"".into(),
                    key_value_data: file_path,
                },
                $(RegistryKeyInfo {
                    root,
                    key_path: ::std::ffi::CString::new(format!("{prefix}{}", $crate::xll_utils::__for_macros::server_name_key_path($server_name))).unwrap(),
                    key_value_name: c"".into(),
                    key_value_data: format!("{{{}}}", $class_id)
                },
                RegistryKeyInfo {
                    root,
                    key_path: ::std::ffi::CString::new(format!("{prefix}{}", $crate::xll_utils::__for_macros::class_key_path($class_id))).unwrap(),
                    key_value_name: c"".into(),
                    key_value_data: String::from($server_name),
                },
                RegistryKeyInfo {
                    root,
                    key_path: ::std::ffi::CString::new(format!("{prefix}{}", $crate::xll_utils::__for_macros::class_inproc_key_path($class_id))).unwrap(),
                    key_value_name: c"".into(),
                    key_value_data: file_path,
                }),*
            ]
        }
    };
}
