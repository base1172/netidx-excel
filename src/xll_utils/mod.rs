mod registration;

#[doc(hidden)]
pub use registration::__for_macros;

pub(crate) mod udf;

mod xlcall;

pub(crate) mod xloper12;

use windows::{
    core::PCSTR,
    Win32::System::LibraryLoader::{GetModuleHandleA, GetProcAddress},
};

pub(crate) use xlcall::{LPXLOPER12, XLOPER12};
pub(crate) use xloper12::XLOper12;

#[allow(unused)]
#[allow(non_camel_case_types)]
#[repr(u32)]
pub(crate) enum Xlfn {
    xlfCaller = xlcall::xlfCaller,
    xlGetName = xlcall::xlGetName,
    xlfRtd = xlcall::xlfRtd,
    xlfRegister = xlcall::xlfRegister,
    xlAsyncReturn = xlcall::xlAsyncReturn,
}

impl Into<i32> for Xlfn {
    fn into(self) -> i32 {
        self as i32
    }
}

#[allow(unused)]
#[repr(u32)]
pub(crate) enum XlErr {
    Null = xlcall::xlerrNull,
    Div0 = xlcall::xlerrDiv0,
    Value = xlcall::xlerrValue,
    Ref = xlcall::xlerrRef,
    Name = xlcall::xlerrName,
    Num = xlcall::xlerrNum,
    NA = xlcall::xlerrNA,
    GettingData = xlcall::xlerrGettingData,
}

impl Into<u32> for XlErr {
    fn into(self) -> u32 {
        self as u32
    }
}

const EXCEL12ENTRYPT: PCSTR = windows::core::s!("MdCallBack12");
const XLCALL32DLL: PCSTR = windows::core::s!("XLCall32");
const XLCALL32ENTRYPT: PCSTR = windows::core::s!("GetExcel12EntryPt");

type EXCEL12PROC = extern "system" fn(
    xlfn: ::std::os::raw::c_int,
    count: ::std::os::raw::c_int,
    rgpxloper12: *const LPXLOPER12,
    xloper12res: LPXLOPER12,
) -> ::std::os::raw::c_int;

static PEXCEL12: std::sync::LazyLock<isize> = std::sync::LazyLock::new(|| {
    match unsafe { GetModuleHandleA(XLCALL32DLL) } {
        Err(_) => {}
        Ok(xlcall_hmodule) => {
            match unsafe { GetProcAddress(xlcall_hmodule, XLCALL32ENTRYPT) } {
                None => {}
                Some(entry_pt) => unsafe {
                    return entry_pt();
                },
            }
        }
    }

    match unsafe { GetModuleHandleA(PCSTR::null()) } {
        Err(_) => 0,
        Ok(xlcall_hmodule) => {
            match unsafe { GetProcAddress(xlcall_hmodule, EXCEL12ENTRYPT) } {
                None => 0,
                Some(f) => f as isize,
            }
        }
    }
});

fn excel12(xlfn: Xlfn, opers: &mut [XLOper12]) -> Result<XLOper12, XLOper12> {
    let args: Vec<LPXLOPER12> =
        opers.into_iter().map(|oper| oper.as_lpxloper12()).collect();
    let mut result = XLOper12::empty();
    let res = excel12v(xlfn, result.as_mut_xloper12(), &args);
    match res {
        0 => Ok(result),
        _nonzero_ => Err(result),
    }
}

pub(crate) fn excel12v(xlfn: Xlfn, px_res: &mut XLOPER12, opers: &[LPXLOPER12]) -> i32 {
    match *PEXCEL12 {
        0 => xlcall::xlretFailed as i32,
        pexcel12 => unsafe {
            std::mem::transmute::<isize, EXCEL12PROC>(pexcel12)(
                xlfn as i32,
                opers.len() as i32,
                opers.as_ptr(),
                px_res,
            )
        },
    }
}

fn excel_free(xloper: LPXLOPER12) -> i32 {
    match *PEXCEL12 {
        0 => xlcall::xlretFailed as i32,
        pexcel12 => unsafe {
            std::mem::transmute::<isize, EXCEL12PROC>(pexcel12)(
                xlcall::xlFree as i32,
                1,
                &xloper,
                std::ptr::null_mut(),
            )
        },
    }
}
