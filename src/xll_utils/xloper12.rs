#![allow(non_snake_case, non_camel_case_types, non_upper_case_globals)]

use super::{xlcall::*, XlErr};

#[derive(Debug)]
#[repr(u32)]
pub enum XLType {
    Num = xltypeNum,
    Str = xltypeStr, // Length-prefixed UTF-16 string. Heap allocated.
    Bool = xltypeBool,
    Ref = xltypeRef, // Reference to multiple ranges. Heap allocated
    Err = xltypeErr,
    Flow = xltypeFlow,
    Multi = xltypeMulti, // Heap allocated
    Missing = xltypeMissing,
    Nil = xltypeNil,
    SRef = xltypeSRef, // Reference to a single range. Not heap allocated.
    Int = xltypeInt,
    BigData = xltypeBigData,
    Unknown(u32),
}

impl xloper12 {
    const fn xltype(&self) -> XLType {
        match self.xltype & !(xlbitDLLFree | xlbitXLFree) {
            xltypeNum => XLType::Num,
            xltypeStr => XLType::Str,
            xltypeBool => XLType::Bool,
            xltypeRef => XLType::Ref,
            xltypeErr => XLType::Err,
            xltypeFlow => XLType::Flow,
            xltypeMulti => XLType::Multi,
            xltypeMissing => XLType::Missing,
            xltypeNil => XLType::Nil,
            xltypeSRef => XLType::SRef,
            xltypeInt => XLType::Int,
            xltypeBigData => XLType::BigData,
            n => XLType::Unknown(n),
        }
    }
}

/// A wrapper around XLOPER12 that automatically frees allocated resources when dropped
#[repr(transparent)]
pub(crate) struct XLOper12(XLOPER12);

impl XLOper12 {
    pub const fn xltype(&self) -> XLType {
        self.0.xltype()
    }

    /// Construct an XLOper12 of type xltypeNil
    pub const fn empty() -> Self {
        XLOper12(XLOPER12 { xltype: xltypeNil, val: xloper12__bindgen_ty_1 { w: 0 } })
    }

    /// Construct an XLOper12 of xltypeMissing. Used to represent arguments that are missing in the invocation of a UDF.
    pub const fn missing() -> Self {
        XLOper12(XLOPER12 { xltype: xltypeMissing, val: xloper12__bindgen_ty_1 { w: 0 } })
    }

    /// Construct an XLOper12 containing an error.
    pub const fn error(xlerr: XlErr) -> Self {
        XLOper12(XLOPER12 {
            xltype: xltypeErr,
            val: xloper12__bindgen_ty_1 { err: xlerr as i32 },
        })
    }

    pub const fn is_err(&self, xlerr: u32) -> bool {
        match self.xltype() {
            XLType::Err => (unsafe { self.0.val.err }) as u32 == xlerr,
            _ => false,
        }
    }

    pub const fn as_mut_xloper12(&mut self) -> &mut XLOPER12 {
        &mut self.0
    }

    pub const fn as_lpxloper12(&self) -> LPXLOPER12 {
        (&self.0) as *const xloper12 as LPXLOPER12
    }

    // Construct an XLOper12 from an LPXLOPER12 without taking ownership. Intentionally NOT public
    // because creating an XLOper12 without taking ownership has potential to be misused
    const fn from_lpxloper12(xloper: LPXLOPER12) -> Self {
        let mut result = Self(unsafe { *xloper });
        result.0.xltype &= !(xlbitDLLFree | xlbitXLFree); // no ownership bits
        result
    }
}

impl Clone for XLOper12 {
    fn clone(&self) -> Self {
        // Create a DLL-owned clone
        let mut clone = XLOper12(self.0);
        clone.0.xltype &= !xlbitXLFree;
        clone.0.xltype |= xlbitDLLFree;

        match clone.xltype() {
            XLType::Str => unsafe {
                let ptr = clone.0.val.str_;
                let len = *ptr as usize + 1;
                let mut s: Box<[u16]> = Box::from(std::slice::from_raw_parts(ptr, len));
                clone.0.val.str_ = s.as_mut_ptr();
                std::mem::forget(s);
            },
            XLType::Multi => unsafe {
                let ptr = self.0.val.array.lparray as *mut XLOper12;
                let len = (self.0.val.array.rows * self.0.val.array.columns) as usize;
                let mut s: Box<[XLOper12]> =
                    Box::from(std::slice::from_raw_parts(ptr, len));
                clone.0.val.array.lparray = s.as_mut_ptr() as LPXLOPER12;
                std::mem::forget(s);
            },
            _ => {}
        }

        clone
    }
}

impl Drop for XLOper12 {
    fn drop(&mut self) {
        match self.0.xltype & xlbitXLFree {
            0 => {
                const DLL_ALLOCATED_STRING: u32 = xlbitDLLFree | xltypeStr;
                const DLL_ALLOCATED_MULTI: u32 = xlbitDLLFree | xltypeMulti;
                const DLL_ALLOCATED_REF: u32 = xlbitDLLFree | xltypeRef;
                match self.0.xltype {
                    DLL_ALLOCATED_STRING => unsafe {
                        let ptr = self.0.val.str_;
                        let len = *ptr as usize + 1;
                        drop(Box::<[u16]>::from_raw(std::slice::from_raw_parts_mut(
                            ptr, len,
                        )));
                    },
                    DLL_ALLOCATED_MULTI => unsafe {
                        let ptr = self.0.val.array.lparray as *mut XLOper12;
                        let len =
                            (self.0.val.array.rows * self.0.val.array.columns) as usize;
                        drop(Box::<[XLOper12]>::from_raw(
                            std::slice::from_raw_parts_mut(ptr, len),
                        ));
                    },
                    // We don't currently create values of type xltypeRef, so this should never raise
                    DLL_ALLOCATED_REF => unimplemented!(),
                    _ => {}
                }
            }
            _nonzero_ => {
                crate::xll_utils::excel_free(&mut self.0);
            }
        }
    }
}

#[no_mangle]
pub extern "system" fn xlAutoFree12(ptr: LPXLOPER12) {
    drop(unsafe { Box::<XLOPER12>::from_raw(ptr) });
}

impl From<XLOper12> for LPXLOPER12 {
    fn from(XLOper12(inner): XLOper12) -> LPXLOPER12 {
        Box::<XLOPER12>::into_raw(Box::new(inner))
    }
}

#[derive(Debug)]
pub enum ToStringError {
    Utf16Error(std::string::FromUtf16Error),
    InvalidType(XLType),
}
impl std::error::Error for ToStringError {}
impl std::fmt::Display for ToStringError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Debug::fmt(self, f)
    }
}

impl From<std::string::FromUtf16Error> for ToStringError {
    fn from(e: std::string::FromUtf16Error) -> Self {
        Self::Utf16Error(e)
    }
}

impl TryFrom<&xloper12> for String {
    type Error = ToStringError;

    fn try_from(v: &xloper12) -> Result<String, Self::Error> {
        match v.xltype() {
      | XLType::Num => Ok(unsafe { v.val.num }.to_string()),
      | XLType::Int => Ok(unsafe { v.val.w }.to_string()),
      | XLType::Str => {
        let bytes = unsafe {
          let ptr: *const u16 = v.val.str_;
          let len = *ptr.offset(0) as usize;
          std::slice::from_raw_parts(ptr.offset(1), len)
        };
        String::from_utf16(bytes).map_err(ToStringError::from)
      }
      | XLType::Multi => unsafe {
        let p = v.val.array.lparray;
        String::try_from(&XLOper12::from_lpxloper12(p.offset(0)))
      },
      | XLType::Bool => Ok(unsafe { v.val.xbool == 1 }.to_string()),
      | typ @ (
        | XLType::Nil
        | XLType::Ref // TODO: Consider handling this case
        | XLType::SRef // TODO: Consider handling this case
        | XLType::Err
        | XLType::Flow
        | XLType::Missing
        | XLType::BigData
        | XLType::Unknown(_)
      ) => Err(ToStringError::InvalidType(typ)),
    }
    }
}

impl<'a> TryFrom<&'a XLOper12> for String {
    type Error = <String as TryFrom<&'a xloper12>>::Error;

    fn try_from(v: &XLOper12) -> Result<String, Self::Error> {
        String::try_from(&v.0)
    }
}

impl TryFrom<&xloper12> for f64 {
    type Error = ();

    fn try_from(v: &xloper12) -> Result<f64, ()> {
        match v.xltype() {
            XLType::Num => Ok(unsafe { v.val.num }),
            XLType::Int => Ok(unsafe { v.val.w as f64 }),
            XLType::Str => Err(()),
            XLType::Bool => Ok((unsafe { v.val.xbool == 1 }) as i64 as f64),
            XLType::Multi => unsafe {
                let p = v.val.array.lparray;
                f64::try_from(&*p.offset(0))
            },
            XLType::Ref | XLType::SRef => Err(()), // CR-soon wgross: Consider handling these cases
            XLType::Nil
            | XLType::Err
            | XLType::Flow
            | XLType::Missing
            | XLType::BigData
            | XLType::Unknown(_) => Err(()),
        }
    }
}

impl<'a> TryFrom<&'a XLOper12> for f64 {
    type Error = <f64 as TryFrom<&'a xloper12>>::Error;

    fn try_from(v: &'a XLOper12) -> Result<f64, Self::Error> {
        f64::try_from(&v.0)
    }
}

impl TryFrom<&xloper12> for i64 {
    type Error = ();

    fn try_from(v: &xloper12) -> Result<i64, ()> {
        match v.xltype() {
            XLType::Num => Ok(unsafe { v.val.num as i64 }),
            XLType::Int => Ok(unsafe { v.val.w.into() }),
            XLType::Str => Err(()),
            XLType::Bool => Ok((unsafe { v.val.xbool == 1 }) as i64),
            XLType::Multi => unsafe {
                let p = v.val.array.lparray;
                i64::try_from(&*p.offset(0))
            },
            XLType::Ref | XLType::SRef => Err(()), // TODO: Consider handling these cases
            XLType::Nil
            | XLType::Err
            | XLType::Flow
            | XLType::Missing
            | XLType::BigData
            | XLType::Unknown(_) => Err(()),
        }
    }
}

impl TryFrom<&XLOper12> for i64 {
    type Error = ();

    fn try_from(v: &XLOper12) -> Result<i64, ()> {
        i64::try_from(&v.0)
    }
}

impl TryFrom<&xloper12> for bool {
    type Error = ();

    fn try_from(v: &xloper12) -> Result<bool, Self::Error> {
        match v.xltype() {
            XLType::Num => Ok(unsafe { v.val.num != 0.0 }),
            XLType::Int => Ok(unsafe { v.val.w != 0 }),
            XLType::Str => Err(()),
            XLType::Bool => Ok(unsafe { v.val.xbool != 0 }),
            XLType::Multi => unsafe {
                let p = v.val.array.lparray;
                bool::try_from(&*p.offset(0))
            },
            XLType::Ref | XLType::SRef => Err(()), // TODO: Consider handling these cases
            XLType::Nil
            | XLType::Err
            | XLType::Flow
            | XLType::Missing
            | XLType::BigData
            | XLType::Unknown(_) => Err(()),
        }
    }
}

impl TryFrom<&XLOper12> for bool {
    type Error = ();

    fn try_from(v: &XLOper12) -> Result<bool, Self::Error> {
        bool::try_from(&v.0)
    }
}

impl From<&xloper12> for netidx::subscriber::Value {
    fn from(v: &xloper12) -> netidx::subscriber::Value {
        use netidx::subscriber::Value;
        match v.xltype() {
            XLType::Missing | XLType::Nil => Value::Null,
            XLType::Num => Value::F64(unsafe { v.val.num }),
            XLType::Int => Value::I64(unsafe { v.val.w.into() }),
            XLType::Str => {
                let bytes = unsafe {
                    let ptr: *const u16 = v.val.str_;
                    let len = *ptr.offset(0) as usize;
                    std::slice::from_raw_parts(ptr.offset(1), len)
                };
                match String::from_utf16(bytes) {
                    Ok(s) => Value::String(s.into()),
                    Err(e) => Value::Error(e.to_string().into()),
                }
            }
            XLType::Bool => unsafe { v.val.xbool != 0 }.into(),
            XLType::Err => match unsafe { v.val.err } as u32 {
                xlerrNull => Value::Error("#NULL!".into()),
                xlerrDiv0 => Value::Error("#DIV/0!".into()),
                xlerrValue => Value::Error("#VALUE!".into()),
                xlerrRef => Value::Error("#REF!".into()),
                xlerrName => Value::Error("#NAME?".into()),
                xlerrNum => Value::Error("#NUM!".into()),
                xlerrNA => Value::Error("#N/A".into()),
                xlerrGettingData => Value::Error("#GETTING_DATA".into()),
                code => Value::Error(format!("#ERR{code}").into()),
            },
            XLType::Multi => unsafe {
                let p = v.val.array.lparray;
                netidx::subscriber::Value::from(&*p.offset(0))
            },
            XLType::Ref => Value::Error(format!("#UNSUPPORTED_REF").into()),
            XLType::Flow => Value::Error(format!("#UNSUPPORTED_FLOW").into()),
            XLType::SRef => Value::Error(format!("#UNSUPPORTED_SREF").into()),
            XLType::BigData => Value::Error(format!("#UNSUPPORTED_BIGDATA").into()),
            XLType::Unknown(typ) => Value::Error(format!("#UNKNOWN{typ}").into()),
        }
    }
}

impl From<f64> for XLOper12 {
    fn from(v: f64) -> Self {
        use std::num::FpCategory::*;
        match v.classify() {
            Nan | Infinite => XLOper12::error(XlErr::NA),
            Zero | Subnormal | Normal => XLOper12(XLOPER12 {
                xltype: xltypeNum,
                val: xloper12__bindgen_ty_1 { num: v },
            }),
        }
    }
}

impl From<bool> for XLOper12 {
    fn from(v: bool) -> XLOper12 {
        XLOper12(XLOPER12 {
            xltype: xltypeBool,
            val: xloper12__bindgen_ty_1 { xbool: v as i32 },
        })
    }
}

/// Construct an XLOper12 containing an int (i32)
impl From<i32> for XLOper12 {
    fn from(w: i32) -> Self {
        XLOper12(XLOPER12 { xltype: xltypeInt, val: xloper12__bindgen_ty_1 { w } })
    }
}

/// XLOper12 string length is limited to 32767 characters. Returns xlerrValue if a string longer than this is provided.
impl From<&str> for XLOper12 {
    fn from(s: &str) -> Self {
        let mut wstr: Box<[u16]> = [0].into_iter().chain(s.encode_utf16()).collect();
        match wstr.len() <= 32768 {
            true => {
                wstr[0] = wstr.len() as u16 - 1;
                let str_ = wstr.as_mut_ptr();
                std::mem::forget(wstr);
                XLOper12(XLOPER12 {
                    xltype: xltypeStr | xlbitDLLFree,
                    val: xloper12__bindgen_ty_1 { str_ },
                })
            }
            false => XLOper12::error(XlErr::Value),
        }
    }
}

impl From<String> for XLOper12 {
    fn from(s: String) -> XLOper12 {
        s.as_str().into()
    }
}

impl std::fmt::Display for XLOper12 {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match self.xltype() {
            XLType::Err => match unsafe { self.0.val.err } as u32 {
                xlerrNull => write!(f, "#NULL!"),
                xlerrDiv0 => write!(f, "#DIV/0!"),
                xlerrValue => write!(f, "#VALUE!"),
                xlerrRef => write!(f, "#REF!"),
                xlerrName => write!(f, "#NAME?"),
                xlerrNum => write!(f, "#NUM!"),
                xlerrNA => write!(f, "#N/A"),
                xlerrGettingData => write!(f, "#GETTING_DATA"),
                code => write!(f, "#ERR{code}"),
            },
            XLType::Int => write!(f, "{}", unsafe { self.0.val.w }),
            XLType::Missing => write!(f, "#MISSING"),
            XLType::Multi => write!(f, "#MULTI"),
            XLType::Nil => write!(f, "#NIL"),
            XLType::Num => write!(f, "{}", unsafe { self.0.val.num }),
            XLType::Bool => write!(f, "{}", unsafe { self.0.val.xbool != 0 }),
            XLType::Str => match String::try_from(&self.clone()) {
                Ok(s) => write!(f, "{}", s),
                Err(e) => write!(f, "#STRING_ERR: {}", e),
            },
            XLType::SRef => {
                write!(
                    f,
                    "Sref:({},{}) -> ({},{})",
                    unsafe { self.0.val.sref.ref_.rwFirst },
                    unsafe { self.0.val.sref.ref_.colFirst },
                    unsafe { self.0.val.sref.ref_.rwLast },
                    unsafe { self.0.val.sref.ref_.colLast }
                )
            }
            XLType::Ref => write!(f, "#REF"),
            XLType::Flow => write!(f, "#FLOW"),
            XLType::BigData => write!(f, "#BIG_DATA"),
            XLType::Unknown(typ) => write!(f, "#UNKNOWN{typ}"),
        }
    }
}

#[macro_export]
macro_rules! xloper12_const_string {
  ($s:expr) => {{
    const __STRING: &'static str = $s;

    const __LEN: usize = 1 + __STRING.len();

    const __BUF: [u16; __LEN] = {
      let mut result = [0; __LEN];

      let mut i: usize = 1; // Skip the first [u16]; we will store the UTF16 string length there

      let mut iterator = $crate::xloper12::__private::CodePointIterator::new(__STRING.as_bytes());

      while let Some((next, mut code)) = iterator.next() {
        iterator = next;

        if (code & 0xFFFF) == code {
          result[i] = code as u16;
          i += 1;
        } else {
          // Supplementary planes break into surrogates.
          code -= 0x1_0000;
          result[i] = 0xD800 | ((code >> 10) as u16);
          result[i + 1] = 0xDC00 | ((code as u16) & 0x3FF);
          i += 2;
        }
      }
      if i > 32768 {
        panic!("string too long (encoded UTF-16 is limited to 32767 characters")
      }

      result[0] = i as u16 - 1;

      result
    };

    $crate::xloper12::__private::xloper12_from_const_utf16(&__BUF)
  }};
}

#[doc(hidden)]
pub(crate) mod __private {
    use super::*;

    pub const fn xloper12_from_const_utf16(s: &'static [u16]) -> XLOper12 {
        XLOper12(XLOPER12 {
            xltype: xltypeStr, /* This is a pointer to a const slice, so we set neither xlbitdllfree NOR xlbitxlfree */
            val: xloper12__bindgen_ty_1 { str_: s.as_ptr() as *mut u16 },
        })
    }

    const CONT_MASK: u8 = 0b0011_1111;

    const fn utf8_first_byte(byte: u8, width: u32) -> u32 {
        (byte & (0x7F >> width)) as u32
    }

    const fn utf8_acc_cont_byte(ch: u32, byte: u8) -> u32 {
        (ch << 6) | (byte & CONT_MASK) as u32
    }

    pub struct CodePointIterator<'a> {
        buffer: &'a [u8],

        offset: usize,
    }

    impl<'a> CodePointIterator<'a> {
        pub const fn new(buffer: &'a [u8]) -> Self {
            Self::new_with_offset(buffer, 0)
        }

        pub const fn new_with_offset(buffer: &'a [u8], offset: usize) -> Self {
            Self { buffer, offset }
        }

        pub const fn next(self) -> Option<(Self, u32)> {
            if let Some((codepont, num_utf8_bytes)) =
                next_code_point(self.buffer, self.offset)
            {
                Some((
                    Self::new_with_offset(self.buffer, self.offset + num_utf8_bytes),
                    codepont,
                ))
            } else {
                None
            }
        }
    }

    /// Adapted from Rust core (https://github.com/rust-lang/rust/blob/7e2032390cf34f3ffa726b7bd890141e2684ba63/library/core/src/str/validations.rs#L40-L68).
    const fn next_code_point(bytes: &[u8], start: usize) -> Option<(u32, usize)> {
        const fn get_or(slice: &[u8], index: usize, default: u8) -> u8 {
            if slice.len() > index {
                slice[index]
            } else {
                default
            }
        }

        if bytes.len() == start {
            return None;
        }

        let mut num_bytes = 1;

        let x = bytes[start + 0];

        if x < 128 {
            return Some((x as u32, num_bytes));
        }

        let init = utf8_first_byte(x, 2);

        let y = get_or(bytes, start + 1, 0);

        if y != 0 {
            num_bytes += 1;
        }

        let mut ch = utf8_acc_cont_byte(init, y);

        if x >= 0xE0 {
            let z = get_or(bytes, start + 2, 0);
            if z != 0 {
                num_bytes += 1;
            }
            let y_z = utf8_acc_cont_byte((y & CONT_MASK) as u32, z);
            ch = init << 12 | y_z;
            if x >= 0xF0 {
                let w = get_or(bytes, start + 3, 0);
                if w != 0 {
                    num_bytes += 1;
                }
                ch = (init & 7) << 18 | utf8_acc_cont_byte(y_z, w);
            }
        }
        Some((ch, num_bytes))
    }
}
