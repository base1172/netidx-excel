#[macro_use]
extern crate serde_derive;
mod comglue;
mod server;
mod xll_utils;
use anyhow::Result;
use comglue::{glue::NetidxRTD, interface::CLSID};
use xll_utils::xloper12;
mod setter;

#[no_mangle]
extern "system" fn NetGet(path: xll_utils::LPXLOPER12) -> xll_utils::LPXLOPER12 {
    use xll_utils::*;
    let mut res = XLOper12::error(XlErr::GettingData);
    const CLASS_NAME: XLOper12 = xloper12_const_string!("NetidxRTD");

    match excel12v(
        Xlfn::xlfRtd,
        res.as_mut_xloper12(),
        &[CLASS_NAME.as_lpxloper12(), XLOper12::missing().as_lpxloper12(), path],
    ) {
        0 => {
            res.set_xlfree();
            res.into()
        }
        _nonzero_ => XLOper12::error(XlErr::NA).into(),
    }
}

#[no_mangle]
extern "system" fn NetSet(
    path: *const std::ffi::c_char,
    value: xll_utils::LPXLOPER12,
    ty: *const std::ffi::c_char,
) -> xll_utils::LPXLOPER12 {
    use netidx::subscriber::Value;
    use std::ffi::CStr;
    use xll_utils::*;

    const SET: XLOper12 = xloper12_const_string!("#SET");

    static SETTER: std::sync::LazyLock<Option<setter::Setter>> =
        std::sync::LazyLock::new(|| match setter::Setter::new() {
            Err(e) => {
                log::error!("Error creating Netidx setter: {e}");
                None
            }
            Ok(setter) => Some(setter),
        });

    /// The type of data to publish
    enum NetSetType {
        Auto,
        F64,
        I64,
        Null,
        Time,
        String,
        Bool,
    }

    impl NetSetType {
        fn apply(&self, raw: LPXLOPER12) -> Value {
            match self {
                NetSetType::Auto => Value::from(&unsafe { *raw }),
                NetSetType::F64 => match f64::try_from(&unsafe { *raw }) {
                    Ok(v) => Value::F64(v),
                    Err(()) => Value::Error("#TYPE!".into()),
                },
                NetSetType::I64 => match i64::try_from(&unsafe { *raw }) {
                    Ok(v) => Value::I64(v),
                    Err(()) => Value::Error("#TYPE!".into()),
                },
                NetSetType::Null => Value::Null,
                NetSetType::String => match String::try_from(&unsafe { *raw }) {
                    Ok(s) => s.into(),
                    Err(e) => Value::Error(e.to_string().into()),
                },
                NetSetType::Bool => match bool::try_from(&unsafe { *raw }) {
                    Ok(v) => v.into(),
                    Err(()) => Value::Error("#TYPE!".into()),
                },
                NetSetType::Time => {
                    match f64::try_from(&unsafe { *raw }) {
                        Ok(mut v) => {
                            use chrono::{
                                offset::LocalResult::*, Duration, NaiveDateTime,
                                TimeZone as _,
                            };
                            // Excel's epoch is midnight on 1900-01-00 (i.e., 1899-12-31)
                            const EXCEL_EPOCH: chrono::NaiveDateTime =
                                chrono::NaiveDate::from_ymd_opt(1899, 12, 31)
                                    .expect("never raises")
                                    .and_hms_opt(0, 0, 0)
                                    .expect("never raises");
                            if v < 0.0 {
                                Value::Error("#VALUE!".into())
                            } else {
                                if v > 59.0 {
                                    // Due to a legacy bug, Excel treats Feb 1900 as having 29 days, so we need to subtract 1 for dates above this
                                    // [https://learn.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year]
                                    v -= 1.0;
                                }
                                let date: chrono::NaiveDateTime =
                                    EXCEL_EPOCH + chrono::Duration::days(v as i64);
                                let milliseconds =
                                    (v.fract() * 86_400.0 * 1_000.0) as i64; // convert to milliseconds * 24.0 * 60.0 * 60.0 * 1000
                                let naive_time: NaiveDateTime =
                                    date + Duration::milliseconds(milliseconds);

                                match chrono::Local.from_local_datetime(&naive_time) {
                                    Single(time) => Value::DateTime(time.to_utc()),
                                    Ambiguous(_, _) => {
                                        Value::Error("#AMBIGUOUS_TIME".into())
                                    }
                                    None => Value::Error("#VALUE!".into()),
                                }
                            }
                        }
                        Err(()) => Value::Error("#TYPE!".into()),
                    }
                }
            }
        }
    }

    impl TryFrom<*const std::ffi::c_char> for NetSetType {
        type Error = ();

        fn try_from(ptr: *const std::ffi::c_char) -> Result<Self, ()> {
            match ptr.is_null() {
                true => Ok(Self::Auto),
                false => match unsafe { CStr::from_ptr(ptr) }.to_bytes() {
                    b"" | b"auto" => Ok(Self::Auto),
                    b"f64" => Ok(Self::F64),
                    b"i64" => Ok(Self::I64),
                    b"null" => Ok(Self::Null),
                    b"time" => Ok(Self::Time),
                    b"string" => Ok(Self::String),
                    b"bool" => Ok(Self::Bool),
                    _ => Err(()),
                },
            }
        }
    }

    match unsafe { CStr::from_ptr(path) }.to_str() {
        Err(_) => XLOper12::error(XlErr::NA).into(),
        Ok(s) => match NetSetType::try_from(ty) {
            Err(()) => XLOper12::error(XlErr::NA).into(),
            Ok(typ) => {
                let path: netidx::path::Path = Into::<netidx::path::Path>::into(s);
                let value = typ.apply(value);
                match *SETTER {
                    None => XLOper12::error(XlErr::NA).into(),
                    Some(ref setter) => match setter.set(path, value) {
                        Ok(()) => SET.as_lpxloper12(),
                        Err(tokio::sync::mpsc::error::SendError((path, value))) => {
                            log::error!("failure setting {value} at netidx path {path}");
                            XLOper12::error(XlErr::NA).into()
                        }
                    },
                }
            }
        },
    }
}

fn register_udfs() -> Result<()> {
    xll_udf!("NetSet", NetSet).register(
        "QCQC$", // Q for the return value, C for the path, Q for the LPXLOPER12 value, C for the type, $ for thread-safe
        "path,value,[type]",
        "Netidx",
        "Write to a Netidx container",
        &[],
    )?;
    Ok(())
}

#[no_mangle]
extern "system" fn xlAutoOpen() -> i32 {
    let hr = DllRegisterServer();
    if hr.is_err() {
        log::debug!("DllRegisterServer failed: HRESULT {hr}");
    }

    // register all the functions we are exporting to Excel
    match register_udfs() {
        Ok(()) => {}
        Err(e) => log::error!("Could not register UDFs: {e}"),
    }

    1 // Per Excel SDK docs, this function must return [1]
}

#[no_mangle]
extern "system" fn xlAutoClose() -> i32 {
    1 // Per Excel SDK docs, this function must return [1]
}

register_xll_module![("NetidxRTD", CLSID, NetidxRTD),];
