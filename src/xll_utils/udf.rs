use super::{XLOper12, Xlfn};

#[macro_export]
macro_rules! xll_udf {
    ($name:literal, $f:ident) => {
        $crate::xll_utils::udf::Udf { name: $name, exported_function: stringify!($f) }
    };
}

pub struct Udf {
    pub name: &'static str,
    pub exported_function: &'static str,
}

impl Udf {
    pub fn register(
        &self,
        arg_types: &str,
        arg_text: &str,
        category: &str,
        help_text: &str,
        arg_help: &[&str],
    ) -> anyhow::Result<()> {
        match super::excel12(Xlfn::xlGetName, &mut []) {
            Err(_) => anyhow::bail!("xlGetName failed"),
            Ok(dll_name) => {
                let mut opers = vec![
                    dll_name,
                    XLOper12::from(self.exported_function),
                    XLOper12::from(arg_types),
                    XLOper12::from(self.name),
                    XLOper12::from(arg_text),
                    XLOper12::from(1),
                    XLOper12::from(category),
                    XLOper12::missing(),
                    XLOper12::missing(),
                    XLOper12::from(help_text),
                ];

                // append any argument help strings
                for arg in arg_help.iter() {
                    opers.push(XLOper12::from(*arg));
                }

                match super::excel12(Xlfn::xlfRegister, opers.as_mut_slice()) {
                    Ok(xloper) if !xloper.is_err(super::xlcall::xlerrValue) => Ok(()),
                    Ok(xloper) | Err(xloper) => {
                        anyhow::bail!("registration failed with error {xloper}")
                    }
                }
            }
        }
    }
}
