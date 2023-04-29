//! Utilities for using [`IDispatch`] from Rust in an ergonomic fashion

use crate::convert::VariantConversionError;
use crate::variant::Variant;
use thiserror::Error;
use widestring::U16CString;
use windows::core::{Error as WinError, GUID, PCWSTR};
use windows::Win32::Foundation::{DISP_E_EXCEPTION, DISP_E_PARAMNOTFOUND, DISP_E_TYPEMISMATCH};
use windows::Win32::System::Com::{
    IDispatch, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPATCH_PROPERTYPUT,
    DISPPARAMS, EXCEPINFO, VARIANT,
};
use windows::Win32::System::Ole::DISPID_PROPERTYPUT;

pub trait IDispatchExt {
    fn get(&self, name: &str) -> Result<Variant, IDispatchError>;
    fn put(&self, name: &str, value: Variant) -> Result<(), IDispatchError>;
    fn call(&self, name: &str, args: Vec<Variant>) -> Result<Variant, IDispatchError>;
}

#[derive(Error, Debug)]
pub enum IDispatchError {
    #[error("Couldn't convert args to VARIANT")]
    VariantConversion(#[from] VariantConversionError),
    #[error("Couldn't convert string to BSTR")]
    StringConversion(#[from] widestring::error::ContainsNul<u16>),
    #[error("Win32 error")]
    GenericWin32(#[from] WinError),
    #[error("COM exception")]
    Exception(EXCEPINFO),
    #[error("COM argument error {0} for argument {1}")]
    Argument(ComArgumentError, usize),
}

#[derive(Error, Debug)]
pub enum ComArgumentError {
    #[error("The value's type does not match the expected type for the parameter")]
    TypeMismatch,
    #[error("A required parameter was missing")]
    ParameterNotFound,
}

const LOCALE_USER_DEFAULT: u32 = 0x0400;
const LOCALE_SYSTEM_DEFAULT: u32 = 0x0800;

fn invoke(
    obj: &IDispatch,
    name: &str,
    dp: &mut DISPPARAMS,
    flags: DISPATCH_FLAGS,
) -> Result<Variant, IDispatchError> {
    let mut name = U16CString::from_str(name).map_err(IDispatchError::StringConversion)?;
    let mut id = 0i32;
    unsafe {
        obj.GetIDsOfNames(
            &GUID::default(),
            &PCWSTR(name.as_mut_ptr()),
            1,
            LOCALE_USER_DEFAULT,
            (&mut id) as *mut i32,
        )
    }
    .map_err(IDispatchError::GenericWin32)?;

    let mut excep = EXCEPINFO::default();
    let mut arg_err = 0;
    let mut result = VARIANT::default();

    let res = unsafe {
        obj.Invoke(
            id,
            &GUID::default(),
            LOCALE_SYSTEM_DEFAULT,
            flags,
            dp,
            Some(&mut result),
            Some(&mut excep),
            Some(&mut arg_err),
        )
    };

    match res {
        Ok(_) => result.try_into().map_err(Into::into),
        Err(e) => Err(match e.code() {
            DISP_E_EXCEPTION => IDispatchError::Exception(excep),
            DISP_E_TYPEMISMATCH => {
                IDispatchError::Argument(ComArgumentError::TypeMismatch, arg_err as usize)
            }
            DISP_E_PARAMNOTFOUND => {
                IDispatchError::Argument(ComArgumentError::ParameterNotFound, arg_err as usize)
            }
            _ => IDispatchError::GenericWin32(e),
        }),
    }
}

/// Get a property from the COM object
///
/// # Example
/// ```
/// use variant_rs::get;
/// let x = get!(com_object, SomeProp)?;
/// ```
#[macro_export]
macro_rules! get {
    ($obj:expr, $name:ident) => {{
        use variant_rs::dispatch::IDispatchExt;
        $obj.get(stringify!($name))
    }};
}

/// Set a property on the COM object
///
/// # Example
/// ```
/// use variant_rs::put;
/// put!(com_object, SomeProp, 10)?;
/// ```
#[macro_export]
macro_rules! put {
    ($obj:expr, $name:ident, $value:expr) => {{
        use variant_rs::dispatch::IDispatchExt;
        let val: Variant = $value.into();
        $obj.put(stringify!($name), val)
    }};
}

/// Call a method on the COM object
///
/// # Example
/// ```
/// use variant_rs::call;
/// let x = call!(com_object, SomeMethod(10, "hello"))?;
/// ```
#[macro_export]
macro_rules! call {
    ($obj:expr, $name:ident($($arg:expr),*)) => {
        {
            use variant_rs::dispatch::IDispatchExt;
            let args = vec![$((&$arg).into()),*];
            $obj.call(stringify!($name), args)
        }
    };
}

impl IDispatchExt for IDispatch {
    /// Get a property from a COM object
    ///
    /// Note: consider using the [`get!`] macro
    fn get(&self, name: &str) -> Result<Variant, IDispatchError> {
        let mut dp = DISPPARAMS::default();
        invoke(self, name, &mut dp, DISPATCH_PROPERTYGET)
    }

    /// Set a property on a COM object
    ///
    /// Note: consider using the [`put!`] macro
    fn put(&self, name: &str, value: Variant) -> Result<(), IDispatchError> {
        let mut value = value.try_into()?;
        let mut dp = DISPPARAMS {
            cArgs: 1,
            rgvarg: &mut value,
            cNamedArgs: 1,
            ..Default::default()
        };
        let mut id = DISPID_PROPERTYPUT;
        dp.rgdispidNamedArgs = &mut id as *mut _;
        invoke(self, name, &mut dp, DISPATCH_PROPERTYPUT)?;
        Ok(())
    }

    /// Call a method on a COM object
    ///
    /// Note: consider using the [`call!`] macro
    fn call(&self, name: &str, args: Vec<Variant>) -> Result<Variant, IDispatchError> {
        let mut dp = DISPPARAMS::default();
        let args: Vec<VARIANT> = args
            .into_iter()
            .rev()
            .map(|v| v.try_into().map_err(IDispatchError::from))
            .collect::<Result<_, _>>()?;
        dp.cArgs = args.len() as u32;
        dp.rgvarg = args.as_ptr() as *mut _;
        invoke(self, name, &mut dp, DISPATCH_METHOD)
    }
}
