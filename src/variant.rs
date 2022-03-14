use chrono::NaiveDateTime;
use rust_decimal::Decimal;
use winapi::shared::ntdef::HRESULT;
use winapi::um::oaidl::{IDispatch, VARIANT};
use crate::{ComBool, PtrWrapper};
use winapi::um::unknwnbase::IUnknown;
use enumn::N;
use crate::com_types::currency::ComCurrency;
use crate::com_types::date::ComDate;
use crate::com_types::string::ComString;

#[derive(Debug, PartialEq)]
#[allow(clippy::enum_variant_names)]
pub enum Variant
{
    Empty,
    Null,

    Bool(bool),
    BoolRef(&'static mut ComBool),

    I8(i8),
    I8Ref(&'static mut i8),
    I16(i16),
    I16Ref(&'static mut i16),
    I32(i32),
    I32Ref(&'static mut i32),
    I64(i64),
    I64Ref(&'static mut i64),

    U8(u8),
    U8Ref(&'static mut u8),
    U16(u16),
    U16Ref(&'static mut u16),
    U32(u32),
    U32Ref(&'static mut u32),
    U64(u64),
    U64Ref(&'static mut u64),

    F32(f32),
    F32Ref(&'static mut f32),
    F64(f64),
    F64Ref(&'static mut f64),

    Currency(Decimal),
    CurrencyRef(&'static mut ComCurrency),

    Date(NaiveDateTime),
    DateRef(&'static mut ComDate),

    String(String),
    StringRef(&'static mut ComString),

    Dispatch(PtrWrapper<IDispatch>),
    Unknown(PtrWrapper<IUnknown>),

    Error(HRESULT),
    ErrorRef(&'static mut HRESULT),

    VariantRef(PtrWrapper<VARIANT>),

    Record(),
}

#[derive(N, Debug, PartialEq)]
#[allow(non_camel_case_types)]
pub enum VariantType
{
    VT_EMPTY = 0,
    VT_NULL = 1,
    VT_I2 = 2,
    VT_I4 = 3,
    VT_R4 = 4,
    VT_R8 = 5,
    VT_CY = 6,
    VT_DATE = 7,
    VT_BSTR = 8,
    VT_DISPATCH = 9,
    VT_ERROR = 10,
    VT_BOOL = 11,
    VT_VARIANT = 12,
    VT_UNKNOWN = 13,
    VT_DECIMAL = 14,
    VT_I1 = 16,
    VT_UI1 = 17,
    VT_UI2 = 18,
    VT_UI4 = 19,
    VT_I8 = 20,
    VT_UI8 = 21,
    VT_INT = 22,
    VT_UINT = 23,
    VT_VOID = 24,
    VT_HRESULT = 25,
    VT_PTR = 26,
    VT_SAFEARRAY = 27,
    VT_CARRAY = 28,
    VT_USERDEFINED = 29,
    VT_LPSTR = 30,
    VT_LPWSTR = 31,
    VT_RECORD = 36,
    VT_INT_PTR = 37,
    VT_UINT_PTR = 38,
}

pub const VT_BYREF: u16 = 16384;

impl VariantType
{
    pub fn byref(self) -> u16
    {
        self as u16 | VT_BYREF
    }
}