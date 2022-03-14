#![allow(unused_unsafe)] // rustc bug #94912

use crate::variant::*;
use crate::com_types::bool::ComBool;
use crate::com_types::ptr_wrapper::PtrWrapper;

mod variant;
mod convert;
mod com_types;

#[macro_export]
macro_rules! variant
{
    ( $type: expr ) =>
    {
        unsafe {
            let mut variant: VARIANT = std::mem::zeroed();
            variant.n1.n2_mut().vt = $type as u16;
            variant
        }
    };

    ( $type: expr, $field: ident, $val: expr ) =>
    {
        unsafe {
            let mut variant: VARIANT = variant!($type);
            *variant.n1.n2_mut().n3.$field() = $val;
            variant
        }
    };

    ( $type: expr, ($field: ident), $val: expr ) =>
    {
        unsafe {
            let mut variant: VARIANT = variant!($type);
            *variant.n1.$field() = $val;
            variant
        }
    };
}

#[cfg(test)]
mod tests
{
    use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
    use winapi::um::oaidl::VARIANT;
    use rust_decimal_macros::dec;
    use widestring::U16CString;
    use crate::{Variant, VariantType};
    use winapi::shared::wtypes::{CY, DECIMAL};

    macro_rules! roundtrip
    {
        ( ($t: ident $(, $($tts:tt)*)?), $b: expr ) =>
        {
            #[test]
            #[allow(non_snake_case)]
            fn $t()
            {
                let conv = variant!(VariantType::$t $(, $($tts)*)?).try_into();
                assert_eq!(conv, Ok($b), "COM to Rust");
                let cv: VARIANT = $b.try_into().unwrap();
                assert_eq!(cv.try_into(), Ok($b), "Rust to COM");
            }
        }
    }

    roundtrip!((VT_EMPTY), Variant::Empty);
    roundtrip!((VT_NULL), Variant::Null);

    roundtrip!((VT_BOOL, boolVal_mut, !0i16), Variant::Bool(true));

    roundtrip!((VT_I1, cVal_mut, 0x55), Variant::I8(0x55));
    roundtrip!((VT_I2, iVal_mut, 0x55aa), Variant::I16(0x55aa));
    roundtrip!((VT_I4, lVal_mut, 0x55aa55aa), Variant::I32(0x55aa55aa));
    roundtrip!((VT_I8, llVal_mut, 0x55aa55aa55aa55aa), Variant::I64(0x55aa55aa55aa55aa));

    roundtrip!((VT_UI1, bVal_mut, 0x55), Variant::U8(0x55));
    roundtrip!((VT_UI2, uiVal_mut, 0x55aa), Variant::U16(0x55aa));
    roundtrip!((VT_UI4, ulVal_mut, 0x55aa55aa), Variant::U32(0x55aa55aa));
    roundtrip!((VT_UI8, ullVal_mut, 0x55aa55aa55aa55aa), Variant::U64(0x55aa55aa55aa55aa));

    roundtrip!((VT_R4, fltVal_mut, 0.5f32), Variant::F32(0.5f32));
    roundtrip!((VT_R8, dblVal_mut, 0.5f64), Variant::F64(0.5f64));

    roundtrip!((VT_CY, cyVal_mut, CY { int64: 123456 }),
                   Variant::Currency(dec!(12.3456)));

    roundtrip!((VT_DECIMAL, (decVal_mut), DECIMAL { wReserved: VariantType::VT_DECIMAL as u16, scale: 4, sign: 0, Hi32: 0, Lo64: 123456 }),
                   Variant::Decimal(dec!(12.3456)));

    roundtrip!((VT_DATE, date_mut, 5.25),
                   Variant::Date(
                       NaiveDateTime::new(
                           NaiveDate::from_ymd(1900, 1, 4),
                           NaiveTime::from_hms(6, 0, 0))));

    roundtrip!((VT_BSTR, bstrVal_mut, U16CString::from_str_truncate("Hello, world!").into_raw()),
                   Variant::String("Hello, world!".to_string()));
}
