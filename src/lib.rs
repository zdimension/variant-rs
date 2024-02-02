#![allow(unused_unsafe)] // rustc bug #94912
#![doc = include_str!("../README.md")]

use crate::com_types::bool::ComBool;
use crate::com_types::ptr_wrapper::PtrWrapper;
pub use crate::variant::*;

pub use windows::Win32::System::Variant::{VARENUM, VARIANT};

pub mod com_types;
pub mod convert;
pub mod dispatch;
pub mod variant;

#[doc(hidden)]
#[macro_export]
macro_rules! variant {
    ( $type: expr ) => {
        unsafe {
            let mut variant: VARIANT = std::mem::zeroed();
            (*variant.Anonymous.Anonymous).vt = VARENUM($type as u16);
            variant
        }
    };

    ( $type: expr, $field: ident, $val: expr ) => {
        unsafe {
            let mut variant: VARIANT = variant!($type);
            (*variant.Anonymous.Anonymous).Anonymous.$field = $val;
            variant
        }
    };

    ( $type: expr, ($field: ident), $val: expr ) => {
        unsafe {
            let mut variant: VARIANT = variant!($type);
            variant.Anonymous.$field = $val;
            variant
        }
    };
}

#[cfg(test)]
mod tests {
    use crate::{ToVariant, Variant, VariantType};
    use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
    use rust_decimal_macros::dec;

    use std::mem::ManuallyDrop;
    use windows::core::BSTR;
    use windows::Win32::Foundation::{DECIMAL, DECIMAL_0, DECIMAL_0_0, DECIMAL_1, VARIANT_BOOL};
    use windows::Win32::System::Com::CY;
    use windows::Win32::System::Variant::{VARENUM, VARIANT};

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

    roundtrip!((VT_BOOL, boolVal, VARIANT_BOOL(!0i16)), Variant::Bool(true));

    roundtrip!((VT_I1, cVal, 0x55), Variant::I8(0x55));
    roundtrip!((VT_I2, iVal, 0x55aa), Variant::I16(0x55aa));
    roundtrip!((VT_I4, lVal, 0x55aa55aa), Variant::I32(0x55aa55aa));
    roundtrip!(
        (VT_I8, llVal, 0x55aa55aa55aa55aa),
        Variant::I64(0x55aa55aa55aa55aa)
    );

    roundtrip!((VT_UI1, bVal, 0x55), Variant::U8(0x55));
    roundtrip!((VT_UI2, uiVal, 0x55aa), Variant::U16(0x55aa));
    roundtrip!((VT_UI4, ulVal, 0x55aa55aa), Variant::U32(0x55aa55aa));
    roundtrip!(
        (VT_UI8, ullVal, 0x55aa55aa55aa55aa),
        Variant::U64(0x55aa55aa55aa55aa)
    );

    roundtrip!((VT_R4, fltVal, 0.5f32), Variant::F32(0.5f32));
    roundtrip!((VT_R8, dblVal, 0.5f64), Variant::F64(0.5f64));

    roundtrip!(
        (VT_CY, cyVal, CY { int64: 123456 }),
        Variant::Currency(dec!(12.3456).into())
    );

    roundtrip!(
        (
            VT_DECIMAL,
            (decVal),
            DECIMAL {
                wReserved: VariantType::VT_DECIMAL as u16,
                Anonymous1: DECIMAL_0 {
                    Anonymous: DECIMAL_0_0 { scale: 4, sign: 0 }
                },
                Hi32: 0,
                Anonymous2: DECIMAL_1 { Lo64: 123456 }
            }
        ),
        Variant::Decimal(dec!(12.3456))
    );

    roundtrip!(
        (VT_DATE, date, 5.25),
        Variant::Date(NaiveDateTime::new(
            NaiveDate::from_ymd_opt(1900, 1, 4).unwrap(),
            NaiveTime::from_hms_opt(6, 0, 0).unwrap()
        ))
    );

    roundtrip!(
        (
            VT_BSTR,
            bstrVal,
            ManuallyDrop::new(BSTR::from("Hello, world!"))
        ),
        Variant::String(BSTR::from("Hello, world!"))
    );

    #[test]
    fn main() {
        let v1 = Variant::I32(123); // manual instanciation
        let v2 = 123.to_variant(); // ToVariant trait
        let v3 = 123.into(); // From / Into traits
        assert_eq!(v1, v2);
        assert_eq!(v1, v3);

        let bstr: Variant = "Hello, world!".into();
        let ptr: VARIANT = bstr.clone().try_into().unwrap(); // convert to COM VARIANT
        let back: Variant = ptr.try_into().unwrap(); // convert back
        assert_eq!(bstr, back);
    }
}
