//! Rust wrapper for the [`VARIANT`] type

use crate::com_types::currency::{ComCurrency, Currency};
use crate::com_types::date::ComDate;
use crate::com_types::decimal::ComDecimal;
//use crate::com_types::string::ComString;
use crate::{ComBool, PtrWrapper};
use chrono::NaiveDateTime;
use enumn::N;
use rust_decimal::Decimal;

use paste::paste;
use windows::core::IUnknown;
use windows::core::HRESULT;
use windows::core::{BSTR, PSTR};
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

macro_rules! variant_enum {
    (@impl $name:ident) => {};

    (@impl $name:ident (&'static $($type:tt)+)) => {};

    (@impl $name:ident (@@ $type:ty)) => {};

    (@impl $name:ident (Option<$type:ty>)) => {
        impl ToVariant for $type {
            fn to_variant(self) -> Variant {
                Variant::$name(Some(self))
            }
        }

        variant_enum!(@impl @opt $name (Option<$type>));
    };

    (@impl $(@opt)? $name:ident ($type:ty)) => {
        impl ToVariant for $type {
            fn to_variant(self) -> Variant {
                Variant::$name(self)
            }
        }
    };

    (@match $self:ident, $name:ident, $type:ty) => {
        match $self {
            Self::$name(v) => Ok(v),
            _ => Err($self),
        }
    };

    (@match $self:ident, $name:ident, ) => {
        match $self {
            Self::$name => Ok(()),
            _ => Err($self),
        }
    };

    (@enum $($name:ident $(( $(@@)? $( $type:ty )+ ) )?),* $(,)?) => {
        #[derive(Debug, PartialEq)]
        #[allow(clippy::enum_variant_names)]
        pub enum Variant {
            $($name $(($($type)+))?),*
        }

        paste! {
            #[allow(unused_parens)]
            impl Variant {
                $(
                    pub fn [< try_ $name:lower >](self) -> Result<( $($( $type )+)? ), Variant> {
                        variant_enum!(@match self, $name, $($( $type )+)?)
                    }

                    pub fn [< expect_ $name:lower >](self) -> ( $($( $type )+)? ) {
                        match self.[< try_ $name:lower >]() {
                            Ok(v) => v,
                            Err(self_) => panic!("Expected variant type {} but got {:?}", stringify!($name), self_),
                        }
                    }
                )*
            }
        }
    };

    ($($name:ident $(( $($type:tt)+ ) )?),* $(,)?) => {
        variant_enum!{@enum $($name $(( $($type)+ ))?),*}

        $(variant_enum!{@impl $name $(($($type)+) )?})*
    };
}

variant_enum! {
    Empty,
    Null,

    Bool(bool),
    BoolRef(&'static mut ComBool),

    I8(i8),
    I8Ref(PSTR),
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

    Currency(Currency),
    CurrencyRef(&'static mut ComCurrency),

    Decimal(Decimal),
    DecimalRef(&'static mut ComDecimal),

    Date(NaiveDateTime),
    DateRef(&'static mut ComDate),

    String(BSTR),
    StringRef(&'static mut BSTR),

    Dispatch(Option<IDispatch>),
    Unknown(Option<IUnknown>),

    Error(@@ HRESULT),
    ErrorRef(&'static mut HRESULT),

    VariantRef(PtrWrapper<VARIANT>),

    //Record(),
}

impl Clone for Variant {
    fn clone(&self) -> Self {
        use Variant::*;
        match self {
            BoolRef(_) | I8Ref(_) | I16Ref(_) | I32Ref(_) | I64Ref(_) | U8Ref(_) | U16Ref(_)
            | U32Ref(_) | U64Ref(_) | F32Ref(_) | F64Ref(_) | CurrencyRef(_) | DecimalRef(_)
            | DateRef(_) | StringRef(_) | ErrorRef(_) | VariantRef(_) => {
                panic!("Cannot clone a reference variant")
            }

            Empty => Empty,
            Null => Null,
            Bool(x) => Bool(*x),
            I8(x) => I8(*x),
            I16(x) => I16(*x),
            I32(x) => I32(*x),
            I64(x) => I64(*x),
            U8(x) => U8(*x),
            U16(x) => U16(*x),
            U32(x) => U32(*x),
            U64(x) => U64(*x),
            F32(x) => F32(*x),
            F64(x) => F64(*x),
            Currency(x) => Currency(*x),
            Decimal(x) => Decimal(*x),
            Date(x) => Date(*x),
            String(x) => String(x.clone()),
            Dispatch(x) => Dispatch(x.clone()),
            Unknown(x) => Unknown(x.clone()),
            Error(x) => Error(*x),
        }
    }
}

pub trait ToVariant {
    fn to_variant(self) -> Variant;
}

impl ToVariant for () {
    fn to_variant(self) -> Variant {
        Variant::Null
    }
}

impl ToVariant for &str {
    fn to_variant(self) -> Variant {
        Variant::String(BSTR::from(self))
    }
}

impl ToVariant for String {
    fn to_variant(self) -> Variant {
        Variant::String(BSTR::from(self))
    }
}

impl<T: Clone + ToVariant> ToVariant for &T {
    fn to_variant(self) -> Variant {
        self.clone().to_variant()
    }
}

impl<T: ToVariant> From<T> for Variant {
    fn from(t: T) -> Self {
        t.to_variant()
    }
}

#[derive(N, Debug, PartialEq, Eq, Copy, Clone)]
#[allow(non_camel_case_types)]
pub enum VariantType {
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

impl VariantType {
    pub fn byref(self) -> u16 {
        self as u16 | VT_BYREF
    }
}
