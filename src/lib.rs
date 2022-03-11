mod bool;
mod ptr_wrapper;

use enumn::N;
use std::fmt::Debug;
use std::ops::Add;
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

use rust_decimal::Decimal;
use winapi::shared::ntdef::HRESULT;
use winapi::shared::wtypes::VARTYPE;
use winapi::um::oaidl::{IDispatch, VARIANT};
use winapi::um::unknwnbase::IUnknown;
use VariantType::*;
use crate::bool::ComBool;
use crate::ptr_wrapper::PtrWrapper;


#[derive(Debug, PartialEq)]
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
    CurrencyRef(&'static mut i64),
    Date(NaiveDateTime),
    DateRef(&'static mut f64),
    String(String),
    StringRef(&'static mut &'static mut u16),
    Dispatch(PtrWrapper<IDispatch>),
    Error(HRESULT),
    ErrorRef(&'static mut HRESULT),
    Unknown(PtrWrapper<IUnknown>),
    VariantRef(PtrWrapper<VARIANT>),
    Record()
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
    VT_UINT_PTR = 38
}

const VT_BYREF: u16 = 16384;

#[derive(Debug, PartialEq)]
pub enum VariantConversionError
{
    /// An error occured while converting the COM String to a Rust string.
    StringConversionError,
    /// The specified variant type is known but not supported.
    Unimplemented(VariantType),
    /// A reference-only variant type was used without VT_BYREF.
    InvalidDirect(VariantType),
    /// An invalid variant type was used in conjunction with VT_BYREF.
    InvalidReference(VariantType),
    /// The specified type can only be used in a TYPEDESC structure.
    TypeDescOnly(VariantType),
    /// The specified variant type is unknown.
    Unknown(VARTYPE),
}

#[macro_export]
macro_rules! vt {
    ( $is_ref:ident, $res:expr ) => {
        if $is_ref.1 { Err(VariantConversionError::InvalidReference(VariantType::n($is_ref.2).unwrap())) }
        else { Ok($res) }
    };

    ( $is_ref:ident, @, $atype: ident => ($ares: expr) ) => {
        if $is_ref.1 { Ok(Variant::$atype($ares)) }
        else { Err(VariantConversionError::InvalidDirect(VariantType::n($is_ref.2).unwrap())) }
    };

    ( $is_ref:ident, $vtype:ident => ($res:expr), $atype: ident => ($ares: expr) ) => {
        Ok(if $is_ref.1 { Variant::$atype($ares) } else { Variant::$vtype($res) })
    };

     ( $is_ref:ident, $vtype:ident => ($res:expr), $atype: ident => $ares: ident ) => {
        Ok(if $is_ref.1 { Variant::$atype((*$is_ref.0.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype($res) })
    };

    ( $is_ref:ident, $vtype:ident => $res:ident, $atype: ident => $ares: ident ) => {
        Ok(if $is_ref.1 { Variant::$atype((*$is_ref.0.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype(*$is_ref.0.n3.$res()) })
    };

    ( $is_ref:ident, $vtype:ident => $res:ident $op:tt $opexpr:expr, $atype: ident => $ares: ident ) => {
        Ok(if $is_ref.1 { Variant::$atype((*$is_ref.0.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype(*$is_ref.0.n3.$res() $op $opexpr) })
    };


}

#[macro_export]
macro_rules! types
{
    ( @vt $t: expr, $val:expr, $is_ref:expr, $res:expr ) => {
        if $is_ref { Err(VariantConversionError::InvalidReference(VariantType::n($t).unwrap())) }
        else { Ok($res) }
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, @, $atype: ident => ($ares: expr) ) => {
        if $is_ref { Ok(Variant::$atype($ares)) }
        else { Err(VariantConversionError::InvalidDirect(VariantType::n($t).unwrap())) }
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => ($res:expr), $atype: ident => ($ares: expr) ) => {
        Ok(if $is_ref { Variant::$atype($ares) } else { Variant::$vtype($res) })
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => ($res:expr), $atype: ident => $ares: ident ) => {
        Ok(if $is_ref { Variant::$atype((*$val.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype($res) })
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => $res:ident, $atype: ident => $ares: ident ) => {
        Ok(if $is_ref { Variant::$atype((*$val.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype(*$val.n3.$res()) })
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => $res:ident $op:tt $opexpr:expr, $atype: ident => $ares: ident ) => {
        Ok(if $is_ref { Variant::$atype((*$val.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype(*$val.n3.$res() $op $opexpr) })
    };

    ($val:expr, $is_ref:expr, $t:expr, [$( $name:ident : ( $($tts:tt)* ) ),*], [ $( $($pat:ident)|+ => $expr:expr ),*]) => {
        match $t
        {
            $(Some($name) => types!(@vt $t, $val, $is_ref, $($tts)*) ,)*
            $($(Some($pat))+ => $expr,)*
            None => Err(VariantConversionError::Unknown($val.vt))
        }
    }

    /*($val:expr, $is_ref:expr, $t:expr, [$( $name:ident : ( $vtype:ident => $res:ident, $atype: ident => $ares: ident ) ),*]) => {
        match $t
        {
            $(Some($name) => Ok(if $is_ref { Variant::$atype((*$val.n3.$ares()).as_mut::<'static>().unwrap().into()) } else { Variant::$vtype(*$val.n3.$res()) }) ,)*
            None => Err(VariantConversionError::Unknown($val.vt))
        }
    }*/
}

impl TryInto<Variant> for VARIANT
{
    type Error = VariantConversionError;

    fn try_into(self) -> Result<Variant, VariantConversionError>
    {
        unsafe {
            let val = self.n1.n2();

            let (unrefd, is_ref) = if val.vt & VT_BYREF as u16 != 0 {
                (val.vt & !(VT_BYREF as u16), true)
            } else {
                (val.vt, false)
            };

            let var = (val, is_ref, unrefd);

            types!(val, is_ref, VariantType::n(unrefd), [
                VT_EMPTY : (Variant::Empty),
                VT_NULL : (Variant::Null),

                VT_BOOL : (Bool => boolVal != 0, BoolRef => pboolVal),

                VT_I1 : (I8 => cVal, I8Ref => pcVal),
                VT_I2 : (I16 => iVal, I16Ref => piVal),
                VT_I4 : (I32 => lVal, I32Ref => plVal),
                VT_I8 : (I64 => llVal, I64Ref => pllVal),
                VT_UI1 : (U8 => bVal, U8Ref => pbVal),
                VT_UI2 : (U16 => uiVal, U16Ref => puiVal),
                VT_UI4 : (U32 => ulVal, U32Ref => pulVal),
                VT_UI8 : (U64 => ullVal, U64Ref => pullVal),
                VT_INT : (I32 => lVal, I32Ref => plVal),
                VT_UINT : (U32 => ulVal, U32Ref => pulVal),

                VT_R4 : (F32 => fltVal, F32Ref => pfltVal),
                VT_R8 : (F64 => dblVal, F64Ref => pdblVal),

                VT_CY : (Currency => (Decimal::new((*val.n3.cyVal()).int64, 4)), CurrencyRef => (&mut ((*(*val.n3.pcyVal())).int64))),

                VT_ERROR : (Error => scode, ErrorRef => pscode),
                VT_DISPATCH : (Variant::Dispatch(val.n3.pdispVal().try_into().unwrap())),
                VT_UNKNOWN : (Variant::Unknown(val.n3.punkVal().try_into().unwrap())),
                VT_VARIANT : (@, VariantRef => (PtrWrapper((*val.n3.pvarVal()).as_mut().unwrap())))
            ], [

            ]);

            match VariantType::n(unrefd)
            {
                Some(t) => match t
                {
                    VT_BOOL => vt!(var, Bool => boolVal != 0, BoolRef => pboolVal),
                    VT_I1 => vt!(var, I8  => cVal,  I8Ref  => pcVal),
                    VT_I2 => vt!(var, I16 => iVal,  I16Ref => piVal),
                    VT_I4 | VT_INT => vt!(var, I32 => lVal,  I32Ref => plVal),
                    VT_I8 => vt!(var, I64 => llVal, I64Ref => pllVal),
                    VT_UI1 => vt!(var, U8  => bVal,   U8Ref  => pbVal),
                    VT_UI2 => vt!(var, U16 => uiVal,  U16Ref => puiVal),
                    VT_UI4 | VT_UINT => vt!(var, U32 => ulVal,  U32Ref => pulVal),
                    VT_UI8 => vt!(var, U64 => ullVal, U64Ref => pullVal),
                    VT_R4 => vt!(var, F32 => fltVal, F32Ref => pfltVal),
                    VT_R8 => vt!(var, F64 => dblVal, F64Ref => pdblVal),
                    VT_CY => vt!(var, Currency => (Decimal::new((*val.n3.cyVal()).int64, 4)), CurrencyRef => (&mut ((*(*val.n3.pcyVal())).int64))),
                    VT_DATE => Ok(Variant::Date(NaiveDateTime::new(NaiveDate::from_ymd(1899, 12, 30), NaiveTime::from_hms(0, 0, 0))
                        .add(Duration::milliseconds((*val.n3.date() * 24.0 * 60.0 * 60.0 * 1000.0) as i64)))),
                    VT_EMPTY => vt!(var, Variant::Empty),
                    VT_NULL => vt!(var, Variant::Null),
                    VT_BSTR => if is_ref
                    {
                        Err(VariantConversionError::InvalidReference(t))
                    }
                    else
                    {
                        widestring::U16CString::from_ptr_str(*val.n3.bstrVal()).to_string()
                            .map(|s| Variant::String(s))
                            .map_err(|_| VariantConversionError::StringConversionError)
                    },
                    VT_ERROR => vt!(var, Error => scode, ErrorRef => pscode),
                    VT_DISPATCH => vt!(var, Variant::Dispatch(val.n3.pdispVal().try_into().unwrap())),
                    VT_UNKNOWN => vt!(var, Variant::Unknown(val.n3.punkVal().try_into().unwrap())),
                    VT_VARIANT => vt!(var, @, VariantRef => (PtrWrapper((*val.n3.pvarVal()).as_mut().unwrap()))),
                    VT_DECIMAL | VT_RECORD => Err(VariantConversionError::Unimplemented(t)),

                    VT_VOID |
                    VT_HRESULT |
                    VT_PTR |
                    VT_SAFEARRAY |
                    VT_CARRAY |
                    VT_USERDEFINED |
                    VT_LPSTR |
                    VT_LPWSTR |
                    VT_INT_PTR |
                    VT_UINT_PTR => Err(VariantConversionError::TypeDescOnly(t)),
                },
                None => Err(VariantConversionError::Unknown(val.vt))
            }
        }
    }
}

#[cfg(test)]
mod tests {
    use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
    use winapi::um::oaidl::{VARIANT};
    use crate::{Variant, VariantType};

    #[macro_export]
    macro_rules! variant {
        ( $type: ident, $field: ident, $val: expr ) => {
            unsafe {
                let mut variant: VARIANT = std::mem::zeroed();
                variant.n1.n2_mut().vt = VariantType::$type as u16;
                *variant.n1.n2_mut().n3.$field() = $val;
                variant.try_into()
            }
        };
    }

    #[test]
    fn it_works() {
        assert_eq!(variant!(VT_DATE, date_mut, 5.25),
                   Ok(Variant::Date(NaiveDateTime::new(NaiveDate::from_ymd(1900, 01, 04), NaiveTime::from_hms(6, 0, 0)))));


    }
}
