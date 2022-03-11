use crate::{Variant, VariantType, VT_BYREF, PtrWrapper};
use winapi::shared::wtypes::VARTYPE;
use winapi::um::oaidl::VARIANT;
use rust_decimal::Decimal;
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use std::ops::Add;
use std::string::FromUtf16Error;
use crate::VariantType::*;

#[derive(Debug, PartialEq)]
pub enum VariantConversionError
{
    /// An error occured while converting the COM String to a Rust string.
    StringConversionError,
    /// An unknown occured while converting the value of the Variant object.
    GenericConversionError,
    /// The specified variant type is known but not supported.
    Unimplemented(VariantType),
    /// A reference-only variant type was used without VT_BYREF.
    InvalidDirect(VariantType),
    /// An invalid variant type was used in conjunction with VT_BYREF.
    InvalidReference(VariantType),
    /// The specified type can only be used in a TYPEDESC structure.
    TypeDescOnly(VariantType),
    /// The specified variant type is unknown.
    UnknownType(VARTYPE),
}

impl From<FromUtf16Error> for VariantConversionError
{
    fn from(_: FromUtf16Error) -> VariantConversionError
    {
        VariantConversionError::StringConversionError
    }
}

impl From<()> for VariantConversionError
{
    fn from(_: ()) -> VariantConversionError
    {
        VariantConversionError::GenericConversionError
    }
}

#[macro_export]
macro_rules! types
{
    // Direct-only type; custom expression
    // VT_DISPATCH : (Dispatch => (...), /),
    ( @vt $t: ident, $val:expr, $is_ref:expr, $vtype:ident => ($res:expr), / ) => {
        if $is_ref { Err(VariantConversionError::InvalidReference($t)) }
        else { $res.map(Variant::$vtype).map_err(Into::into) }
    };

    // VT_EMPTY : (Empty, /)
    ( @vt $t: ident, $val:expr, $is_ref:expr, $vtype:ident, / ) => {
        if $is_ref { Err(VariantConversionError::InvalidReference($t)) }
        else { Ok(Variant::$vtype) }
    };

    // Reference-only type, custom expression
    // VT_VARIANT : (/, VariantRef => ...)
    ( @vt $t: ident, $val:expr, $is_ref:expr, /, $atype: ident => ($ares: expr) ) => {
        if $is_ref { $ares.map(Variant::$atype).map_err(Into::into) }
        else { Err(VariantConversionError::InvalidDirect($t)) }
    };

    ( @ref $val: expr, $atype: ident, $ares: ident ) => {
        (*$val.n3.$ares()).as_mut::<'static>().ok_or(()).map(Into::into).map(Variant::$atype).map_err(Into::into)
    };

    ( @ref $val: expr, $atype: ident, $ares: expr ) => {
        $ares.map(Variant::$atype)
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => ($res:expr), $atype: ident => $ares: expr ) => {
        if $is_ref { types!(@ref $val, $atype, $ares) } else { $res.map(Variant::$vtype) }
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => $res:ident, $atype: ident => $ares: ident ) => {
        if $is_ref { types!(@ref $val, $atype, $ares) } else { Ok(Variant::$vtype(*$val.n3.$res())) }
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => $res:ident $op:tt $opexpr:expr, $atype: ident => $ares: ident ) => {
        if $is_ref { types!(@ref $val, $atype, $ares) } else { Ok(Variant::$vtype(*$val.n3.$res() $op $opexpr)) }
    };

    ($val:expr, $is_ref:expr, $t:expr, [$( $name:ident : ( $($tts:tt)* ) ),*], [ $( $($pat:ident)|+ => $expr:expr ),*], [ $([ $( $u:ident ),* ] => $err:ident),* ]) => {
        match $t
        {
            Some(t) => match t
            {
                $($name => types!(@vt $name, $val, $is_ref, $($tts)*) ,)*
                $($($pat => $expr ,)*)*
                $($($u => Err(VariantConversionError::$err($u)),)*)*
            },
            None => Err(VariantConversionError::UnknownType($val.vt))
        }
    };
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

            types!(val, is_ref, VariantType::n(unrefd), [
                VT_EMPTY : (Empty, /),
                VT_NULL : (Null, /),

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

                VT_CY : (
                    Currency => (Ok(Decimal::new((*val.n3.cyVal()).int64, 4))),
                    CurrencyRef => (Ok(&mut ((**val.n3.pcyVal()).int64)))),

                VT_ERROR : (Error => scode, ErrorRef => pscode),
                VT_DISPATCH : (Dispatch => (val.n3.pdispVal().try_into()), /),
                VT_UNKNOWN : (Unknown => (val.n3.punkVal().try_into()), /),

                VT_VARIANT : (/, VariantRef => ((*val.n3.pvarVal()).as_mut().map(PtrWrapper).ok_or(()))),

                VT_DATE : (
                    Date => (Ok(NaiveDateTime::new(NaiveDate::from_ymd(1899, 12, 30), NaiveTime::from_hms(0, 0, 0))
                        .add(Duration::milliseconds((*val.n3.date() * 24.0 * 60.0 * 60.0 * 1000.0) as i64)))),
                    DateRef => (Ok(&mut **val.n3.pdate()))),

                VT_BSTR : (String => (widestring::U16CString::from_ptr_str(*val.n3.bstrVal()).to_string()), /)
            ], [

            ], [
                [VT_DECIMAL, VT_RECORD] => Unimplemented,
                [
                    VT_VOID, VT_HRESULT,
                    VT_SAFEARRAY, VT_CARRAY,
                    VT_USERDEFINED,
                    VT_LPSTR, VT_LPWSTR,
                    VT_PTR, VT_INT_PTR, VT_UINT_PTR
                ] => TypeDescOnly
            ])
        }
    }
}
