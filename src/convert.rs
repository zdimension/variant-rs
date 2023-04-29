//! Conversion between native [`VARIANT`] and Rust [`Variant`]

use crate::com_types::currency::ComCurrency;
use crate::com_types::date::ComDate;
use crate::com_types::decimal::ComDecimal;
use crate::Variant::*;
use crate::VariantType::*;
use crate::{variant, ComBool, PtrWrapper, Variant, VariantType, VT_BYREF};
use std::string::FromUtf16Error;

use std::convert::Infallible;
use std::mem::ManuallyDrop;
use thiserror::Error;
use windows::core::HRESULT;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Com::{VARENUM, VARIANT};

#[derive(Debug, PartialEq, Eq, Error)]
pub enum VariantConversionError {
    #[error("An error occured while converting the COM String to a Rust string.")]
    StringConversionError,
    #[error("An unknown occured while converting the value of the Variant object.")]
    GenericConversionError,
    #[error("The specified variant type is known but not supported.")]
    Unimplemented(VariantType),
    #[error("A reference-only variant type was used without VT_BYREF.")]
    InvalidDirect(VariantType),
    #[error("An invalid variant type was used in conjunction with VT_BYREF.")]
    InvalidReference(VariantType),
    #[error("The specified type can only be used in a TYPEDESC structure.")]
    TypeDescOnly(VariantType),
    #[error("The specified variant type is unknown.")]
    UnknownType(VARENUM),
}

impl From<Infallible> for VariantConversionError {
    fn from(p: Infallible) -> Self {
        match p {}
    }
}

impl From<FromUtf16Error> for VariantConversionError {
    fn from(_: FromUtf16Error) -> VariantConversionError {
        VariantConversionError::StringConversionError
    }
}

impl From<()> for VariantConversionError {
    fn from(_: ()) -> VariantConversionError {
        VariantConversionError::GenericConversionError
    }
}

#[doc(hidden)]
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
        ($val.Anonymous.Anonymous.$ares).as_mut::<'static>().ok_or(()).map(Into::into).map(Variant::$atype).map_err(Into::into)
    };

    ( @ref $val: expr, $atype: ident, $ares: expr ) => {
        $ares.map(Variant::$atype)
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => ($res:expr), $atype: ident => $ares: expr ) => {
        if $is_ref { types!(@ref $val, $atype, $ares) } else { $res.map(Variant::$vtype) }
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => $res:ident, $atype: ident => $ares: ident ) => {
        if $is_ref { types!(@ref $val, $atype, $ares) } else { Ok(Variant::$vtype($val.Anonymous.Anonymous.$res)) }
    };

    ( @vt $t: expr, $val:expr, $is_ref:expr, $vtype:ident => $res:ident $op:tt $opexpr:expr, $atype: ident => $ares: ident ) => {
        if $is_ref { types!(@ref $val, $atype, $ares) } else { Ok(Variant::$vtype($val.Anonymous.Anonymous.$res $op $opexpr)) }
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
            None => Err(VariantConversionError::UnknownType($val.Anonymous.vt))
        }
    };
}

impl TryInto<Variant> for VARIANT {
    type Error = VariantConversionError;

    fn try_into(self) -> Result<Variant, VariantConversionError> {
        unsafe {
            let val = self.Anonymous;

            let (unrefd, is_ref) = if val.Anonymous.vt.0 & VT_BYREF != 0 {
                (val.Anonymous.vt.0 & !VT_BYREF, true)
            } else {
                (val.Anonymous.vt.0, false)
            };

            types!(val, is_ref, VariantType::n(unrefd), [
                VT_EMPTY : (Empty, /),
                VT_NULL : (Null, /),

                VT_BOOL : (Bool => (Ok(val.Anonymous.Anonymous.boolVal.0 != 0)), BoolRef => (Ok(&mut *(val.Anonymous.Anonymous.pboolVal as *mut ComBool)))),

                VT_I1 : (I8 => (Ok(val.Anonymous.Anonymous.cVal as i8)), I8Ref => (Ok(val.Anonymous.Anonymous.pcVal))),
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
                    Currency => (Ok(ComCurrency::from(val.Anonymous.Anonymous.cyVal).into())),
                    CurrencyRef => (Ok(<&mut ComCurrency>::from(val.Anonymous.Anonymous.pcyVal)))),

                VT_DECIMAL : (
                    Decimal => (Ok(ComDecimal(val.decVal).into())),
                    DecimalRef => (Ok(<&mut ComDecimal>::from(val.Anonymous.Anonymous.pdecVal)))),

                VT_DATE : (
                    Date => (Ok(ComDate(val.Anonymous.Anonymous.date).into())),
                    DateRef => (Ok(<&mut ComDate>::from(val.Anonymous.Anonymous.pdate)))),

                VT_BSTR : (String => (Ok::<_, Infallible>(ManuallyDrop::into_inner(ManuallyDrop::into_inner(val.Anonymous).Anonymous.bstrVal))), /),

                VT_DISPATCH : (Dispatch => (Ok::<_, Infallible>((*ManuallyDrop::into_inner(val.Anonymous).Anonymous.pdispVal).take())), /),
                VT_UNKNOWN : (Unknown => (Ok::<_, Infallible>((*ManuallyDrop::into_inner(val.Anonymous).Anonymous.punkVal).take())), /),

                VT_ERROR : (Error => (Ok(HRESULT(val.Anonymous.Anonymous.scode))), ErrorRef => (Ok((val.Anonymous.Anonymous.pscode as *mut HRESULT).as_mut::<'static>().unwrap()))),

                VT_VARIANT : (/, VariantRef => (PtrWrapper::try_from(&val.Anonymous.Anonymous.pvarVal)))
            ], [

            ], [
                [VT_RECORD] => Unimplemented,
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

impl TryInto<VARIANT> for Variant {
    type Error = VariantConversionError;

    fn try_into(self) -> Result<VARIANT, VariantConversionError> {
        match self {
            Empty => Ok(variant!(VT_EMPTY)),
            Null => Ok(variant!(VT_NULL)),

            Bool(b) => Ok(variant!(VT_BOOL, boolVal, VARIANT_BOOL(ComBool::from(b) as i16))),
            BoolRef(b) => Ok(variant!(VT_BOOL, pboolVal, b as *mut ComBool as *mut VARIANT_BOOL)),

            I8(i) => Ok(variant!(VT_I1, cVal, i as u8)),
            I8Ref(i) => Ok(variant!(VT_I1.byref(), pcVal, i)),
            I16(i) => Ok(variant!(VT_I2, iVal, i)),
            I16Ref(i) => Ok(variant!(VT_I2.byref(), piVal, i)),
            I32(i) => Ok(variant!(VT_I4, lVal, i)),
            I32Ref(i) => Ok(variant!(VT_I4.byref(), plVal, i)),
            I64(i) => Ok(variant!(VT_I8, llVal, i)),
            I64Ref(i) => Ok(variant!(VT_I8.byref(), pllVal, i)),

            U8(i) => Ok(variant!(VT_UI1, bVal, i)),
            U8Ref(i) => Ok(variant!(VT_UI1.byref(), pbVal, i)),
            U16(i) => Ok(variant!(VT_UI2, uiVal, i)),
            U16Ref(i) => Ok(variant!(VT_UI2.byref(), puiVal, i)),
            U32(i) => Ok(variant!(VT_UI4, ulVal, i)),
            U32Ref(i) => Ok(variant!(VT_UI4.byref(), pulVal, i)),
            U64(i) => Ok(variant!(VT_UI8, ullVal, i)),
            U64Ref(i) => Ok(variant!(VT_UI8.byref(), pullVal, i)),

            F32(f) => Ok(variant!(VT_R4, fltVal, f)),
            F32Ref(f) => Ok(variant!(VT_R4.byref(), pfltVal, f)),
            F64(f) => Ok(variant!(VT_R8, dblVal, f)),
            F64Ref(f) => Ok(variant!(VT_R8.byref(), pdblVal, f)),

            Currency(d) => Ok(variant!(VT_CY, cyVal, ComCurrency::from(d).into())),
            CurrencyRef(r) => Ok(variant!(VT_CY.byref(), pcyVal, r.as_mut_ptr())),

            Decimal(d) => Ok(variant!(VT_DECIMAL, (decVal), ComDecimal::from(d).0)),
            DecimalRef(d) => Ok(variant!(VT_DECIMAL.byref(), pdecVal, d.as_mut_ptr())),

            Date(d) => Ok(variant!(VT_DATE, date, ComDate::from(d).0)),
            DateRef(d) => Ok(variant!(VT_DATE.byref(), pdate, d.as_mut_ptr())),

            String(s) => Ok(variant!(VT_BSTR, bstrVal, ManuallyDrop::new(s))),
            StringRef(s) => Ok(variant!(VT_BSTR.byref(), pbstrVal, s)),

            Dispatch(ptr) => Ok(variant!(VT_DISPATCH, pdispVal, ManuallyDrop::new(ptr))),
            Unknown(ptr) => Ok(variant!(VT_UNKNOWN, punkVal, ManuallyDrop::new(ptr))),

            Error(code) => Ok(variant!(VT_ERROR, scode, code.0)),
            ErrorRef(code) => Ok(variant!(VT_ERROR.byref(), pscode, &mut code.0)),

            VariantRef(ptr) => Ok(variant!(VT_VARIANT.byref(), pvarVal, ptr.0)),
            //_ => Err(VariantConversionError::GenericConversionError),
        }
    }
}
