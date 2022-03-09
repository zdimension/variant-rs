use std::string::FromUtf16Error;
use enumn::N;
use rust_decimal::Decimal;
use winapi::shared::wtypes::VARTYPE;
use winapi::um::oaidl::VARIANT;
use winapi::um::winnt::HRESULT;

#[derive(PartialEq, Debug)]
pub enum Variant
{
    Empty,
    Null,
    Bool(bool),
    I8(i8),
    I16(i16),
    I32(i32),
    I64(i64),
    U8(u8),
    U16(u16),
    U32(u32),
    U64(u64),
    F32(f32),
    F64(f64),
    Currency(Decimal),
    String(String),
    Error(HRESULT)
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

pub enum VariantConversionError
{
    FromUtf16Error(FromUtf16Error),
    Unimplemented(VariantType),
    Unknown(VARTYPE),
}

impl TryInto<Variant> for VARIANT
{
    type Error = VariantConversionError;

    fn try_into(self) -> Result<Variant, VariantConversionError>
    {
        unsafe {
            let val = self.n1.n2();

            match VariantType::n(val.vt)
            {
                Some(t) => match t
                {
                    VariantType::VT_EMPTY => Ok(Variant::Empty),
                    VariantType::VT_NULL => Ok(Variant::Null),
                    VariantType::VT_BOOL => Ok(Variant::Bool(*val.n3.boolVal() != 0)),
                    VariantType::VT_I1 => Ok(Variant::I8(*val.n3.cVal())),
                    VariantType::VT_I2 => Ok(Variant::I16(*val.n3.iVal())),
                    VariantType::VT_I4 => Ok(Variant::I32(*val.n3.lVal())),
                    VariantType::VT_I8 => Ok(Variant::I64(*val.n3.llVal())),
                    VariantType::VT_UI1 => Ok(Variant::U8(*val.n3.bVal())),
                    VariantType::VT_UI2 => Ok(Variant::U16(*val.n3.uiVal())),
                    VariantType::VT_UI4 => Ok(Variant::U32(*val.n3.ulVal())),
                    VariantType::VT_UI8 => Ok(Variant::U64(*val.n3.ullVal())),
                    VariantType::VT_R4 => Ok(Variant::F32(*val.n3.fltVal())),
                    VariantType::VT_R8 => Ok(Variant::F64(*val.n3.dblVal())),
                    VariantType::VT_CY => Ok(Variant::Currency(Decimal::new((*val.n3.cyVal()).int64, 4))),
                    VariantType::VT_BSTR => widestring::U16CString::from_ptr_str(*val.n3.bstrVal()).to_string()
                        .map(|s| Variant::String(s))
                        .map_err(VariantConversionError::FromUtf16Error),
                    VariantType::VT_ERROR => Ok(Variant::Error(*val.n3.scode())),
                    _ => Err(VariantConversionError::Unimplemented(t))
                },
                None => Err(VariantConversionError::Unknown(val.vt))
            }
        }
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn it_works() {
        let result = 2 + 2;
        assert_eq!(result, 4);
    }
}
