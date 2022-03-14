use std::fmt::{Debug, Display};
use rust_decimal::Decimal;
use winapi::shared::wtypes::{DECIMAL, DECIMAL_NEG};
use crate::VariantType;

#[derive(Copy)]
pub struct ComDecimal(pub DECIMAL);

impl Display for ComDecimal
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "{}", Into::<Decimal>::into(self))
    }
}

impl Debug for ComDecimal
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "ComDecimal({})", self)
    }
}

impl Clone for ComDecimal
{
    fn clone(&self) -> Self
    {
        ComDecimal(DECIMAL
        {
            wReserved: self.0.wReserved,
            scale: self.0.scale,
            sign: self.0.sign,
            Hi32: self.0.Hi32,
            Lo64: self.0.Lo64,
        })
    }
}

impl From<&Decimal> for ComDecimal
{
    fn from(dec: &Decimal) -> Self
    {
        let unpack = dec.unpack();
        ComDecimal(DECIMAL {
            wReserved: VariantType::VT_DECIMAL as u16,
            scale: dec.scale() as u8,
            sign: if dec.is_sign_positive() { 0 } else { DECIMAL_NEG },
            Hi32: unpack.hi,
            Lo64: ((unpack.mid as u64) << 32) + unpack.lo as u64,
        })
    }
}

impl From<Decimal> for ComDecimal
{
    fn from(dec: Decimal) -> Self
    {
        Self::from(&dec)
    }
}

impl From<&ComDecimal> for Decimal
{
    fn from(dec: &ComDecimal) -> Self
    {
        let num = dec.0;
        Decimal::from_parts((num.Lo64 & 0xFFFFFFFF) as u32,
                            ((num.Lo64 >> 32) & 0xFFFFFFFF) as u32,
                            num.Hi32,
                            num.sign == DECIMAL_NEG,
                            num.scale as u32)
    }
}

impl From<ComDecimal> for Decimal
{
    fn from(dec: ComDecimal) -> Self
    {
        Self::from(&dec)
    }
}

impl From<*mut DECIMAL> for &mut ComDecimal
{
    fn from(dec: *mut DECIMAL) -> Self
    {
        unsafe { &mut *(dec as *mut ComDecimal) }
    }
}

impl PartialEq for ComDecimal
{
    fn eq(&self, other: &ComDecimal) -> bool
    {
        (self.0.scale, self.0.sign, self.0.Hi32, self.0.Lo64) ==
            (other.0.scale, other.0.sign, other.0.Hi32, other.0.Lo64)
    }
}

impl ComDecimal
{
    pub fn as_mut_ptr(&mut self) -> *mut DECIMAL
    {
        (&mut self.0) as *mut DECIMAL
    }
}