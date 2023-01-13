//! Wrapper type for [`DECIMAL`]

use crate::VariantType;
use rust_decimal::Decimal;
use std::fmt::{Debug, Display};
use windows::Win32::Foundation::{DECIMAL, DECIMAL_0, DECIMAL_0_0, DECIMAL_1};

const DECIMAL_NEG: u8 = 0x80;

/// Transparent wrapper around a [`DECIMAL`] value
#[derive(Copy, Clone, PartialEq, Eq)]
pub struct ComDecimal(pub DECIMAL);

impl Display for ComDecimal {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", Into::<Decimal>::into(self))
    }
}

impl Debug for ComDecimal {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "ComDecimal({})", self)
    }
}

impl From<&Decimal> for ComDecimal {
    fn from(dec: &Decimal) -> Self {
        let unpack = dec.unpack();
        ComDecimal(DECIMAL {
            wReserved: VariantType::VT_DECIMAL as u16,
            Anonymous1: DECIMAL_0 {
                Anonymous: DECIMAL_0_0 {
                    scale: dec.scale() as u8,
                    sign: if dec.is_sign_positive() {
                        0
                    } else {
                        DECIMAL_NEG
                    },
                },
            },
            Hi32: unpack.hi,
            Anonymous2: DECIMAL_1 {
                Lo64: ((unpack.mid as u64) << 32) + unpack.lo as u64,
            },
        })
    }
}

impl From<Decimal> for ComDecimal {
    fn from(dec: Decimal) -> Self {
        Self::from(&dec)
    }
}

impl From<&ComDecimal> for Decimal {
    fn from(dec: &ComDecimal) -> Self {
        let num = dec.0;
        unsafe {
            Decimal::from_parts(
                (num.Anonymous2.Lo64 & 0xFFFFFFFF) as u32,
                ((num.Anonymous2.Lo64 >> 32) & 0xFFFFFFFF) as u32,
                num.Hi32,
                num.Anonymous1.Anonymous.sign == DECIMAL_NEG,
                num.Anonymous1.Anonymous.scale as u32,
            )
        }
    }
}

impl From<ComDecimal> for Decimal {
    fn from(dec: ComDecimal) -> Self {
        Self::from(&dec)
    }
}

impl From<*mut DECIMAL> for &mut ComDecimal {
    fn from(dec: *mut DECIMAL) -> Self {
        unsafe { &mut *(dec as *mut ComDecimal) }
    }
}

impl ComDecimal {
    pub fn as_mut_ptr(&mut self) -> *mut DECIMAL {
        (&mut self.0) as *mut DECIMAL
    }
}
