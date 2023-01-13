//! Wrapper type for [`CY`]

use rust_decimal::Decimal;
use std::fmt::{Debug, Display};
use windows::Win32::System::Com::CY;

/// Transparent wrapper around a [`CY`] value stored as an [`i64`]
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd)]
pub struct ComCurrency(pub i64);

/// Wrapper around COM [`CY`] using [`Decimal`].
#[derive(Clone, Copy, PartialEq, Eq, PartialOrd, Debug)]
pub struct Currency(pub Decimal);

impl AsRef<Decimal> for Currency {
    fn as_ref(&self) -> &Decimal {
        &self.0
    }
}

impl Display for ComCurrency {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "{}", Into::<Decimal>::into(*self))
    }
}

impl Debug for ComCurrency {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        write!(f, "ComCurrency({}, {})", self.0, self)
    }
}

impl From<Decimal> for ComCurrency {
    fn from(cy: Decimal) -> Self {
        let mut g = cy;
        g.rescale(4);
        ComCurrency(g.mantissa() as i64)
    }
}

impl From<Currency> for ComCurrency {
    fn from(cy: Currency) -> Self {
        Self::from(cy.0)
    }
}

impl From<ComCurrency> for Decimal {
    fn from(cy: ComCurrency) -> Self {
        Decimal::new(cy.0, 4)
    }
}

impl From<ComCurrency> for Currency {
    fn from(cy: ComCurrency) -> Self {
        Currency(cy.into())
    }
}

impl From<Decimal> for Currency {
    fn from(cy: Decimal) -> Self {
        Currency(cy)
    }
}

impl From<ComCurrency> for CY {
    fn from(cy: ComCurrency) -> Self {
        CY { int64: cy.0 }
    }
}

impl From<CY> for ComCurrency {
    fn from(cy: CY) -> Self {
        ComCurrency(unsafe { cy.int64 })
    }
}

impl From<*mut CY> for &mut ComCurrency {
    fn from(cy: *mut CY) -> Self {
        unsafe { &mut *(cy as *mut ComCurrency) }
    }
}

impl ComCurrency {
    pub fn as_mut_ptr(&mut self) -> *mut CY {
        (&mut self.0) as *mut i64 as *mut CY
    }
}
