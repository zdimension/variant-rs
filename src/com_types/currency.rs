use std::fmt::{Debug, Display};
use rust_decimal::Decimal;
use winapi::shared::wtypes::CY;

#[derive(Clone, Copy, PartialEq, PartialOrd)]
pub struct ComCurrency(pub i64);

impl Display for ComCurrency
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "{}", Into::<Decimal>::into(*self))
    }
}

impl Debug for ComCurrency
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "ComCurrency({}, {})", self.0, self)
    }
}

impl From<Decimal> for ComCurrency
{
    fn from(cy: Decimal) -> Self
    {
        let mut g = cy;
        g.rescale(4);
        ComCurrency(g.mantissa() as i64)
    }
}

impl From<ComCurrency> for Decimal
{
    fn from(cy: ComCurrency) -> Self
    {
        Decimal::new(cy.0, 4)
    }
}

impl From<ComCurrency> for CY
{
    fn from(cy: ComCurrency) -> Self
    {
        CY { int64: cy.0 }
    }
}

impl From<CY> for ComCurrency
{
    fn from(cy: CY) -> Self
    {
        ComCurrency(cy.int64)
    }
}

impl From<*mut CY> for &mut ComCurrency
{
    fn from(cy: *mut CY) -> Self
    {
        unsafe { &mut *(cy as *mut ComCurrency) }
    }
}

impl ComCurrency
{
    pub fn as_mut_ptr(&mut self) -> *mut CY
    {
        (&mut self.0) as *mut i64 as *mut CY
    }
}