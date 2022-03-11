use crate::variant::*;

use crate::bool::ComBool;
use crate::ptr_wrapper::PtrWrapper;

mod bool;
mod ptr_wrapper;
mod variant;
mod convert;


#[cfg(test)]
mod tests
{
    use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
    use winapi::um::oaidl::VARIANT;

    use crate::{Variant, VariantType};

    #[macro_export]
    macro_rules! variant
    {
        ( $type: ident, $field: ident, $val: expr ) =>
        {
            unsafe {
                let mut variant: VARIANT = std::mem::zeroed();
                variant.n1.n2_mut().vt = VariantType::$type as u16;
                *variant.n1.n2_mut().n3.$field() = $val;
                variant.try_into()
            }
        };
    }

    #[test]
    fn it_works()
    {
        assert_eq!(variant!(VT_DATE, date_mut, 5.25),
                   Ok(Variant::Date(
                       NaiveDateTime::new(
                           NaiveDate::from_ymd(1900, 01, 04),
                           NaiveTime::from_hms(6, 0, 0)))));
    }
}
