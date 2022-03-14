use std::fmt::{Debug, Display};
use std::ops::Sub;
use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};

#[derive(Clone, Copy, PartialEq, PartialOrd)]
pub struct ComDate(pub f64);

impl Display for ComDate
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "{}", match TryInto::<NaiveDateTime>::try_into(*self)
        {
            Ok(s) => s.to_string(),
            Err(_) => "<invalid>".to_owned(),
        })
    }
}

impl Debug for ComDate
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "ComString({}, {})", self.0, self)
    }
}

macro_rules! com_epoch
{
    () => { NaiveDateTime::new(NaiveDate::from_ymd(1899, 12, 30), NaiveTime::from_hms(0, 0, 0)) }
}

impl From<NaiveDateTime> for ComDate
{
    fn from(date: NaiveDateTime) -> Self
    {
        ComDate(date.sub(com_epoch!()).num_milliseconds() as f64 / 24.0 / 60.0 / 60.0 / 1000.0)
    }
}

impl From<ComDate> for NaiveDateTime
{
    fn from(date: ComDate) -> Self
    {
        com_epoch!() + Duration::milliseconds((date.0 * 24.0 * 60.0 * 60.0 * 1000.0) as i64)
    }
}