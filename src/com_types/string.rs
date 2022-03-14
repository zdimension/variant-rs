use std::fmt::{Debug, Display};
use std::string::FromUtf16Error;
use widestring::error::ContainsNul;

#[derive(Clone, Copy, Eq, Hash, Ord, PartialEq, PartialOrd)]
pub struct ComString(pub *mut u16);

impl Display for ComString
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "{}", match TryInto::<String>::try_into(*self)
        {
            Ok(s) => s,
            Err(_) => "<invalid>".to_string(),
        })
    }
}

impl Debug for ComString
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "ComString({:p}, {})", self.0, match TryInto::<String>::try_into(*self)
        {
            Ok(s) => format!("\"{}\"", s),
            Err(_) => "<invalid>".to_string(),
        })
    }
}

impl TryFrom<String> for ComString
{
    type Error = ContainsNul<u16>;

    fn try_from(s: String) -> Result<Self, Self::Error>
    {
        widestring::U16CString::from_str(s).map(|s| unsafe { ComString(s.into_raw()) })
    }
}

impl TryFrom<ComString> for String
{
    type Error = FromUtf16Error;

    fn try_from(s: ComString) -> Result<Self, Self::Error>
    {
        unsafe { widestring::U16CString::from_ptr_str(s.0).to_string() }
    }
}

impl ComString
{
    pub fn as_mut_ptr(&mut self) -> *mut *mut u16
    {
        &mut self.0
    }
}