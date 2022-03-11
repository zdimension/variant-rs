use enumn::N;

#[derive(N, PartialEq, Debug)]
#[repr(i16)]
pub enum ComBool
{
    False = 0i16,
    True = !0i16,
}

impl From<&mut i16> for &'static mut ComBool
{
    fn from(value: &mut i16) -> &'static mut ComBool
    {
        unsafe { &mut *(value as *mut i16 as *mut ComBool) }
    }
}

impl From<ComBool> for bool
{
    fn from(value: ComBool) -> bool
    {
        value != ComBool::False
    }
}

impl From<bool> for ComBool
{
    fn from(value: bool) -> ComBool
    {
        if value { ComBool::True } else { ComBool::False }
    }
}