use std::fmt::Debug;

pub struct PtrWrapper<T: 'static>(pub &'static mut T);

impl <T: 'static> Debug for PtrWrapper<T>
{
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result
    {
        write!(f, "ComIPtr({:p})", self.0)
    }
}

impl <T: 'static> PartialEq for PtrWrapper<T>
{
    fn eq(&self, other: &PtrWrapper<T>) -> bool
    {
        std::ptr::eq(self.0, other.0)
    }
}

impl <T: 'static> TryFrom<&*mut T> for PtrWrapper<T>
{
    type Error = ();

    fn try_from(ptr: &*mut T) -> Result<Self, Self::Error>
    {
        unsafe { ptr.as_mut() }
            .map(|ptr| PtrWrapper(ptr))
            .ok_or(())
    }
}