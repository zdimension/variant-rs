//! Wrapper type for [`BOOL`]

#![allow(unused_imports)]
use windows::Win32::Foundation::{BOOL, VARIANT_BOOL};

use enumn::N;

/// Enum equivalent to COM [`BOOL`]. False is 0 and True is all ones (-1)
#[derive(N, Clone, Copy, Debug, Eq, Hash, Ord, PartialEq, PartialOrd)]
#[repr(i16)]
pub enum ComBool {
    False = 0i16,
    True = !0i16,
}

impl From<&mut i16> for &'static mut ComBool {
    fn from(value: &mut i16) -> &'static mut ComBool {
        unsafe { &mut *(value as *mut i16 as *mut ComBool) }
    }
}

impl From<&VARIANT_BOOL> for &'static mut ComBool {
    fn from(value: &VARIANT_BOOL) -> &'static mut ComBool {
        unsafe { &mut *(value as *const VARIANT_BOOL as *mut ComBool) }
    }
}

impl From<ComBool> for bool {
    fn from(value: ComBool) -> bool {
        value != ComBool::False
    }
}

impl From<bool> for ComBool {
    fn from(value: bool) -> ComBool {
        if value {
            ComBool::True
        } else {
            ComBool::False
        }
    }
}
