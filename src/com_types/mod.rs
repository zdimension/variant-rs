//! Utilities for handling COM types ([`BOOL`], [`CY`], [`DECIMAL`], etc.)

#![allow(unused_imports)]
use windows::Win32::Foundation::{BOOL, DECIMAL};
use windows::Win32::System::Com::CY;

pub mod bool;
pub mod currency;
pub mod date;
pub mod decimal;
pub mod ptr_wrapper;
