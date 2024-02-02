# variant-rs

[![Crates.io](https://img.shields.io/crates/v/variant-rs)](https://crates.io/crates/variant-rs)
[![Crates.io](https://img.shields.io/crates/d/variant-rs)](https://crates.io/crates/variant-rs)
[![License](https://img.shields.io/badge/license-Apache%202.0-blue)](https://github.com/zdimension/variant-rs/blob/master/LICENSE-APACHE)
[![License](https://img.shields.io/badge/license-MIT-blue)](https://github.com/zdimension/variant-rs/blob/master/LICENSE-MIT)

`variant-rs` is a Rust crate that provides idiomatic handling of COM `VARIANT` types. Rust supports discriminated
union types out of the box, so although `VARIANT`s are usually a pain to work with, Rust makes it easy to encode and
decode them.

The crate is designed to work with the `VARIANT` type from the [`winapi`](https://crates.io/crates/winapi) crate.

## Basic usage
```rust
use variant_rs::*;

fn main() {
    let v1 = Variant::I32(123); // manual instanciation
    let v2 = 123i32.to_variant(); // ToVariant trait
    let v3 = 123.into(); // From / Into traits
    assert_eq!(v1, v2);
    assert_eq!(v1, v3);
  
    let bstr: Variant = "Hello, world!".into();
    let ptr: VARIANT = bstr.clone().try_into().unwrap(); // convert to COM VARIANT
    let back: Variant = ptr.try_into().unwrap(); // convert back
    assert_eq!(bstr, back);
}
```

## Supported `VARIANT` types and corresponding types
| `VARIANT` type | Rust type           | Rust type (BY_REF)             |
|----------------|---------------------|--------------------------------|
| `VT_EMPTY`     | `()`                | N/A                            |
| `VT_NULL`      | `()`                | N/A                            |
| `VT_I1`        | `i8`                | `PSTR`                         |
| `VT_I2`        | `i16`               | `&'static mut i16`             |
| `VT_I4`        | `i32`               | `&'static mut i32`             |
| `VT_I8`        | `i64`               | `&'static mut i64`             |
| `VT_UI1`       | `u8`                | `&'static mut u8`              |
| `VT_UI2`       | `u16`               | `&'static mut u16`             |
| `VT_UI4`       | `u32`               | `&'static mut u32`             |
| `VT_UI8`       | `u64`               | `&'static mut u64`             |
| `VT_INT`       | `i32`               | `&'static mut i32`             |
| `VT_UINT`      | `u32`               | `&'static mut u32`             |
| `VT_R4`        | `f32`               | `&'static mut f32`             |
| `VT_R8`        | `f64`               | `&'static mut f64`             |
| `VT_BOOL`      | `bool`              | `&'static mut ComBool`         |
| `VT_BSTR`      | `BSTR`              | `&'static mut BSTR`            |
| `VT_ERROR`     | `HRESULT` (`i32`)   | `&'static mut HRESULT` (`i32`) |
| `VT_CY`        | `Currency`          | `&'static mut ComCurrency`     |
| `VT_DATE`      | `NaiveDateTime`     | `&'static mut ComDate`         |
| `VT_DECIMAL`   | `Decimal`           | `&'static mut ComDecimal`      |
| `VT_UNKNOWN`   | `Option<IUnknown>`  | N/A                            |
| `VT_DISPATCH`  | `Option<IDispatch>` | N/A                            |
| `VT_VARIANT`   | N/A                 | `PtrWrapper<VARIANT>`          |

## Wrapper types

### `ComBool`
`i16`-backed enum.

### `ComCurrency`
Maps COM's `i64` currency data [`CY`](https://docs.microsoft.com/en-us/windows/win32/api/wtypes/ns-wtypes-cy-r1) to [`Decimal`](https://docs.rs/rust_decimal/latest/rust_decimal/struct.Decimal.html).

### `ComDecimal`
Maps COM's 96-bit decimals [`DECIMAL`](https://docs.microsoft.com/en-us/windows/win32/api/wtypes/ns-wtypes-decimal-r1) to [`Decimal`](https://docs.rs/rust_decimal/latest/rust_decimal/struct.Decimal.html).

### `ComData`
Maps COM's [`DATE`](https://docs.microsoft.com/en-us/cpp/atl-mfc-shared/date-type?view=msvc-170) (`f64` milliseconds from 1899-12-30) to [`NaiveDateTime`](https://docs.rs/chrono/latest/chrono/naive/struct.NaiveDateTime.html).

### `PtrWrapper`
Safe wrapper around COM interface pointers.

## Installation
Add this to your `Cargo.toml`:
```toml
[dependencies]
variant-rs = "0.4.0"
```

## License
This project is licensed under either of
* Apache License, Version 2.0, ([LICENSE-APACHE](LICENSE-APACHE) or
  <https://www.apache.org/licenses/LICENSE-2.0>)
* MIT license ([LICENSE-MIT](LICENSE-MIT) or
  <https://opensource.org/licenses/MIT>)
  at your option.