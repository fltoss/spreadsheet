//! Encoding [`CellValue`]s as Elixir terms: `nil`, floats, booleans, binaries,
//! `{:error, code}` tuples and `%NaiveDateTime{}`, `%Time{}` and `%Duration{}`
//! structs built with the fields Elixir's own constructors produce.

use crate::cell::{CellValue, Duration};
use chrono::{Datelike, NaiveDateTime, NaiveTime, Timelike};
use rustler::{Atom, Encoder, Env, NifStruct, Term};

mod atoms {
    rustler::atoms! {
        calendar_iso = "Elixir.Calendar.ISO",
    }
}

#[derive(NifStruct)]
#[module = "NaiveDateTime"]
struct ElixirNaiveDateTime {
    calendar: Atom,
    year: i32,
    month: u32,
    day: u32,
    hour: u32,
    minute: u32,
    second: u32,
    microsecond: (u32, u32),
}

#[derive(NifStruct)]
#[module = "Time"]
struct ElixirTime {
    calendar: Atom,
    hour: u32,
    minute: u32,
    second: u32,
    microsecond: (u32, u32),
}

#[derive(NifStruct)]
#[module = "Duration"]
struct ElixirDuration {
    year: i64,
    month: i64,
    week: i64,
    day: i64,
    hour: i64,
    minute: i64,
    second: i64,
    microsecond: (i64, u32),
}

/// Encodes one cell. `nil` is passed in so callers can encode it once per sheet.
pub fn encode_cell<'a>(env: Env<'a>, value: CellValue<'_>, nil: Term<'a>) -> Term<'a> {
    match value {
        CellValue::Nil => nil,
        CellValue::Float(number) => number.encode(env),
        CellValue::Bool(boolean) => boolean.encode(env),
        CellValue::Text(text) => text.encode(env),
        CellValue::NaiveDateTime(date_time, precision) => {
            naive_date_time(date_time, precision).encode(env)
        }
        CellValue::Time(time, precision) => elixir_time(time, precision).encode(env),
        CellValue::Duration(duration) => elixir_duration(duration).encode(env),
        CellValue::Error(code) => (rustler::types::atom::error(), code).encode(env),
    }
}

fn microsecond(nanosecond: u32, precision: u32) -> (u32, u32) {
    (nanosecond / 1000, precision)
}

fn naive_date_time(date_time: NaiveDateTime, precision: u32) -> ElixirNaiveDateTime {
    ElixirNaiveDateTime {
        calendar: atoms::calendar_iso(),
        year: date_time.year(),
        month: date_time.month(),
        day: date_time.day(),
        hour: date_time.hour(),
        minute: date_time.minute(),
        second: date_time.second(),
        microsecond: microsecond(date_time.nanosecond(), precision),
    }
}

fn elixir_time(time: NaiveTime, precision: u32) -> ElixirTime {
    ElixirTime {
        calendar: atoms::calendar_iso(),
        hour: time.hour(),
        minute: time.minute(),
        second: time.second(),
        microsecond: microsecond(time.nanosecond(), precision),
    }
}

fn elixir_duration(duration: Duration) -> ElixirDuration {
    ElixirDuration {
        year: duration.year,
        month: duration.month,
        week: duration.week,
        day: duration.day,
        hour: duration.hour,
        minute: duration.minute,
        second: duration.second,
        microsecond: (duration.microsecond, duration.precision),
    }
}
