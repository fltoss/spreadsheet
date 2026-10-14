//! Converting calamine cells into the Elixir values `Spreadsheet` returns.
//!
//! The conversion is pure (no BEAM terms) so it can be unit tested; `encode`
//! turns a [`CellValue`] into a term.

use calamine::{CellErrorType, Data, DataRef, ExcelDateTime};
use chrono::{Datelike, NaiveDate, NaiveDateTime, NaiveTime, Timelike};

const MILLISECONDS_PER_DAY: f64 = 86_400_000.0;

/// Largest millisecond count converted to a duration. Beyond 2^53 an `f64`
/// no longer holds whole milliseconds exactly.
const MAX_DURATION_MILLISECONDS: f64 = 9_007_199_254_740_992.0;

/// The value a cell becomes in Elixir.
#[derive(Debug, PartialEq)]
pub enum CellValue<'a> {
    /// `nil`
    Nil,
    /// A float. Every number is a float, whatever the source format.
    Float(f64),
    /// `true` / `false`
    Bool(bool),
    /// A binary
    Text(&'a str),
    /// `%NaiveDateTime{}`
    NaiveDateTime(NaiveDateTime, Precision),
    /// `%Time{}`
    Time(NaiveTime, Precision),
    /// `%Duration{}`
    Duration(Duration),
    /// `{:error, "#DIV/0!"}` and friends
    Error(&'static str),
}

/// Number of fractional-second digits, as in Elixir's `{microsecond, precision}`.
pub type Precision = u32;

/// The fields of an Elixir `%Duration{}`. Every component carries the sign of
/// the duration.
#[derive(Debug, Default, PartialEq)]
pub struct Duration {
    pub year: i64,
    pub month: i64,
    pub week: i64,
    pub day: i64,
    pub hour: i64,
    pub minute: i64,
    pub second: i64,
    pub microsecond: i64,
    pub precision: Precision,
}

pub trait ToCellValue {
    fn to_cell_value(&self) -> CellValue<'_>;
}

impl ToCellValue for Data {
    fn to_cell_value(&self) -> CellValue<'_> {
        match self {
            Data::Int(value) => float(*value as f64),
            Data::Float(value) => float(*value),
            Data::String(value) => CellValue::Text(value),
            Data::Bool(value) => CellValue::Bool(*value),
            Data::DateTime(value) => excel_date_time(value),
            Data::DateTimeIso(value) => iso_date_time(value),
            Data::DurationIso(value) => iso_duration(value),
            Data::Error(error) => CellValue::Error(error_code(error)),
            Data::Empty => CellValue::Nil,
        }
    }
}

impl ToCellValue for DataRef<'_> {
    fn to_cell_value(&self) -> CellValue<'_> {
        match self {
            DataRef::Int(value) => float(*value as f64),
            DataRef::Float(value) => float(*value),
            DataRef::String(value) => CellValue::Text(value),
            DataRef::SharedString(value) => CellValue::Text(value),
            DataRef::Bool(value) => CellValue::Bool(*value),
            DataRef::DateTime(value) => excel_date_time(value),
            DataRef::DateTimeIso(value) => iso_date_time(value),
            DataRef::DurationIso(value) => iso_duration(value),
            DataRef::Error(error) => CellValue::Error(error_code(error)),
            DataRef::Empty => CellValue::Nil,
        }
    }
}

/// NaN and infinity cannot be represented as Elixir floats, so they surface as
/// Excel's `#NUM!` error.
fn float(value: f64) -> CellValue<'static> {
    if value.is_finite() {
        CellValue::Float(value)
    } else {
        CellValue::Error(error_code(&CellErrorType::Num))
    }
}

fn error_code(error: &CellErrorType) -> &'static str {
    match error {
        CellErrorType::Div0 => "#DIV/0!",
        CellErrorType::NA => "#N/A",
        CellErrorType::Name => "#NAME?",
        CellErrorType::Null => "#NULL!",
        CellErrorType::Num => "#NUM!",
        CellErrorType::Ref => "#REF!",
        CellErrorType::Value => "#VALUE!",
        CellErrorType::GettingData => "#DATA!",
    }
}

/// Excel (xlsx/xls/xlsb) serial date-times.
///
/// - Duration-formatted cells (e.g. `[h]:mm:ss`) become a `%Duration{}`,
///   keeping their sign.
/// - Serials in `[0, 1)` are a time of day and become a `%Time{}`.
/// - Other serials become a `%NaiveDateTime{}`.
/// - Serials that are negative or fall outside years 0..=9999 cannot be
///   represented, so the raw serial is returned as a float.
fn excel_date_time(value: &ExcelDateTime) -> CellValue<'static> {
    let serial = value.as_f64();
    if !serial.is_finite() {
        return float(serial);
    }

    let milliseconds = (serial * MILLISECONDS_PER_DAY).round();

    if value.is_duration() {
        return if milliseconds.abs() < MAX_DURATION_MILLISECONDS {
            CellValue::Duration(duration_from_milliseconds(milliseconds as i64))
        } else {
            CellValue::Float(serial)
        };
    }

    if serial < 0.0 {
        return CellValue::Float(serial);
    }

    if milliseconds < MILLISECONDS_PER_DAY {
        let milliseconds = milliseconds as u32;
        let time = NaiveTime::from_num_seconds_from_midnight_opt(
            milliseconds / 1000,
            (milliseconds % 1000) * 1_000_000,
        );
        if let Some(time) = time {
            return CellValue::Time(time, millisecond_precision(time.nanosecond()));
        }
    }

    match value.as_datetime() {
        Some(date_time) if (0..=9999).contains(&date_time.year()) => {
            CellValue::NaiveDateTime(date_time, millisecond_precision(date_time.nanosecond()))
        }
        _ => CellValue::Float(serial),
    }
}

/// Excel stores milliseconds at most; whole seconds keep Elixir's default
/// precision of 0, as `~N[2024-01-01 00:00:00]` does.
fn millisecond_precision(nanosecond: u32) -> Precision {
    if nanosecond == 0 {
        0
    } else {
        3
    }
}

fn duration_from_milliseconds(milliseconds: i64) -> Duration {
    let sign = if milliseconds < 0 { -1 } else { 1 };
    let magnitude = milliseconds.abs();
    let remainder = magnitude % 1000;

    Duration {
        hour: sign * (magnitude / 3_600_000),
        minute: sign * (magnitude / 60_000 % 60),
        second: sign * (magnitude / 1000 % 60),
        microsecond: sign * remainder * 1000,
        precision: if remainder == 0 { 0 } else { 3 },
        ..Duration::default()
    }
}

/// ODS `office:date-value`: an ISO 8601 date or date-time. Values that do not
/// parse (or fall outside years 0..=9999) are returned as the raw string.
fn iso_date_time(value: &str) -> CellValue<'_> {
    match parse_iso_date_time(value) {
        Some((date_time, precision)) => CellValue::NaiveDateTime(date_time, precision),
        None => CellValue::Text(value),
    }
}

/// ODS `office:time-value`: an ISO 8601 duration such as `PT36H00M00S`.
/// ODS does not distinguish a time of day from a duration, so these are always
/// durations. Values that do not parse are returned as the raw string.
fn iso_duration(value: &str) -> CellValue<'_> {
    match parse_iso_duration(value) {
        Some(duration) => CellValue::Duration(duration),
        None => CellValue::Text(value),
    }
}

/// Parses `YYYY-MM-DD`, optionally followed by `THH:MM:SS[.fraction]` (or a
/// space instead of `T`) and a `Z` or `±HH:MM` offset, which is ignored as
/// `NaiveDateTime.from_iso8601/1` does. A date alone becomes midnight.
pub fn parse_iso_date_time(value: &str) -> Option<(NaiveDateTime, Precision)> {
    let value = strip_utc_offset(value);
    let (date, time) = match value.len() {
        10 => (value, None),
        length if length > 11 && matches!(value.as_bytes()[10], b'T' | b' ') => {
            (&value[..10], Some(&value[11..]))
        }
        _ => return None,
    };

    let date = parse_iso_date(date)?;
    let (time, precision) = match time {
        Some(time) => parse_iso_time(time)?,
        None => (NaiveTime::MIN, 0),
    };
    Some((date.and_time(time), precision))
}

fn strip_utc_offset(value: &str) -> &str {
    if let Some(stripped) = value.strip_suffix('Z') {
        return stripped;
    }
    let bytes = value.as_bytes();
    let length = bytes.len();
    // Only look for an offset after a time part, so "2024-01-15" is untouched.
    if length > 16
        && matches!(bytes[length - 6], b'+' | b'-')
        && bytes[length - 3] == b':'
        && all_digits(&value[length - 5..length - 3])
        && all_digits(&value[length - 2..])
    {
        return &value[..length - 6];
    }
    value
}

fn parse_iso_date(date: &str) -> Option<NaiveDate> {
    let bytes = date.as_bytes();
    if bytes.len() != 10 || bytes[4] != b'-' || bytes[7] != b'-' {
        return None;
    }
    NaiveDate::from_ymd_opt(
        parse_digits(date.get(0..4)?)? as i32,
        parse_digits(date.get(5..7)?)? as u32,
        parse_digits(date.get(8..10)?)? as u32,
    )
}

fn parse_iso_time(time: &str) -> Option<(NaiveTime, Precision)> {
    let bytes = time.as_bytes();
    if bytes.len() < 8 || bytes[2] != b':' || bytes[5] != b':' {
        return None;
    }
    let hour = parse_digits(time.get(0..2)?)? as u32;
    let minute = parse_digits(time.get(3..5)?)? as u32;
    let second = parse_digits(time.get(6..8)?)? as u32;
    let (microsecond, precision) = match time.get(8..)? {
        "" => (0, 0),
        fraction if fraction.starts_with(['.', ',']) => parse_fraction(&fraction[1..])?,
        _ => return None,
    };
    if second > 59 {
        return None;
    }
    let time = NaiveTime::from_hms_micro_opt(hour, minute, second, microsecond as u32)?;
    Some((time, precision))
}

/// Parses `[-]P[nY][nM][nW][nD][T[nH][nM][n[.fraction]S]]`. A leading `-`
/// negates every component.
pub fn parse_iso_duration(value: &str) -> Option<Duration> {
    let (sign, rest) = match value.as_bytes().first()? {
        b'-' => (-1, &value[1..]),
        b'+' => (1, &value[1..]),
        _ => (1, value),
    };
    let rest = rest.strip_prefix('P')?;
    let (date_part, time_part) = match rest.split_once('T') {
        Some((date_part, time_part)) if !time_part.is_empty() => (date_part, Some(time_part)),
        Some(_) => return None,
        None => (rest, None),
    };
    if date_part.is_empty() && time_part.is_none() {
        return None;
    }

    let mut duration = Duration::default();
    for (number, unit) in duration_components(date_part)? {
        let number = parse_signed_integer(number)?;
        match unit {
            'Y' => duration.year = number,
            'M' => duration.month = number,
            'W' => duration.week = number,
            'D' => duration.day = number,
            _ => return None,
        }
    }
    for (number, unit) in duration_components(time_part.unwrap_or(""))? {
        match unit {
            'H' => duration.hour = parse_signed_integer(number)?,
            'M' => duration.minute = parse_signed_integer(number)?,
            'S' => {
                let (whole, fraction) = match number.split_once(['.', ',']) {
                    Some((whole, fraction)) => (whole, Some(fraction)),
                    None => (number, None),
                };
                duration.second = parse_signed_integer(whole)?;
                if let Some(fraction) = fraction {
                    let (microsecond, precision) = parse_fraction(fraction)?;
                    duration.microsecond = if whole.starts_with('-') {
                        -microsecond
                    } else {
                        microsecond
                    };
                    duration.precision = precision;
                }
            }
            _ => return None,
        }
    }

    if sign < 0 {
        duration.year = -duration.year;
        duration.month = -duration.month;
        duration.week = -duration.week;
        duration.day = -duration.day;
        duration.hour = -duration.hour;
        duration.minute = -duration.minute;
        duration.second = -duration.second;
        duration.microsecond = -duration.microsecond;
    }
    Some(duration)
}

/// Splits `36H00M00S` into `[("36", 'H'), ("00", 'M'), ("00", 'S')]`.
fn duration_components(part: &str) -> Option<Vec<(&str, char)>> {
    let mut components = Vec::new();
    let mut start = 0;
    for (index, character) in part.char_indices() {
        if character.is_ascii_alphabetic() {
            if index == start {
                return None;
            }
            components.push((&part[start..index], character));
            start = index + 1;
        }
    }
    if start != part.len() {
        return None;
    }
    Some(components)
}

fn parse_signed_integer(number: &str) -> Option<i64> {
    let (negative, digits) = match number.strip_prefix('-') {
        Some(digits) => (true, digits),
        None => (false, number),
    };
    let magnitude = parse_digits(digits)?;
    Some(if negative { -magnitude } else { magnitude })
}

fn parse_digits(digits: &str) -> Option<i64> {
    if digits.is_empty() || !all_digits(digits) {
        return None;
    }
    digits.parse().ok()
}

fn all_digits(text: &str) -> bool {
    text.bytes().all(|byte| byte.is_ascii_digit())
}

/// Fractional seconds: precision is the number of digits (at most 6), and
/// digits beyond microseconds are truncated, as Elixir does.
fn parse_fraction(fraction: &str) -> Option<(i64, Precision)> {
    if fraction.is_empty() || !all_digits(fraction) {
        return None;
    }
    let digits = &fraction[..fraction.len().min(6)];
    let precision = digits.len() as Precision;
    let microsecond = digits.parse::<i64>().ok()? * 10_i64.pow(6 - precision);
    Some((microsecond, precision))
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::ExcelDateTimeType;

    fn date_time(serial: f64) -> CellValue<'static> {
        excel_date_time(&ExcelDateTime::new(
            serial,
            ExcelDateTimeType::DateTime,
            false,
        ))
    }

    fn duration(serial: f64) -> CellValue<'static> {
        excel_date_time(&ExcelDateTime::new(
            serial,
            ExcelDateTimeType::TimeDelta,
            false,
        ))
    }

    fn naive(text: &str) -> NaiveDateTime {
        NaiveDateTime::parse_from_str(text, "%Y-%m-%d %H:%M:%S%.f").unwrap()
    }

    #[test]
    fn serials_become_naive_date_times() {
        // 2025-10-13 12:00:00 is serial 45943.5.
        assert_eq!(
            date_time(45943.5),
            CellValue::NaiveDateTime(naive("2025-10-13 12:00:00"), 0)
        );
        assert_eq!(
            date_time(45943.5 + 0.123 / 86_400.0),
            CellValue::NaiveDateTime(naive("2025-10-13 12:00:00.123"), 3)
        );
    }

    #[test]
    fn epoch_1904_serials_are_offset() {
        let value = ExcelDateTime::new(43830.0, ExcelDateTimeType::DateTime, true);
        assert_eq!(
            excel_date_time(&value),
            CellValue::NaiveDateTime(naive("2024-01-01 00:00:00"), 0)
        );
    }

    #[test]
    fn fractions_of_a_day_become_times() {
        assert_eq!(
            date_time(0.5),
            CellValue::Time(NaiveTime::from_hms_opt(12, 0, 0).unwrap(), 0)
        );
        assert_eq!(date_time(0.0), CellValue::Time(NaiveTime::MIN, 0));
        assert_eq!(
            date_time(0.75 + 0.25 / 86_400.0),
            CellValue::Time(NaiveTime::from_hms_milli_opt(18, 0, 0, 250).unwrap(), 3)
        );
    }

    #[test]
    fn unrepresentable_dates_return_the_serial() {
        assert_eq!(date_time(-1.0), CellValue::Float(-1.0));
        assert_eq!(date_time(3_000_000.0), CellValue::Float(3_000_000.0));
        assert_eq!(date_time(1.0e300), CellValue::Float(1.0e300));
        // 9999-12-31 is the last representable day.
        assert_eq!(
            date_time(2_958_465.0),
            CellValue::NaiveDateTime(naive("9999-12-31 00:00:00"), 0)
        );
        assert_eq!(date_time(2_958_466.0), CellValue::Float(2_958_466.0));
    }

    #[test]
    fn durations_keep_their_sign() {
        assert_eq!(
            duration(0.0625),
            CellValue::Duration(Duration {
                hour: 1,
                minute: 30,
                ..Duration::default()
            })
        );
        assert_eq!(
            duration(-0.0625),
            CellValue::Duration(Duration {
                hour: -1,
                minute: -30,
                ..Duration::default()
            })
        );
        assert_eq!(
            duration(1.0625 + 1.5 / 86_400.0),
            CellValue::Duration(Duration {
                hour: 25,
                minute: 30,
                second: 1,
                microsecond: 500_000,
                precision: 3,
                ..Duration::default()
            })
        );
        assert_eq!(duration(1.0e300), CellValue::Float(1.0e300));
    }

    #[test]
    fn numbers_are_floats() {
        assert_eq!(Data::Int(1234).to_cell_value(), CellValue::Float(1234.0));
        assert_eq!(Data::Float(0.5).to_cell_value(), CellValue::Float(0.5));
        assert_eq!(
            Data::Float(f64::NAN).to_cell_value(),
            CellValue::Error("#NUM!")
        );
        assert_eq!(
            Data::Float(f64::INFINITY).to_cell_value(),
            CellValue::Error("#NUM!")
        );
    }

    #[test]
    fn scalars() {
        assert_eq!(Data::Empty.to_cell_value(), CellValue::Nil);
        assert_eq!(Data::Bool(true).to_cell_value(), CellValue::Bool(true));
        assert_eq!(
            Data::String("Zoë".into()).to_cell_value(),
            CellValue::Text("Zoë")
        );
        assert_eq!(
            DataRef::SharedString("shared").to_cell_value(),
            CellValue::Text("shared")
        );
        assert_eq!(
            Data::Error(CellErrorType::Div0).to_cell_value(),
            CellValue::Error("#DIV/0!")
        );
        assert_eq!(
            Data::Error(CellErrorType::NA).to_cell_value(),
            CellValue::Error("#N/A")
        );
    }

    #[test]
    fn iso_date_times() {
        let parse = |text: &str| parse_iso_date_time(text);
        assert_eq!(
            parse("1987-05-08T20:10:12"),
            Some((naive("1987-05-08 20:10:12"), 0))
        );
        assert_eq!(parse("2024-01-15"), Some((naive("2024-01-15 00:00:00"), 0)));
        assert_eq!(
            parse("1987-05-08T20:10:12.5"),
            Some((naive("1987-05-08 20:10:12.5"), 1))
        );
        assert_eq!(
            parse("2024-01-15T10:00:00.123456789"),
            Some((naive("2024-01-15 10:00:00.123456"), 6))
        );
        assert_eq!(
            parse("2024-01-15 10:00:00Z"),
            Some((naive("2024-01-15 10:00:00"), 0))
        );
        assert_eq!(
            parse("2024-01-15T10:00:00+02:00"),
            Some((naive("2024-01-15 10:00:00"), 0))
        );
        assert_eq!(parse("not-a-date"), None);
        assert_eq!(parse("2024-02-30"), None);
        assert_eq!(parse("2024-01-15T25:00:00"), None);
        assert_eq!(parse("12024-01-15"), None);
        assert_eq!(parse("2024-01-15T"), None);
        assert_eq!(parse("2024-01-15T10:00:0é"), None);
    }

    #[test]
    fn unparseable_iso_values_stay_strings() {
        assert_eq!(
            Data::DateTimeIso("not-a-date".into()).to_cell_value(),
            CellValue::Text("not-a-date")
        );
        assert_eq!(
            Data::DurationIso("12:00".into()).to_cell_value(),
            CellValue::Text("12:00")
        );
    }

    #[test]
    fn iso_durations() {
        assert_eq!(
            parse_iso_duration("PT36H00M00S"),
            Some(Duration {
                hour: 36,
                ..Duration::default()
            })
        );
        assert_eq!(
            parse_iso_duration("-PT01H30M00S"),
            Some(Duration {
                hour: -1,
                minute: -30,
                ..Duration::default()
            })
        );
        assert_eq!(
            parse_iso_duration("PT0H0M1.5S"),
            Some(Duration {
                second: 1,
                microsecond: 500_000,
                precision: 1,
                ..Duration::default()
            })
        );
        assert_eq!(
            parse_iso_duration("-PT1.25S"),
            Some(Duration {
                second: -1,
                microsecond: -250_000,
                precision: 2,
                ..Duration::default()
            })
        );
        assert_eq!(
            parse_iso_duration("P1Y2M3W4DT5H"),
            Some(Duration {
                year: 1,
                month: 2,
                week: 3,
                day: 4,
                hour: 5,
                ..Duration::default()
            })
        );
        assert_eq!(parse_iso_duration("P"), None);
        assert_eq!(parse_iso_duration("PT"), None);
        assert_eq!(parse_iso_duration("PTH"), None);
        assert_eq!(parse_iso_duration("PT1X"), None);
        assert_eq!(parse_iso_duration("PT1H30"), None);
        assert_eq!(parse_iso_duration("12:00"), None);
    }
}
