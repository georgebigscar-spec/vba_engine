//! `VBA.Math` and `VBA.Financial`.

use rust_decimal::{Decimal, RoundingStrategy};

use super::*;

/// Constants of this library (`name` is lowercase).
pub(super) fn constant(name: &str) -> Option<Value> {
    let _ = name;
    None
}

pub(super) fn call(
    lname: &str,
    name: &str,
    args: &Args,
    host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    Ok(match lname {
        // ---- math ----
        "abs" => abs(arg(args, 0)?)?,
        "sgn" => Value::Integer(n(args, 0)?.signum() as i16 * i16::from(n(args, 0)? != 0.0)),
        "int" => int_fix(arg(args, 0)?, f64::floor)?,
        "fix" => int_fix(arg(args, 0)?, f64::trunc)?,
        "sqr" => {
            let x = n(args, 0)?;
            if x < 0.0 {
                return Err(VbaError::invalid_call());
            }
            Value::Double(x.sqrt())
        }
        "exp" => Value::Double(finite(n(args, 0)?.exp())?),
        "log" => {
            let x = n(args, 0)?;
            if x <= 0.0 {
                return Err(VbaError::invalid_call());
            }
            Value::Double(x.ln())
        }
        "sin" => Value::Double(n(args, 0)?.sin()),
        "cos" => Value::Double(n(args, 0)?.cos()),
        "tan" => Value::Double(n(args, 0)?.tan()),
        "atn" => Value::Double(n(args, 0)?.atan()),
        "round" => {
            let digits = opt(args, 1).map(|v| v.to_i64()).transpose()?.unwrap_or(0);
            if digits < 0 {
                return Err(VbaError::invalid_call());
            }
            round(arg(args, 0)?, digits)?
        }
        "rnd" => rnd(opt(args, 0).map(|v| v.to_f64()).transpose()?.unwrap_or(1.0) as f32),
        "randomize" => {
            let seed = match opt(args, 0) {
                Some(v) => v.to_f64()?,
                None => (host.now().fract() * 86_400.0) as f32 as f64,
            };
            randomize(seed);
            Value::Empty
        }
        "rgb" => Value::Long(
            (int(args, 0)? & 0xFF | (int(args, 1)? & 0xFF) << 8 | (int(args, 2)? & 0xFF) << 16)
                as i32,
        ),
        "partition" => match arg(args, 0)? {
            Value::Null => Value::Null,
            v => Value::String(partition(
                v.to_i64()?,
                int(args, 1)?,
                int(args, 2)?,
                int(args, 3)?,
            )?),
        },

        // ---- financial ----
        "sln" => {
            let life = n(args, 2)?;
            if life == 0.0 {
                return Err(VbaError::invalid_call());
            }
            Value::Double((n(args, 0)? - n(args, 1)?) / life)
        }
        "syd" => {
            let (cost, salvage, life, per) = (n(args, 0)?, n(args, 1)?, n(args, 2)?, n(args, 3)?);
            if life <= 0.0 || per <= 0.0 || per > life || salvage < 0.0 {
                return Err(VbaError::invalid_call());
            }
            Value::Double((cost - salvage) * (life - per + 1.0) * 2.0 / (life * (life + 1.0)))
        }
        "ddb" => {
            let (cost, salvage, life, per) = (n(args, 0)?, n(args, 1)?, n(args, 2)?, n(args, 3)?);
            let factor = opt_n(args, 4, 2.0)?;
            if cost < 0.0
                || salvage < 0.0
                || life <= 0.0
                || per <= 0.0
                || per > life
                || factor <= 0.0
            {
                return Err(VbaError::invalid_call());
            }
            let rate = (factor / life).min(1.0);
            let book = cost * (1.0 - rate).powf(per - 1.0);
            Value::Double((book * rate).min(book - salvage).max(0.0))
        }
        "pmt" => {
            let (fv, due) = (opt_n(args, 3, 0.0)?, opt_n(args, 4, 0.0)?);
            Value::Double(pmt(n(args, 0)?, n(args, 1)?, n(args, 2)?, fv, due)?)
        }
        "fv" => {
            let (pv, due) = (opt_n(args, 3, 0.0)?, opt_n(args, 4, 0.0)?);
            Value::Double(future_value(n(args, 0)?, n(args, 1)?, n(args, 2)?, pv, due))
        }
        "pv" => {
            let (rate, nper, pmt) = (n(args, 0)?, n(args, 1)?, n(args, 2)?);
            let (fv, due) = (opt_n(args, 3, 0.0)?, opt_n(args, 4, 0.0)?);
            Value::Double(if rate == 0.0 {
                -(fv + pmt * nper)
            } else {
                let g = (1.0 + rate).powf(nper);
                -(fv + pmt * (1.0 + rate * due) * (g - 1.0) / rate) / g
            })
        }
        "nper" => {
            let (rate, pmt, pv) = (n(args, 0)?, n(args, 1)?, n(args, 2)?);
            let (fv, due) = (opt_n(args, 3, 0.0)?, opt_n(args, 4, 0.0)?);
            let r = if rate == 0.0 {
                -(pv + fv) / pmt
            } else {
                let z = pmt * (1.0 + rate * due) / rate;
                ((z - fv) / (pv + z)).ln() / (1.0 + rate).ln()
            };
            Value::Double(finite(r).map_err(|_| VbaError::invalid_call())?)
        }
        "rate" => {
            let (nper, pmt, pv) = (n(args, 0)?, n(args, 1)?, n(args, 2)?);
            let (fv, due) = (opt_n(args, 3, 0.0)?, opt_n(args, 4, 0.0)?);
            let guess = opt_n(args, 5, 0.1)?;
            if nper <= 0.0 {
                return Err(VbaError::invalid_call());
            }
            Value::Double(solve(guess, |r| {
                fv + pv * (1.0 + r).powf(nper) + pmt * (1.0 + r * due) * annuity(r, nper)
            })?)
        }
        "ipmt" | "ppmt" => {
            let (rate, per, nper, pv) = (n(args, 0)?, n(args, 1)?, n(args, 2)?, n(args, 3)?);
            let (fv, due) = (opt_n(args, 4, 0.0)?, opt_n(args, 5, 0.0)?);
            if per < 1.0 || per >= nper + 1.0 {
                return Err(VbaError::invalid_call());
            }
            let total = pmt(rate, nper, pv, fv, due)?;
            let interest = if due != 0.0 && per == 1.0 {
                0.0
            } else {
                future_value(rate, per - 1.0, total, pv, due) * rate / (1.0 + rate * due)
            };
            Value::Double(if lname == "ipmt" {
                interest
            } else {
                total - interest
            })
        }
        "npv" => {
            let rate = n(args, 0)?;
            let flows = cash_flows(arg(args, 1)?)?;
            if rate == -1.0 || flows.is_empty() {
                return Err(VbaError::invalid_call());
            }
            Value::Double(npv(rate, &flows) / (1.0 + rate))
        }
        "irr" => {
            let flows = cash_flows(arg(args, 0)?)?;
            let guess = opt_n(args, 1, 0.1)?;
            if !flows.iter().any(|&v| v < 0.0) || !flows.iter().any(|&v| v > 0.0) {
                return Err(VbaError::invalid_call());
            }
            Value::Double(solve(guess, |r| npv(r, &flows))?)
        }
        "mirr" => {
            let flows = cash_flows(arg(args, 0)?)?;
            let (finance, reinvest) = (n(args, 1)?, n(args, 2)?);
            let len = flows.len() as f64;
            let pos: Vec<f64> = flows.iter().map(|&v| v.max(0.0)).collect();
            let neg: Vec<f64> = flows.iter().map(|&v| v.min(0.0)).collect();
            let (pos, neg) = (npv(reinvest, &pos), npv(finance, &neg));
            if finance == -1.0 || reinvest == -1.0 || len < 2.0 || pos == 0.0 || neg == 0.0 {
                return Err(VbaError::invalid_call());
            }
            let r = (-pos * (1.0 + reinvest).powf(len - 1.0) / neg).powf(1.0 / (len - 1.0)) - 1.0;
            Value::Double(finite(r).map_err(|_| VbaError::invalid_call())?)
        }
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}

/// Optional numeric argument `i`, `default` when omitted.
fn opt_n(args: &Args, i: usize, default: f64) -> VbaResult<f64> {
    opt(args, i).map_or(Ok(default), Value::to_f64)
}

/// Overflow (error 6) unless `x` is finite.
fn finite(x: f64) -> VbaResult<f64> {
    if x.is_finite() {
        Ok(x)
    } else {
        Err(VbaError::overflow())
    }
}

/// `Abs`: keeps the numeric type; overflows on the most negative integer.
fn abs(v: &Value) -> VbaResult<Value> {
    Ok(match v {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Boolean(b) => Value::Integer(i16::from(*b)),
        Value::Byte(_) => v.clone(),
        Value::Integer(i) => Value::Integer(i.checked_abs().ok_or_else(VbaError::overflow)?),
        Value::Long(i) => Value::Long(i.checked_abs().ok_or_else(VbaError::overflow)?),
        Value::LongLong(i) => Value::LongLong(i.checked_abs().ok_or_else(VbaError::overflow)?),
        Value::Single(x) => Value::Single(x.abs()),
        Value::Currency(c) => Value::Currency(c.checked_abs().ok_or_else(VbaError::overflow)?),
        Value::Decimal(d) => Value::Decimal(d.abs()),
        Value::Date(x) => Value::Date(x.abs()),
        other => Value::Double(other.to_f64()?.abs()),
    })
}

/// `Int` / `Fix`: keep the numeric type (strings become Double).
fn int_fix(v: &Value, f: fn(f64) -> f64) -> VbaResult<Value> {
    Ok(match v {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Boolean(b) => Value::Integer(-i16::from(*b)),
        Value::Byte(_) | Value::Integer(_) | Value::Long(_) | Value::LongLong(_) => v.clone(),
        Value::Single(x) => Value::Single(f(*x as f64) as f32),
        Value::Currency(c) => Value::Currency(f(*c as f64 / 10_000.0) as i64 * 10_000),
        Value::Decimal(d) => {
            // `f` of the fraction alone is 0 or -1 (Int of a negative fraction).
            let t = d.trunc();
            Value::Decimal(t + Decimal::from(f(Value::Decimal(d - t).to_f64()?) as i64))
        }
        Value::Date(x) => Value::Date(f(*x)),
        other => Value::Double(f(other.to_f64()?)),
    })
}

/// `Round` with banker's rounding to `digits` decimal places, keeping the type.
fn round(v: &Value, digits: i64) -> VbaResult<Value> {
    let round_f = |x: f64| {
        let f = 10f64.powi(digits.min(308) as i32);
        let r = (x * f).round_ties_even() / f;
        if r.is_finite() {
            r
        } else {
            x
        }
    };
    Ok(match v {
        Value::Null => Value::Null,
        Value::Empty => Value::Integer(0),
        Value::Boolean(b) => Value::Integer(-i16::from(*b)),
        Value::Byte(_) | Value::Integer(_) | Value::Long(_) | Value::LongLong(_) => v.clone(),
        Value::Single(x) => Value::Single(round_f(*x as f64) as f32),
        Value::Currency(c) if digits < 4 => {
            let unit = 10i64.pow(4 - digits as u32);
            let q = (*c as f64 / unit as f64).round_ties_even() as i64;
            Value::Currency(q * unit)
        }
        Value::Currency(_) => v.clone(),
        Value::Decimal(d) => Value::Decimal(
            d.round_dp_with_strategy(digits.min(28) as u32, RoundingStrategy::MidpointNearestEven),
        ),
        other => Value::Double(round_f(other.to_f64()?)),
    })
}

/// `Rnd`: VBA's 24-bit linear congruential generator. A negative argument
/// reseeds from its Single bit pattern, zero repeats the last number.
fn rnd(x: f32) -> Value {
    RND_STATE.with(|st| {
        let mut seed = st.get();
        if x < 0.0 {
            let bits = x.to_bits();
            seed = bits.wrapping_add(bits >> 24) & 0xFF_FFFF;
        }
        if x != 0.0 {
            seed = seed.wrapping_mul(0x43FD_43FD).wrapping_add(0xC3_9EC3) & 0xFF_FFFF;
        }
        st.set(seed);
        Value::Single(seed as f32 / 16_777_216.0)
    })
}

/// `Randomize seed`: mixes the high word of the Double into the middle bytes
/// of the generator state, as VBA does.
fn randomize(seed: f64) {
    let hi = (seed.to_bits() >> 32) as u32;
    let mixed = ((hi & 0xFFFF) ^ (hi >> 16)) << 8;
    RND_STATE.with(|st| st.set((st.get() & 0xFF) | (mixed & 0xFF_FF00)));
}

/// `Partition`: the `"lower:upper"` range `number` falls in.
fn partition(number: i64, start: i64, stop: i64, interval: i64) -> VbaResult<String> {
    if start < 0 || stop <= start || interval < 1 {
        return Err(VbaError::invalid_call());
    }
    let width = (stop + 1)
        .to_string()
        .len()
        .max((start - 1).to_string().len())
        + 1;
    let pad = |v: Option<i64>| match v {
        Some(v) => format!("{v:>width$}"),
        None => " ".repeat(width),
    };
    let (lo, hi) = if number < start {
        (None, Some(start - 1))
    } else if number > stop {
        (Some(stop + 1), None)
    } else {
        let lo = start + (number - start) / interval * interval;
        (Some(lo), Some((lo + interval - 1).min(stop)))
    };
    Ok(format!("{}:{}", pad(lo), pad(hi)))
}

/// `((1 + rate)^nper - 1) / rate`, the future value of a unit annuity.
fn annuity(rate: f64, nper: f64) -> f64 {
    if rate == 0.0 {
        nper
    } else {
        ((1.0 + rate).powf(nper) - 1.0) / rate
    }
}

fn pmt(rate: f64, nper: f64, pv: f64, fv: f64, due: f64) -> VbaResult<f64> {
    if nper == 0.0 {
        return Err(VbaError::invalid_call());
    }
    let due = if rate == 0.0 { 0.0 } else { due };
    let r = -(fv + pv * (1.0 + rate).powf(nper)) / ((1.0 + rate * due) * annuity(rate, nper));
    finite(r).map_err(|_| VbaError::invalid_call())
}

fn future_value(rate: f64, nper: f64, pmt: f64, pv: f64, due: f64) -> f64 {
    let due = if rate == 0.0 { 0.0 } else { due };
    -(pv * (1.0 + rate).powf(nper) + pmt * (1.0 + rate * due) * annuity(rate, nper))
}

/// Net present value with the first flow at time 0.
fn npv(rate: f64, flows: &[f64]) -> f64 {
    flows
        .iter()
        .rev()
        .fold(0.0, |acc, &v| acc / (1.0 + rate) + v)
}

/// The numeric elements of a cash-flow array argument.
fn cash_flows(v: &Value) -> VbaResult<Vec<f64>> {
    match v {
        Value::Array(a) => a.data.iter().map(Value::to_f64).collect(),
        _ => Err(VbaError::type_mismatch()),
    }
}

/// Finds a root of `f` near `guess` by the secant method (error 5 if none).
fn solve(guess: f64, f: impl Fn(f64) -> f64) -> VbaResult<f64> {
    let (mut x0, mut x1) = (guess, guess + 1e-4);
    let (mut f0, mut f1) = (f(x0), f(x1));
    for _ in 0..100 {
        if f1 == f0 {
            break;
        }
        let x2 = x1 - f1 * (x1 - x0) / (f1 - f0);
        if !x2.is_finite() || x2 <= -1.0 {
            break;
        }
        if (x2 - x1).abs() < 1e-12 * x2.abs().max(1e-3) {
            return Ok(x2);
        }
        (x0, f0, x1) = (x1, f1, x2);
        f1 = f(x1);
    }
    Err(VbaError::invalid_call())
}
