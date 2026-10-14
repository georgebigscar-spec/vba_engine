//! `VBA.Math` and `VBA.Financial`.

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
    _host: &mut dyn Host,
    _text_compare: bool,
) -> VbaResult<Value> {
    Ok(match lname {
        // ---- math ----
        "abs" => match arg(args, 0)? {
            Value::Null => Value::Null,
            v if value::compare(v, &Value::Integer(0), false)?
                == Some(std::cmp::Ordering::Less) =>
            {
                value::neg(v)?
            }
            v => v.clone(),
        },
        "sgn" => Value::Integer(n(args, 0)?.signum() as i16 * i16::from(n(args, 0)? != 0.0)),
        "int" => Value::Double(n(args, 0)?.floor()),
        "fix" => Value::Double(n(args, 0)?.trunc()),
        "sqr" => {
            let x = n(args, 0)?;
            if x < 0.0 {
                return Err(VbaError::invalid_call());
            }
            Value::Double(x.sqrt())
        }
        "exp" => Value::Double(n(args, 0)?.exp()),
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
            let f = 10f64.powi(digits as i32);
            Value::Double((n(args, 0)? * f).round_ties_even() / f)
        }
        "rnd" => RND_STATE.with(|st| {
            // VBA's 24-bit LCG.
            let next = (st.get().wrapping_mul(0x43FD43FD).wrapping_add(0xC39EC3)) & 0xFF_FFFF;
            st.set(next);
            Value::Single(next as f32 / 16_777_216.0)
        }),
        "randomize" => Value::Empty,
        "rgb" => Value::Long(
            (int(args, 0)? & 0xFF | (int(args, 1)? & 0xFF) << 8 | (int(args, 2)? & 0xFF) << 16)
                as i32,
        ),
        _ => return Err(VbaError::sub_not_defined(name)),
    })
}
