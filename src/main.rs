mod simple;

use std::env;
use std::fmt::{Debug, Display, Formatter};

#[derive(Debug)]
pub enum ImportError {
    Excel(calamine::Error),
    Serialize(serde_json::error::Error),
    UnknownImport(String),
    SheetNotFound(String),
    ValueError(usize, String, String),
    UnknownHeader(String),
    MissingHeader(String),
}

impl Display for ImportError {
    fn fmt(&self, f: &mut Formatter<'_>) -> std::fmt::Result {
        match *self {
            ImportError::Excel(ref err) => std::fmt::Display::fmt(&err, f),
            ImportError::Serialize(ref err) => std::fmt::Display::fmt(&err, f),
            ImportError::UnknownImport(ref import) => write!(f, "Unknown import type {}", import),
            ImportError::SheetNotFound(ref name) => {
                write!(f, "Could not find sheet with name {}", name)
            }
            ImportError::ValueError(ref i, ref name, ref message) => {
                write!(f, "{} at row {} column {}", message, i, name)
            }
            ImportError::UnknownHeader(ref header) => {
                write!(f, "Unknown header name {} found", header)
            }
            ImportError::MissingHeader(ref header) => {
                write!(f, "Missing header name {}", header)
            }
        }
    }
}

impl From<calamine::Error> for ImportError {
    fn from(err: calamine::Error) -> Self {
        ImportError::Excel(err)
    }
}

fn main() {
    let excel_type = env::args().nth(1).unwrap();
    let path = env::args().nth(2).unwrap();

    match run(excel_type, path) {
        Ok(_) => {}
        Err(err) => {
            println!("{}", err);
            std::process::exit(1);
        }
    }
}

fn run(excel_type: String, path: String) -> Result<(), ImportError> {
    match excel_type.as_str() {
        "simple" => {
            let rows = simple::run(path)?;
            let j = serde_json::to_string(&rows).map_err(|err| ImportError::Serialize(err))?;
            println!("{}", j)
        }
        _ => Err(ImportError::UnknownImport(excel_type))?,
    }

    Ok(())
}
