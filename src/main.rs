mod commission;
mod datentraeger;
mod generated;
mod sap;
mod simple;

use std::env;
use std::fmt::{Debug, Display, Formatter};
use std::fs::File;
use std::path::Path;

#[derive(Debug)]
pub enum ImportError {
    Excel(calamine::Error),
    Serialize(serde_json::error::Error),
    UnknownImport(String),
    SheetNotFound(String),
    ValueError(usize, String, String),
    UnknownHeader(String),
    MissingHeader(String),
    IoError(std::io::Error),
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
            ImportError::IoError(ref err) => std::fmt::Display::fmt(&err, f),
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
    let json_path = Path::new(path.clone().as_str()).with_extension("json");

    let json_writer = &File::create(json_path).map_err(|e| ImportError::IoError(e))?;

    match excel_type.as_str() {
        "simple" => {
            let rows = simple::run(path)?;
            let j = serde_json::to_string(&rows).map_err(|err| ImportError::Serialize(err))?;
            println!("{}", j)
        }
        "generated" => {
            let rows = generated::run(path)?;

            let j = serde_json::to_string(&rows).map_err(|err| ImportError::Serialize(err))?;
            println!("{}", j);
        }
        "mye_datentraeger" => {
            let rows = datentraeger::run(path)?;

            serde_json::to_writer(json_writer, &rows).map_err(|err| ImportError::Serialize(err))?;
        }
        "mye_commission" => {
            let rows = commission::run(path)?;

            serde_json::to_writer(json_writer, &rows).map_err(|err| ImportError::Serialize(err))?;
        }
        "mye_sap" => {
            let rows = sap::run(path)?;

            serde_json::to_writer(json_writer, &rows).map_err(|err| ImportError::Serialize(err))?;
        }
        _ => Err(ImportError::UnknownImport(excel_type))?,
    }

    Ok(())
}
