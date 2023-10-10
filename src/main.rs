mod commission;
mod contact_attempt;
mod customer_tag;
mod datentraeger;
mod kam;
mod meterpoint_value;
mod sap;

use chrono::Utc;
use regex::Regex;
use serde::Serialize;
use std::collections::HashMap;
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
    Error(String),
}

#[derive(Debug, Serialize)]
struct Schema<T: Serialize> {
    messages: T,
    meta: HashMap<String, String>,
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
                write!(f, "Value '{}' at row {} column {}", message, i, name)
            }
            ImportError::UnknownHeader(ref header) => {
                write!(f, "Unknown header name {} found", header)
            }
            ImportError::MissingHeader(ref header) => {
                write!(f, "Missing header name {}", header)
            }
            ImportError::Error(ref err) => {
                write!(f, "{}", err)
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
        "mye_datentraeger" => {
            let rows = datentraeger::run(path)?;
            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());
            let s: Schema<Vec<&Vec<datentraeger::Row>>> = Schema {
                messages: rows.values().collect(),
                meta,
            };
            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        "mye_commission" => {
            let rows = commission::run(path.clone())?;

            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());

            let re = Regex::new(r".*commissions-enelteco-(?P<timeframe>[0-9]{4}-[0-9]{2})\.xlsx?")
                .unwrap();

            let timeframe = re.captures(path.as_str()).and_then(|cap| {
                cap.name("timeframe")
                    .map(|timeframe| timeframe.as_str().to_string())
            });

            match timeframe {
                Some(timeframe) => {
                    meta.insert("timeframe".to_string(), timeframe);
                }
                None => {
                    return Err(ImportError::Error(
                        "Could not extract timeframe from path".to_string(),
                    ))
                }
            }

            let s: Schema<Vec<&Vec<commission::Row>>> = Schema {
                messages: rows.values().collect(),
                meta,
            };

            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        "mye_sap" => {
            let rows = sap::run(path)?;

            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());
            let s: Schema<Vec<&Vec<sap::Row>>> = Schema {
                messages: rows.values().collect(),
                meta,
            };

            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        "mye_kam" => {
            let rows = kam::run(path.clone())?;

            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());

            let re = Regex::new(r".*enelteco-kam-(?P<timeframe>[0-9]{4}-[0-9]{2}-[0-9]{2})\.xlsx?")
                .unwrap();

            let timeframe = re.captures(path.as_str()).and_then(|cap| {
                cap.name("timeframe")
                    .map(|timeframe| timeframe.as_str().to_string())
            });

            match timeframe {
                Some(timeframe) => {
                    meta.insert("timeframe".to_string(), timeframe);
                }
                None => {
                    return Err(ImportError::Error(
                        "Could not extract timeframe from path".to_string(),
                    ))
                }
            }

            let s: Schema<Vec<&Vec<kam::Row>>> = Schema {
                messages: rows.values().collect(),
                meta,
            };

            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        "customer_tag" => {
            let rows = customer_tag::run(path)?;

            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());
            let s: Schema<Vec<&Vec<customer_tag::Row>>> = Schema {
                messages: rows.values().collect(),
                meta,
            };
            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        "mye_meterpoint_value" => {
            let rows = meterpoint_value::run(path)?;

            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());
            let s = Schema {
                messages: vec![rows],
                meta,
            };
            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        "contact_attempt" => {
            let rows = contact_attempt::run(path)?;

            let mut meta: HashMap<String, String> = HashMap::new();
            meta.insert("created_at".to_string(), Utc::now().to_string());
            let s: Schema<Vec<&Vec<contact_attempt::Row>>> = Schema {
                messages: rows.values().collect(),
                meta,
            };
            serde_json::to_writer(json_writer, &s).map_err(|err| ImportError::Serialize(err))?;
        }
        _ => Err(ImportError::UnknownImport(excel_type))?,
    }

    Ok(())
}
