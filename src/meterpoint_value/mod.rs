use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Range, Reader};
use chrono::NaiveDateTime;
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

mod myelectric;
mod wiener_netze;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    columns: Vec<String>,
    index: Vec<NaiveDateTime>,
    data: Vec<Vec<Option<f64>>>,
}

pub enum Schema {
    MyElectric,
    WienerNetze,
    Unknown,
}

pub fn detect_schema(sheet: Range<DataType>) -> Schema {
    // wiener netze
    let ts_check = sheet
        .get_value((1, 0))
        .map(|v| v.to_string().trim().to_string());
    let mp_check = sheet
        .get_value((1, 1))
        .map(|v| v.to_string().trim().to_string());
    let mp2_check = sheet
        .get_value((6, 1))
        .map(|v| v.to_string().trim().to_string());
    let kwh_check = sheet
        .get_value((13, 1))
        .map(|v| v.to_string().trim().to_string());

    if ts_check == Some("Zeitpunkt".to_string())
        && mp_check == Some("Abnahmestelle".to_string())
        && mp2_check == Some("Zählpunkt".to_string())
        && kwh_check == Some("Wirkverbrauch_kWh".to_string())
    {
        return Schema::WienerNetze;
    }
    // myElectric
    let ts_check = sheet
        .get_value((0, 0))
        .map(|v| v.to_string().trim().to_string());

    if ts_check == Some("Timestamp".to_string()) {
        return Schema::MyElectric;
    }
    return Schema::Unknown;
}

pub fn run<P: AsRef<std::path::Path>>(path: P) -> Result<HashMap<String, Vec<Row>>, ImportError> {
    let mut excel = open_workbook_auto(path)?;

    // todo input
    let sheet_names = excel.sheet_names().to_vec();
    let sheet_name = sheet_names.first().unwrap();

    let sheet = excel
        .worksheet_range(sheet_name)
        .ok_or_else(|| ImportError::SheetNotFound(sheet_name.to_string()))??;

    return match detect_schema(sheet.clone()) {
        Schema::WienerNetze => wiener_netze::run(sheet),
        Schema::MyElectric => myelectric::run(sheet),
        Schema::Unknown => Err(ImportError::Error(
            "Could not detect schema for meterpoint_value import".to_string(),
        )),
    };
}
