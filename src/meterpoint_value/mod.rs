use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Range, Reader};
use chrono::NaiveDateTime;
use serde::Serialize;
use std::fmt::Debug;

mod myelectric;
mod netze_noe;
mod netze_ooe;
mod wiener_netze;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Data {
    columns: Vec<String>,
    index: Vec<NaiveDateTime>,
    data: Vec<Vec<Option<f64>>>,
}

pub enum Schema {
    MyElectric,
    WienerNetze,
    NetzeOoe,
    NetzeNoe,
    Unknown,
}

pub fn detect_schema(sheet: &Range<DataType>) -> Schema {
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

    let head_check = sheet
        .get_value((0, 0))
        .map(|v| v.to_string().trim().to_string());
    let unit_head_check = sheet
        .get_value((6, 0))
        .map(|v| v.to_string().trim().to_string());
    let unit_check = sheet
        .get_value((6, 1))
        .map(|v| v.to_string().trim().to_string());
    let date_check = sheet
        .get_value((1, 3))
        .map(|v| v.to_string().trim().to_string());
    let time_check = sheet
        .get_value((1, 4))
        .map(|v| v.to_string().trim().to_string());
    let value_check = sheet
        .get_value((1, 5))
        .map(|v| v.to_string().trim().to_string());
    if head_check == Some("Kopfdaten des Profils".to_string())
        && unit_head_check == Some("Maßeinheit".to_string())
        && unit_check == Some("kW".to_string())
        && date_check == Some("Ab-Datum".to_string())
        && time_check == Some("Ab-Zeit".to_string())
        && value_check == Some("Profilwert".to_string())
    {
        return Schema::NetzeOoe;
    }

    // myElectric
    let ts_check = sheet
        .get_value((0, 0))
        .map(|v| v.to_string().trim().to_string());

    if ts_check == Some("Timestamp".to_string()) {
        return Schema::MyElectric;
    }
    let head_check = sheet
        .get_value((0, 0))
        .map(|v| v.to_string().trim().to_string());

    if head_check == Some("Werte in kW".to_string()) {
        return Schema::NetzeNoe;
    }
    return Schema::Unknown;
}

pub fn run(path: String) -> Result<Data, ImportError> {
    let mut excel = open_workbook_auto(path.clone())?;

    if let Some(Ok(sheet)) = excel.worksheet_range_at(0) {
        return match detect_schema(&sheet) {
            Schema::WienerNetze => wiener_netze::run(sheet),
            Schema::MyElectric => myelectric::run(sheet),
            Schema::NetzeOoe => netze_ooe::run(sheet, path),
            Schema::NetzeNoe => netze_noe::run(sheet),
            Schema::Unknown => Err(ImportError::Error(
                "Could not detect schema for meterpoint_value import".to_string(),
            )),
        };
    }

    return Err(ImportError::Error(
        "Could not find any sheet in excel".to_string(),
    ));
}
