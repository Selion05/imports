use crate::ImportError;
use calamine::{open_workbook_auto, Reader};
use chrono::{NaiveDateTime, SubsecRound};
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

#[cfg(test)]
mod tests;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    columns: Vec<String>,
    index: Vec<NaiveDateTime>,
    data: Vec<Vec<Option<f64>>>,
}

pub fn run<P: AsRef<std::path::Path>>(path: P) -> Result<HashMap<String, Vec<Row>>, ImportError> {
    let mut excel = open_workbook_auto(path)?;

    // todo input
    let header_row = 0;
    let data_start_row = 1;
    let sheet_names = excel.sheet_names().to_vec();
    let sheet_name = sheet_names.first().unwrap();

    let sheet = excel
        .worksheet_range(sheet_name)
        .ok_or_else(|| ImportError::SheetNotFound(sheet_name.to_string()))??;

    let ts = sheet
        .rows()
        .nth(header_row)
        .unwrap()
        .iter()
        .nth(0)
        .unwrap()
        .to_string();

    if ts.trim() != "Timestamp" {
        return Err(ImportError::UnknownHeader(String::from(
            "Expected Timestamp in A1",
        )));
    }

    let headers: Vec<String> = sheet
        .rows()
        .nth(header_row)
        .unwrap()
        .iter()
        .skip(1)
        .map(|header_cell| {
            return match header_cell.get_string() {
                None => String::from(""),
                Some(v) => v[..33].to_string(),
            };
        })
        .collect();

    let mut groups: HashMap<String, Vec<Row>> = HashMap::new();

    let header_cnt = headers.len();
    let mut r = Row {
        columns: headers,
        index: vec![],
        data: vec![],
    };

    for (i, row) in sheet.rows().enumerate().skip(data_start_row) {
        let summary_row = row[0].to_string();
        if summary_row == "Summe" || summary_row == "Sum" {
            println!("found");
            // meterpoint_value files contain a summary row
            break;
        }
        let date = row[0]
            .as_datetime()
            .ok_or_else(|| {
                ImportError::ValueError(
                    i,
                    "Timestamp".to_string(),
                    "could not parse datetime".to_string(),
                )
            })?
            .round_subsecs(0);

        let values = row
            .iter()
            .skip(1)
            .take(header_cnt)
            .map(|v| {
                return v.get_float();
            })
            .collect();

        r.index.push(date);
        r.data.push(values);
    }

    let k = "all".to_string();
    groups.insert(k.clone(), vec![r]);

    Ok(groups)
}
