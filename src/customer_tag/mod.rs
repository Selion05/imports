#[cfg(test)]
mod tests;

use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Reader};
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    customer_id: String,
    tag_id: String,
    tag_value: String,
}

#[derive(Debug)]
enum Column {
    CustomerId,
    TagId,
    TagValue,
}

impl Into<usize> for Column {
    fn into(self) -> usize {
        self as usize
    }
}

impl Column {
    fn usize(self) -> usize {
        self as usize
    }
}

fn get_column_map(headers: Vec<String>) -> Result<Vec<usize>, ImportError> {
    let mut map: Vec<Option<usize>> = vec![None; 3];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "kunden id" => {
                map[<Column as Into<usize>>::into(Column::CustomerId)] = Some(i);
            }

            "tag id" => {
                map[<Column as Into<usize>>::into(Column::TagId)] = Some(i);
            }

            "tag wert" => {
                map[<Column as Into<usize>>::into(Column::TagValue)] = Some(i);
            }

            _ => return Err(ImportError::UnknownHeader(header.clone())),
        }
    }

    for (_i, h) in map.clone().into_iter().enumerate() {
        // todo how to get enum from index _i?
        match h {
            Some(_) => {}
            None => return Err(ImportError::MissingHeader("todo".to_string())),
        }
    }

    Ok(map.into_iter().flatten().collect())
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

    let headers: Vec<String> = sheet
        .rows()
        .nth(header_row)
        .unwrap()
        .iter()
        .map(|header_cell| {
            return match header_cell.get_string() {
                None => String::from(""),
                Some(v) => v.to_string(),
            };
        })
        .collect();
    let column_map: Vec<usize> = get_column_map(headers)?;

    let mut groups: HashMap<String, Vec<Row>> = HashMap::new();
    for (i, row) in sheet.rows().enumerate().skip(data_start_row) {
        let r = transform_row(&column_map, row, i)?;
        let k = r.customer_id.clone();
        if !groups.contains_key(&*k) {
            groups.insert(k.clone(), Vec::new());
        }
        groups.get_mut(&*k).unwrap().push(r);
        // rows.push(r);
    }

    Ok(groups)
}

fn transform_row(
    column_map: &[usize],
    row: &[DataType],
    _row_number: usize,
) -> Result<Row, ImportError> {
    let r = Row {
        customer_id: row[column_map[Column::CustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        tag_id: row[column_map[Column::TagId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        tag_value: row[column_map[Column::TagValue.usize()]]
            .to_string()
            .trim()
            .to_string(),
    };

    Ok(r)
}
