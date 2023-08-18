mod contact_type;
mod rating;
mod result;
mod status;

use contact_type::ContactType;
use rating::Rating;
use result::Result_;
use status::Status;

use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Reader};
use chrono::{NaiveDate, NaiveDateTime, NaiveTime};
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    #[serde(skip_serializing)]
    retry_time: NaiveTime,
    #[serde(skip_serializing)]
    retry_date: NaiveDate,
    retry: Option<NaiveDateTime>,
    rating: Option<Rating>,
    project_contact_id: String,
    status: Status,
    result: Result_,
    contact_type: ContactType,
    created_by: String,
    feedback: String,
}

#[derive(Debug)]
enum Column {
    RetryTime,
    RetryDate,
    Rating,
    ProjectContactId,
    Status,
    Result,
    ContactType,
    CreatedBy,
    Feedback,
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
    let mut map: Vec<Option<usize>> = vec![None; 9];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "wiedervorlage zeit" => {
                map[<Column as Into<usize>>::into(Column::RetryTime)] = Some(i);
            }

            "wiedervorlage datum" => {
                map[<Column as Into<usize>>::into(Column::RetryDate)] = Some(i);
            }

            "bewertung" => {
                map[<Column as Into<usize>>::into(Column::Rating)] = Some(i);
            }

            "projectcontactid" => {
                map[<Column as Into<usize>>::into(Column::ProjectContactId)] = Some(i);
            }

            "status" => {
                map[<Column as Into<usize>>::into(Column::Status)] = Some(i);
            }

            "ergebnis" => {
                map[<Column as Into<usize>>::into(Column::Result)] = Some(i);
            }

            "contacttype" => {
                map[<Column as Into<usize>>::into(Column::ContactType)] = Some(i);
            }

            "createdby" => {
                map[<Column as Into<usize>>::into(Column::CreatedBy)] = Some(i);
            }

            "rückmeldung" => {
                map[<Column as Into<usize>>::into(Column::Feedback)] = Some(i);
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
        let mut r = transform_row(&column_map, row, i)?;

        r.retry = Some(NaiveDateTime::new(r.retry_date, r.retry_time));

        let k = r.project_contact_id.clone();
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
    row_number: usize,
) -> Result<Row, ImportError> {
    let r = Row {
        retry_time: row[column_map[Column::RetryTime.usize()]]
            .as_time()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Wiedervorlage Datum".to_string(),
                    "Could not read time".to_string(),
                )
            })?,

        retry_date: row[column_map[Column::RetryDate.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Wiedervorlage Zeit".to_string(),
                    "Could not read date".to_string(),
                )
            })?,
        retry: None,

        rating: Rating::from_excel_value(row[column_map[Column::Rating.usize()]].to_string())
            .map_err(|e| ImportError::ValueError(row_number, "Bewertung".to_string(), e))?,

        project_contact_id: row[column_map[Column::ProjectContactId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        status: Status::from_excel_value(row[column_map[Column::Status.usize()]].to_string())
            .map_err(|e| ImportError::ValueError(row_number, "status".to_string(), e))?,

        result: Result_::from_excel_value(row[column_map[Column::Result.usize()]].to_string())
            .map_err(|e| ImportError::ValueError(row_number, "Ergebnis".to_string(), e))?,

        contact_type: ContactType::from_excel_value(
            row[column_map[Column::ContactType.usize()]].to_string(),
        )
        .map_err(|e| ImportError::ValueError(row_number, "contactType".to_string(), e))?,

        created_by: row[column_map[Column::CreatedBy.usize()]]
            .to_string()
            .trim()
            .to_string(),

        feedback: row[column_map[Column::Feedback.usize()]]
            .to_string()
            .trim()
            .to_string(),
    };

    Ok(r)
}
