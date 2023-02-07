use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Reader};
use chrono::NaiveDate;
use serde::Serialize;
use std::fmt::Debug;

#[derive(Debug, Serialize)]
enum SomeEnum {
    #[serde(rename = "yes")]
    Yes,
    #[serde(rename = "no")]
    No,
}

#[derive(Debug, Serialize)]
pub struct Row {
    name: Option<String>,
    value: i64,
    some_date: NaiveDate,
    weird_date: Option<NaiveDate>,
    some_enum: SomeEnum,
}

#[derive(Debug)]
enum Column {
    Name,
    Value,
    SomeDate,
    WeirdDate,
    SomeEnum,
}

impl Into<usize> for Column {
    fn into(self) -> usize {
        self as usize
    }
}

fn get_column_map(headers: Vec<String>) -> Result<Vec<usize>, ImportError> {
    let mut map: Vec<Option<usize>> = vec![None; 5];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "name" => {
                map[<Column as Into<usize>>::into(Column::Name)] = Some(i);
            }
            "value" => {
                map[<Column as Into<usize>>::into(Column::Value)] = Some(i);
            }
            "somedate" => {
                map[<Column as Into<usize>>::into(Column::SomeDate)] = Some(i);
            }
            "weirddate" => {
                map[<Column as Into<usize>>::into(Column::WeirdDate)] = Some(i);
            }
            "someenum" => {
                map[<Column as Into<usize>>::into(Column::SomeEnum)] = Some(i);
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

pub fn run<P: AsRef<std::path::Path>>(path: P) -> Result<Vec<Row>, ImportError> {
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

    let mut rows: Vec<Row> = Vec::new();
    for (i, row) in sheet.rows().enumerate().skip(data_start_row) {
        let r = transform_row(&column_map, row, i)?;
        println!("{}: {:?}", i, r);
        rows.push(r);
    }

    Ok(rows)
}

fn transform_row(
    column_map: &[usize],
    row: &[DataType],
    row_number: usize,
) -> Result<Row, ImportError> {
    let r = Row {
        name: row[column_map[0]]
            .get_string()
            .map(|s| s.trim().to_string()),
        value: row[column_map[1]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "value".to_string(),
                    "Cell has no value".to_string(),
                )
            })?
            .round() as i64,
        some_date: row[column_map[2]].as_date().ok_or_else(|| {
            ImportError::ValueError(
                row_number,
                "some_date".to_string(),
                "Cell has no value".to_string(),
            )
        })?,
        weird_date: row[column_map[3]].as_date().and_then(|d| {
            // this is how none is represented in some excel files of mine
            if NaiveDate::from_ymd_opt(9999, 12, 31).unwrap() == d {
                None
            } else {
                Some(d)
            }
        }),
        some_enum: match row[column_map[4]].get_string().unwrap_or("").trim() {
            "y" => Ok(SomeEnum::Yes),
            "n" => Ok(SomeEnum::No),
            _ => Err(ImportError::ValueError(
                row_number,
                "some_enum".to_string(),
                "Unknown value in enum".to_string(),
            )),
        }?,
    };

    Ok(r)
}

#[cfg(test)]
mod tests {
    use crate::simple::get_column_map;

    #[test]
    fn test_get_column_map_success_with_ordered_columns() {
        let result = get_column_map(vec![
            "Name".to_string(),
            "Value".to_string(),
            "SomeDate".to_string(),
            "WeirdDate".to_string(),
            "SomeEnum".to_string(),
        ]);
        assert!(result.is_ok());
        let expected: Vec<usize> = vec![0, 1, 2, 3, 4];
        assert_eq!(expected, result.unwrap())
    }

    #[test]
    fn test_get_column_map_success_with_unordered_columns() {
        let result = get_column_map(vec![
            "SomeDate".to_string(),
            "WeirdDate".to_string(),
            "Value".to_string(),
            "Name".to_string(),
            "SomeEnum".to_string(),
        ]);
        assert!(result.is_ok());
        let expected: Vec<usize> = vec![3, 2, 0, 1, 4];
        assert_eq!(expected, result.unwrap())
    }
    #[test]
    fn test_get_column_map_fails_on_missing_column() {
        let result = get_column_map(vec![
            "Name".to_string(),
            "Value".to_string(),
            "SomeDate".to_string(),
            "WeirdDate".to_string(),
        ]);
        assert!(result.is_err());
        // todo how to test that it is a ImportError with missing column SomeEnum
        // let expected = ImportError::MissingHeader("SomeEnum".to_string());
        // I can not derive PartialEq on ImportError because calamine::Error and serde_json::error:Error do not implement it
        // assert_eq!(expected, result.unwrap_err());
    }
}
