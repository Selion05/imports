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

// fn transform_row(
//     column_map: &[usize],
//     row: &[DataType],
//     row_number: usize,
// ) -> Result<Row, ImportError> {
//     let r = Row {
//         _type: row[column_map[Column::Type.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         billing_amount: row[column_map[Column::BillingAmount.usize()]]
//             .get_float()
//             .ok_or_else(|| {
//                 ImportError::ValueError(
//                     row_number,
//                     "Abrmenge".to_string(),
//                     "Cell has no value".to_string(),
//                 )
//             })?,
//
//         contract_account: row[column_map[Column::ContractAccount.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         currency: row[column_map[Column::Currency.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         entry_date: row[column_map[Column::EntryDate.usize()]]
//             .as_date()
//             .ok_or_else(|| {
//                 ImportError::ValueError(
//                     row_number,
//                     "Buch.dat.".to_string(),
//                     "Cell has no value".to_string(),
//                 )
//             })?,
//
//         meterpoint: row[column_map[Column::Meterpoint.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         name: row[column_map[Column::Name.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         net_amount: row[column_map[Column::NetAmount.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         price: row[column_map[Column::Price.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         print_receipt: row[column_map[Column::PrintReceipt.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         stgrbt: row[column_map[Column::Stgrbt.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         supplier_customer_id: row[column_map[Column::SupplierCustomerId.usize()]]
//             .to_string()
//             .trim()
//             .to_string(),
//
//         valid_from: row[column_map[Column::ValidFrom.usize()]]
//             .as_date()
//             .ok_or_else(|| {
//                 ImportError::ValueError(
//                     row_number,
//                     "Gültig ab".to_string(),
//                     "Cell has no value".to_string(),
//                 )
//             })?,
//
//         valid_to: row[column_map[Column::ValidTo.usize()]]
//             .as_date()
//             .ok_or_else(|| {
//                 ImportError::ValueError(
//                     row_number,
//                     "Gültig bis".to_string(),
//                     "Cell has no value".to_string(),
//                 )
//             })?,
//     };
//
//     Ok(r)
// }
