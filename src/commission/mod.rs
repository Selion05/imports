use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Reader};
use chrono::NaiveDate;
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

#[cfg(test)]
mod tests;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    _type: String,
    billing_amount: f64,
    contract_account: String,
    currency: String,
    entry_date: NaiveDate,
    meterpoint: String,
    name: String,
    net_amount: String,
    price: String,
    print_receipt: String,
    stgrbt: String,
    supplier_customer_id: String,
    valid_from: NaiveDate,
    valid_to: NaiveDate,
}

#[derive(Debug)]
enum Column {
    Type,
    BillingAmount,
    ContractAccount,
    Currency,
    EntryDate,
    Meterpoint,
    Name,
    NetAmount,
    Price,
    PrintReceipt,
    Stgrbt,
    SupplierCustomerId,
    ValidFrom,
    ValidTo,
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
    let mut map: Vec<Option<usize>> = vec![None; 14];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "sparte" => {
                map[<Column as Into<usize>>::into(Column::Type)] = Some(i);
            }

            "abrmenge" => {
                map[<Column as Into<usize>>::into(Column::BillingAmount)] = Some(i);
            }

            "vertragskonto" => {
                map[<Column as Into<usize>>::into(Column::ContractAccount)] = Some(i);
            }

            "twährg" => {
                map[<Column as Into<usize>>::into(Column::Currency)] = Some(i);
            }

            "buch.dat." => {
                map[<Column as Into<usize>>::into(Column::EntryDate)] = Some(i);
            }

            "zählpunkt" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "zp" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "name" => {
                map[<Column as Into<usize>>::into(Column::Name)] = Some(i);
            }

            "nettobetrag" => {
                map[<Column as Into<usize>>::into(Column::NetAmount)] = Some(i);
            }

            "preisbetrag" => {
                map[<Column as Into<usize>>::into(Column::Price)] = Some(i);
            }

            "druckbeleg" => {
                map[<Column as Into<usize>>::into(Column::PrintReceipt)] = Some(i);
            }

            "stgrbt" => {
                map[<Column as Into<usize>>::into(Column::Stgrbt)] = Some(i);
            }

            "geschäftspartner" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerId)] = Some(i);
            }

            "gültig ab" => {
                map[<Column as Into<usize>>::into(Column::ValidFrom)] = Some(i);
            }

            "gültig bis" => {
                map[<Column as Into<usize>>::into(Column::ValidTo)] = Some(i);
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
    let header_row = 6;
    let data_start_row = 7;
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
        let summary_row = row[column_map[Column::Stgrbt.usize()]].to_string();
        if summary_row == "Management Fee".to_string() {
            println!("found");
            // commission files contain a summary row
            break;
        }
        let r = transform_row(&column_map, row, i)?;
        let k = r.supplier_customer_id.clone();
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
        _type: row[column_map[Column::Type.usize()]]
            .to_string()
            .trim()
            .to_string(),

        billing_amount: row[column_map[Column::BillingAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Abrmenge".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        contract_account: row[column_map[Column::ContractAccount.usize()]]
            .to_string()
            .trim()
            .to_string(),

        currency: row[column_map[Column::Currency.usize()]]
            .to_string()
            .trim()
            .to_string(),

        entry_date: row[column_map[Column::EntryDate.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Buch.dat.".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        meterpoint: row[column_map[Column::Meterpoint.usize()]]
            .to_string()
            .trim()
            .to_string(),

        name: row[column_map[Column::Name.usize()]]
            .to_string()
            .trim()
            .to_string(),

        net_amount: row[column_map[Column::NetAmount.usize()]]
            .to_string()
            .trim()
            .to_string(),

        price: row[column_map[Column::Price.usize()]]
            .to_string()
            .trim()
            .to_string(),

        print_receipt: row[column_map[Column::PrintReceipt.usize()]]
            .to_string()
            .trim()
            .to_string(),

        stgrbt: row[column_map[Column::Stgrbt.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_id: row[column_map[Column::SupplierCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        valid_from: row[column_map[Column::ValidFrom.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Gültig ab".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        valid_to: row[column_map[Column::ValidTo.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Gültig bis".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,
    };

    Ok(r)
}
