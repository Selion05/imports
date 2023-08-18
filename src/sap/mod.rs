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
    read_unit: String,
    price_amount: f64,
    tariff: String,
    net_amount: f64,
    currency: String,
    billing_amount: f64,
    line_id: String,
    energy_type: String,
    ba: String,
    valid_from: NaiveDate,
    contract: String,
    entry_date: NaiveDate,
    invoice_id: String,
    supplier_customer_id: String,
    contract_account: String,
    valid_to: NaiveDate,
    meterpoint: String,
}

#[derive(Debug)]
enum Column {
    ReadUnit,
    PriceAmount,
    Tariff,
    NetAmount,
    Currency,
    BillingAmount,
    LineId,
    EnergyType,
    Ba,
    ValidFrom,
    Contract,
    EntryDate,
    InvoiceId,
    SupplierCustomerId,
    ContractAccount,
    ValidTo,
    Meterpoint,
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
    let mut map: Vec<Option<usize>> = vec![None; 17];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "ableinh." => {
                map[<Column as Into<usize>>::into(Column::ReadUnit)] = Some(i);
            }
            "preisbetrag" => {
                map[<Column as Into<usize>>::into(Column::PriceAmount)] = Some(i);
            }
            "tariftyp" => {
                map[<Column as Into<usize>>::into(Column::Tariff)] = Some(i);
            }
            "nettobetrag" => {
                map[<Column as Into<usize>>::into(Column::NetAmount)] = Some(i);
            }
            "twährg" => {
                map[<Column as Into<usize>>::into(Column::Currency)] = Some(i);
            }
            "abrmenge" => {
                map[<Column as Into<usize>>::into(Column::BillingAmount)] = Some(i);
            }
            "bart" => {
                map[<Column as Into<usize>>::into(Column::LineId)] = Some(i);
            }
            "sp" => {
                map[<Column as Into<usize>>::into(Column::EnergyType)] = Some(i);
            }
            "ba" => {
                map[<Column as Into<usize>>::into(Column::Ba)] = Some(i);
            }
            "gültig ab" => {
                map[<Column as Into<usize>>::into(Column::ValidFrom)] = Some(i);
            }
            "vertrag" => {
                map[<Column as Into<usize>>::into(Column::Contract)] = Some(i);
            }
            "buch.dat." => {
                map[<Column as Into<usize>>::into(Column::EntryDate)] = Some(i);
            }
            "druckbeleg" => {
                map[<Column as Into<usize>>::into(Column::InvoiceId)] = Some(i);
            }
            "geschäftspartner" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerId)] = Some(i);
            }
            "vertragskont" => {
                map[<Column as Into<usize>>::into(Column::ContractAccount)] = Some(i);
            }
            "gültig bis" => {
                map[<Column as Into<usize>>::into(Column::ValidTo)] = Some(i);
            }
            "zp" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
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
        let k = r.invoice_id.clone();
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
        read_unit: row[column_map[Column::ReadUnit.usize()]]
            .to_string()
            .trim()
            .to_string(),

        price_amount: row[column_map[Column::PriceAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Preisbetrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        tariff: row[column_map[Column::Tariff.usize()]]
            .to_string()
            .trim()
            .to_string(),

        net_amount: row[column_map[Column::NetAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Nettobetrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        currency: row[column_map[Column::Currency.usize()]]
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

        line_id: row[column_map[Column::LineId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        energy_type: row[column_map[Column::EnergyType.usize()]]
            .to_string()
            .trim()
            .to_string(),

        ba: row[column_map[Column::Ba.usize()]]
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

        contract: row[column_map[Column::Contract.usize()]]
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

        invoice_id: row[column_map[Column::InvoiceId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_id: row[column_map[Column::SupplierCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        contract_account: row[column_map[Column::ContractAccount.usize()]]
            .to_string()
            .trim()
            .to_string(),

        valid_to: row[column_map[Column::ValidTo.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Gültig bis".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        meterpoint: row[column_map[Column::Meterpoint.usize()]]
            .to_string()
            .trim()
            .to_string(),
    };

    Ok(r)
}
