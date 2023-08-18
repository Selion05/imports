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
    affiliate: String,
    bill_addition: String,
    bill_city: String,
    bill_number: String,
    bill_street: String,
    bill_zip: String,
    billing_type: String,
    consumption_at_change: f64,
    consumption_forecast: f64,
    contract: String,
    contract_account: String,
    corresponding_bill_receiver: String,
    deviant_bill_receiver: String,
    e_invoice: String,
    energy_type: String,
    grid_billing_integrated: String,
    in_date: NaiveDate,
    meterpoint: String,
    mp_addition: String,
    mp_city: String,
    mp_number: String,
    mp_street: String,
    mp_zip: String,
    name: String,
    name_add: String,
    out_date: Option<NaiveDate>,
    pool_customer_id: String,
    profile: String,
    read_unit: String,
    sepa: String,
    sepa_blocked: String,
    supplier_customer_group_id: String,
    supplier_customer_group_name: String,
    supplier_customer_id: String,
    supplier_meterpoint_id: String,
    tariff_typ: String,
}

#[derive(Debug)]
enum Column {
    Affiliate,
    BillAddition,
    BillCity,
    BillNumber,
    BillStreet,
    BillZip,
    BillingType,
    ConsumptionAtChange,
    ConsumptionForecast,
    Contract,
    ContractAccount,
    CorrespondingBillReceiver,
    DeviantBillReceiver,
    EInvoice,
    EnergyType,
    GridBillingIntegrated,
    InDate,
    Meterpoint,
    MpAddition,
    MpCity,
    MpNumber,
    MpStreet,
    MpZip,
    Name,
    NameAdd,
    OutDate,
    PoolCustomerId,
    Profile,
    ReadUnit,
    Sepa,
    SepaBlocked,
    SupplierCustomerGroupId,
    SupplierCustomerGroupName,
    SupplierCustomerId,
    SupplierMeterpointId,
    TariffTyp,
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
    let mut map: Vec<Option<usize>> = vec![None; 36];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "zugehörigkeit" => {
                map[<Column as Into<usize>>::into(Column::Affiliate)] = Some(i);
            }

            "r-zusatz" => {
                map[<Column as Into<usize>>::into(Column::BillAddition)] = Some(i);
            }

            "r-ort" => {
                map[<Column as Into<usize>>::into(Column::BillCity)] = Some(i);
            }

            "r-hausnummer" => {
                map[<Column as Into<usize>>::into(Column::BillNumber)] = Some(i);
            }

            "r-straße" => {
                map[<Column as Into<usize>>::into(Column::BillStreet)] = Some(i);
            }

            "r-postleitzahl" => {
                map[<Column as Into<usize>>::into(Column::BillZip)] = Some(i);
            }

            "r-plz" => {
                map[<Column as Into<usize>>::into(Column::BillZip)] = Some(i);
            }

            "abrechnungsklasse" => {
                map[<Column as Into<usize>>::into(Column::BillingType)] = Some(i);
            }

            "wert bei wechsel kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionAtChange)] = Some(i);
            }

            "jahresverbrauch bei wechsel kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionAtChange)] = Some(i);
            }

            "aktueller prognosewert kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionForecast)] = Some(i);
            }

            "prognosewert kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionForecast)] = Some(i);
            }

            "vertrag" => {
                map[<Column as Into<usize>>::into(Column::Contract)] = Some(i);
            }

            "vertragskonto" => {
                map[<Column as Into<usize>>::into(Column::ContractAccount)] = Some(i);
            }

            "korresp. empfänger" => {
                map[<Column as Into<usize>>::into(Column::CorrespondingBillReceiver)] = Some(i);
            }

            "korrespondenzempfänger" => {
                map[<Column as Into<usize>>::into(Column::CorrespondingBillReceiver)] = Some(i);
            }

            "abw.rechnungsempfg" => {
                map[<Column as Into<usize>>::into(Column::DeviantBillReceiver)] = Some(i);
            }

            "abw rechnungsempfänger" => {
                map[<Column as Into<usize>>::into(Column::DeviantBillReceiver)] = Some(i);
            }

            "e-rechnung zpdf=ja" => {
                map[<Column as Into<usize>>::into(Column::EInvoice)] = Some(i);
            }

            "e-mail rechnung zpdf=ja" => {
                map[<Column as Into<usize>>::into(Column::EInvoice)] = Some(i);
            }

            "serviceart" => {
                map[<Column as Into<usize>>::into(Column::EnergyType)] = Some(i);
            }

            "sparte" => {
                map[<Column as Into<usize>>::into(Column::EnergyType)] = Some(i);
            }

            "vorleistungsmodellnetz" => {
                map[<Column as Into<usize>>::into(Column::GridBillingIntegrated)] = Some(i);
            }

            "einzugsdatum" => {
                map[<Column as Into<usize>>::into(Column::InDate)] = Some(i);
            }

            "zählpunktbezeichnung" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "zählpunkt" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "lagezusatz" => {
                map[<Column as Into<usize>>::into(Column::MpAddition)] = Some(i);
            }

            "a-ort" => {
                map[<Column as Into<usize>>::into(Column::MpCity)] = Some(i);
            }

            "a-hausnummer" => {
                map[<Column as Into<usize>>::into(Column::MpNumber)] = Some(i);
            }

            "a-straße" => {
                map[<Column as Into<usize>>::into(Column::MpStreet)] = Some(i);
            }

            "a-postleitzahl" => {
                map[<Column as Into<usize>>::into(Column::MpZip)] = Some(i);
            }

            "a-plz" => {
                map[<Column as Into<usize>>::into(Column::MpZip)] = Some(i);
            }

            "name1" => {
                map[<Column as Into<usize>>::into(Column::Name)] = Some(i);
            }

            "name2" => {
                map[<Column as Into<usize>>::into(Column::NameAdd)] = Some(i);
            }

            "auszugsdatum" => {
                map[<Column as Into<usize>>::into(Column::OutDate)] = Some(i);
            }

            "poolbetreiber-kundennummer" => {
                map[<Column as Into<usize>>::into(Column::PoolCustomerId)] = Some(i);
            }

            "bez. des profils" => {
                map[<Column as Into<usize>>::into(Column::Profile)] = Some(i);
            }

            "lastprofil" => {
                map[<Column as Into<usize>>::into(Column::Profile)] = Some(i);
            }

            "ableseeinheit" => {
                map[<Column as Into<usize>>::into(Column::ReadUnit)] = Some(i);
            }

            "sepa hinterlegt e=ja" => {
                map[<Column as Into<usize>>::into(Column::Sepa)] = Some(i);
            }

            "sepa" => {
                map[<Column as Into<usize>>::into(Column::Sepa)] = Some(i);
            }

            "sepa gesperrt" => {
                map[<Column as Into<usize>>::into(Column::SepaBlocked)] = Some(i);
            }

            "sepa sperre" => {
                map[<Column as Into<usize>>::into(Column::SepaBlocked)] = Some(i);
            }

            "gruppenkopf_av" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerGroupId)] = Some(i);
            }

            "gruppenkopf" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerGroupId)] = Some(i);
            }

            "gruppenkopf_av_name" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerGroupName)] = Some(i);
            }

            "name-gruppenkopf" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerGroupName)] = Some(i);
            }

            "geschäftspartner" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerId)] = Some(i);
            }

            "anlage" => {
                map[<Column as Into<usize>>::into(Column::SupplierMeterpointId)] = Some(i);
            }

            "tariftyp" => {
                map[<Column as Into<usize>>::into(Column::TariffTyp)] = Some(i);
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

        if r.out_date == NaiveDate::from_ymd_opt(9999, 12, 31) {
            r.out_date = None;
        }

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
        affiliate: row[column_map[Column::Affiliate.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_addition: row[column_map[Column::BillAddition.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_city: row[column_map[Column::BillCity.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_number: row[column_map[Column::BillNumber.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_street: row[column_map[Column::BillStreet.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_zip: row[column_map[Column::BillZip.usize()]]
            .to_string()
            .trim()
            .to_string(),

        billing_type: row[column_map[Column::BillingType.usize()]]
            .to_string()
            .trim()
            .to_string(),

        consumption_at_change: row[column_map[Column::ConsumptionAtChange.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Wert bei Wechsel kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        consumption_forecast: row[column_map[Column::ConsumptionForecast.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "aktueller Prognosewert kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        contract: row[column_map[Column::Contract.usize()]]
            .to_string()
            .trim()
            .to_string(),

        contract_account: row[column_map[Column::ContractAccount.usize()]]
            .to_string()
            .trim()
            .to_string(),

        corresponding_bill_receiver: row[column_map[Column::CorrespondingBillReceiver.usize()]]
            .to_string()
            .trim()
            .to_string(),

        deviant_bill_receiver: row[column_map[Column::DeviantBillReceiver.usize()]]
            .to_string()
            .trim()
            .to_string(),

        e_invoice: row[column_map[Column::EInvoice.usize()]]
            .to_string()
            .trim()
            .to_string(),

        energy_type: row[column_map[Column::EnergyType.usize()]]
            .to_string()
            .trim()
            .to_string(),

        grid_billing_integrated: row[column_map[Column::GridBillingIntegrated.usize()]]
            .to_string()
            .trim()
            .to_string(),

        in_date: row[column_map[Column::InDate.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Einzugsdatum".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        meterpoint: row[column_map[Column::Meterpoint.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_addition: row[column_map[Column::MpAddition.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_city: row[column_map[Column::MpCity.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_number: row[column_map[Column::MpNumber.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_street: row[column_map[Column::MpStreet.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_zip: row[column_map[Column::MpZip.usize()]]
            .to_string()
            .trim()
            .to_string(),

        name: row[column_map[Column::Name.usize()]]
            .to_string()
            .trim()
            .to_string(),

        name_add: row[column_map[Column::NameAdd.usize()]]
            .to_string()
            .trim()
            .to_string(),

        out_date: row[column_map[Column::OutDate.usize()]].as_date(),

        pool_customer_id: row[column_map[Column::PoolCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        profile: row[column_map[Column::Profile.usize()]]
            .to_string()
            .trim()
            .to_string(),

        read_unit: row[column_map[Column::ReadUnit.usize()]]
            .to_string()
            .trim()
            .to_string(),

        sepa: row[column_map[Column::Sepa.usize()]]
            .to_string()
            .trim()
            .to_string(),

        sepa_blocked: row[column_map[Column::SepaBlocked.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_group_id: row[column_map[Column::SupplierCustomerGroupId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_group_name: row[column_map[Column::SupplierCustomerGroupName.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_id: row[column_map[Column::SupplierCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_meterpoint_id: row[column_map[Column::SupplierMeterpointId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        tariff_typ: row[column_map[Column::TariffTyp.usize()]]
            .to_string()
            .trim()
            .to_string(),
    };

    Ok(r)
}
