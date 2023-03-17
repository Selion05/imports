use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Reader};
use chrono::NaiveDate;
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    #[serde(rename = "bill_street")]
    bill_street: String,
    #[serde(rename = "bill_zip")]
    bill_zip: String,
    #[serde(rename = "bill_number")]
    bill_number: String,
    #[serde(rename = "bill_addition")]
    bill_addition: String,
    #[serde(rename = "bill_city")]
    bill_city: String,
    #[serde(rename = "mp_addition")]
    mp_addition: String,
    #[serde(rename = "mp_city")]
    mp_city: String,
    #[serde(rename = "mp_zip")]
    mp_zip: String,
    #[serde(rename = "mp_number")]
    mp_number: String,
    #[serde(rename = "mp_street")]
    mp_street: String,
    corresponding_bill_receiver: String,
    out_date: Option<NaiveDate>,
    supplier_customer_group_id: String,
    supplier_customer_group_name: String,
    sepa: String,
    consumption_forecast: f64,
    grid_billing_integrated: String,
    affiliate: String,
    contract_account: String,
    tariff_typ: String,
    energy_type: String,
    e_invoice: String,
    deviant_bill_receiver: String,
    meterpoint: String,
    supplier_meterpoint_id: String,
    pool_customer_id: String,
    billing_type: String,
    sepa_blocked: String,
    contract: String,
    consumption_at_change: f64,
    name_add: String,
    name: String,
    in_date: NaiveDate,
    read_unit: String,
    supplier_customer_id: String,
    profile: String,
}

#[derive(Debug)]
enum Column {
    MpAddition,
    OutDate,
    BillStreet,
    BillZip,
    CorrespondingBillReceiver,
    SupplierCustomerGroupId,
    SupplierCustomerGroupName,
    Sepa,
    ConsumptionForecast,
    GridBillingIntegrated,
    BillNumber,
    Affiliate,
    MpNumber,
    ContractAccount,
    TariffTyp,
    EnergyType,
    EInvoice,
    MpCity,
    DeviantBillReceiver,
    Meterpoint,
    SupplierMeterpointId,
    PoolCustomerId,
    MpZip,
    BillingType,
    SepaBlocked,
    Contract,
    BillCity,
    ConsumptionAtChange,
    NameAdd,
    Name,
    InDate,
    MpStreet,
    ReadUnit,
    BillAddition,
    SupplierCustomerId,
    Profile,
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
            "lagezusatz" => {
                map[<Column as Into<usize>>::into(Column::MpAddition)] = Some(i);
            }

            "auszugsdatum" => {
                map[<Column as Into<usize>>::into(Column::OutDate)] = Some(i);
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

            "korresp. empfänger" => {
                map[<Column as Into<usize>>::into(Column::CorrespondingBillReceiver)] = Some(i);
            }

            "korrespondenzempfänger" => {
                map[<Column as Into<usize>>::into(Column::CorrespondingBillReceiver)] = Some(i);
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

            "sepa hinterlegt e=ja" => {
                map[<Column as Into<usize>>::into(Column::Sepa)] = Some(i);
            }

            "sepa" => {
                map[<Column as Into<usize>>::into(Column::Sepa)] = Some(i);
            }

            "aktueller prognosewert kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionForecast)] = Some(i);
            }

            "prognosewert kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionForecast)] = Some(i);
            }

            "vorleistungsmodellnetz" => {
                map[<Column as Into<usize>>::into(Column::GridBillingIntegrated)] = Some(i);
            }

            "r-hausnummer" => {
                map[<Column as Into<usize>>::into(Column::BillNumber)] = Some(i);
            }

            "zugehörigkeit" => {
                map[<Column as Into<usize>>::into(Column::Affiliate)] = Some(i);
            }

            "a-hausnummer" => {
                map[<Column as Into<usize>>::into(Column::MpNumber)] = Some(i);
            }

            "vertragskonto" => {
                map[<Column as Into<usize>>::into(Column::ContractAccount)] = Some(i);
            }

            "tariftyp" => {
                map[<Column as Into<usize>>::into(Column::TariffTyp)] = Some(i);
            }

            "serviceart" => {
                map[<Column as Into<usize>>::into(Column::EnergyType)] = Some(i);
            }

            "sparte" => {
                map[<Column as Into<usize>>::into(Column::EnergyType)] = Some(i);
            }

            "e-rechnung zpdf=ja" => {
                map[<Column as Into<usize>>::into(Column::EInvoice)] = Some(i);
            }

            "e-mail rechnung zpdf=ja" => {
                map[<Column as Into<usize>>::into(Column::EInvoice)] = Some(i);
            }

            "a-ort" => {
                map[<Column as Into<usize>>::into(Column::MpCity)] = Some(i);
            }

            "abw.rechnungsempfg" => {
                map[<Column as Into<usize>>::into(Column::DeviantBillReceiver)] = Some(i);
            }

            "abw rechnungsempfänger" => {
                map[<Column as Into<usize>>::into(Column::DeviantBillReceiver)] = Some(i);
            }

            "zählpunktbezeichnung" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "zählpunkt" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "anlage" => {
                map[<Column as Into<usize>>::into(Column::SupplierMeterpointId)] = Some(i);
            }

            "poolbetreiber-kundennummer" => {
                map[<Column as Into<usize>>::into(Column::PoolCustomerId)] = Some(i);
            }

            "a-postleitzahl" => {
                map[<Column as Into<usize>>::into(Column::MpZip)] = Some(i);
            }

            "a-plz" => {
                map[<Column as Into<usize>>::into(Column::MpZip)] = Some(i);
            }

            "abrechnungsklasse" => {
                map[<Column as Into<usize>>::into(Column::BillingType)] = Some(i);
            }

            "sepa gesperrt" => {
                map[<Column as Into<usize>>::into(Column::SepaBlocked)] = Some(i);
            }

            "sepa sperre" => {
                map[<Column as Into<usize>>::into(Column::SepaBlocked)] = Some(i);
            }

            "vertrag" => {
                map[<Column as Into<usize>>::into(Column::Contract)] = Some(i);
            }

            "r-ort" => {
                map[<Column as Into<usize>>::into(Column::BillCity)] = Some(i);
            }

            "wert bei wechsel kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionAtChange)] = Some(i);
            }

            "jahresverbrauch bei wechsel kwh" => {
                map[<Column as Into<usize>>::into(Column::ConsumptionAtChange)] = Some(i);
            }

            "name2" => {
                map[<Column as Into<usize>>::into(Column::NameAdd)] = Some(i);
            }

            "name1" => {
                map[<Column as Into<usize>>::into(Column::Name)] = Some(i);
            }

            "einzugsdatum" => {
                map[<Column as Into<usize>>::into(Column::InDate)] = Some(i);
            }

            "a-straße" => {
                map[<Column as Into<usize>>::into(Column::MpStreet)] = Some(i);
            }

            "ableseeinheit" => {
                map[<Column as Into<usize>>::into(Column::ReadUnit)] = Some(i);
            }

            "r-zusatz" => {
                map[<Column as Into<usize>>::into(Column::BillAddition)] = Some(i);
            }

            "geschäftspartner" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerId)] = Some(i);
            }

            "bez. des profils" => {
                map[<Column as Into<usize>>::into(Column::Profile)] = Some(i);
            }

            "lastprofil" => {
                map[<Column as Into<usize>>::into(Column::Profile)] = Some(i);
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
        mp_addition: row[column_map[Column::MpAddition.usize()]]
            .to_string()
            .trim()
            .to_string(),

        out_date: row[column_map[Column::OutDate.usize()]].as_date(),

        bill_street: row[column_map[Column::BillStreet.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_zip: row[column_map[Column::BillZip.usize()]]
            .to_string()
            .trim()
            .to_string(),

        corresponding_bill_receiver: row[column_map[Column::CorrespondingBillReceiver.usize()]]
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

        sepa: row[column_map[Column::Sepa.usize()]]
            .to_string()
            .trim()
            .to_string(),

        consumption_forecast: row[column_map[Column::ConsumptionForecast.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "aktueller Prognosewert kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_billing_integrated: row[column_map[Column::GridBillingIntegrated.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_number: row[column_map[Column::BillNumber.usize()]]
            .to_string()
            .trim()
            .to_string(),

        affiliate: row[column_map[Column::Affiliate.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_number: row[column_map[Column::MpNumber.usize()]]
            .to_string()
            .trim()
            .to_string(),

        contract_account: row[column_map[Column::ContractAccount.usize()]]
            .to_string()
            .trim()
            .to_string(),

        tariff_typ: row[column_map[Column::TariffTyp.usize()]]
            .to_string()
            .trim()
            .to_string(),

        energy_type: row[column_map[Column::EnergyType.usize()]]
            .to_string()
            .trim()
            .to_string(),

        e_invoice: row[column_map[Column::EInvoice.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_city: row[column_map[Column::MpCity.usize()]]
            .to_string()
            .trim()
            .to_string(),

        deviant_bill_receiver: row[column_map[Column::DeviantBillReceiver.usize()]]
            .to_string()
            .trim()
            .to_string(),

        meterpoint: row[column_map[Column::Meterpoint.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_meterpoint_id: row[column_map[Column::SupplierMeterpointId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        pool_customer_id: row[column_map[Column::PoolCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        mp_zip: row[column_map[Column::MpZip.usize()]]
            .to_string()
            .trim()
            .to_string(),

        billing_type: row[column_map[Column::BillingType.usize()]]
            .to_string()
            .trim()
            .to_string(),

        sepa_blocked: row[column_map[Column::SepaBlocked.usize()]]
            .to_string()
            .trim()
            .to_string(),

        contract: row[column_map[Column::Contract.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_city: row[column_map[Column::BillCity.usize()]]
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

        name_add: row[column_map[Column::NameAdd.usize()]]
            .to_string()
            .trim()
            .to_string(),

        name: row[column_map[Column::Name.usize()]]
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

        mp_street: row[column_map[Column::MpStreet.usize()]]
            .to_string()
            .trim()
            .to_string(),

        read_unit: row[column_map[Column::ReadUnit.usize()]]
            .to_string()
            .trim()
            .to_string(),

        bill_addition: row[column_map[Column::BillAddition.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_id: row[column_map[Column::SupplierCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        profile: row[column_map[Column::Profile.usize()]]
            .to_string()
            .trim()
            .to_string(),
    };

    Ok(r)
}

#[cfg(test)]
mod tests {
    use crate::kam::run;

    #[test]
    fn test_get_column_map_success_with_ordered_columns() {
        let result = run("var/kam.xlsx");
        assert!(result.is_ok());

        let result = result.unwrap();

        let rows = result.get("10051234");

        assert!(rows.is_some());

        let rows = rows.unwrap();

        assert_eq!(rows.len(), 2);

        let rows = result.get("15091234");

        assert!(rows.is_some());

        let rows = rows.unwrap();

        assert_eq!(rows.len(), 1);
    }
}
