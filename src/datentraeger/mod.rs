#[cfg(test)]
mod tests;

use crate::ImportError;
use calamine::{open_workbook_auto, DataType, Reader};
use serde::Serialize;
use std::collections::HashMap;
use std::fmt::Debug;

use chrono::NaiveDate;

#[derive(Debug, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Row {
    addition: String,
    city: String,
    commission_price: f64,
    contract_account: String,
    energy_amount: f64,
    energy_base_amount: f64,
    energy_base_price: f64,
    energy_consumption: f64,
    energy_contract: String,
    energy_fee: f64,
    energy_law_amount: f64,
    energy_law_price: f64,
    energy_tranche_net_to_pay: f64,
    energy_usage_fee: f64,
    entry_exit_amount: f64,
    entry_exit_price: f64,
    grid_base_amount: f64,
    grid_base_price: f64,
    grid_consumption: f64,
    grid_consumption_ht: f64,
    grid_consumption_nt: f64,
    grid_contract: String,
    grid_fee: f64,
    grid_operator: String,
    grid_power: f64,
    grid_power_amount: f64,
    grid_tranche_net_to_pay: f64,
    handling_fee_amount: f64,
    handling_fee_price: f64,
    invoice: String,
    invoice_date: NaiveDate,
    invoice_type: String,
    meter_fee: f64,
    meterpoint: String,
    name: String,
    net_due: NaiveDate,
    number: String,
    price_zone: Option<f64>,
    proof_of_origin_amount: f64,
    proof_of_origin_price: f64,
    reactive_energy_consumption_amount: f64,
    reactive_energy_consumption_to_pay: f64,
    street: String,
    supplier_customer_id: String,
    total_vat: f64,
    valid_from: NaiveDate,
    valid_to: NaiveDate,
    vat: f64,
    working_price: f64,
    zip: String,
}

#[derive(Debug)]
enum Column {
    Addition,
    City,
    CommissionPrice,
    ContractAccount,
    EnergyAmount,
    EnergyBaseAmount,
    EnergyBasePrice,
    EnergyConsumption,
    EnergyContract,
    EnergyFee,
    EnergyLawAmount,
    EnergyLawPrice,
    EnergyTrancheNetToPay,
    EnergyUsageFee,
    EntryExitAmount,
    EntryExitPrice,
    GridBaseAmount,
    GridBasePrice,
    GridConsumption,
    GridConsumptionHt,
    GridConsumptionNt,
    GridContract,
    GridFee,
    GridOperator,
    GridPower,
    GridPowerAmount,
    GridTrancheNetToPay,
    HandlingFeeAmount,
    HandlingFeePrice,
    Invoice,
    InvoiceDate,
    InvoiceType,
    MeterFee,
    Meterpoint,
    Name,
    NetDue,
    Number,
    PriceZone,
    ProofOfOriginAmount,
    ProofOfOriginPrice,
    ReactiveEnergyConsumptionAmount,
    ReactiveEnergyConsumptionToPay,
    Street,
    SupplierCustomerId,
    TotalVat,
    ValidFrom,
    ValidTo,
    Vat,
    WorkingPrice,
    Zip,
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
    let mut map: Vec<Option<usize>> = vec![None; 50];
    for (i, header) in headers.iter().enumerate() {
        match header.to_lowercase().trim() {
            "zusatz" => {
                map[<Column as Into<usize>>::into(Column::Addition)] = Some(i);
            }

            "ort" => {
                map[<Column as Into<usize>>::into(Column::City)] = Some(i);
            }

            "provisionspreis energie cent/kwh" => {
                map[<Column as Into<usize>>::into(Column::CommissionPrice)] = Some(i);
            }

            "vk mye" => {
                map[<Column as Into<usize>>::into(Column::ContractAccount)] = Some(i);
            }

            "energie €/betrag" => {
                map[<Column as Into<usize>>::into(Column::EnergyAmount)] = Some(i);
            }

            "grundpreis energie €/betrag" => {
                map[<Column as Into<usize>>::into(Column::EnergyBaseAmount)] = Some(i);
            }

            "grundpreis energie €/monat" => {
                map[<Column as Into<usize>>::into(Column::EnergyBasePrice)] = Some(i);
            }

            "energie kwh" => {
                map[<Column as Into<usize>>::into(Column::EnergyConsumption)] = Some(i);
            }

            "e-vertrag" => {
                map[<Column as Into<usize>>::into(Column::EnergyContract)] = Some(i);
            }

            "energieabgabe betr." => {
                map[<Column as Into<usize>>::into(Column::EnergyFee)] = Some(i);
            }

            "eeffg €/betrag" => {
                map[<Column as Into<usize>>::into(Column::EnergyLawAmount)] = Some(i);
            }

            "eeffg cent/kwh" => {
                map[<Column as Into<usize>>::into(Column::EnergyLawPrice)] = Some(i);
            }

            "off.tb.ener.netto" => {
                map[<Column as Into<usize>>::into(Column::EnergyTrancheNetToPay)] = Some(i);
            }

            "gebrauchsabgabe energie" => {
                map[<Column as Into<usize>>::into(Column::EnergyUsageFee)] = Some(i);
            }

            "entry exit entgelt €/betrag" => {
                map[<Column as Into<usize>>::into(Column::EntryExitAmount)] = Some(i);
            }

            "entry exit entgelt cent/kwh" => {
                map[<Column as Into<usize>>::into(Column::EntryExitPrice)] = Some(i);
            }

            "grundpreis netz €/betrag" => {
                map[<Column as Into<usize>>::into(Column::GridBaseAmount)] = Some(i);
            }

            "grundpreis netz €/jahr" => {
                map[<Column as Into<usize>>::into(Column::GridBasePrice)] = Some(i);
            }

            "netzverbrauch gesamt" => {
                map[<Column as Into<usize>>::into(Column::GridConsumption)] = Some(i);
            }

            "netzverbrauch ht" => {
                map[<Column as Into<usize>>::into(Column::GridConsumptionHt)] = Some(i);
            }

            "netzverbrauch nt" => {
                map[<Column as Into<usize>>::into(Column::GridConsumptionNt)] = Some(i);
            }

            "n-vertrag" => {
                map[<Column as Into<usize>>::into(Column::GridContract)] = Some(i);
            }

            "netzarb./netzverl./gebr.abg." => {
                map[<Column as Into<usize>>::into(Column::GridFee)] = Some(i);
            }

            "netzbetreiber" => {
                map[<Column as Into<usize>>::into(Column::GridOperator)] = Some(i);
            }

            "netzleistung kw" => {
                map[<Column as Into<usize>>::into(Column::GridPower)] = Some(i);
            }

            "netzleistung betrag" => {
                map[<Column as Into<usize>>::into(Column::GridPowerAmount)] = Some(i);
            }

            "off.tb.netz netto" => {
                map[<Column as Into<usize>>::into(Column::GridTrancheNetToPay)] = Some(i);
            }

            "abwicklungsbeitrag betrag" => {
                map[<Column as Into<usize>>::into(Column::HandlingFeeAmount)] = Some(i);
            }

            "abwicklungsbeitrag kwh" => {
                map[<Column as Into<usize>>::into(Column::HandlingFeePrice)] = Some(i);
            }

            "einzelrechnung" => {
                map[<Column as Into<usize>>::into(Column::Invoice)] = Some(i);
            }

            "belegdatum" => {
                map[<Column as Into<usize>>::into(Column::InvoiceDate)] = Some(i);
            }

            "belegart" => {
                map[<Column as Into<usize>>::into(Column::InvoiceType)] = Some(i);
            }

            "messpreis/mieten" => {
                map[<Column as Into<usize>>::into(Column::MeterFee)] = Some(i);
            }

            "zählpunktbezeichnung" => {
                map[<Column as Into<usize>>::into(Column::Meterpoint)] = Some(i);
            }

            "name" => {
                map[<Column as Into<usize>>::into(Column::Name)] = Some(i);
            }

            "nettofälligkeit" => {
                map[<Column as Into<usize>>::into(Column::NetDue)] = Some(i);
            }

            "hausnummer" => {
                map[<Column as Into<usize>>::into(Column::Number)] = Some(i);
            }

            "preiszonentrennung €/betrag" => {
                map[<Column as Into<usize>>::into(Column::PriceZone)] = Some(i);
            }

            "hkn €/betrag" => {
                map[<Column as Into<usize>>::into(Column::ProofOfOriginAmount)] = Some(i);
            }

            "hkn cent/kwh" => {
                map[<Column as Into<usize>>::into(Column::ProofOfOriginPrice)] = Some(i);
            }

            "blindverbrauch betrag" => {
                map[<Column as Into<usize>>::into(Column::ReactiveEnergyConsumptionAmount)] =
                    Some(i);
            }

            "blindverbrauch verr." => {
                map[<Column as Into<usize>>::into(Column::ReactiveEnergyConsumptionToPay)] =
                    Some(i);
            }

            "straße" => {
                map[<Column as Into<usize>>::into(Column::Street)] = Some(i);
            }

            "gp mye" => {
                map[<Column as Into<usize>>::into(Column::SupplierCustomerId)] = Some(i);
            }

            "netto ustpf." => {
                map[<Column as Into<usize>>::into(Column::TotalVat)] = Some(i);
            }

            "gültig ab" => {
                map[<Column as Into<usize>>::into(Column::ValidFrom)] = Some(i);
            }

            "gültig bis" => {
                map[<Column as Into<usize>>::into(Column::ValidTo)] = Some(i);
            }

            "umsatzsteuer" => {
                map[<Column as Into<usize>>::into(Column::Vat)] = Some(i);
            }

            "arbeitspreis energie cent/kwh" => {
                map[<Column as Into<usize>>::into(Column::WorkingPrice)] = Some(i);
            }

            "plz" => {
                map[<Column as Into<usize>>::into(Column::Zip)] = Some(i);
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
        addition: row[column_map[Column::Addition.usize()]]
            .to_string()
            .trim()
            .to_string(),

        city: row[column_map[Column::City.usize()]]
            .to_string()
            .trim()
            .to_string(),

        commission_price: row[column_map[Column::CommissionPrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Provisionspreis Energie Cent/kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        contract_account: row[column_map[Column::ContractAccount.usize()]]
            .to_string()
            .trim()
            .to_string(),

        energy_amount: row[column_map[Column::EnergyAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Energie €/Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_base_amount: row[column_map[Column::EnergyBaseAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Grundpreis Energie €/Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_base_price: row[column_map[Column::EnergyBasePrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Grundpreis Energie €/Monat".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_consumption: row[column_map[Column::EnergyConsumption.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Energie kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_contract: row[column_map[Column::EnergyContract.usize()]]
            .to_string()
            .trim()
            .to_string(),

        energy_fee: row[column_map[Column::EnergyFee.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Energieabgabe Betr.".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_law_amount: row[column_map[Column::EnergyLawAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "EEffG €/Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_law_price: row[column_map[Column::EnergyLawPrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "EEffG Cent/kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_tranche_net_to_pay: row[column_map[Column::EnergyTrancheNetToPay.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Off.Tb.Ener.Netto".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        energy_usage_fee: row[column_map[Column::EnergyUsageFee.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Gebrauchsabgabe Energie".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        entry_exit_amount: row[column_map[Column::EntryExitAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Entry Exit Entgelt €/Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        entry_exit_price: row[column_map[Column::EntryExitPrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Entry Exit Entgelt Cent/kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_base_amount: row[column_map[Column::GridBaseAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Grundpreis Netz €/Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_base_price: row[column_map[Column::GridBasePrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Grundpreis Netz €/Jahr".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_consumption: row[column_map[Column::GridConsumption.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netzverbrauch Gesamt".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_consumption_ht: row[column_map[Column::GridConsumptionHt.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netzverbrauch HT".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_consumption_nt: row[column_map[Column::GridConsumptionNt.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netzverbrauch NT".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_contract: row[column_map[Column::GridContract.usize()]]
            .to_string()
            .trim()
            .to_string(),

        grid_fee: row[column_map[Column::GridFee.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netzarb./Netzverl./Gebr.abg.".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_operator: row[column_map[Column::GridOperator.usize()]]
            .to_string()
            .trim()
            .to_string(),

        grid_power: row[column_map[Column::GridPower.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netzleistung kW".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_power_amount: row[column_map[Column::GridPowerAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netzleistung Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        grid_tranche_net_to_pay: row[column_map[Column::GridTrancheNetToPay.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Off.Tb.Netz Netto".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        handling_fee_amount: row[column_map[Column::HandlingFeeAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Abwicklungsbeitrag Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        handling_fee_price: row[column_map[Column::HandlingFeePrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Abwicklungsbeitrag kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        invoice: row[column_map[Column::Invoice.usize()]]
            .to_string()
            .trim()
            .to_string(),

        invoice_date: row[column_map[Column::InvoiceDate.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Belegdatum".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        invoice_type: row[column_map[Column::InvoiceType.usize()]]
            .to_string()
            .trim()
            .to_string(),

        meter_fee: row[column_map[Column::MeterFee.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Messpreis/Mieten".to_string(),
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

        net_due: row[column_map[Column::NetDue.usize()]]
            .as_date()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Nettofälligkeit".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        number: row[column_map[Column::Number.usize()]]
            .to_string()
            .trim()
            .to_string(),

        price_zone: row[column_map[Column::PriceZone.usize()]].get_float(),

        proof_of_origin_amount: row[column_map[Column::ProofOfOriginAmount.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "HKN €/Betrag".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        proof_of_origin_price: row[column_map[Column::ProofOfOriginPrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "HKN Cent/kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        reactive_energy_consumption_amount: row
            [column_map[Column::ReactiveEnergyConsumptionAmount.usize()]]
        .get_float()
        .ok_or_else(|| {
            ImportError::ValueError(
                row_number,
                "Blindverbrauch Betrag".to_string(),
                "Cell has no value".to_string(),
            )
        })?,

        reactive_energy_consumption_to_pay: row
            [column_map[Column::ReactiveEnergyConsumptionToPay.usize()]]
        .get_float()
        .ok_or_else(|| {
            ImportError::ValueError(
                row_number,
                "Blindverbrauch Verr.".to_string(),
                "Cell has no value".to_string(),
            )
        })?,

        street: row[column_map[Column::Street.usize()]]
            .to_string()
            .trim()
            .to_string(),

        supplier_customer_id: row[column_map[Column::SupplierCustomerId.usize()]]
            .to_string()
            .trim()
            .to_string(),

        total_vat: row[column_map[Column::TotalVat.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Netto Ustpf.".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

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

        vat: row[column_map[Column::Vat.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Umsatzsteuer".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        working_price: row[column_map[Column::WorkingPrice.usize()]]
            .get_float()
            .ok_or_else(|| {
                ImportError::ValueError(
                    row_number,
                    "Arbeitspreis Energie Cent/kWh".to_string(),
                    "Cell has no value".to_string(),
                )
            })?,

        zip: row[column_map[Column::Zip.usize()]]
            .to_string()
            .trim()
            .to_string(),
    };

    Ok(r)
}
