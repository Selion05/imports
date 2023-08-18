use crate::commission::run;
use chrono::NaiveDate;

#[test]
fn test_get_column_map_success_with_ordered_columns() {
    let result = run("var/commission.xlsx");
    assert!(result.is_ok());

    let result = result.unwrap();

    let rows = result.get("123456789");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 4);

    assert_eq!(rows[0].billing_amount, 30343.0);
    assert_eq!(rows[0].contract_account, "123456789");
    assert_eq!(rows[0].currency, "EUR");
    assert_eq!(
        rows[0].entry_date,
        NaiveDate::from_ymd_opt(2022, 08, 16).unwrap()
    );
    assert_eq!(rows[0].meterpoint, "AT0010000000000000001000004107355");
    assert_eq!(rows[0].name, "Company");
    assert_eq!(rows[0].net_amount, "13.35092".to_string());
    assert_eq!(rows[0].price, "0.00044".to_string());
    assert_eq!(rows[0].print_receipt, "300051234");
    assert_eq!(rows[0].stgrbt, "SBPROV");
    assert_eq!(rows[0].supplier_customer_id, "123456789");
    assert_eq!(rows[0]._type, "Strom");
    assert_eq!(
        rows[0].valid_from,
        NaiveDate::from_ymd_opt(2021, 05, 01).unwrap()
    );
    assert_eq!(
        rows[0].valid_to,
        NaiveDate::from_ymd_opt(2021, 12, 31).unwrap()
    );

    assert_eq!(rows[1].net_amount, "8".to_string());
    assert_eq!(rows[2].net_amount, "87.1299".to_string());
    assert_eq!(rows[3].net_amount, "159.8535".to_string());
}
