use crate::meterpoint_value::run;
use chrono::NaiveDate;

#[test]
fn test_get_column_map_success_with_ordered_columns() {
    let result = run("var/meterpoint_value.xlsx");
    assert!(result.is_ok());

    let result = result.unwrap();

    // let rows = result.get("2022-10-01");
    let rows = result.get("all");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 1);

    let row = rows.get(0).unwrap();

    assert_eq!(
        row.columns,
        vec![
            String::from("AT0020000000000000000000100003400"),
            String::from("AT0031000000000000000000141934000")
        ]
    );

    let d = NaiveDate::from_ymd_opt(2022, 10, 1).unwrap();
    assert_eq!(
        row.index,
        vec![
            d.and_hms_opt(0, 0, 0).unwrap(),
            d.and_hms_opt(0, 15, 0).unwrap(),
            d.and_hms_opt(0, 30, 0).unwrap()
        ]
    );

    assert_eq!(
        row.data,
        vec![
            vec![Some(2.04), Some(0.435)],
            vec![Some(2.28), Some(0.435)],
            vec![Some(1.68), Some(0.42)],
        ]
    )
}
