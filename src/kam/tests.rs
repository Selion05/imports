use crate::kam::run;
use chrono::NaiveDate;

#[test]
fn test_get_column_map_success_with_ordered_columns() {
    let result = run("var/kam.xlsx");
    assert!(result.is_ok());

    let result = result.unwrap();

    let rows = result.get("10051234");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 2);

    assert_eq!(rows[0].out_date, None);
    assert_eq!(rows[1].out_date, NaiveDate::from_ymd_opt(2023, 12, 31));

    let rows = result.get("15091234");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 1);
}
