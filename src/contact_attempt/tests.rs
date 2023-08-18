use crate::contact_attempt::run;
use chrono::NaiveDate;

#[test]
fn test_get_column_map_success_with_ordered_columns() {
    let result = run("var/contact_attempt.xlsx");
    assert!(result.is_ok());

    let result = result.unwrap();

    let rows = result.get("1");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 2);
}
