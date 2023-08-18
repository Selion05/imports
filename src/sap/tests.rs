use crate::sap::run;

#[test]
fn test_get_column_map_success_with_ordered_columns() {
    let result = run("var/sap.xlsx");
    assert!(result.is_ok());

    let result = result.unwrap();

    let rows = result.get("310651234");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 4);
}
