use crate::datentraeger::run;

#[test]
fn test_get_column_map_success_with_ordered_columns() {
    let result = run("var/datentraeger.xlsx");
    assert!(result.is_ok());

    let result = result.unwrap();

    let rows = result.get("15205123");

    assert!(rows.is_some());

    let rows = rows.unwrap();

    assert_eq!(rows.len(), 4);

    assert_eq!(rows[0].invoice_type, "Monatsabrechnung");
    assert_eq!(rows[1].invoice_type, "Schlussabrechnung");
    assert_eq!(rows[2].invoice_type, "Jahresabrechnung");
    assert_eq!(rows[3].invoice_type, "TZB");
}
