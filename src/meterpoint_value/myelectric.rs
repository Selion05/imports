use crate::meterpoint_value::Row;
use crate::ImportError;
use calamine::{DataType, Range};
use chrono::SubsecRound;
use std::collections::HashMap;

pub fn run(sheet: Range<DataType>) -> Result<HashMap<String, Vec<Row>>, ImportError> {
    let header_row = 0;
    let data_start_row = 1;

    let headers: Vec<String> = sheet
        .rows()
        .nth(header_row)
        .unwrap()
        .iter()
        .skip(1)
        .map(|header_cell| {
            return match header_cell.get_string() {
                None => String::from(""),
                Some(v) => v[..33].to_string(),
            };
        })
        .collect();

    let mut groups: HashMap<String, Vec<Row>> = HashMap::new();

    let header_cnt = headers.len();
    let mut r = Row {
        columns: headers,
        index: vec![],
        data: vec![],
    };

    for (i, row) in sheet.rows().enumerate().skip(data_start_row) {
        let summary_row = row[0].to_string();
        if summary_row == "Summe" || summary_row == "Sum" {
            println!("found");
            // meterpoint_value files contain a summary row
            break;
        }
        let date = row[0]
            .as_datetime()
            .ok_or_else(|| {
                ImportError::ValueError(
                    i,
                    "Timestamp".to_string(),
                    "could not parse datetime".to_string(),
                )
            })?
            .round_subsecs(0);

        let values = row
            .iter()
            .skip(1)
            .take(header_cnt)
            .map(|v| {
                return v.get_float();
            })
            .collect();

        r.index.push(date);
        r.data.push(values);
    }

    let k = "all".to_string();
    groups.insert(k.clone(), vec![r]);

    Ok(groups)
}

#[cfg(test)]
mod tests {
    use crate::meterpoint_value::run;
    use chrono::NaiveDate;

    #[test]
    fn test_parse_myelectric_is_successful() {
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
}
