use crate::meterpoint_value::Data;
use crate::ImportError;
use calamine::{DataType, Range};
use chrono::SubsecRound;

pub fn run(sheet: Range<DataType>) -> Result<Data, ImportError> {
    let headers: Vec<String> = sheet
        .rows()
        .nth(6)
        .unwrap()
        .iter()
        .skip(2)
        .map(|header_cell| {
            return match header_cell.get_string() {
                None => String::from(""),
                Some(v) => v.to_string(),
            };
        })
        .collect();

    let header_cnt = headers.len();
    let mut r = Data {
        columns: headers,
        index: vec![],
        data: vec![],
    };

    for (i, row) in sheet.rows().enumerate().skip(14) {
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
            .skip(2)
            .take(header_cnt)
            .map(|v| {
                return v.get_float().map(|f| f / 4.0);
            })
            .collect();

        r.index.push(date);
        r.data.push(values);
    }

    Ok(r)
}

#[cfg(test)]
mod tests {
    use crate::meterpoint_value::run;
    use chrono::NaiveDate;

    #[test]
    fn test_parse_wiener_netze_is_successful() {
        let result = run("var/meterpoint_value_wiener_netze.xlsx".to_string());
        assert!(result.is_ok());

        let data = result.unwrap();

        assert_eq!(
            data.columns,
            vec![
                String::from("AT0010000000000000001000001234567"),
                String::from("AT0010000000000000001000001234568"),
            ]
        );

        let d = NaiveDate::from_ymd_opt(2021, 1, 1).unwrap();
        assert_eq!(
            data.index,
            vec![
                d.and_hms_opt(0, 15, 0).unwrap(),
                d.and_hms_opt(0, 30, 0).unwrap(),
                d.and_hms_opt(0, 45, 0).unwrap(),
                d.and_hms_opt(1, 00, 0).unwrap(),
            ]
        );

        assert_eq!(
            data.data,
            vec![
                vec![Some(1.32 / 4.0), Some(2.32 / 4.0)],
                vec![Some(0.51 / 4.0), Some(1.51 / 4.0)],
                vec![Some(0.36 / 4.0), Some(1.36 / 4.0)],
                vec![Some(0.39 / 4.0), Some(1.39 / 4.0)],
            ]
        )
    }
}
