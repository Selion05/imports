use crate::meterpoint_value::Data;
use crate::ImportError;
use calamine::{DataType, Range};
use chrono::{NaiveDateTime, SubsecRound};

pub fn run(sheet: Range<DataType>) -> Result<Data, ImportError> {
    let headers: Vec<String> = sheet
        .rows()
        .nth(0)
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

    let header_cnt = headers.len();
    let mut r = Data {
        columns: headers,
        index: vec![],
        data: Vec::with_capacity(header_cnt),
    };

    // there is an edge case where this does not work but with real data it should never appear
    // the schema has at lease 13 rows so if one file has less than ~11 rows of values it tries
    // to parse empty cells to date which then fails
    for (i, row) in sheet.rows().enumerate().skip(1) {
        let mut date = row[0].as_datetime();

        if None == date {
            date =
                NaiveDateTime::parse_from_str(row[0].to_string().as_str(), "%d.%m.%Y %H:%M").ok();
        }
        let date = date
            .ok_or_else(|| {
                ImportError::ValueError(
                    i,
                    "Timestamp".to_string(),
                    // "could not parse datetime".to_string(),
                    format!("Could not parse datetime {}", row[0].to_string()),
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

    Ok(r)
}

#[cfg(test)]
mod tests {
    use crate::meterpoint_value::run;
    use chrono::{NaiveDate, NaiveDateTime};

    #[test]
    fn test_parse_is_successful() {
        let result = run("var/meterpoint_value_netze_noe.xlsx".to_string());

        if !result.is_ok() {
            println!("{:?}", result);
        }
        assert!(result.is_ok());

        let data = result.unwrap();

        assert_eq!(
            data.columns,
            vec![
                String::from("AT0020000000000000000000100151234"),
                String::from("AT0020000000000000000000100251234"),
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
                d.and_hms_opt(1, 15, 0).unwrap(),
                d.and_hms_opt(1, 30, 0).unwrap(),
                d.and_hms_opt(1, 45, 0).unwrap(),
                d.and_hms_opt(2, 00, 0).unwrap(),
            ]
        );

        assert_eq!(
            data.data,
            vec![
                vec![Some(6.870), Some(3.180)],
                vec![Some(6.270), Some(3.240)],
                vec![Some(5.670), Some(3.660)],
                vec![Some(5.520), Some(3.060)],
                vec![Some(4.320), Some(4.260)],
                vec![Some(4.770), Some(4.860)],
                vec![Some(4.770), Some(3.720)],
                vec![Some(4.230), Some(3.360)],
            ]
        );
    }
}
