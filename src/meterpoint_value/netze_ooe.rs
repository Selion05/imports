use crate::meterpoint_value::Data;
use crate::ImportError;
use calamine::{DataType, Range};
use chrono::SubsecRound;
use regex::Regex;

fn meterpoint_label(path: String) -> Result<String, String> {
    let p = std::path::Path::new(path.as_str());
    let filename: String = p
        .file_name()
        .ok_or_else(|| "path has no filename".to_string())?
        .to_str()
        .ok_or_else(|| "could not convert path to str".to_string())?
        .trim()
        .to_uppercase();
    println!("{:?}", filename);

    let re = Regex::new(r"^LASTPROFIL[-\s]?([A-Z]{2}[A-Z0-9]{31}).*\.XLSX$").unwrap();

    if let Some(captures) = re.captures(filename.as_str()) {
        if let Some(meterpoint) = captures.get(1) {
            return Ok(meterpoint.as_str().to_string());
        }
    }

    let meterpoint: String = filename
        .strip_prefix("LASTPROFIL")
        .ok_or_else(|| "filename has no prefix lastprofil".to_string())?
        .strip_suffix(".XLSX")
        .ok_or_else(|| "filename has no suffix xlsx".to_string())?
        .trim()
        .chars()
        .filter(|c| c.is_ascii_alphanumeric())
        .collect();
    if meterpoint.len() != 33 {
        return Err(format!("could not get meterpoint from filename '{path}'",));
    }

    return Ok(meterpoint);
}

pub fn run(sheet: Range<DataType>, path: String) -> Result<Data, ImportError> {
    let meterpoint = meterpoint_label(path).map_err(|e| ImportError::Error(e))?;

    println!("{:?}", meterpoint);
    let headers: Vec<String> = vec![meterpoint];

    let mut r = Data {
        columns: headers,
        index: vec![],
        data: vec![],
    };

    // there is an edge case where this does not work but with real data it should never appear
    // the schema has at lease 13 rows so if one file has less than ~11 rows of values it tries
    // to parse empty cells to date which then fails
    for (i, row) in sheet.rows().enumerate().skip(2) {
        let date = row[3].as_date().ok_or_else(|| {
            ImportError::ValueError(
                i,
                "Timestamp".to_string(),
                "could not parse date".to_string(),
            )
        })?;
        let time = row[4]
            .as_time()
            .ok_or_else(|| {
                ImportError::ValueError(
                    i,
                    "Timestamp".to_string(),
                    "could not parse time".to_string(),
                )
            })?
            .round_subsecs(0);
        let value = row[5].get_float();
        r.index.push(date.and_time(time));
        r.data.push(vec![value]);
    }

    Ok(r)
}

#[cfg(test)]
mod tests {
    use crate::meterpoint_value::netze_ooe::meterpoint_label;
    use crate::meterpoint_value::run;
    use chrono::NaiveDate;

    #[test]
    fn test_parse_is_successful() {
        let result = run("var/Lastprofil AT0030000000000000000000000001234.xlsx".to_string());

        println!("{:?}", result);
        assert!(result.is_ok());

        let data = result.unwrap();

        assert_eq!(
            data.columns,
            vec![String::from("AT0030000000000000000000000001234"),]
        );

        let d = NaiveDate::from_ymd_opt(2023, 1, 5).unwrap();
        assert_eq!(
            data.index,
            vec![
                d.and_hms_opt(0, 00, 0).unwrap(),
                d.and_hms_opt(0, 15, 0).unwrap(),
                d.and_hms_opt(0, 30, 0).unwrap(),
                d.and_hms_opt(0, 45, 0).unwrap(),
                d.and_hms_opt(1, 00, 0).unwrap(),
                d.and_hms_opt(1, 15, 0).unwrap(),
                d.and_hms_opt(1, 30, 0).unwrap(),
                d.and_hms_opt(1, 45, 0).unwrap(),
                d.and_hms_opt(2, 00, 0).unwrap(),
                d.and_hms_opt(2, 15, 0).unwrap(),
                d.and_hms_opt(2, 30, 0).unwrap(),
            ]
        );

        assert_eq!(
            data.data,
            vec![
                vec![Some(1.623)],
                vec![Some(1.482)],
                vec![Some(1.494)],
                vec![Some(1.557)],
                vec![Some(1.536)],
                vec![Some(1.494)],
                vec![Some(1.596)],
                vec![Some(1.557)],
                vec![Some(1.485)],
                vec![Some(1.818)],
                vec![Some(4.743)],
            ]
        )
    }
    #[test]
    fn test_meterpoint_extraction() {
        for (input, expected) in vec![
            (
                "test/lastprofil AT1234567891234567891234567891234.xlsx".to_string(),
                Ok("AT1234567891234567891234567891234".to_string()),
            ),
            (
                "test/lastprofil-at1234567891234567891234567891234.xlsx".to_string(),
                Ok("AT1234567891234567891234567891234".to_string()),
            ),
            (
                "lastprofil-at1234567891234567891234567891234.xlsx".to_string(),
                Ok("AT1234567891234567891234567891234".to_string()),
            ),
            (
                "test/AT1234567891234567891234567891234.xlsx".to_string(),
                Err("filename has no prefix lastprofil".to_string()),
            ),
            (
                "test/lastprofil AT1234567891234567891234567891234.csv".to_string(),
                Err("filename has no suffix xlsx".to_string()),
            ),
            (
                "test/lastprofil AT123456789123456789123456789123.xlsx".to_string(),
                Err("could not get meterpoint from filename 'test/lastprofil AT123456789123456789123456789123.xlsx'".to_string()),
            ),
            (
                "lastprofil-at0300000000000000000000001234567-2020.xlsx".to_string(),
                Ok("AT0300000000000000000000001234567".to_string()),
            ),
            (
                "lastprofil-at0030000000000000000000000012345-2019(1).xlsx".to_string(),
                Ok("AT0030000000000000000000000012345".to_string()),
            ),
            ("".to_string(), Err("path has no filename".to_string())),
        ] {
            let result = meterpoint_label(input);
            assert_eq!(result, expected);
        }
    }
}
