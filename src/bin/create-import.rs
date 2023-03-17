extern crate serde_json;

use convert_case::{Case, Casing};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::env;
use std::fs::File;
use tera::Context;
use tera::Tera;

#[derive(Debug, Deserialize, Serialize)]
struct ColumnDefinition {
    #[serde(rename = "name")]
    header_name: String,
    #[serde(rename = "name2")]
    #[serde(default)]
    header_name2: String,
    #[serde(rename = "type")]
    kind: String,
    #[serde(default)]
    optional: bool,
    #[serde(default)]
    key: String,
}

impl ColumnDefinition {
    fn type_hint(&self) -> String {
        let mut t = match self.kind.as_str() {
            "date" => "NaiveDate".to_string(),
            "float" => "f64".to_string(),
            "string" => "String".to_string(),
            _ => {
                panic!("unknown type {}", self.kind)
            }
        };

        if self.optional {
            t = format!("Option<{}>", t);
        }
        return t;
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct Column {
    field_name: String,
    header_name: String,
    kind: String,
    type_hint: String,
    enum_name: String,
    match_string: String,
    match_string2: String,
    optional: bool,
}

impl From<ColumnDefinition> for Column {
    fn from(d: ColumnDefinition) -> Self {
        let type_hint = d.type_hint();
        let mut field_name = d.key.to_case(Case::Snake);
        // todo quick fix...
        if field_name == "type".to_string() {
            field_name = "_type".to_string();
        }
        Column {
            field_name: field_name,
            header_name: d.header_name.clone(),
            type_hint: type_hint,
            kind: d.kind,
            enum_name: d.key.to_case(Case::Pascal),
            match_string: d.header_name.to_lowercase().trim().to_string(),
            match_string2: d.header_name2.to_lowercase().trim().to_string(),
            optional: d.optional,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct TemplateContext {
    columns: Vec<Column>,
}

#[derive(Debug, Deserialize, Serialize)]
struct Definition {
    #[serde(default)]
    #[serde(rename = "dataStartRowNumber")]
    data_start_row_number: usize,
    #[serde(default)]
    #[serde(rename = "headerRowNumber")]
    header_row_number: usize,
    columns: HashMap<String, ColumnDefinition>,
}

fn main() {
    let path = env::args().nth(1).unwrap();
    // let path = "var/datentraeger.columns.json";
    let file = File::open(path).unwrap();
    let definition: Definition = serde_json::from_reader(file).unwrap();

    let mut data = TemplateContext { columns: vec![] };

    for (key, column) in definition.columns.iter() {
        // todo stupido
        let c = ColumnDefinition {
            header_name: column.header_name.to_string(),
            header_name2: column.header_name2.to_string(),
            kind: column.kind.to_string(),
            optional: column.optional,
            key: key.to_string(),
        };

        data.columns.push(Column::from(c))
    }

    let tera = match Tera::new("templates/**/*.tera") {
        Ok(t) => t,
        Err(e) => {
            println!("Parsing error(s): {}", e);
            std::process::exit(1);
        }
    };

    let s = tera.render("importer.rs.tera", &Context::from_serialize(&data).unwrap());

    println!("{}", s.unwrap())
}
