extern crate serde_json;

use convert_case::{Case, Casing};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::env;
use std::fs::File;
use tera::Context;
use tera::Tera;

#[derive(Serialize, Deserialize, Debug, Clone, PartialEq)]
enum Kind {
    #[serde(rename = "date")]
    Date,
    #[serde(rename = "time")]
    Time,
    #[serde(rename = "string")]
    String,
    #[serde(rename = "float")]
    Float,
    #[serde(rename = "enum")]
    Enum,
}

#[derive(Debug, Deserialize, Serialize)]
struct ColumnDefinition {
    #[serde(rename = "name")]
    header_name: String,
    #[serde(rename = "name2")]
    #[serde(default)]
    header_name2: String,
    #[serde(rename = "type")]
    kind: Kind,
    #[serde(default)]
    optional: bool,
    #[serde(default)]
    key: String,
    #[serde(rename = "enum")]
    enum_: Option<String>,
}

impl ColumnDefinition {
    fn type_hint(&self) -> String {
        let mut t = match self.kind {
            Kind::Date => {
                return if self.optional {
                    "Option<NaiveDate>".to_string()
                } else {
                    "NaiveDate".to_string()
                }
            }
            Kind::Time => "NaiveTime".to_string(),
            Kind::Float => "f64".to_string(),
            Kind::String => "String".to_string(),
            Kind::Enum => self
                .enum_
                .clone()
                .expect(&format!("could not get enum path for {}", self.key))
                .split("::")
                .last()
                .expect(&format!("could not get enum name for {}", self.key))
                .to_string(),
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
    kind: Kind,
    type_hint: String,
    enum_name: String,
    enum_: Option<String>,
    enum_path: Option<String>,
    enum_mod: Option<String>,
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

        let mut e: Option<String> = None;
        let mut enum_mod: Option<String> = None;
        if d.kind == Kind::Enum {
            e = Some(
                d.enum_
                    .clone()
                    .unwrap()
                    .split("::")
                    .last()
                    .unwrap()
                    .to_string(),
            );
            enum_mod = Some(
                d.enum_
                    .clone()
                    .unwrap()
                    .split("::")
                    .next()
                    .unwrap()
                    .to_string(),
            );
        };
        Column {
            field_name: field_name,
            header_name: d.header_name.clone(),
            type_hint: type_hint,
            kind: d.kind,
            enum_name: d.key.to_case(Case::Pascal),
            match_string: d.header_name.to_lowercase().trim().to_string(),
            match_string2: d.header_name2.to_lowercase().trim().to_string(),
            optional: d.optional,
            enum_: e,
            enum_path: d.enum_.clone(),
            enum_mod,
        }
    }
}

#[derive(Debug, Deserialize, Serialize)]
struct TemplateContext {
    modules: Vec<String>,
    uses: Vec<String>,
    columns: Vec<Column>,
    header_row_number: usize,
    data_start_row_number: usize,
    group_key: String,
    kind: Kind,
}

#[derive(Debug, Deserialize, Serialize)]
struct Definition {
    #[serde(rename = "groupKey")]
    group_key: String,
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

    let mut columns: Vec<Column> = Vec::new();
    let mut modules: Vec<String> = Vec::new();
    let mut uses: Vec<String> = Vec::new();

    for (key, column) in definition.columns.iter() {
        // todo stupido
        let c = ColumnDefinition {
            header_name: column.header_name.to_string(),
            header_name2: column.header_name2.to_string(),
            kind: column.kind.clone(),
            optional: column.optional,
            key: key.to_string(),
            enum_: column.enum_.clone(),
        };

        let column2 = Column::from(c);

        match column2.kind {
            Kind::Enum => {
                modules.push(column2.enum_mod.clone().unwrap());
                uses.push(column2.enum_path.clone().unwrap());
            }
            Kind::Date => uses.push("chrono::NaiveDate".to_string()),
            Kind::Time => uses.push("chrono::NaiveTime".to_string()),
            _ => {}
        }

        modules.sort();
        modules.dedup();
        uses.sort();
        uses.dedup();
        columns.push(column2)
    }
    columns.sort_by_key(|c| c.field_name.clone());

    let data = TemplateContext {
        columns,
        header_row_number: definition.header_row_number - 1,
        data_start_row_number: definition.data_start_row_number - 1,
        group_key: definition.group_key.to_case(Case::Snake).trim().to_string(),
        modules,
        uses,
        kind: Kind::String,
    };

    let tera = match Tera::new("templates/**/*.tera") {
        Ok(t) => t,
        Err(e) => {
            println!("Parsing error(s): {}", e);
            std::process::exit(1);
        }
    };

    let s = tera.render("importer.rs.tera", &Context::from_serialize(&data).unwrap());

    println!("{}", s.unwrap());
}
