use serde::Serialize;

#[derive(Serialize, Debug)]
pub(crate) enum Status {
    #[serde(rename = "done")]
    Done,
    #[serde(rename = "todo")]
    Todo,
}
impl Status {
    pub(crate) fn from_excel_value(v: String) -> Result<Option<Status>, String> {
        match v.to_lowercase().trim() {
            "erledigt" => Ok(Some(Status::Done)),
            "zu erledigen" => Ok(Some(Status::Todo)),
            "" => Ok(None),
            &_ => Err(v),
        }
    }
}
