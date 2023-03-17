use serde::Serialize;

#[derive(Serialize, Debug)]
pub(crate) enum Status {
    #[serde(rename = "done")]
    Done,
    #[serde(rename = "todo")]
    Todo,
}
impl Status {
    pub(crate) fn from_excel_value(v: String) -> Result<Status, String> {
        match v.to_lowercase().trim() {
            "erledigt" => Ok(Status::Done),
            "zu erledigen" => Ok(Status::Todo),
            &_ => Err(v),
        }
    }
}
