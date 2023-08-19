use serde::Serialize;

#[derive(Serialize, Debug)]
pub(crate) enum ContactType {
    #[serde(rename = "email")]
    Email,
    #[serde(rename = "personally")]
    Personally,
    #[serde(rename = "phone")]
    Phone,
}
impl ContactType {
    pub(crate) fn from_excel_value(v: String) -> Result<Option<ContactType>, String> {
        match v.to_lowercase().trim() {
            "email" => Ok(Some(ContactType::Email)),
            "persönlich" => Ok(Some(ContactType::Personally)),
            "telefon" => Ok(Some(ContactType::Phone)),
            "" => Ok(None),
            &_ => Err(v),
        }
    }
}
