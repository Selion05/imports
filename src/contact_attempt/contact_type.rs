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
    pub(crate) fn from_excel_value(v: String) -> Result<ContactType, String> {
        match v.to_lowercase().trim() {
            "email" => Ok(ContactType::Email),
            "persönlich" => Ok(ContactType::Personally),
            "telefon" => Ok(ContactType::Phone),
            &_ => Err(v),
        }
    }
}
