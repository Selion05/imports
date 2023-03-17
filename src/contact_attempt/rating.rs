use serde_repr::*;

#[derive(Serialize_repr, Deserialize_repr, PartialEq, Debug)]
#[repr(u8)]
pub(crate) enum Rating {
    Bad = 1,
    RatherBad = 2,
    RatherGood = 3,
    Good = 4,
}

impl Rating {
    pub(crate) fn from_excel_value(v: String) -> Result<Option<Rating>, String> {
        match v.to_lowercase().trim() {
            "schlecht" => Ok(Some(Rating::Bad)),
            "eher schlecht" => Ok(Some(Rating::RatherBad)),
            "eher gut" => Ok(Some(Rating::RatherGood)),
            "gut" => Ok(Some(Rating::Good)),
            "" => Ok(None),
            &_ => Err(v),
        }
    }
}
