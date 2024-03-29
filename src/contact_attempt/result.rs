use serde::Serialize;

#[derive(Serialize, Debug)]
pub(crate) enum Result_ {
    #[serde(rename = "appointment")]
    Appointment,
    #[serde(rename = "contact-again")]
    ContactAgain,
    #[serde(rename = "customer-acquired")]
    CustomerAcquired,
    #[serde(rename = "email-sent")]
    EmailSent,
    #[serde(rename = "gdpr-ban")]
    GdprBan,
    #[serde(rename = "reached-new-contact")]
    ReachedNewContact,
    #[serde(rename = "reached-no-interest")]
    ReachedNoInterest,
    #[serde(rename = "not-reached")]
    NotReached,
    #[serde(rename = "reached-interest")]
    ReachedInterest,
    #[serde(rename = "wrong")]
    Wrong,
    #[serde(rename = "abort-by-customer")]
    AbortByCustomer,
    #[serde(rename = "abort-by-sales")]
    AbortBySales,
    #[serde(rename = "create-new-appointment")]
    CreateNewAppointment,
    #[serde(rename = "data-request")]
    DataRequest,
    #[serde(rename = "data-send")]
    DataSend,
    #[serde(rename = "offer-created")]
    OfferCreated,
}
impl Result_ {
    pub(crate) fn from_excel_value(v: String) -> Result<Option<Result_>, String> {
        match v.to_lowercase().trim() {
            "termin vereinbart" => Ok(Some(Result_::Appointment)),
            "wieder kontaktieren" => Ok(Some(Result_::ContactAgain)),
            "kunde gewonnen" => Ok(Some(Result_::CustomerAcquired)),
            "email versendet" => Ok(Some(Result_::EmailSent)),
            "gesperrt (dsgvo)" => Ok(Some(Result_::GdprBan)),
            "neuer kontakt" => Ok(Some(Result_::ReachedNewContact)),
            "kein interesse" => Ok(Some(Result_::ReachedNoInterest)),
            "nicht erreicht" => Ok(Some(Result_::NotReached)),
            "erreicht - interesse" => Ok(Some(Result_::ReachedInterest)),
            "falscher kontakt" => Ok(Some(Result_::Wrong)),
            "abgebrochen durch kunde" => Ok(Some(Result_::AbortByCustomer)),
            "abgebrochen durch vertrieb" => Ok(Some(Result_::AbortBySales)),
            "neuen termin vereinbaren" => Ok(Some(Result_::CreateNewAppointment)),
            "daten angefragt" => Ok(Some(Result_::DataRequest)),
            "datenanforderung" => Ok(Some(Result_::DataRequest)),
            "daten übermittelt" => Ok(Some(Result_::DataSend)),
            "angebot erstellt" => Ok(Some(Result_::OfferCreated)),
            "" => Ok(None),
            &_ => Err(v),
        }
    }
}
