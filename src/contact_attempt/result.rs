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
    pub(crate) fn from_excel_value(v: String) -> Result<Result_, String> {
        match v.to_lowercase().trim() {
            "termin vereinbart" => Ok(Result_::Appointment),
            "wieder kontaktieren" => Ok(Result_::ContactAgain),
            "kunde gewonnen" => Ok(Result_::CustomerAcquired),
            "email versendet" => Ok(Result_::EmailSent),
            "gesperrt (dsgvo)" => Ok(Result_::GdprBan),
            "neuer kontakt" => Ok(Result_::ReachedNewContact),
            "kein interesse" => Ok(Result_::ReachedNoInterest),
            "nicht erreicht" => Ok(Result_::NotReached),
            "erreicht - interesse" => Ok(Result_::ReachedInterest),
            "falscher kontakt" => Ok(Result_::Wrong),
            "abgebrochen durch kunde" => Ok(Result_::AbortByCustomer),
            "abgebrochen durch vertrieb" => Ok(Result_::AbortBySales),
            "neuen termin vereinbaren" => Ok(Result_::CreateNewAppointment),
            "daten angefragt" => Ok(Result_::DataRequest),
            "datenanforderung" => Ok(Result_::DataRequest),
            "daten übermittelt" => Ok(Result_::DataSend),
            "angebot erstellt" => Ok(Result_::OfferCreated),
            &_ => Err(v),
        }
    }
}
