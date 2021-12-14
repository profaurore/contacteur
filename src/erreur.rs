#[derive(Debug)]
pub enum Erreur {
    Abandonné,
    Arbre(String),
    ÉtatInconnu(reqwest::StatusCode),
    IdentifiantsInvalides,
    Ods(spreadsheet_ods::OdsError),
    Requête(reqwest::Error),
    Sqlite(rusqlite::Error)
}

pub type Result<T, E = Erreur> = std::result::Result<T, E>;

impl From<rusqlite::Error> for Erreur {
    fn from(err: rusqlite::Error) -> Erreur {
        Erreur::Sqlite(err)
    }
}

impl From<reqwest::Error> for Erreur {
    fn from(err: reqwest::Error) -> Erreur {
        Erreur::Requête(err)
    }
}

impl From<spreadsheet_ods::OdsError> for Erreur {
    fn from(err: spreadsheet_ods::OdsError) -> Erreur {
        Erreur::Ods(err)
    }
}

impl std::fmt::Display for Erreur {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match &self {
            Erreur::Abandonné => f.write_str("Le travail a été abandonné."),
            Erreur::Arbre(e) => f.write_str(e),
            Erreur::ÉtatInconnu(état) => f.write_str(&format!("Le code d'état ({}) de la requête envoyé est inattendu.", état)),
            Erreur::IdentifiantsInvalides => f.write_str("Le nom d'utilisateur et le mot de passe sont invalides."),
            Erreur::Ods(e) => f.write_str(&format!("Une erreur s'est produite lors du traitement d'un fichier ODS: {}.", e)),
            Erreur::Requête(e) => f.write_str(&format!("Une erreur s'est produite lors de l'envoie de la requête: {}.", e)),
            Erreur::Sqlite(e) => f.write_str(&format!("Une erreur SQLite s'est produite: {}.", e))
        }
    }
}