use reqwest::{
    StatusCode
};
use rusqlite::{params, Connection};
use spreadsheet_ods::{Value, read_ods, xmltree::XmlContent, WorkBook, Sheet, write_ods, CellStyle, style::FontFaceDecl, style::units::Length};
use regex::Regex;
use std::collections::HashMap;
use std::path::Path;
use chrono::Local;


#[derive(Debug)]
pub enum Erreur {
    Sqlite(rusqlite::Error),
    IdentifiantsInvalides,
    ÉtatInconnu(StatusCode),
    Requête(reqwest::Error),
    Csv(csv::Error),
    Ods(spreadsheet_ods::OdsError),
    Abandonné
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
            Erreur::Sqlite(e) => f.write_str(&format!("Une erreur SQLite s'est produite: {}.", e)),
            Erreur::IdentifiantsInvalides => f.write_str("Le nom d'utilisateur et le mot de passe sont invalides."),
            Erreur::ÉtatInconnu(état) => f.write_str(&format!("Le code d'état ({}) de la requête envoyé est inattendu.", état)),
            Erreur::Requête(e) => f.write_str(&format!("Une erreur s'est produite lors de l'envoie de la requête: {}.", e)),
            Erreur::Csv(e) => f.write_str(&format!("Une erreur s'est produite lors de l'envoie de la requête: {}.", e)),
            Erreur::Ods(e) => f.write_str(&format!("Une erreur s'est produite lors du traitement d'un fichier ODS: {}.", e)),
            Erreur::Abandonné => f.write_str("Le travail a été abandonné.")
        }
    }
}


fn ouvrir_base_de_données(fichier: Option<&str>) -> Result<Connection> {
    let conn = match fichier {
        Some(f) => Connection::open(Path::new(f))?,
        None => Connection::open_in_memory()?
    };

    conn.execute_batch("
        BEGIN TRANSACTION;
        CREATE TABLE IF NOT EXISTS cours (
            id INTEGER PRIMARY KEY,
            code TEXT NOT NULL,
            nom TEXT
        );
        CREATE TABLE IF NOT EXISTS étiquette (
            id INTEGER PRIMARY KEY,
            nom TEXT NOT NULL,

            CONSTRAINT u_nom UNIQUE (nom)
        );
        INSERT OR IGNORE INTO étiquette(nom) VALUES
            ('Virtuel'),
            ('AP');
        CREATE TABLE IF NOT EXISTS élève (
            id INTEGER PRIMARY KEY,
            prénom_préféré TEXT,
            prénom TEXT NOT NULL,
            nom TEXT NOT NULL,
            id_cours INTEGER NOT NULL,

            CONSTRAINT f_cours FOREIGN KEY (id_cours) REFERENCES cours(id),
            CONSTRAINT u_élève UNIQUE (prénom, nom, id_cours)
        );
        CREATE TABLE IF NOT EXISTS élève_étiquette (
            id_élève INTEGER NOT NULL,
            id_étiquette INTEGER NOT NULL,

            CONSTRAINT f_élève FOREIGN KEY (id_élève) REFERENCES élève(id),
            CONSTRAINT f_étiquette FOREIGN KEY (id_étiquette) REFERENCES étiquette(id)
        );
        CREATE TABLE IF NOT EXISTS élève_contact (
            id INTEGER PRIMARY KEY,
            id_élève INTEGER NOT NULL,
            nom_complet TEXT NOT NULL,
            relation TEXT,
            correspondance INTEGER NOT NULL,
            automatique INTEGER NOT NULL,
            ordre INTEGER,

            CONSTRAINT u_nom UNIQUE (id_élève, nom_complet),
            CONSTRAINT c_correspondance CHECK (correspondance = 0 OR correspondance = 1),
            CONSTRAINT c_automatique CHECK (automatique = 0 OR automatique = 1)
        );
        CREATE TABLE IF NOT EXISTS élève_contact_type (
            id INTEGER PRIMARY KEY,
            type TEXT NOT NULL,

            CONSTRAINT u_type UNIQUE (type)
        );
        INSERT OR IGNORE INTO élève_contact_type(type) VALUES
            ('Courriel'),
            ('Téléphone cellulaire'),
            ('Téléphone au domicile'),
            ('Téléphone au travail');
        CREATE TABLE IF NOT EXISTS élève_contact_item (
            id_contact INTEGER NOT NULL,
            id_type INTEGER NOT NULL,
            coordonnée TEXT NOT NULL,
            automatique INTEGER NOT NULL,

            CONSTRAINT f_contact FOREIGN KEY (id_contact) REFERENCES élève_contact(id),
            CONSTRAINT f_type FOREIGN KEY (id_type) REFERENCES élève_contact_type(id),
            CONSTRAINT u_coordonnée UNIQUE (id_contact, coordonnée, automatique),
            CONSTRAINT c_automatique CHECK (automatique = 0 OR automatique = 1)
        );
        CREATE TABLE IF NOT EXISTS échelle_niveaux (
            id_échelle INTEGER NOT NULL,
            nom TEXT NOT NULL,
            min REAL NOT NULL,
            max REAL NOT NULL,

            CONSTRAINT f_échelle FOREIGN KEY (id_échelle) REFERENCES échelle(id),
            CONSTRAINT c_étendu CHECK (min <= max)
        );
        CREATE TABLE IF NOT EXISTS échelle (
            id INTEGER PRIMARY KEY,
            nom TEXT NOT NULL,
            précision INTEGER NOT NULL,
            min REAL NOT NULL,
            max REAL NOT NULL,

            CONSTRAINT c_étendu CHECK (min <= max)
        );
        CREATE TABLE IF NOT EXISTS évaluation_item (
            id INTEGER PRIMARY KEY,
            nom TEXT NOT NULL,
            id_parent INTEGER,
            id_précédent INTEGER,
            id_échelle INTEGER,
            formule TEXT,

            CONSTRAINT u_position UNIQUE (id_parent, id_précédent),
            CONSTRAINT c_type CHECK (id_parent IS NULL OR id_échelle IS NULL),
            CONSTRAINT c_formule CHECK (id_échelle IS NOT NULL OR formule IS NULL),
            CONSTRAINT f_parent FOREIGN KEY (id_parent) REFERENCES évaluation_item(id),
            CONSTRAINT f_précédent FOREIGN KEY (id_précédent) REFERENCES évaluation_item(id),
            CONSTRAINT f_échelle FOREIGN KEY (id_échelle) REFERENCES échelle(id)
        );
        CREATE TABLE IF NOT EXISTS évaluation_reprise (
            id INTEGER PRIMARY KEY,
            exclus INTEGER NOT NULL DEFAULT 0,
            temps TEXTE NOT NULL,

            CONSTRAINT c_exclus CHECK (exclus = 0 OR exclus = 1)
        );
        CREATE TABLE IF NOT EXISTS évaluation_résultat (
            id_item INTEGER NOT NULL,
            id_reprise INTEGER,
            id_élève INTEGER NOT NULL,
            résultat REAL,
            résultat_auto REAL,

            CONSTRAINT u_item UNIQUE (id_item, id_reprise, id_élève),
            CONSTRAINT c_résultat CHECK (résultat IS NOT NULL OR résultat_auto IS NOT NULL),
            CONSTRAINT f_item FOREIGN KEY (id_item) REFERENCES évaluation_item(id),
            CONSTRAINT f_reprise FOREIGN KEY (id_reprise) REFERENCES évaluation_reprise(id),
            CONSTRAINT f_élève FOREIGN KEY (id_élève) REFERENCES élève(id)
        );
        COMMIT;
    "
    /*
    "
        CREATE TRIGGER u_échelle AFTER UPDATE OF précision, min, max ON échelle
            WHEN new.précision <> old.précision OR new.min <> old.min OR new.max <> old.max
            BEGIN
                UPDATE évaluation_résultat
                    SET résultat = ROUND(
                        CASE
                            WHEN old.min < old.max THEN
                                (évaluation_résultat.résultat - old.min) / (old.max - old.min) * (new.max - new.min) + new.min
                            ELSE
                                new.max
                        END, new.précision)
                    FROM évaluation_résultat JOIN évaluation_item ON évaluation_résultat.id_item = évaluation_item.id
                    WHERE évaluation_item.id_échelle = new.id AND résultat IS NOT NULL;
                SELECT évaluer_formule_échelle(new.id);
            END;

        CREATE TRIGGER i_résultats BEFORE INSERT ON évaluation_résultat
            WHEN new.résultat IS NOT NULL OR new.résultat_auto IS NOT NULL
            BEGIN
                SELECT
                    CASE 
                        WHEN new.résultat IS NOT NULL AND (new.résultat < échelle.min OR new.résultat > échelle.max) THEN
                            RAISE(ABORT, 'Résultat hors bornes.')
                        WHEN new.résultat_auto IS NOT NULL AND évaluation_item.formule IS NOT NULL THEN
                            RAISE(ABORT, 'Résultat automatique n\'\'a pas une formule.')
                        WHEN new.résultat_auto IS NOT NULL AND (new.résultat_auto < échelle.min OR new.résultat_auto > échelle.max) THEN
                            RAISE(ABORT, 'Résultat automatique hors bornes.')
                    END
                    FROM évaluation_item
                    LEFT JOIN échelle ON échelle.id = évaluation_item.id_échelle
                    WHERE évaluation_item.id = new.id_item
                    LIMIT 1;
            END;
        CREATE TRIGGER u_résultats BEFORE UPDATE ON évaluation_résultat
            WHEN new.résultat IS NOT NULL OR new.résultat_auto IS NOT NULL
            BEGIN
                SELECT 
                    CASE 
                        WHEN new.résultat IS NOT NULL AND (new.résultat < échelle.min OR new.résultat > échelle.max) THEN
                            RAISE(ABORT, 'Résultat hors bornes.')
                        WHEN new.résultat_auto IS NOT NULL AND évaluation_item.formule IS NOT NULL THEN
                            RAISE(ABORT, 'Résultat automatique n\'\'a pas une formule.')
                        WHEN new.résultat_auto IS NOT NULL AND (new.résultat_auto < échelle.min OR new.résultat_auto > échelle.max) THEN
                            RAISE(ABORT, 'Résultat automatique hors bornes.')
                    END
                    FROM évaluation_item
                    LEFT JOIN échelle ON échelle.id = évaluation_item.id_échelle
                    WHERE évaluation_item.id = new.id_item
                    LIMIT 1;
            END;

        CREATE TRIGGER u_item_formule AFTER UPDATE OF formule ON évaluation_item
            WHEN new.formule IS NOT old.formule AND new.formule IS NOT NULL
            BEGIN
                SELECT évaluer_formule_item(new.id);
            END;
        CREATE TRIGGER u_item_formule_vide AFTER UPDATE OF formule ON évaluation_item
            WHEN new.formule IS NOT old.formule AND new.formule IS NULL
            BEGIN
                UPDATE évaluation_résultat
                    SET résultat_auto = NULL
                    WHERE id_item = new.id;
            END;

        COMMIT;
    "*/
    )?;

    Ok(conn)
}


mod ilc {
    use rpassword::read_password;
    use std::io::Write;
    use super::Erreur;

    pub fn boucler_travail<F, X>(mut travail: F) -> Result<X, ()>
    where
        F: FnMut() -> Result<X, Erreur> {
        loop {
            match travail() {
                Result::Ok(résultat) => {
                    return Ok(résultat)
                },
                Result::Err(err) => {
                    println!("Erreur: {}", err);
                    let mut réponse = String::new();
                    loop {
                        print!("Essayer à nouveau (o ou n)? ");
                        std::io::stdout().flush().unwrap();
                        std::io::stdin().read_line(&mut réponse).unwrap();
                
                        let réponse = réponse.trim();
                        match réponse {
                            "o" => break,
                            "n" => return Err(()),
                            _ => ()
                        }
                    }
                }
            }
        }
    }

    pub fn obtenir_identifiants(service: &str) -> (String, String) {
        println!("***{}***", service);

        let mut utilisateur = String::new();
        loop {
            print!("Nom d'utilisateur: ");
            std::io::stdout().flush().unwrap();
            std::io::stdin().read_line(&mut utilisateur).unwrap();
            utilisateur = utilisateur.trim().to_string();
            if utilisateur != "" {
                break
            }
        }

        let mut mot_de_passe;
        loop {
            print!("Mot de passe: ");
            std::io::stdout().flush().unwrap();
            mot_de_passe = read_password().unwrap();
            if mot_de_passe != "" {
                break
            }
        }

        (utilisateur, mot_de_passe)
    }
}


use encompass::ClientEncompass;
mod encompass {
    use html_escape::decode_html_entities;
    use percent_encoding::percent_decode_str;
    use regex::Regex;
    use super::{ilc, Erreur, Result};
    use reqwest::{
        blocking::Client,
        redirect::Policy,
        StatusCode
    };
    use std::collections::HashMap;
    use chrono::naive::NaiveDate;

    pub struct ClientEncompass {
        client: Option<Client>
    }


    struct Groupe {
        id: i32,
        code: String
    }


    #[derive(Clone)]
    pub struct Cours {
        id_groupe: i32,
        pub code: String,
        pub élèves: Vec<Élève>
    }


    impl Cours {
        fn new<S: Into<String>>(id_groupe: i32, code: S) -> Self {
            Self {
                id_groupe: id_groupe,
                code: code.into(),
                élèves: Vec::new()
            }
        }
    }


    #[derive(Clone)]
    pub struct Élève {
        id: i32,
        pub prénom: String,
        pub nom: String,
        pub naissance: Option<NaiveDate>,
        pub contacts: Vec<Contact>
    }


    impl Élève {
        fn new<S: Into<String>>(id: i32, prénom: S, nom: S) -> Self {
            Self {
                id: id,
                prénom: prénom.into(),
                nom: nom.into(),
                naissance: None,
                contacts: Vec::new()
            }
        }
    }


    #[derive(Clone)]
    pub struct Contact {
        pub nom_complet: String,
        pub relation: Option<String>,
        pub tel_domicile: Option<String>,
        pub tel_travail: Option<String>,
        pub tel_cellulaire: Option<String>,
        pub courriel: Option<String>,
        pub correspondance: bool,
        pub ordre: Option<u32>
    }


    impl ClientEncompass {
        pub fn new() -> Self {
            Self {
                client: None
            }
        }

        fn init_client(&mut self) -> Result<&Client> {
            let connecter_encompass = |client: &mut Client, utilisateur: &str, mot_de_passe: &str| {
                let res = client
                    .post("https://french.compassforsuccess.ca/portal/auth/login.do")
                    .form(&[
                        ("username", &utilisateur),
                        ("password", &mot_de_passe)
                    ])
                    .send();

                match res {
                    Result::Ok(res) => {
                        match res.status() {
                            StatusCode::OK => Err(Erreur::IdentifiantsInvalides),
                            StatusCode::FOUND => Ok(()),
                            _ => Err(Erreur::ÉtatInconnu(res.status()))
                        }
                    },
                    Result::Err(err) => Err(Erreur::Requête(err))
                }
            };

            if self.client.is_some() {
                return Ok(self.client.as_ref().unwrap())
            }

            match Client::builder()
                .cookie_store(true)
                .redirect(Policy::none())
                .build() {
                Ok(mut client) => {
                    println!("Connexion...");
                    if ilc::boucler_travail(
                            || {
                                let (utilisateur, mot_de_passe) = ilc::obtenir_identifiants("EnCompass");
                                connecter_encompass(&mut client, &utilisateur, &mot_de_passe)
                            }).is_ok() {
                        println!("Connexion réussie!");
                        self.client = Some(client);
                        Ok(self.client.as_ref().unwrap())
                    } else {
                        Err(Erreur::Abandonné)
                    }
                },
                Err(e) => {
                    println!("Erreur: {}", e);
                    Err(Erreur::Requête(e))
                }
            }
        }

        fn obtenir_groupes(&mut self) -> Result<Vec<Groupe>> {
            let obtenir_cours_web = |client: &Client| {
                let res = client
                    .get("https://french.compassforsuccess.ca/portal/class/search.do?text=")
                    .send();

                let r_classes = Regex::new(r"classID=([^&]+).+className=([^&]+)").unwrap();
                match res {
                    Result::Ok(res) if res.status() == StatusCode::OK => {
                        let page = res.text().unwrap_or(String::new());
                        Ok(r_classes.captures_iter(&page).map(|c| Groupe {
                            id: c[1].parse().unwrap(),
                            code: percent_decode_str(&c[2]).decode_utf8_lossy().to_string()
                        }).collect())
                    },
                    Result::Ok(res) => Err(Erreur::ÉtatInconnu(res.status())),
                    Result::Err(err) => Err(Erreur::Requête(err))
                }
            };

            let client = self.init_client()?;

            println!("Obtention de la liste des groupes...");
            if let Ok(cours) = ilc::boucler_travail(|| obtenir_cours_web(client)) {
                println!("Obtention réussie!");
                Ok(cours)
            } else {
                Err(Erreur::Abandonné)
            }
        }

        fn obtenir_élèves(&mut self) -> Result<Vec<Cours>> {
            let obtenir_élèves_web = |client: &Client, id_groupe: i32| -> Result<Vec<Cours>> {
                let res = client
                    .get(format!("https://french.compassforsuccess.ca/portal/studentsuccess/studentSuccessMonitoringTable.do?classId={}", id_groupe))
                    .send();

                let r_élèves = Regex::new(r#"(?s:<tr +data-id="([0-9]+)".+?([^>]+)</a>.+?<td.+?>([^<>]+).+?(?:<td.+?){5}([0-9]*)</td>.+?([^>]+)</td>)"#).unwrap();
                match res {
                    Result::Ok(res) if res.status() == StatusCode::OK => {
                        let page = res.text().unwrap_or(String::new());

                        let mut élèves = HashMap::new();
                        r_élèves.captures_iter(&page).for_each(|c| {
                            let mut code = decode_html_entities(&c[5]).to_string();
                            if !c[4].is_empty() {
                                code.push_str("-");
                                code.push_str(&format!("{:0>2}", decode_html_entities(&c[4])));
                            }
                            
                            élèves
                                .entry(code.clone())
                                .or_insert_with(|| Cours::new(id_groupe, code))
                                .élèves
                                .push(Élève::new(
                                    c[1].parse().unwrap(),
                                    decode_html_entities(&c[2]),
                                    decode_html_entities(&c[3])
                                ))
                        });

                        Ok(élèves.values().cloned().collect())
                    },
                    Result::Ok(res) => Err(Erreur::ÉtatInconnu(res.status())),
                    Result::Err(err) => Err(Erreur::Requête(err))
                }
            };
            
            let groupes = self.obtenir_groupes()?;
            let client = self.init_client()?;

            let mut cours = Vec::new();
            for c in groupes {
                println!("Obtention des élèves pour {}...", c.code);
                if let Ok(élèves_cours) = ilc::boucler_travail(|| obtenir_élèves_web(client, c.id)) {
                    cours.extend(élèves_cours);
                } else {
                    return Err(Erreur::Abandonné)
                }
                println!("Obtention réussie!")
            }

            Ok(cours)
        }

        // 18+: https://french.compassforsuccess.ca/portal/student/130823/profile.do?referral=MHF4U-UQ.2
        pub fn obtenir_infos(&mut self) -> Result<Vec<Cours>> {
            let obtenir_infos_web = |client: &Client, id: i32| -> Result<(Option<NaiveDate>, Vec<Contact>)> {
                let res = client
                    .get(format!("https://french.compassforsuccess.ca/portal/gb/student/{}/gbInfo.do", id))
                    .send();

                // Incertain de "sept"
                let mois = vec!["janv", "févr", "mars", "avr", "mai", "juin", "juil", "août", "sept", "oct", "nov", "déc"];
                let r_date = Regex::new(r#"(?s:<th>Date de naissance</th>\s+<td>.+?>([0-9]+) ([^<>]+)\. ([0-9]+)</span>)"#).unwrap();
                let r_majeur = Regex::new(r#"<STRONG>([^<>]+)</STRONG>"#).unwrap();
                let r_contact = Regex::new(r#"(?s:<th>Nom.+?>([^<>]+)</span.+?([^<>]*)</span.+?Domicile.+?<td>([^<>]*).+?Travail.+?<td>([^<>]*).+?Cellulaire.+?<td>([^<>]*).+?Courriel.+?>([^<>]*)</a>.+?Correspondance.+?(green|red).+?Priorité de fermeture.+?>([0-9]+)</td>)"#).unwrap();
                match res {
                    Result::Ok(res) if res.status() == StatusCode::OK => {
                        let page = res.text().unwrap_or(String::new());

                        let naissance = r_date
                            .captures(&page)
                            .map(|c|
                                NaiveDate::from_ymd(
                                    c[3].parse().unwrap(),
                                    mois.iter().position(|&m| m == &c[2]).unwrap() as u32 + 1,
                                    c[1].parse().unwrap()
                                ));

                        let mut contacts: Vec<_> = r_contact.captures_iter(&page).map(|c| {
                            Contact {
                                nom_complet: decode_html_entities(&c[1]).into(),
                                relation: if !c[2].is_empty() && &c[2] != "Unknown" { Some(decode_html_entities(&c[2]).into()) } else { None },
                                tel_domicile: if !c[3].is_empty() { Some(decode_html_entities(&c[3]).into()) } else { None },
                                tel_travail: if !c[4].is_empty() { Some(decode_html_entities(&c[4]).into()) } else { None },
                                tel_cellulaire: if !c[5].is_empty() { Some(decode_html_entities(&c[5]).into()) } else { None },
                                courriel: if !c[6].is_empty() { Some(decode_html_entities(&c[6]).into()) } else { None },
                                correspondance: &c[7] == "green",
                                ordre: Some(c[8].parse().unwrap())
                            }
                        }).collect();

                        if page.contains("Student is 18") {
                            let contacts_majeur: Vec<_> = r_majeur.captures_iter(&page).filter_map(|c| {
                                if &c[1] == "NONE" {
                                    None
                                } else {
                                    Some(decode_html_entities(&c[1]).to_string())
                                }
                            }).collect();
                            contacts = contacts.iter().cloned().filter(|c| contacts_majeur.contains(&c.nom_complet)).collect();
                        }

                        Ok((naissance, contacts))
                    },
                    Result::Ok(res) => Err(Erreur::ÉtatInconnu(res.status())),
                    Result::Err(err) => Err(Erreur::Requête(err))
                }
            };

            let mut cours = self.obtenir_élèves()?;
            let client = self.init_client()?;

            for c in &mut cours {
                for élève in &mut c.élèves {
                    println!("Obtention des contacts pour {} {} {}...", c.code, élève.prénom, élève.nom);
                    if let Ok((naissance, contacts_élève)) = ilc::boucler_travail(|| obtenir_infos_web(client, élève.id)) {
                        élève.naissance = naissance;
                        élève.contacts = contacts_élève;
                    } else {
                        return Err(Erreur::Abandonné)
                    }
                    println!("Obtention réussie!")
                }
            }

            Ok(cours)
        }
    }
}


fn charge_élèves_encompass(conn: &mut Connection, encompass: &mut ClientEncompass) -> Result<()> {
    let insérer_contact = |id_contact, item, type_item| -> Result<()> {
        if let Some(val) = item {
            conn.execute("
                INSERT OR IGNORE INTO élève_contact_item(id_contact, id_type, coordonnée, automatique)
                    SELECT ?1, id, ?2, 1
                        FROM élève_contact_type WHERE type = ?3;
            ", params![id_contact, val, type_item])?;
        }

        Ok(())
    };

    if conn.query_row("SELECT COUNT(*) FROM élève;", [], |r| r.get(0)).unwrap_or(0) == 0 {
        // TODO: 1. delete contacts_item automatiques
        //     2. delete contacts automatiques qui n'ont aucun item
        for cours in encompass.obtenir_infos()? {
            conn.execute("INSERT OR IGNORE INTO cours(code) VALUES (?1)", [&cours.code])?;
            let id_cours: i64 = conn.query_row("SELECT id FROM cours WHERE code = ?1", [&cours.code], |r| r.get(0))?;

            for élève in cours.élèves {
                conn.execute("INSERT OR IGNORE INTO élève(prénom, nom, id_cours) VALUES (?1, ?2, ?3)", params![élève.prénom, élève.nom, id_cours])?;
                let id_élève: i64 = conn.query_row("SELECT id FROM élève WHERE prénom = ?1 AND nom = ?2 AND id_cours = ?3", params![élève.prénom, élève.nom, id_cours], |r| r.get(0))?;

                for contact in élève.contacts {
                    if contact.tel_domicile.is_none() && contact.tel_travail.is_none() && contact.tel_cellulaire.is_none() && contact.courriel.is_none() {
                        continue
                    }

                    conn.execute(
                        "INSERT OR IGNORE INTO élève_contact(id_élève, nom_complet, relation, correspondance, automatique, ordre) VALUES (?1, ?2, ?3, ?4, 1, ?5)",
                        params![id_élève, contact.nom_complet, contact.relation, contact.correspondance, contact.ordre]
                    )?;
                    let id_contact: i64 = conn.query_row("SELECT id FROM élève_contact WHERE id_élève = ?1 AND nom_complet = ?2", params![id_élève, contact.nom_complet], |r| r.get(0))?;

                    insérer_contact(id_contact, contact.tel_domicile, &"Téléphone au domicile")?;
                    insérer_contact(id_contact, contact.tel_travail, &"Téléphone au travail")?;
                    insérer_contact(id_contact, contact.tel_cellulaire, &"Téléphone cellulaire")?;
                    insérer_contact(id_contact, contact.courriel, &"Courriel")?;
                }
            }
        }
    }

    Ok(())
}


fn exporter_élèves_excel(conn: &Connection) -> Result<()> {
    let mut wb = WorkBook::new();

    let mut fonte = FontFaceDecl::new_with_name("Palatino Linotype");
    fonte.set_font_family("Palatino Linotype");
    fonte.set_font_family_generic("roman");
    wb.add_font(fonte);

    let mut défaut = CellStyle::empty();
    défaut.set_name("Défaut");
    défaut.set_font_name("Palatino Linotype");
    défaut.set_font_size(Length::Pt(12.));
    let défaut_ref = wb.add_cellstyle(défaut);

    let mut gras = CellStyle::empty();
    gras.set_name("Gras");
    gras.set_font_name("Palatino Linotype");
    gras.set_font_bold();
    gras.set_font_size(Length::Pt(12.));
    let gras_ref = wb.add_cellstyle(gras);

    let mut f_élèves = Sheet::new_with_name("Élèves");
    vec!["Cours", "Prénom", "Nom"]
        .iter()
        .enumerate()
        .for_each(|(i, titre)| f_élèves.set_styled_value(0, i as u32, *titre, &gras_ref));

    let mut f_tout = Sheet::new_with_name("Contacts");
    vec!["Cours", "Prénom", "Nom", "Contact", "Relation", "Priorité", "Courriel", "Domicile", "Travail", "Cellulaire"]
        .iter()
        .enumerate()
        .for_each(|(i, titre)| f_tout.set_styled_value(0, i as u32, *titre, &gras_ref));

    let mut f_courriels = Sheet::new_with_name("Courriels");
    vec!["Cours", "Prénom", "Nom", "Contact", "Relation", "Priorité", "Courriel"]
        .iter()
        .enumerate()
        .for_each(|(i, titre)| f_courriels.set_styled_value(0, i as u32, *titre, &gras_ref));

    let mut f_téléphones = Sheet::new_with_name("Téléphones");
    vec!["Cours", "Prénom", "Nom", "Contact", "Relation", "Priorité", "Domicile", "Travail", "Cellulaire"]
        .iter()
        .enumerate()
        .for_each(|(i, titre)| f_téléphones.set_styled_value(0, i as u32, *titre, &gras_ref));

    let mut stmt = conn.prepare("
        SELECT cours.code, élève.prénom, élève.nom
            FROM élève
            LEFT JOIN cours ON élève.id_cours = cours.id
            ORDER BY cours.code, élève.prénom, élève.nom;")?;
    let mut req = stmt.query([])?;
    let mut ligne = 1;
    while let Ok(Some(r)) = req.next() {
        for i in 0..3 {
            f_élèves.set_styled_value(ligne, i, r.get::<_, String>(i as usize).unwrap(), &défaut_ref);
        }
        ligne += 1;
    }

    let mut stmt = conn.prepare("
        SELECT cours.code, é.prénom, é.nom, COALESCE(c.nom_complet, ''), COALESCE(c.relation, ''), COALESCE(CAST(c.ordre as text), ''), COALESCE(i1.coordonnée, ''), COALESCE(i2.coordonnée, ''), COALESCE(i3.coordonnée, ''), COALESCE(i4.coordonnée, '')
            FROM élève as é
            LEFT JOIN cours ON cours.id = é.id_cours
            LEFT JOIN élève_contact AS c ON c.id_élève = é.id
            LEFT JOIN élève_contact_item AS i1 ON i1.id_contact = c.id AND i1.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Courriel')
            LEFT JOIN élève_contact_item AS i2 ON i2.id_contact = c.id AND i2.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Téléphone au domicile')
            LEFT JOIN élève_contact_item AS i3 ON i3.id_contact = c.id AND i3.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Téléphone au travail')
            LEFT JOIN élève_contact_item AS i4 ON i4.id_contact = c.id AND i4.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Téléphone cellulaire')
            WHERE c.correspondance = 1
            ORDER BY cours.code, é.prénom, é.nom, c.ordre, c.nom_complet;")?;
    let mut req = stmt.query([])?;
    let mut ligne = 1;
    while let Ok(Some(r)) = req.next() {
        for i in 0..10 {
            f_tout.set_styled_value(ligne, i, r.get::<_, String>(i as usize).unwrap(), &défaut_ref);
        }
        ligne += 1;
    }

    let mut stmt = conn.prepare("
        SELECT cours.code, é.prénom, é.nom, COALESCE(c.nom_complet, ''), COALESCE(c.relation, ''), COALESCE(CAST(c.ordre as text), ''), COALESCE(i.coordonnée, '')
            FROM élève as é
            LEFT JOIN cours ON cours.id = é.id_cours
            LEFT JOIN élève_contact AS c ON c.id_élève = é.id
            LEFT JOIN élève_contact_item AS i ON i.id_contact = c.id
            LEFT JOIN élève_contact_type AS t ON t.id = i.id_type
            WHERE c.correspondance = 1 AND t.type = 'Courriel'
            ORDER BY cours.code, é.prénom, é.nom, c.ordre, c.nom_complet;")?;
    let mut req = stmt.query([])?;
    let mut ligne = 1;
    while let Ok(Some(r)) = req.next() {
        for i in 0..7 {
            f_courriels.set_styled_value(ligne, i, r.get::<_, String>(i as usize).unwrap(), &défaut_ref);
        }
        ligne += 1;
    }

    let mut stmt = conn.prepare("
        SELECT cours.code, é.prénom, é.nom, COALESCE(c.nom_complet, ''), COALESCE(c.relation, ''), COALESCE(CAST(c.ordre as text), ''), COALESCE(i1.coordonnée, ''), COALESCE(i2.coordonnée, ''), COALESCE(i3.coordonnée, '')
            FROM élève as é
            LEFT JOIN cours ON cours.id = é.id_cours
            LEFT JOIN élève_contact AS c ON c.id_élève = é.id
            LEFT JOIN élève_contact_item AS i1 ON i1.id_contact = c.id AND i1.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Téléphone au domicile')
            LEFT JOIN élève_contact_item AS i2 ON i2.id_contact = c.id AND i2.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Téléphone au travail')
            LEFT JOIN élève_contact_item AS i3 ON i3.id_contact = c.id AND i3.id_type = (SELECT id FROM élève_contact_type WHERE type = 'Téléphone cellulaire')
            WHERE c.correspondance = 1
            ORDER BY cours.code, é.prénom, é.nom, c.ordre, c.nom_complet;")?;
    let mut req = stmt.query([])?;
    let mut ligne = 1;
    while let Ok(Some(r)) = req.next() {
        for i in 0..9 {
            f_téléphones.set_styled_value(ligne, i, r.get::<_, String>(i as usize).unwrap(), &défaut_ref);
        }
        ligne += 1;
    }

    wb.push_sheet(f_élèves);
    wb.push_sheet(f_tout);
    wb.push_sheet(f_courriels);
    wb.push_sheet(f_téléphones);

    let date = Local::now().format("%Y-%m-%d_%H-%M-%S");
    write_ods(&mut wb, format!["élèves_{}.ods", date])?;

    Ok(())
}


fn main() -> Result<()> {
    let mut encompass = ClientEncompass::new();
    let mut conn = ouvrir_base_de_données(None)?; // "contacteur.db3"

    charge_élèves_encompass(&mut conn, &mut encompass)?;

    println!("Exportation des données à un fichier...");
    exporter_élèves_excel(&conn)?;
    println!("Exportation réussie!");

    return Ok(());

    // Contacts + Date de naissance:
    // https://french.compassforsuccess.ca/portal/gb/student/128633/gbInfo.do

    fn simplifier_nom(nom: &str) -> String {
        nom.replace("ⱽ", "").replace("ᴬᴾ", "")
    }

    let re_code_cours = Regex::new(r"[A-Z]{3}[1-4][A-Z][0-9]?").unwrap();
    let re_nom_élève = Regex::new(r"([^,]+), ([^,]+)").unwrap();
    let doc = read_ods(std::path::Path::new(r"évaluations.ods")).unwrap();
    let mut sections_cours = HashMap::new();
    for i in 0..doc.num_sheets() {
        let feuille = doc.sheet(i);
        let feuille_nom = feuille.name();
        if re_code_cours.is_match(feuille_nom) {
            conn.execute("
                INSERT INTO cours(code) VALUES (?1);
            ", params![feuille_nom]).unwrap();
            let id_cours = conn.last_insert_rowid();

            let (n_lignes, n_colonnes) = feuille.used_grid_size();

            // Batir la liste des élèves
            let mut idx_élèves = Vec::new();
            for ligne in 3..n_lignes {
                let prénom_préféré = cellule_str(feuille.value(ligne, 0));
                if prénom_préféré.is_empty() {
                    continue
                }
                let est_virtuel = prénom_préféré.contains("ⱽ");
                let est_ap = prénom_préféré.contains("ᴬᴾ");
                let prénom_préféré = simplifier_nom(&prénom_préféré);

                let nom_élève = match re_nom_élève.captures(feuille.value(ligne, 1).as_str_or("").trim()) {
                    Some(n) => n,
                    None => continue
                };

                let nom = &nom_élève[1];
                let prénom = &nom_élève[2];
                
                conn.execute("
                    INSERT INTO élève(prénom_préféré, prénom, nom, id_cours) VALUES (?1, ?2, ?3, ?4);
                ", params![prénom_préféré, prénom, nom, id_cours]).unwrap();
                let id_élève = conn.last_insert_rowid();
                idx_élèves.push((ligne, id_élève));

                if est_virtuel {
                    conn.execute("
                        INSERT INTO élève_étiquette(id_élève, id_étiquette)
                            SELECT ?1, étiquette.id
                            FROM étiquette
                            WHERE étiquette.id = 'Virtuel';
                    ", params![id_élève]).unwrap();
                }

                if est_ap {
                    conn.execute("
                        INSERT INTO élève_étiquette(id_élève, id_étiquette)
                            SELECT ?1, étiquette.id
                            FROM étiquette
                            WHERE étiquette.id = 'AP';
                    ", params![id_élève]).unwrap();
                }
            }

            // Batir la liste des évaluations
            let mut id_évaluation = None;
            let mut idx_évaluation = -1;
            let mut id_section = None;
            let mut section = "".to_string();
            let mut id_composant = None;
            for colonne in 2..n_colonnes {
                let nom_évaluation = feuille.value(0, colonne).as_str_or("").trim();
                if !nom_évaluation.is_empty() {
                    conn.execute("
                        INSERT INTO évaluation_item(nom) VALUES (?1);
                    ", params![nom_évaluation]).unwrap();
                    idx_évaluation += 1;
                    id_évaluation = Some(conn.last_insert_rowid());
                    id_section = None;
                    id_composant = None;
                }

                let nom_section = cellule_str(feuille.value(1, colonne));
                if id_évaluation.is_some() && !nom_section.is_empty() {
                    conn.execute("
                        INSERT INTO évaluation_item(nom, id_parent, id_précédent)
                            VALUES (?1, ?2, ?3);
                    ", params![nom_section, if id_section.is_some() { None } else { id_évaluation }, id_section]).unwrap();
                    section = nom_section;
                    id_section = Some(conn.last_insert_rowid());
                    id_composant = None;
                }

                let nom_composant = feuille.value(2, colonne).as_str_or("").trim();
                if id_section.is_some() && !nom_composant.is_empty() {
                    conn.execute("
                        INSERT INTO évaluation_item(nom, id_parent, id_précédent)
                            VALUES (?1, ?2, ?3);
                    ", params![nom_composant, if id_composant.is_some() { None } else { id_section }, id_composant]).unwrap();
                    id_composant = Some(conn.last_insert_rowid());
                    
                    sections_cours
                        .entry(feuille_nom)
                        .or_insert_with(|| HashMap::new())
                        .entry(idx_évaluation)
                        .or_insert_with(|| HashMap::new())
                        .entry(section.clone())
                        .or_insert_with(|| Vec::new())
                        .push(colonne);
                }

                if id_composant.is_some() {
                    for (ligne, id_élève) in &idx_élèves {
                        if let Value::Number(n) = feuille.value(*ligne, colonne) {
                            conn.execute("
                                INSERT INTO évaluation_résultat(id_item, id_élève, résultat)
                                    VALUES (?1, ?2, ?3);
                            ", params![id_composant, id_élève, n]).unwrap();
                        }
                    }
                }
            }

            //for ligne in 0..n_lignes {
            //    for colonne in 0..n_colonnes {
            //    }
            //}
        }
    }

    fn xml_str(v: &Vec<XmlContent>) -> String {
        v.iter()
            .map(|x| match x {
                XmlContent::Text(t) => t.clone(),
                XmlContent::Tag(t) => xml_str(t.content())
            })
            .reduce(|a, b| a + " " + &b)
            .unwrap()
    }

    fn cellule_str(v: &Value) -> String {
        match v {
            Value::Text(t) => t.trim().to_string(),
            Value::TextXml(t) => xml_str(&t.iter().map(|x| XmlContent::Tag(x.clone())).collect()),
            Value::Boolean(b) => if *b { "v" } else { "f" }.to_string(),
            Value::Number(n) => n.to_string(),
            Value::Percentage(n) => (n * 100.).to_string() + "%",
            Value::Currency(n, devise) => n.to_string() + &std::str::from_utf8(devise).unwrap(),
            Value::DateTime(t) => t.format("%Y-%m-%d %H:%M:%S").to_string(),
            Value::TimeDuration(d) => d.to_std().unwrap().as_secs_f32().to_string(),
            _ => "".to_string()
        }
    }

    let feuille = doc.sheet(doc.sheet_idx("Reprises").unwrap());
    for ligne in 1..feuille.used_rows() {
        let temps = match feuille.value(ligne, 0) {
            Value::DateTime(t) => t.format("%Y-%m-%d %H:%M:%S").to_string(),
            _ => continue
        };

        let cours = cellule_str(feuille.value(ligne, 1));
        if cours.is_empty() {
            continue
        }

        let nom_élève = simplifier_nom(&cellule_str(feuille.value(ligne, 3)));

        let évaluation = match feuille.value(ligne, 4) {
            Value::Number(n) => *n as i32,
            _ => continue
        };

        let section = cellule_str(feuille.value(ligne, 5));
        if section.is_empty() {
            continue
        }

        let mut notes_orig = Vec::new();
        for i in 0..4 {
            match feuille.value(ligne, 6 + i) {
                Value::Number(n) => notes_orig.push(n),
                _ => break
            };
        }

        if let Value::Empty = feuille.value(ligne + 1, 0) {
            for i in 0..4 {
                match feuille.value(ligne + 1, 6 + i) {
                    Value::Number(n) => notes_orig.push(n),
                    _ => break
                };
            }
        }

        conn.execute("
            INSERT INTO évaluation_reprise(temps, exclus)
                VALUES (?1, ?2);
        ", params![temps, if nom_élève.ends_with(" (x)") { 1 } else { 0 }]).unwrap();
        let id_reprise = conn.last_insert_rowid();

        let ids_items = sections_cours
            .get(&cours)
            .unwrap()
            .get(&(évaluation - 1))
            .unwrap()
            .get(&*section)
            .unwrap();
        for i in 0..notes_orig.len() {
            conn.execute("
                INSERT INTO évaluation_résultat(id_item, id_reprise, id_élève, résultat)
                    SELECT ?1, ?2, id, ?3
                    FROM élève
                    WHERE prénom_préféré = ?4;
            ", params![ids_items[i], id_reprise, notes_orig[i], nom_élève.replace(" (x)", "")]).unwrap();
        }
    }

    let feuille = doc.sheet(doc.sheet_idx("Courriels").unwrap());
    for ligne in 1..feuille.used_rows() {
        let cours = cellule_str(feuille.value(ligne, 0));
        if cours.is_empty() {
            continue
        }

        let nom_élève = simplifier_nom(&cellule_str(feuille.value(ligne, 1)));
        if nom_élève.is_empty() {
            continue
        }

        let nom_parent = cellule_str(feuille.value(ligne, 2));
        if nom_parent.is_empty() {
            continue
        }

        let courriel = cellule_str(feuille.value(ligne, 3));
        if courriel.is_empty() {
            continue
        }

        conn.execute("
            INSERT INTO élève_contact(id_élève, nom, automatique)
                SELECT élève.id, ?1, 1
                FROM élève
                WHERE prénom_préféré = ?2;
        ", params![nom_parent, nom_élève]).unwrap();
        let id_contact = conn.last_insert_rowid();

        conn.execute("
            INSERT INTO élève_contact_item(id_contact, id_type, coordonnée, automatique)
                SELECT ?1, élève_contact_type.id, ?2, 1
                FROM élève_contact_type
                WHERE élève_contact_type.type = 'Courriel';
        ", params![id_contact, courriel]).unwrap();
    }






    let test = true;
    if !test {
        #[derive(Debug)]
        struct Cours {
            id: i64,
            code: String,
            nom: Option<String>
        }

        let mut stmt = conn.prepare("SELECT id, code, nom FROM cours").unwrap();
        let cours_iter = stmt.query_map([], |row| {
            Ok(Cours {
                id: row.get(0).unwrap(),
                code: row.get(1).unwrap(),
                nom: row.get(2).unwrap(),
            })
        }).unwrap();
        for cours in cours_iter {
            println!("{:?}", cours);
        }
    }

    if !test {
        #[derive(Debug)]
        struct Élève {
            id: i64,
            prénom_préféré: String,
            prénom: String,
            nom: String,
            cours: i64
        }

        let mut stmt = conn.prepare("SELECT id, prénom_préféré, prénom, nom, id_cours FROM élève").unwrap();
        let élève_iter = stmt.query_map([], |row| {
            Ok(Élève {
                id: row.get(0).unwrap(),
                prénom_préféré: row.get(1).unwrap(),
                prénom: row.get(2).unwrap(),
                nom: row.get(3).unwrap(),
                cours: row.get(4).unwrap(),
            })
        }).unwrap();
        for élève in élève_iter {
            println!("{:?}", élève);
        }
    }

    if !test {
        #[derive(Debug)]
        struct ÉvaluationItem {
            id: i64,
            nom: String,
            id_parent: Option<i64>,
            id_précédent: Option<i64>
        }

        let mut stmt = conn.prepare("SELECT id, nom, id_parent, id_précédent FROM évaluation_item").unwrap();
        let item_iter = stmt.query_map([], |row| {
            Ok(ÉvaluationItem {
                id: row.get(0).unwrap(),
                nom: row.get(1).unwrap(),
                id_parent: row.get(2).unwrap(),
                id_précédent: row.get(3).unwrap(),
            })
        }).unwrap();
        for item in item_iter {
            println!("{:?}", item);
        }
    }

    if !test {
        #[derive(Debug)]
        struct ÉvaluationRésultat {
            id_item: i64,
            id_reprise: Option<i64>,
            id_élève: i64,
            résultat: Option<f64>
        }

        let mut stmt = conn.prepare("
            SELECT COUNT(*), id_reprise, id_élève
            FROM évaluation_résultat
            LEFT JOIN évaluation_reprise ON évaluation_reprise.id = évaluation_résultat.id_reprise
            WHERE id_reprise IS NOT NULL AND évaluation_reprise.exclus = 0
            GROUP BY id_reprise, id_élève
        ").unwrap();
        let résultat_iter = stmt.query_map([], |row| {
            Ok(ÉvaluationRésultat {
                id_item: row.get(0).unwrap(),
                id_reprise: row.get(1).unwrap(),
                id_élève: row.get(2).unwrap(),
                résultat: None
            })
        }).unwrap();
        for résultat in résultat_iter {
            println!("{:?}", résultat);
        }
    }

    Ok(())
}
