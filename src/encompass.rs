use std::collections::HashMap;
use chrono::naive::NaiveDate;
use html_escape::decode_html_entities;
use percent_encoding::percent_decode_str;
use regex::Regex;
use reqwest::{blocking::Client, redirect::Policy, StatusCode};
use crate::erreur::{Erreur, Result};

pub struct ClientEncompass {
    client: Client
}

pub struct Groupe {
    id: i32,
    pub code: String
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
    pub fn new(utilisateur: &str, mot_de_passe: &str) -> Result<Self> {
        let client = match Client::builder()
            .cookie_store(true)
            .redirect(Policy::none())
            .build() {
            Ok(client) => client,
            Err(e) => return Err(Erreur::Requête(e))
        };

        let res = client
            .post("https://french.compassforsuccess.ca/portal/auth/login.do")
            .form(&[
                ("username", &utilisateur),
                ("password", &mot_de_passe)
            ])
            .send()?;

        match res.status() {
            StatusCode::FOUND => Ok(Self { client: client }),
            StatusCode::OK => Err(Erreur::IdentifiantsInvalides),
            _ => Err(Erreur::ÉtatInconnu(res.status()))
        }
    }

    pub fn obtenir_groupes(&mut self) -> Result<Vec<Groupe>> {
        let res = self.client
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
    }

    pub fn obtenir_élèves_groupe(&mut self, groupe: &Groupe) -> Result<Vec<Cours>> {
        let res = self.client
            .get(format!("https://french.compassforsuccess.ca/portal/studentsuccess/studentSuccessMonitoringTable.do?classId={}", groupe.id))
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
                        .or_insert_with(|| Cours::new(groupe.id, code))
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
    }

    pub fn obtenir_données_élève(&self, élève: &Élève) -> Result<(Option<NaiveDate>, Vec<Contact>)> {
        let res = self.client
            .get(format!("https://french.compassforsuccess.ca/portal/gb/student/{}/gbInfo.do", élève.id))
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
    }
}