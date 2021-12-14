use std::collections::HashMap;
use regex::Regex;
use spreadsheet_ods::{CellStyle, read_ods, Sheet, style::{FontFaceDecl, units::Length}, Value, WorkBook, write_ods, xmltree::XmlContent};
use crate::erreur::Result;

const DÉCALAGE_NOTES: u32 = 5;

pub struct ClientClasseur {
    ods: WorkBook
}

#[derive(Clone)]
pub struct Cours {
    idx: usize,
    pub code: String,
    pub évaluations: Vec<Évaluation>,
    pub élèves: Vec<Élève>
}

#[derive(Clone)]
pub struct Élève {
    idx: u32,
    pub prénom: String,
    pub nom: String,
    pub prénom_préféré: String,
    pub étiquettes: Vec<String>,
    pub notes: Vec<Option<f64>>
}

impl Élève {
    pub fn note(&self, composant: &ComposantÉvaluation) -> Option<f64> {
        self.notes[composant.idx as usize]
    }
}

#[derive(Clone)]
pub struct ComposantÉvaluation {
    idx: u32,
    colonne: u32,
    pub nom: String,
}

#[derive(Clone)]
pub struct SectionÉvaluation {
    idx: u32,
    pub nom: String,
    pub composants: Vec<ComposantÉvaluation>
}

#[derive(Clone)]
pub struct Évaluation {
    idx: u32,
    pub nom: String,
    pub sections: Vec<SectionÉvaluation>
}

impl ClientClasseur {
    pub fn new(fichier: &str) -> Result<ClientClasseur> {
        let ods = read_ods(fichier)?;

        Ok(Self { ods: ods })
    }

    fn obtenir_cours(&self) -> Result<Vec<Cours>> {
        let re_code_cours = Regex::new(r"[A-Z]{3}[1-4][A-Z][0-9]?").unwrap();
        Ok((0..self.ods.num_sheets()).filter_map(|idx| {
            let nom = self.ods.sheet(idx).name();
            re_code_cours.is_match(nom).then(|| Cours {
                idx: idx,
                code: nom.into(),
                évaluations: Vec::new(),
                élèves: Vec::new()
            })
        }).collect())
    }

    fn obtenir_élèves(&self, cours: &Cours) -> Result<Vec<Cours>> {
        let feuille = self.ods.sheet(cours.idx);
        let n_lignes = feuille.used_rows();

        let mut cours_nouv = HashMap::new();
        for ligne in 3..n_lignes {
            let prénom_préféré = cellule_str(feuille.value(ligne, 0));
            if prénom_préféré.is_empty() {
                continue
            }

            let étiquettes = cellule_str(feuille.value(ligne, 1));
            let est_virtuel = étiquettes.contains("V");
            let est_ap = étiquettes.contains("AP");

            let nom = cellule_str(feuille.value(ligne, 2));
            let prénom = cellule_str(feuille.value(ligne, 3));
            let nom_cours = cellule_str(feuille.value(ligne, 4));

            let cours_élève = cours_nouv.entry(nom_cours.clone()).or_insert_with(|| cours.clone());
            cours_élève.code = nom_cours;
            cours_élève.élèves.push(Élève {
                idx: ligne,
                prénom: prénom,
                nom: nom,
                prénom_préféré: prénom_préféré,
                étiquettes: vec![est_virtuel.then(|| "Virtuel"), est_ap.then(|| "AP")]
                    .iter()
                    .filter_map(|x| x.map(|x| x.to_string()))
                    .collect(),
                notes: Vec::new()
            });
        }

        Ok(cours_nouv.values().cloned().collect())
    }

    fn obtenir_évaluations(&self, cours: &Cours) -> Result<Vec<Évaluation>> {
        let feuille = self.ods.sheet(cours.idx);
        let n_colonnes = feuille.used_cols();

        let mut idx_dernier = 0;
        let mut évaluations = Vec::new();
        for colonne in DÉCALAGE_NOTES..n_colonnes {
            let nom_évaluation = cellule_str(feuille.value(0, colonne));
            let nom_section = cellule_str(feuille.value(1, colonne));
            let nom_composant = cellule_str(feuille.value(2, colonne));

            if !nom_évaluation.is_empty() {
                évaluations.push(Évaluation {
                    idx: idx_dernier,
                    nom: nom_évaluation,
                    sections: Vec::new()
                });
            }

            if !nom_section.is_empty() {
                if let Some(évaluation) = évaluations.last_mut() {
                    évaluation.sections.push(SectionÉvaluation {
                        idx: idx_dernier,
                        nom: nom_section,
                        composants: Vec::new()
                    });
                }
            }

            if !nom_composant.is_empty() {
                if let Some(évaluation) = évaluations.last_mut() {
                    if let Some(section) = évaluation.sections.last_mut() {
                        section.composants.push(ComposantÉvaluation {
                            idx: idx_dernier,
                            colonne: colonne,
                            nom: nom_composant
                        });
                        idx_dernier += 1;
                    }
                }
            }
        }

        Ok(évaluations)
    }

    fn obtenir_notes(&self, cours: &Cours, élève: &Élève) -> Result<Vec<Option<f64>>> {
        let feuille = self.ods.sheet(cours.idx);
        let ligne = élève.idx;

        let mut notes = Vec::new();

        for évaluation in &cours.évaluations {
            for section in &évaluation.sections {
                for composant in &section.composants {
                    if let Value::Number(n) = feuille.value(ligne, composant.colonne) {
                        notes.push(Some(*n));
                    } else {
                        notes.push(None);
                    }
                }
            }
        }

        Ok(notes)
    }

    pub fn obtenir_données(&self) -> Result<Vec<Cours>> {
        let mut cours = Vec::new();
        for mut c in &mut self.obtenir_cours()? {
            c.évaluations = self.obtenir_évaluations(&c)?;

            let mut sous_cours = self.obtenir_élèves(&c)?;
            for sc in &mut sous_cours {
                for mut élève in &mut sc.élèves {
                    élève.notes = self.obtenir_notes(&c, &élève)?;
                }
            }
            cours.extend(sous_cours);
        }

        Ok(cours)
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
        Value::TextXml(t) => xml_str(&t.iter().map(|x| XmlContent::Tag(x.clone())).collect()).trim().to_string(),
        Value::Boolean(b) => if *b { "v" } else { "f" }.to_string(),
        Value::Number(n) => n.to_string(),
        Value::Percentage(n) => (n * 100.).to_string() + "%",
        Value::Currency(n, devise) => n.to_string() + &std::str::from_utf8(devise).unwrap(),
        Value::DateTime(t) => t.format("%Y-%m-%d %H:%M:%S").to_string(),
        Value::TimeDuration(d) => d.to_std().unwrap().as_secs_f32().to_string(),
        _ => "".to_string()
    }
}
