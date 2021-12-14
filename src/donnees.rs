mod bd;
mod classeur;
mod connecteurs;
mod encompass;
mod erreur;
mod foret;
mod ilc;
mod ilc_encompass;

use std::collections::HashMap;
use regex::Regex;
use rusqlite::params;
use spreadsheet_ods::{read_ods, Value, xmltree::XmlContent};
use crate::erreur::{Erreur, Result};
use crate::foret::Forêt;
use crate::connecteurs::{exporter_contacts_classeur, importer_encompass, importer_notes_classeur};

fn main() -> Result<()> {
    let mut conn = bd::ouvrir(Some("contacteur.db3"))?;

    let sauter = false;
    if sauter {
        importer_encompass(&mut conn)?;

        println!("Exportation des données à un fichier...");
        exporter_contacts_classeur(&conn)?;
        println!("Exportation réussie!");

        println!("Importation des notes d'évaluation...");
        importer_notes_classeur(&mut conn)?;
        println!("Importation réussie!");
    }


    // TODO: exporter au format PDF/html-email approprié
    return Ok(());

    let mut stmt = conn.prepare("
        SELECT id, code FROM cours ORDER BY code;
    ").unwrap();
    let ids_cours = stmt.query_map([], |r| Ok((
        r.get_unwrap::<_, u32>(0),
        r.get_unwrap::<_, String>(1)
    )))?.filter_map(|id| id.ok());
    for (id, code) in ids_cours {
        let mut évaluations = Forêt::new();
        let mut items = HashMap::new();
        println!("---{}", code);
        let mut stmt = conn.prepare("
            WITH RECURSIVE
                arbre(id, nom, id_parent, niveau, indice) AS (
                    SELECT id, nom, id_parent, 0, indice
                        FROM évaluation_item
                        WHERE id_cours = ?1 AND id_parent IS NULL
                    UNION ALL
                    SELECT éi.id, éi.nom, éi.id_parent, arbre.niveau+1, éi.indice
                        FROM évaluation_item AS éi
                        JOIN arbre ON éi.id_parent = arbre.id
                        ORDER BY 3 DESC, éi.indice
                )
            SELECT id, nom, id_parent FROM arbre;
        ").unwrap();
        let mut rangées = stmt.query([id])?;
        while let Some(r) = rangées.next()? {
            println!("{}", r.get_unwrap::<_, String>(1));
            let id: u32 = r.get_unwrap(0);
            let id_parent: Option<u32> = r.get_unwrap(2);
            let id_loc = match id_parent {
                Some(id_parent) => évaluations.ajouter_descendant(items[&id_parent], id)?,
                None => évaluations.créer(id)
            };
            items.insert(id, id_loc);
        }
        println!("Fini!");
    }

    Ok(())
}
/*
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
*/