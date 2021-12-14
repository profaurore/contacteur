use rusqlite::{params, Connection};
use spreadsheet_ods::{Value, read_ods, xmltree::XmlContent};
use regex::Regex;
use std::collections::HashMap;

fn main() {
    let conn = Connection::open_in_memory().unwrap();

    conn.execute_batch("
        BEGIN TRANSACTION;
        CREATE TABLE cours (
            id INTEGER PRIMARY KEY,
            code TEXT NOT NULL,
            nom TEXT
        );
        CREATE TABLE étiquette (
            id INTEGER PRIMARY KEY,
            nom TEXT NOT NULL
        );
        INSERT INTO étiquette(nom) VALUES ('Virtuel');
        INSERT INTO étiquette(nom) VALUES ('AP');
        CREATE TABLE élève (
            id INTEGER PRIMARY KEY,
            prénom_préféré TEXT,
            prénom TEXT NOT NULL,
            nom TEXT NOT NULL,
            id_cours INTEGER NOT NULL,

            CONSTRAINT f_cours FOREIGN KEY (id_cours) REFERENCES cours(id)
        );
        CREATE TABLE élève_étiquette (
            id_élève INTEGER NOT NULL,
            id_étiquette INTEGER NOT NULL,

            CONSTRAINT f_élève FOREIGN KEY (id_élève) REFERENCES élève(id),
            CONSTRAINT f_étiquette FOREIGN KEY (id_étiquette) REFERENCES étiquette(id)
        );
        CREATE TABLE élève_contact (
            id INTEGER PRIMARY KEY,
            id_élève INTEGER NOT NULL,
            nom TEXT NOT NULL,
            relation TEXT,
            automatique INTEGER NOT NULL,

            CONSTRAINT u_nom UNIQUE (id_élève, nom),
            CONSTRAINT c_automatique CHECK (automatique = 0 OR automatique = 1)
        );
        CREATE TABLE élève_contact_type (
            id INTEGER PRIMARY KEY,
            type TEXT NOT NULL,

            CONSTRAINT u_type UNIQUE (type)
        );
        INSERT INTO élève_contact_type(type) VALUES ('Courriel');
        INSERT INTO élève_contact_type(type) VALUES ('Téléphone cellulaire');
        INSERT INTO élève_contact_type(type) VALUES ('Téléphone au domicile');
        INSERT INTO élève_contact_type(type) VALUES ('Téléphone au travail');
        CREATE TABLE élève_contact_item (
            id_contact INTEGER NOT NULL,
            id_type INTEGER NOT NULL,
            coordonnée TEXT NOT NULL,
            automatique INTEGER NOT NULL,

            CONSTRAINT f_contact FOREIGN KEY (id_contact) REFERENCES élève_contact(id),
            CONSTRAINT f_type FOREIGN KEY (id_type) REFERENCES élève_contact_type(id),
            CONSTRAINT u_coordonnée UNIQUE (id_contact, coordonnée, automatique),
            CONSTRAINT c_automatique CHECK (automatique = 0 OR automatique = 1)
        );
        CREATE TABLE échelle_niveaux (
            id_échelle INTEGER NOT NULL,
            nom TEXT NOT NULL,
            min REAL NOT NULL,
            max REAL NOT NULL,

            CONSTRAINT f_échelle FOREIGN KEY (id_échelle) REFERENCES échelle(id),
            CONSTRAINT c_étendu CHECK (min <= max)
        );
        CREATE TABLE échelle (
            id INTEGER PRIMARY KEY,
            nom TEXT NOT NULL,
            précision INTEGER NOT NULL,
            min REAL NOT NULL,
            max REAL NOT NULL,

            CONSTRAINT c_étendu CHECK (min <= max)
        );
        CREATE TABLE évaluation_item (
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
        CREATE TABLE évaluation_reprise (
            id INTEGER PRIMARY KEY,
            exclus INTEGER NOT NULL DEFAULT 0,
            temps TEXTE NOT NULL,

            CONSTRAINT c_exclus CHECK (exclus = 0 OR exclus = 1)
        );
        CREATE TABLE évaluation_résultat (
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
    "*/).unwrap();


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
}
