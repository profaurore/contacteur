use std::collections::HashMap;
use chrono::Local;
use regex::Regex;
use rusqlite::{Connection, params};
use spreadsheet_ods::{CellStyle, read_ods, Sheet, style::{FontFaceDecl, units::Length}, WorkBook, write_ods};
use crate::classeur::ClientClasseur;
use crate::erreur::Result;
use crate::ilc_encompass;

pub fn importer_encompass(conn: &mut Connection) -> Result<()> {
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
        for cours in ilc_encompass::obtenir_contacts()? {
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

pub fn importer_notes_classeur(conn: &mut Connection) -> Result<()> {
    let classeur = ClientClasseur::new(r"évaluations.ods")?;

    conn.execute_batch("
        BEGIN;

        DELETE FROM échelle_niveau;
        DELETE FROM échelle;
        DELETE FROM évaluation_résultat;
        DELETE FROM évaluation_reprise;
        DELETE FROM évaluation_item;

        INSERT INTO échelle(id, nom, précision, min, max) VALUES
            (1, 'Niveau', 4, 0, 4),
            (2, 'Pourcentage', 0, 0, 100);

        COMMIT;
    ")?;

    for cours in classeur.obtenir_données()? {
        let id_cours: i64 = conn.query_row("
            SELECT id FROM cours WHERE code = ?1
        ", [cours.code], |r| r.get(0))?;

        let mut ids_élèves = Vec::new();
        for élève in &cours.élèves {
            let id_élève: i64 = conn.query_row("
                SELECT id FROM élève WHERE prénom = ?1 AND nom = ?2 AND id_cours = ?3;
            ", params![élève.prénom, élève.nom, id_cours], |r| r.get(0))?;
            ids_élèves.push(id_élève);

            conn.execute("
                UPDATE élève SET prénom_préféré = ?1
                    WHERE id = ?2;
            ", params![élève.prénom_préféré, id_élève]).unwrap();

            for étiquette in &élève.étiquettes {
                conn.execute("
                    INSERT INTO élève_étiquette(id_élève, id_étiquette)
                        SELECT ?1, étiquette.id
                        FROM étiquette
                        WHERE étiquette.nom = ?2;
                ", params![id_élève, étiquette]).unwrap();
            }
        }

        for (i, évaluation) in cours.évaluations.iter().enumerate() {
            conn.execute("
                INSERT INTO évaluation_item(nom, id_cours, indice) VALUES (?1, ?2, ?3);
            ", params![évaluation.nom, id_cours, i]).unwrap();
            let id_évaluation = conn.last_insert_rowid();

            for (j, section) in évaluation.sections.iter().enumerate() {
                conn.execute("
                    INSERT INTO évaluation_item(nom, id_cours, id_parent, indice)
                        VALUES (?1, ?2, ?3, ?4);
                ", params![section.nom, id_cours, Some(id_évaluation), j]).unwrap();
                let id_section = conn.last_insert_rowid();

                for (k, composant) in section.composants.iter().enumerate() {
                    //println!("!!! {} {} {}", évaluation.nom, section.nom, composant.nom);
                    conn.execute("
                        INSERT INTO évaluation_item(nom, id_cours, id_parent, indice)
                            VALUES (?1, ?2, ?3, ?4);
                    ", params![composant.nom, id_cours, Some(id_section), k]).unwrap();
                    let id_composant = conn.last_insert_rowid();
                
                    for (id_élève, élève) in ids_élèves.iter().zip(&cours.élèves) {
                        if let Some(n) = élève.note(&composant) {
                            conn.execute("
                                INSERT INTO évaluation_résultat(id_item, id_élève, résultat)
                                    SELECT ?1, élève.id, ?3
                                    FROM élève
                                    WHERE prénom = ?1 AND nom = ?2 AND id_cours = ?3;
                            ", params![Some(id_composant), id_élève, n]).unwrap();
                        }
                    }
                }
            }
        }
    }

    Ok(())
}

pub fn exporter_contacts_classeur(conn: &Connection) -> Result<()> {
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
