use std::path::Path;
use rusqlite::Connection;
use crate::erreur::Result;

pub fn ouvrir(fichier: Option<&str>) -> Result<Connection> {
    let conn = match fichier {
        Some(f) => Connection::open(Path::new(f))?,
        None => Connection::open_in_memory()?
    };

    conn.execute_batch("
        BEGIN;
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
        CREATE TABLE IF NOT EXISTS échelle_niveau (
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
            id_cours INTEGER NOT NULL,
            id_parent INTEGER,
            indice INTEGER NOT NULL,
            id_échelle INTEGER,
            formule TEXT,

            CONSTRAINT u_position UNIQUE (id_cours, id_parent, indice),
            CONSTRAINT c_type CHECK (id_parent IS NULL OR id_échelle IS NULL),
            CONSTRAINT c_formule CHECK (id_échelle IS NOT NULL OR formule IS NULL),
            CONSTRAINT f_cours FOREIGN KEY (id_cours) REFERENCES cours(id),
            CONSTRAINT f_parent FOREIGN KEY (id_parent) REFERENCES évaluation_item(id),
            CONSTRAINT c_indice CHECK (indice >= 0),
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
