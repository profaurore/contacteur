mod bd;
mod classeur;
mod connecteurs;
mod encompass;
mod erreur;
mod ilc;
mod ilc_encompass;

use crate::erreur::Result;
use crate::connecteurs::{exporter_contacts_classeur, importer_encompass};

fn main() -> Result<()> {
    let mut conn = bd::ouvrir(None)?;

    importer_encompass(&mut conn)?;

    println!("Exportation des données à un fichier...");
    exporter_contacts_classeur(&conn)?;
    println!("Exportation réussie!");

    Ok(())
}
