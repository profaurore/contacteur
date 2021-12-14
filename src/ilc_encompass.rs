use crate::encompass::{ClientEncompass, Cours};
use crate::erreur::Result;
use crate::ilc;

pub fn obtenir_contacts() -> Result<Vec<Cours>> {
    println!("Connexion...");
    let mut client = ilc::boucler_travail(|| {
        let (utilisateur, mot_de_passe) = ilc::obtenir_identifiants("EnCompass");
        ClientEncompass::new(&utilisateur, &mot_de_passe)
    })?;
    println!("Connexion réussie!");

    println!("Obtention de la liste des groupes...");
    let groupes = ilc::boucler_travail(|| client.obtenir_groupes())?;
    println!("Obtention réussie!");

    let mut cours = Vec::new();
    for g in groupes {
        println!("Obtention des élèves pour {}...", g.code);
        let élèves_cours = ilc::boucler_travail(|| client.obtenir_élèves_groupe(&g))?;
        cours.extend(élèves_cours);
        println!("Obtention réussie!")
    }

    for c in &mut cours {
        for élève in &mut c.élèves {
            println!("Obtention des contacts pour {} {} {}...", c.code, élève.prénom, élève.nom);
            let (naissance, contacts_élève) = ilc::boucler_travail(|| client.obtenir_données_élève(élève))?;
            élève.naissance = naissance;
            élève.contacts = contacts_élève;
            println!("Obtention réussie!")
        }
    }

    Ok(cours)
}
