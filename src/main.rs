use regex::Regex;
use reqwest::{
    blocking::Client,
    redirect::Policy,
    StatusCode
};
use html_escape::decode_html_entities;
use csv;
use rpassword::read_password;
use std::io::Write;
use std::collections::HashMap;


struct Erreur {
    err: TypeErreur
}


#[derive(Debug)]
enum TypeErreur {
    IdentifiantsInvalides,
    ÉtatInconnu(StatusCode),
    Requête(reqwest::Error),
    Csv(csv::Error)
}


impl Erreur {
    fn new(err: TypeErreur) -> Self {
        Self { err: err }
    }
}


impl std::fmt::Debug for Erreur {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        let mut builder = f.debug_struct("reqwest::Error");
        builder.field("kind", &self.err);
        builder.finish()
    }
}


impl std::fmt::Display for Erreur {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        match &self.err {
            TypeErreur::IdentifiantsInvalides => f.write_str("Le nom d'utilisateur et le mot de passe sont invalides."),
            TypeErreur::ÉtatInconnu(état) => f.write_str(&format!("Le code d'état ({}) de la requête envoyé est inattendu.", état)),
            TypeErreur::Requête(e) => f.write_str(&format!("Une erreur s'est produite lors de l'envoie de la requête: {}.", e)),
            TypeErreur::Csv(e) => f.write_str(&format!("Une erreur s'est produite lors de l'envoie de la requête: {}.", e))
        }
    }
}


fn obtenir_identifiants() -> (String, String) {
    println!("***EnCompass***");

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

fn essayer_nouveau() -> bool {
    let mut réponse = String::new();
    loop {
        print!("Essayer à nouveau (o ou n)? ");
        std::io::stdout().flush().unwrap();
        std::io::stdin().read_line(&mut réponse).unwrap();

        let réponse = réponse.trim();
        match réponse {
            "o" => return true,
            "n" => return false,
            _ => ()
        }
    }
}

fn connecter_encompass(client: &mut Client, utilisateur: &str, mot_de_passe: &str) -> Result<(), Erreur> {
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
                StatusCode::OK => Err(Erreur::new(TypeErreur::IdentifiantsInvalides)),
                StatusCode::FOUND => Ok(()),
                _ => Err(Erreur::new(TypeErreur::ÉtatInconnu(res.status())))
            }
        },
        Result::Err(err) => Err(Erreur::new(TypeErreur::Requête(err)))
    }
}


fn obtenir_liste_cours(client: &mut Client) -> Result<Vec<(String, String)>, Erreur> {
    let res = client
        .get("https://french.compassforsuccess.ca/portal/messaging/searchParents.do?isMessage=false")
        .send();

    let r = Regex::new(r#"<option value="([0-9]+)">([^<]+)</option>"#).unwrap();
    match res {
        Result::Ok(res) => {
            match res.status() {
                StatusCode::OK => {
                    let page = res.text().unwrap_or(String::new());
                    Ok(r.captures_iter(&page)
                        .map(|c| (decode_html_entities(&c[1]).to_string(), decode_html_entities(&c[2]).to_string()))
                        .collect())
                },
                _ => Err(Erreur::new(TypeErreur::ÉtatInconnu(res.status())))
            }
        },
        Result::Err(err) => Err(Erreur::new(TypeErreur::Requête(err)))
    }
}


fn sauvegarder_contacts(fichier: &str, contacts: &HashMap<String, HashMap<String, ContactsÉlève>>) -> Result<(), Erreur> {
    let mut csv = csv::WriterBuilder::new().delimiter(b';').from_path(fichier).map_err(|e| Erreur::new(TypeErreur::Csv(e)))?;
    csv.write_record(["\u{FEFF}Cours", "Élève", "Parent", "Courriel"]).map_err(|e| Erreur::new(TypeErreur::Csv(e)))?;
    for (code_cours, contacts_cours) in contacts {
        for (nom_élève, contacts_élève) in contacts_cours {
            for contact in &contacts_élève.parents {
                csv.write_record([&code_cours, &nom_élève, &contact.nom, &contact.courriel]).map_err(|e| Erreur::new(TypeErreur::Csv(e)))?;
            }
        }
    }
    csv.flush().unwrap();

    Ok(())
}


struct ContactsÉlève {
    parents: Vec<Contact>
}


impl ContactsÉlève {
    pub fn new() -> Self {
        Self { parents: Vec::new() }
    }
}

struct Contact {
    nom: String,
    courriel: String
}


fn obtenir_contacts_cours(client: &mut Client, id_cours: &str) -> Result<HashMap<String, ContactsÉlève>, Erreur> {
    let res = client
        .get(format!("https://french.compassforsuccess.ca/portal/messaging/searchParents.do?isMessage=false&classId={}", id_cours))
        .send();

    let r = Regex::new(r"(?s:<tr[^>]*>\s*<td[^>]*>.*?</td>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>.*?</td>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*<td[^>]*>(.*?)</td>\s*</tr>)").unwrap();
    match res {
        Result::Ok(res) => {
            match res.status() {
                StatusCode::OK => {
                    let page = res.text().unwrap_or(String::new());

                    let mut contacts = HashMap::new();
                    for c in r.captures_iter(&page) {
                        contacts
                            .entry(c[3].to_string())
                            .or_insert_with(|| ContactsÉlève::new())
                            .parents
                            .push(Contact { nom: c[1].to_string(), courriel: c[2].to_string() });
                    }

                    Ok(contacts)
                },
                _ => Err(Erreur::new(TypeErreur::ÉtatInconnu(res.status())))
            }
        },
        Result::Err(err) => Err(Erreur::new(TypeErreur::Requête(err)))
    }
}


fn boucler_travail<F, P, X>(mut travail: F, mut post_travail: P) -> bool
    where
        F: FnMut() -> Result<X, Erreur>,
        P: FnMut(X) -> () {
    loop {
        match travail() {
            Result::Ok(résultat) => {
                post_travail(résultat);
                return true
            },
            Result::Err(err) => {
                println!("Erreur: {}", err);
                if !essayer_nouveau() {
                    return false
                }
            }
        }
    }
}


fn main() {
    let mut client = Client::builder()
        .cookie_store(true)
        .redirect(Policy::none())
        .build().unwrap();

    println!("Connexion...");
    if !boucler_travail(
        || {
            let (utilisateur, mot_de_passe) = obtenir_identifiants();
            connecter_encompass(&mut client, &utilisateur, &mot_de_passe)
        },
        |_| ()) {
        return
    }
    println!("Connexion réussie!");

    println!("Obtention de la liste des cours...");
    let mut liste_cours = Vec::new();
    if !boucler_travail(
        || obtenir_liste_cours(&mut client),
        |l| liste_cours.extend(l)
    ) {
        return
    }
    println!("Obtention réussie!");

    let mut contacts = HashMap::new();
    for (id, code) in liste_cours {
        println!("Obtention des contacts pour {}...", code);
        if !boucler_travail(
            || obtenir_contacts_cours(&mut client, &id),
            |contacts_cours| { let _ = contacts.insert(code.clone(), contacts_cours); }
        ) {
            return
        }
        println!("Obtention réussie!")
    }

    const FICHIER_SAUVEGARDE: &str = "contacts.csv";
    println!("Sauvegarde au fichier {}...", FICHIER_SAUVEGARDE);
    if !boucler_travail(
        || sauvegarder_contacts(FICHIER_SAUVEGARDE, &contacts),
        |_| ()
    ) {
        return
    }
    println!("Sauvegarde réussie");

    let mut vide = String::new();
    print!("Pesez entrer pour quitter...");
    std::io::stdout().flush().unwrap();
    std::io::stdin().read_line(&mut vide).unwrap();
}
