use std::io::Write;
use rpassword::read_password;
use crate::erreur::{Erreur, Result};

pub fn boucler_travail<F, X>(mut travail: F) -> Result<X>
where
    F: FnMut() -> Result<X> {
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
                        "n" => return Err(Erreur::Abandonné),
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