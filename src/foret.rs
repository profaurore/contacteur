use std::collections::HashMap;
use crate::erreur::{Erreur, Result};

// Inspiré de https://rust-leipzig.github.io/architecture/2016/12/20/idiomatic-trees-in-rust/
// https://docs.rs/indextree/4.3.1/indextree/index.html
#[derive(Copy, Clone)]
pub struct IdNoeud {
    idx: usize
}

struct Noeud<T> {
    ascendant: Option<IdNoeud>,
    voisin_précédent: Option<IdNoeud>,
    voisin_prochain: Option<IdNoeud>,
    descendant_premier: Option<IdNoeud>,
    descendant_dernier: Option<IdNoeud>,
    pub val: T
}

pub struct Forêt<T> {
    prochain_idx: usize,
    noeuds: HashMap<usize, Noeud<T>>
}

impl<T> Forêt<T> {
    pub fn new() -> Self {
        Self {
            prochain_idx: 0,
            noeuds: HashMap::new()
        }
    }

    pub fn créer(&mut self, val: T) -> IdNoeud {
        let idx = self.prochain_idx;
        self.prochain_idx += 1;

        self.noeuds.insert(idx, Noeud {
            ascendant: None,
            voisin_précédent: None,
            voisin_prochain: None,
            descendant_premier: None,
            descendant_dernier: None,
            val: val
        });

        IdNoeud { idx: idx }
    }

    fn obtenir_noeud(&self, id: IdNoeud) -> Result<&Noeud<T>> {
        self.noeuds
            .get(&id.idx)
            .ok_or(Erreur::Arbre("Noeud invalide.".into()))
    }

    fn obtenir_noeud_mut(&mut self, id: IdNoeud) -> Result<&mut Noeud<T>> {
        self.noeuds
            .get_mut(&id.idx)
            .ok_or(Erreur::Arbre("Noeud invalide.".into()))
    }

    pub fn ajouter_descendant(&mut self, id_ascendant: IdNoeud, val: T) -> Result<IdNoeud> {
        let id = self.créer(val);
        let vieux_descendant_dernier = self.obtenir_noeud_mut(id_ascendant)?.descendant_dernier.replace(id);
        if let Some(dernier) = vieux_descendant_dernier {
            self.noeuds.get_mut(&dernier.idx).unwrap().voisin_prochain = Some(id);
            self.noeuds.get_mut(&id.idx).unwrap().voisin_précédent = Some(dernier);
        }
        Ok(id)
    }
}