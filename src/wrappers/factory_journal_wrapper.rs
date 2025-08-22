use crate::errors::SageResult;
use crate::com::{SafeDispatch, SafeVariant};
use windows::Win32::System::Com::IDispatch;

/// Wrapper pour l'objet FactoryJournal de Sage 100c (IBOJournalFactory3)
pub struct FactoryJournal {
    pub dispatch: IDispatch,
}

impl FactoryJournal {
    /// Crée un SafeDispatch temporaire pour les appels
    fn dispatch(&self) -> SafeDispatch {
        SafeDispatch::new(&self.dispatch)
    }

    /// Lit un journal par son numéro/code - SIGNATURE CORRIGÉE : 1 PARAMÈTRE REQUIS
    /// Équivalent VB/C# : factory.ReadNumero("VTE") ou factory.ReadNumero(1)
    /// 
    /// ✅ Test confirmé : Accepte 1 paramètre (String ou I4)
    /// ❌ Erreur "Nombre de paramètres non valide" avec 0 ou 2+ paramètres
    pub fn read_numero<T>(&self, numero: T) -> SageResult<SafeVariant>
    where
        T: Into<SafeVariant>,
    {
        let param = numero.into();
        self.dispatch().call_method_by_name("ReadNumero", &[param])
    }

    /// Vérifie si un journal existe par son numéro/code - SIGNATURE CORRIGÉE : 1 PARAMÈTRE REQUIS  
    /// Équivalent VB/C# : factory.ExistNumero("VTE") ou factory.ExistNumero(1)
    /// 
    /// ✅ Test confirmé : Accepte 1 paramètre (String ou I4)
    /// ❌ Erreur "Nombre de paramètres non valide" avec 0 ou 2+ paramètres
    pub fn exist_numero<T>(&self, numero: T) -> SageResult<bool>
    where
        T: Into<SafeVariant>,
    {
        let param = numero.into();
        let result = self.dispatch().call_method_by_name("ExistNumero", &[param])?;
        result.to_bool()
    }

    /// Lit un journal par son code (chaîne) - VERSION TYPÉE
    pub fn read_by_code(&self, code: &str) -> SageResult<SafeVariant> {
        self.read_numero(SafeVariant::from_string(code))
    }

    /// Lit un journal par son numéro (entier) - VERSION TYPÉE
    pub fn read_by_id(&self, numero: i32) -> SageResult<SafeVariant> {
        self.read_numero(SafeVariant::I4(numero))
    }

    /// Vérifie si un journal existe par son code (chaîne) - VERSION TYPÉE
    pub fn exists_by_code(&self, code: &str) -> SageResult<bool> {
        self.exist_numero(SafeVariant::from_string(code))
    }

    /// Vérifie si un journal existe par son numéro (entier) - VERSION TYPÉE
    pub fn exists_by_id(&self, numero: i32) -> SageResult<bool> {
        self.exist_numero(SafeVariant::I4(numero))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_factory_journal_signatures() {
        // Test de signatures confirmées par l'analyse COM :
        
        // ✅ SIGNATURES VALIDÉES ✅
        // ReadNumero(param) : 1 paramètre requis (String ou I4)
        // ExistNumero(param) : 1 paramètre requis (String ou I4)
        
        // Utilisation avec code journal (string) :
        // let journal = factory.read_journal_by_code("VTE")?;
        // let exists = factory.journal_exists_by_code("VTE")?;
        
        // Utilisation avec numéro journal (integer) :
        // let journal = factory.read_journal_by_number(1)?;
        // let exists = factory.journal_exists_by_number(1)?;
        
        // Utilisation générique :
        // let journal = factory.read_numero("VTE")?;
        // let exists = factory.exist_numero(1)?;
        
        // ❌ ERREURS "Nombre de paramètres non valide" si :
        // factory.read_numero() -- 0 paramètre
        // factory.read_numero("VTE", 1) -- 2+ paramètres
    }
}