use crate::errors::{SageError, SageResult};
use crate::com::{ComInstance, SafeDispatch, SafeVariant};
use crate::wrappers::{ILoggable};

/// Wrapper pour l'application BSCPTAApplication avec accès typé aux propriétés
pub struct CialApplication {
    instance: ComInstance,
}

impl CialApplication {
    /// Crée un wrapper CialApplication à partir d'un CLSID
    pub fn new(clsid: &str) -> SageResult<Self> {
        let instance = ComInstance::new(clsid)?;
        
        Ok(CialApplication {
            instance,
        })
    }

    /// Accès à l'instance COM sous-jacente pour les opérations avancées
    pub fn instance(&self) -> &ComInstance {
        &self.instance
    }

    /// Crée un SafeDispatch temporaire pour les appels
    fn dispatch(&'_ self) -> SageResult<SafeDispatch<'_>> {
        let dispatch_ref = self.instance.dispatch()?;
        Ok(SafeDispatch::new(dispatch_ref))
    }

    /// Vérifie si une base de données est ouverte
    pub fn is_open(&self) -> SageResult<bool> {
        self.dispatch()?.call_method_by_name("IsOpen", &[])?
            .to_bool()
    }

    /// Obtient le nom de l'application
    pub fn get_name(&self) -> SageResult<String> {
        self.dispatch()?.call_method_by_name("Name", &[])?
            .to_string()
    }

    /// SETTER - Définit le chemin de la base de données - ÉQUIVALENT _mCpta.Name = "D:\\TMP\\BIJOU.MAE"
    pub fn set_name(&self, database_path: &str) -> SageResult<()> {
        let path_variant = SafeVariant::from_string(database_path);
        // Utilisation de PROPPUT pour définir la propriété
        self.dispatch()?.call_property_put("Name", &[path_variant])?;
        Ok(())
    }

    /// Accède à l'objet Loggable - ÉQUIVALENT .Loggable en C#/VB
    pub fn loggable(&self) -> SageResult<ILoggable> {
        let loggable_variant = self.dispatch()?.call_method_by_name("Loggable", &[])?;
        
        if !loggable_variant.is_object() {
            return Err(SageError::ConversionError {
                from_type: loggable_variant.type_name().to_string(),
                to_type: "IBILoggable".to_string(),
                value: "Propriété Loggable n'est pas un objet COM".to_string(),
            });
        }

        let loggable_dispatch = loggable_variant.to_dispatch()?;
        
        Ok(ILoggable { dispatch: loggable_dispatch })
    }

    /// NOUVELLE MÉTHODE - Ouvre une base de données sans paramètre - ÉQUIVALENT _mCpta.Open()
    /// Utilise le chemin défini précédemment avec set_name()
    pub fn open(&self) -> SageResult<()> {
        self.dispatch()?.call_method_by_name("Open", &[])?;
        Ok(())
    }

    /// Ferme la base de données
    pub fn close(&self) -> SageResult<()> {
        self.dispatch()?.call_method_by_name("Close", &[])?;
        Ok(())
    }

    /// Crée une nouvelle base de données
    pub fn create(&self) -> SageResult<()> {
        self.dispatch()?.call_method_by_name("Create", &[])?;
        Ok(())
    }

    /// Accède à l'objet FactoryArticle - ÉQUIVALENT .FactoryArticle en C#/VB
    pub fn factory_article(&self) -> SageResult<SafeVariant> {
        self.dispatch()?.call_method_by_name("FactoryArticle", &[])
    }

    /// Obtient les informations sur la base de données
    pub fn database_info(&self) -> SageResult<String> {
        self.dispatch()?.call_method_by_name("DatabaseInfo", &[])?
            .to_string()
    }
}

#[cfg(test)]
mod tests {

    #[test]
    fn test_sage_application_creation() {
        // Ce test nécessiterait un environnement Sage pour fonctionner
        // Il sert de documentation pour l'utilisation
    }
}
