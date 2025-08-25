use crate::errors::{SageResult};
use crate::com::{SafeDispatch, SafeVariant, FromDispatch};
use windows::Win32::System::Com::IDispatch;

/// Wrapper pour l'objet IBILoggable avec accès typé aux propriétés
pub struct ILoggable {
    pub dispatch: IDispatch,
}


impl ILoggable {
    /// Crée un SafeDispatch temporaire pour les appels
    fn dispatch(&self) -> SafeDispatch {
        SafeDispatch::new(&self.dispatch)
    }

    /// Vérifie si l'utilisateur actuel est administrateur - ÉQUIVALENT .IsAdministrator
    pub fn is_administrator(&self) -> SageResult<bool> {
        self.dispatch().call_method_by_name("IsAdministrator", &[])?
            .to_bool()
    }

    /// Vérifie si un utilisateur est connecté - ÉQUIVALENT .IsLogged
    pub fn is_logged(&self) -> SageResult<bool> {
        self.dispatch().call_method_by_name("IsLogged", &[])?
            .to_bool()
    }

    /// Obtient le nom d'utilisateur connecté - ÉQUIVALENT .UserName
    pub fn get_user_name(&self) -> SageResult<String> {
        self.dispatch().call_method_by_name("UserName", &[])?
            .to_string()
    }

    /// SETTER - Définit le nom d'utilisateur - ÉQUIVALENT .Loggable.UserName = "<Administrateur>"
    pub fn set_user_name(&self, username: &str) -> SageResult<()> {
        let username_variant = SafeVariant::from_string(username);
        self.dispatch().call_property_put("UserName", &[username_variant])?;
        Ok(())
    }

    /// SETTER - Définit le mot de passe utilisateur - ÉQUIVALENT .Loggable.UserPwd = ""
    pub fn set_user_pwd(&self, password: &str) -> SageResult<()> {
        let pwd_variant = SafeVariant::from_string(password);
        self.dispatch().call_property_put("UserPwd", &[pwd_variant])?;
        Ok(())
    }

    /// Obtient l'ID du service - ÉQUIVALENT .ServiceId
    pub fn service_id(&self) -> SageResult<String> {
        self.dispatch().call_method_by_name("ServiceId", &[])?
            .to_string()
    }

    /// Obtient le nom du service - ÉQUIVALENT .ServiceName
    pub fn service_name(&self) -> SageResult<String> {
        self.dispatch().call_method_by_name("ServiceName", &[])?
            .to_string()
    }

    /// Obtient le token d'application - ÉQUIVALENT .ApplicationToken
    pub fn application_token(&self) -> SageResult<String> {
        self.dispatch().call_method_by_name("ApplicationToken", &[])?
            .to_string()
    }

    /// Accès direct au dispatch pour propriétés non wrappées
    pub fn get_property(&self, name: &str) -> SageResult<SafeVariant> {
        self.dispatch().call_method_by_name(name, &[])
    }

    /// Affiche toutes les informations de l'utilisateur connecté
    pub fn user_info(&self) -> SageResult<String> {
        let is_logged = self.is_logged().unwrap_or(false);
        let is_admin = self.is_administrator().unwrap_or(false);
        let username = self.get_user_name().unwrap_or_else(|_| "Non disponible".to_string());
        
        // ServiceName peut ne pas exister, donc on ne l'utilise pas dans le résumé par défaut
        Ok(format!(
            "Utilisateur: {} | Connecté: {} | Admin: {}",
            username, is_logged, is_admin
        ))
    }
}

impl FromDispatch for ILoggable {
    fn from_dispatch(dispatch: IDispatch) -> SageResult<Self> {
        Ok(ILoggable { dispatch })
    }
}