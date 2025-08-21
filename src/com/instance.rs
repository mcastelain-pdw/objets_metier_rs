use crate::errors::{SageError, SageResult};
use windows::{
    core::*, Win32::System::Com::*, Win32::System::Variant::*,
};
use std::collections::HashMap;

/// Instance COM sûre avec gestion automatique du cycle de vie
pub struct ComInstance {
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    unknown: IUnknown,
    dispatch: Option<IDispatch>,
    initialized_com: bool,
}

#[derive(Debug, Clone)]
pub enum MemberType {
    Method,
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    PropertyGet,
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    PropertyPut,
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    PropertyPutRef,
}

#[derive(Debug, Clone)]
pub struct MemberInfo {
    pub id: i32,
    pub name: String,
    pub member_type: MemberType,
    pub param_count: Option<u32>,
    pub return_type: Option<String>,
}

impl ComInstance {
    /// Crée une nouvelle instance COM en initialisant automatiquement COM si nécessaire
    pub fn new(clsid: &str) -> SageResult<Self> {
        unsafe {
            // Initialiser COM
            let com_result = CoInitializeEx(None, COINIT_APARTMENTTHREADED);
            let initialized_com = com_result.is_ok();

            // Parser le CLSID
            let guid = Self::parse_clsid(clsid)?;

            // Créer l'instance
            let unknown: IUnknown =
                CoCreateInstance(&guid, None, CLSCTX_INPROC_SERVER).map_err(|e| {
                    if initialized_com {
                        CoUninitialize();
                    }
                    SageError::from(e)
                })?;

            // Tenter d'obtenir IDispatch pour l'automation
            let dispatch = unknown.cast::<IDispatch>().ok();

            Ok(ComInstance {
                unknown,
                dispatch,
                initialized_com,
            })
        }
    }

    /// Crée une instance à partir d'un IUnknown existant
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    pub fn from_unknown(unknown: IUnknown) -> Self {
        let dispatch = unknown.cast::<IDispatch>().ok();

        ComInstance {
            unknown,
            dispatch,
            initialized_com: false, // N'a pas initialisé COM
        }
    }

    /// Obtient l'interface IDispatch pour l'automation
    pub fn dispatch(&self) -> SageResult<&IDispatch> {
        self.dispatch.as_ref().ok_or_else(|| {
            SageError::InternalError("Interface IDispatch non disponible".to_string())
        })
    }

    /// Obtient l'interface IUnknown
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    pub fn unknown(&self) -> &IUnknown {
        &self.unknown
    }

    /// Vérifie si l'instance supporte l'automation
    pub fn supports_automation(&self) -> bool {
        self.dispatch.is_some()
    }

    /// Parse un CLSID string en GUID
    fn parse_clsid(clsid_str: &str) -> SageResult<GUID> {
        let clsid_formatted = if clsid_str.starts_with('{') {
            clsid_str.to_string()
        } else {
            format!("{{{}}}", clsid_str)
        };

        let clsid_wide: Vec<u16> = clsid_formatted
            .encode_utf16()
            .chain(std::iter::once(0))
            .collect();

        unsafe { CLSIDFromString(PCWSTR(clsid_wide.as_ptr())).map_err(SageError::from) }
    }

    /// Obtient les informations de type de l'objet COM
    pub fn get_type_info(&self) -> SageResult<String> {
        let dispatch = self.dispatch()?;

        unsafe {
            let type_info_count = dispatch.GetTypeInfoCount()?;

            if type_info_count == 0 {
                return Ok("Aucune information de type disponible".to_string());
            }

            let type_info = dispatch.GetTypeInfo(0, 0)?;

            let mut names = BSTR::default();
            let mut doc_string = BSTR::default();

            type_info.GetDocumentation(
                -1, // MEMBERID_NIL
                Some(&mut names as *mut BSTR),
                Some(&mut doc_string as *mut BSTR),
                std::ptr::null_mut(),
                None,
            )?;

            Ok(format!("Nom: {}, Description: {}", names, doc_string))
        }
    }

    /// Obtient l'interface ITypeInfo brute pour introspection avancée
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    fn get_type_info_raw(&self) -> SageResult<ITypeInfo> {
        let dispatch = self.dispatch()?;

        unsafe {
            let type_info_count = dispatch.GetTypeInfoCount()?;

            if type_info_count == 0 {
                return Err(SageError::InternalError(
                    "Aucune information de type disponible".to_string(),
                ));
            }

            dispatch.GetTypeInfo(0, 0).map_err(SageError::from)
        }
    }

    /// Liste les méthodes disponibles
    pub fn list_methods(&self) -> SageResult<Vec<(i32, String)>> {
        let dispatch = self.dispatch()?;
        let mut methods = Vec::new();

        unsafe {
            let type_info_count = dispatch.GetTypeInfoCount()?;

            if type_info_count > 0 {
                let type_info = dispatch.GetTypeInfo(0, 0)?;

                // Essayer de récupérer les méthodes par ID
                for method_id in 1..=50 {
                    // Limite arbitraire
                    let mut names = BSTR::default();

                    if type_info
                        .GetDocumentation(
                            method_id,
                            Some(&mut names as *mut BSTR),
                            None,
                            std::ptr::null_mut(),
                            None,
                        )
                        .is_ok()
                    {
                        methods.push((method_id, names.to_string()));
                    }
                }
            }
        }

        Ok(methods)
    }

    /// Liste toutes les méthodes et propriétés avec leur type (version intelligente)
    pub fn list_members(&self) -> SageResult<Vec<MemberInfo>> {
        // Utiliser la découverte basique puis analyser intelligemment
        let basic_methods = self.list_methods()?;
        let mut members = Vec::new();
        
        for (id, name) in basic_methods {
            // Analyser le nom et ID pour déterminer le type réel
            let member_type = Self::classify_member_by_name(&name, id);
            
            // Estimer le nombre de paramètres basé sur le nom
            let param_count = Self::estimate_parameter_count(&name, &member_type);
            
            // Deviner le type de retour basé sur le nom
            let return_type = Self::guess_return_type(&name, &member_type);
            
            members.push(MemberInfo {
                id,
                name,
                member_type,
                param_count,
                return_type,
            });
        }
        
        Ok(members)
    }
    
    /// Classifie un membre basé sur son nom et des heuristiques COM Sage
    fn classify_member_by_name(name: &str, _id: i32) -> MemberType {
        // Propriétés évidentes : Factory*
        if name.starts_with("Factory") {
            return MemberType::PropertyGet;
        }
        
        // Propriétés communes dans les objets Sage COM
        let property_names = [
            "Name", "Version", "Database", "User", "Server", "Path", 
            "Description", "Type", "Count", "Value", "Status", "State",
            "Mode", "Level", "Index", "Size", "IsOpen", "Application",
            "Parent", "Collection", "Handle", "ID", "Code", "Reference",
        ];
        
        // Vérification exacte des noms de propriétés
        if property_names.contains(&name) {
            return MemberType::PropertyGet;
        }
        
        // Méthodes évidentes : verbes d'action
        let method_verbs = [
            "Open", "Close", "Create", "Delete", "Add", "Remove", 
            "Update", "Save", "Load", "Connect", "Disconnect",
            "Execute", "Run", "Start", "Stop", "Cancel", "Reset",
            "Clear", "Refresh", "Reload", "Import", "Export",
            "Print", "Preview", "Validate", "Check", "Test",
            "Initialize", "Finalize", "Process", "Calculate",
            "Search", "Find", "Locate", "Get", "Set", "Move",
            "Copy", "Paste", "Cut", "Undo", "Redo", "Backup",
            "Restore", "Synchronize", "Synchro", "ReadFrom"
        ];
        
        // Si le nom commence par un verbe d'action
        if method_verbs.iter().any(|&verb| name.starts_with(verb)) {
            return MemberType::Method;
        }
        
        // Pattern CamelCase sans verbe d'action = probablement propriété
        if name.chars().next().unwrap_or('a').is_uppercase() && 
           !name.contains('(') && !name.contains("Method") {
            // Exceptions : certains noms qui ressemblent à des propriétés mais sont des méthodes
            let method_exceptions = [
                "DatabaseInfo", "ReadFrom", "WriteTo", "ToString"
            ];
            
            if !method_exceptions.contains(&name) {
                return MemberType::PropertyGet;
            }
        }
        
        // Par défaut : méthode
        MemberType::Method
    }
    
    /// Estime le nombre de paramètres basé sur le nom et type
    fn estimate_parameter_count(name: &str, member_type: &MemberType) -> Option<u32> {
        match member_type {
            MemberType::PropertyGet => Some(0), // Les getters n'ont pas de paramètres
            MemberType::PropertyPut | MemberType::PropertyPutRef => Some(1), // Les setters ont 1 paramètre
            MemberType::Method => {
                // Estimation basée sur le nom de méthode
                match name {
                    "IsOpen" | "Close" | "Create" | "Save" | "Clear" | "Refresh" => Some(0),
                    "Open" | "Delete" | "Add" | "Remove" => Some(1),
                    "Update" | "Copy" | "Move" => Some(2),
                    _ => None, // Paramètres inconnus
                }
            }
        }
    }
    
    /// Devine le type de retour basé sur le nom et type
    fn guess_return_type(name: &str, member_type: &MemberType) -> Option<String> {
        match member_type {
            MemberType::PropertyGet => {
                if name.starts_with("Factory") {
                    Some("Object".to_string()) // Les Factory retournent des objets
                } else if name.starts_with("Is") || name.ends_with("ed") {
                    Some("Boolean".to_string()) // Les propriétés booléennes
                } else if name.contains("Count") || name.contains("Size") {
                    Some("Integer".to_string()) // Les propriétés numériques
                } else if name == "Name" || name == "Description" || name == "Path" {
                    Some("String".to_string()) // Les propriétés texte
                } else {
                    Some("Variant".to_string()) // Type générique
                }
            },
            MemberType::PropertyPut | MemberType::PropertyPutRef => {
                Some("void".to_string()) // Les setters ne retournent rien
            },
            MemberType::Method => {
                match name {
                    "IsOpen" => Some("Boolean".to_string()),
                    "Open" | "Close" | "Create" | "Save" | "Delete" => Some("void".to_string()),
                    "DatabaseInfo" => Some("String".to_string()),
                    _ => Some("Variant".to_string()), // Type générique pour les méthodes
                }
            }
        }
    }
    
    /// Filtre uniquement les méthodes
    pub fn list_methods_only(&self) -> SageResult<Vec<MemberInfo>> {
        let members = self.list_members()?;
        Ok(members.into_iter()
            .filter(|m| matches!(m.member_type, MemberType::Method))
            .collect())
    }
    
    /// Filtre uniquement les propriétés
    pub fn list_properties(&self) -> SageResult<Vec<MemberInfo>> {
        let members = self.list_members()?;
        Ok(members.into_iter()
            .filter(|m| matches!(m.member_type, 
                MemberType::PropertyGet | 
                MemberType::PropertyPut | 
                MemberType::PropertyPutRef))
            .collect())
    }
    
    /// Groupe les propriétés par nom (Get/Put/PutRef ensemble)
    pub fn group_properties(&self) -> SageResult<HashMap<String, Vec<MemberInfo>>> {
        let properties = self.list_properties()?;
        let mut grouped = HashMap::new();
        
        for prop in properties {
            grouped.entry(prop.name.clone())
                .or_insert_with(Vec::new)
                .push(prop);
        }
        
        Ok(grouped)
    }
    
    /// Convertit le VARTYPE en nom de type lisible
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    fn get_type_name(vt: VARENUM) -> Option<String> {
        let type_name = match vt {
            VT_EMPTY => "void",
            VT_NULL => "null", 
            VT_I2 => "short",
            VT_I4 => "long",
            VT_R4 => "float",
            VT_R8 => "double",
            VT_CY => "currency",
            VT_DATE => "date",
            VT_BSTR => "string",
            VT_DISPATCH => "object",
            VT_ERROR => "error",
            VT_BOOL => "bool",
            VT_VARIANT => "variant",
            VT_UNKNOWN => "unknown",
            VT_DECIMAL => "decimal",
            VT_UI1 => "byte",
            VT_UI2 => "ushort",
            VT_UI4 => "ulong",
            VT_I8 => "longlong",
            VT_UI8 => "ulonglong",
            VT_HRESULT => "hresult",
            VT_PTR => "pointer",
            VT_SAFEARRAY => "array",
            _ => return None,
        };
        Some(type_name.to_string())
    }
}

impl Drop for ComInstance {
    fn drop(&mut self) {
        // Libérer COM si on l'a initialisé
        if self.initialized_com {
            unsafe {
                CoUninitialize();
            }
        }
    }
}

// Implémentation Send et Sync pour utilisation multi-thread (avec précautions)
unsafe impl Send for ComInstance {}
unsafe impl Sync for ComInstance {}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_clsid_parsing() {
        // Test avec et sans accolades
        let result1 = ComInstance::parse_clsid("309DE0FB-9FB8-4F4E-8295-CC60C60DAA33");
        let result2 = ComInstance::parse_clsid("{309DE0FB-9FB8-4F4E-8295-CC60C60DAA33}");

        assert!(result1.is_ok());
        assert!(result2.is_ok());
    }
}
