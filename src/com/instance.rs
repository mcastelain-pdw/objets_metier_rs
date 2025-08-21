use crate::errors::{SageError, SageResult};
use windows::{Win32::System::Com::*, core::*};

/// Instance COM sûre avec gestion automatique du cycle de vie
pub struct ComInstance {
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    unknown: IUnknown,
    dispatch: Option<IDispatch>,
    initialized_com: bool,
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
