use super::SafeVariant;
use crate::errors::{SageError, SageResult};
use windows::{Win32::System::Com::*, Win32::System::Ole::*, Win32::System::Variant::*, core::*};

/// Wrapper sûr pour les appels IDispatch
pub struct SafeDispatch<'a> {
    dispatch: &'a IDispatch,
}

impl<'a> SafeDispatch<'a> {
    /// Crée un nouveau wrapper SafeDispatch
    pub fn new(dispatch: &'a IDispatch) -> Self {
        SafeDispatch { dispatch }
    }

    /// Appelle une méthode COM par ID avec paramètres
    #[allow(dead_code)] // Sera utilisé dans les futures versions
    pub fn call_method(&self, method_id: i32, method_name: &str) -> SageResult<SafeVariant> {
        self.call_method_with_params(method_id, method_name, &[])
    }

    /// Appelle une méthode avec paramètres
    pub fn call_method_with_params(
        &self,
        method_id: i32,
        method_name: &str,
        params: &[SafeVariant],
    ) -> SageResult<SafeVariant> {
        unsafe {
            let mut result = VARIANT::default();
            let mut excep_info = EXCEPINFO::default();
            let mut arg_err: u32 = 0;

            // Convertir les paramètres SafeVariant en VARIANT
            let mut variant_params = Vec::new();
            for param in params {
                variant_params.push(param.to_variant()?);
            }

            let dispparams = DISPPARAMS {
                rgvarg: if variant_params.is_empty() {
                    std::ptr::null_mut()
                } else {
                    variant_params.as_ptr() as *mut VARIANT
                },
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: variant_params.len() as u32,
                cNamedArgs: 0,
            };

            let hr = self.dispatch.Invoke(
                method_id,
                &GUID::zeroed(),
                0,
                DISPATCH_METHOD | DISPATCH_PROPERTYGET,
                &dispparams,
                Some(&mut result),
                Some(&mut excep_info),
                Some(&mut arg_err),
            );

            match hr {
                Ok(_) => SafeVariant::from_variant(result),
                Err(e) => {
                    // Vérifier si on a des informations d'exception
                    let error_msg = if !excep_info.bstrDescription.is_empty() {
                        excep_info.bstrDescription.to_string()
                    } else {
                        format!("Erreur COM: {}", e.message().to_string_lossy())
                    };

                    Err(SageError::method_call(method_name, method_id, &error_msg))
                }
            }
        }
    }

    /// Obtient la valeur d'une propriété
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn get_property(&self, property_id: i32, property_name: &str) -> SageResult<SafeVariant> {
        unsafe {
            let mut result = VARIANT::default();
            let mut excep_info = EXCEPINFO::default();
            let mut arg_err: u32 = 0;

            let dispparams = DISPPARAMS {
                rgvarg: std::ptr::null_mut(),
                rgdispidNamedArgs: std::ptr::null_mut(),
                cArgs: 0,
                cNamedArgs: 0,
            };

            let hr = self.dispatch.Invoke(
                property_id,
                &GUID::zeroed(),
                0,
                DISPATCH_PROPERTYGET,
                &dispparams,
                Some(&mut result),
                Some(&mut excep_info),
                Some(&mut arg_err),
            );

            match hr {
                Ok(_) => SafeVariant::from_variant(result),
                Err(e) => {
                    let error_msg = if !excep_info.bstrDescription.is_empty() {
                        excep_info.bstrDescription.to_string()
                    } else {
                        format!("Erreur COM: {}", e.message().to_string_lossy())
                    };

                    Err(SageError::method_call(
                        property_name,
                        property_id,
                        &error_msg,
                    ))
                }
            }
        }
    }

    /// Définit la valeur d'une propriété
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn set_property(
        &self,
        property_id: i32,
        property_name: &str,
        value: SafeVariant,
    ) -> SageResult<()> {
        unsafe {
            let mut excep_info = EXCEPINFO::default();
            let mut arg_err: u32 = 0;

            let variant_value = value.to_variant()?;
            let named_arg_id = DISPID_PROPERTYPUT;

            let dispparams = DISPPARAMS {
                rgvarg: &variant_value as *const VARIANT as *mut VARIANT,
                rgdispidNamedArgs: &named_arg_id as *const i32 as *mut i32,
                cArgs: 1,
                cNamedArgs: 1,
            };

            let hr = self.dispatch.Invoke(
                property_id,
                &GUID::zeroed(),
                0,
                DISPATCH_PROPERTYPUT,
                &dispparams,
                None,
                Some(&mut excep_info),
                Some(&mut arg_err),
            );

            match hr {
                Ok(_) => Ok(()),
                Err(e) => {
                    let error_msg = if !excep_info.bstrDescription.is_empty() {
                        excep_info.bstrDescription.to_string()
                    } else {
                        format!("Erreur COM: {}", e.message().to_string_lossy())
                    };

                    Err(SageError::method_call(
                        property_name,
                        property_id,
                        &error_msg,
                    ))
                }
            }
        }
    }

    /// Obtient un ID de méthode par son nom
    pub fn get_method_id(&self, method_name: &str) -> SageResult<i32> {
        unsafe {
            let name_wide: Vec<u16> = method_name
                .encode_utf16()
                .chain(std::iter::once(0))
                .collect();

            let mut dispatch_id = 0i32;
            let name_ptr = PCWSTR(name_wide.as_ptr());

            self.dispatch
                .GetIDsOfNames(&GUID::zeroed(), &name_ptr, 1, 0, &mut dispatch_id)
                .map_err(|e| {
                    SageError::method_call(
                        method_name,
                        -1,
                        &format!(
                            "Méthode '{}' non trouvée: {}",
                            method_name,
                            e.message().to_string_lossy()
                        ),
                    )
                })?;

            Ok(dispatch_id)
        }
    }

    /// Appelle une méthode par son nom
    pub fn call_method_by_name(
        &self,
        method_name: &str,
        params: &[SafeVariant],
    ) -> SageResult<SafeVariant> {
        let method_id = self.get_method_id(method_name)?;
        self.call_method_with_params(method_id, method_name, params)
    }

    /// Obtient une propriété par son nom
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn get_property_by_name(&self, property_name: &str) -> SageResult<SafeVariant> {
        let property_id = self.get_method_id(property_name)?;
        self.get_property(property_id, property_name)
    }

    /// Définit une propriété par son nom
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn set_property_by_name(&self, property_name: &str, value: SafeVariant) -> SageResult<()> {
        let property_id = self.get_method_id(property_name)?;
        self.set_property(property_id, property_name, value)
    }
}

#[cfg(test)]
mod tests {
    // Note: Ces tests nécessitent une instance COM réelle pour être exécutés
    // Ils sont ici à titre d'exemple de structure

    #[test]
    fn test_safe_dispatch_structure() {
        // Test de compilation uniquement
        assert!(true);
    }
}