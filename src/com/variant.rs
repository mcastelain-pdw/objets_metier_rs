use crate::errors::{SageError, SageResult};
use windows::{Win32::System::Variant::*, Win32::Foundation::{VARIANT_BOOL, VARIANT_TRUE, VARIANT_FALSE}, core::*};
use std::mem::ManuallyDrop;

/// Wrapper sûr pour les VARIANT COM
#[derive(Debug, Clone)]
#[allow(dead_code)] // Les variantes seront utilisées dans les futures versions
pub enum SafeVariant {
    Empty,
    Null,
    Bool(bool),
    I2(i16),
    I4(i32),
    R4(f32),
    R8(f64),
    Currency(i64), // CY
    Date(f64),     // DATE
    BStr(String),
    Error(HRESULT),
    I1(i8),
    UI1(u8),
    UI2(u16),
    UI4(u32),
    I8(i64),
    UI8(u64),
}

impl SafeVariant {
    /// Crée un SafeVariant à partir d'un VARIANT Windows - PRAGMATIQUE v0.1.3
    pub fn from_variant(variant: VARIANT) -> SageResult<Self> {
        unsafe {
            // Pour v0.1.3, on accepte que certains types soient convertis en string
            // TODO v0.1.4: Améliorer avec accès direct aux champs VARIANT
            
            // Obtenir le type de variant
            let vt = variant.Anonymous.Anonymous.vt;
            
            match vt {
                VT_EMPTY => Ok(SafeVariant::Empty),
                VT_NULL => Ok(SafeVariant::Null),
                
                VT_BSTR => {
                    // Accès sécurisé au BSTR
                    let bstr = &variant.Anonymous.Anonymous.Anonymous.bstrVal;
                    if bstr.is_empty() {
                        Ok(SafeVariant::BStr(String::new()))
                    } else {
                        let rust_string = bstr.to_string();
                        Ok(SafeVariant::BStr(rust_string))
                    }
                },
                
                VT_BOOL => {
                    let val = variant.Anonymous.Anonymous.Anonymous.boolVal;
                    Ok(SafeVariant::Bool(val.as_bool()))
                },
                
                VT_I2 => {
                    let val = variant.Anonymous.Anonymous.Anonymous.iVal;
                    Ok(SafeVariant::I2(val))
                },
                
                VT_I4 => {
                    let val = variant.Anonymous.Anonymous.Anonymous.lVal;
                    Ok(SafeVariant::I4(val))
                },
                
                VT_R4 => {
                    let val = variant.Anonymous.Anonymous.Anonymous.fltVal;
                    Ok(SafeVariant::R4(val))
                },
                
                VT_R8 => {
                    let val = variant.Anonymous.Anonymous.Anonymous.dblVal;
                    Ok(SafeVariant::R8(val))
                },
                
                _ => {
                    // Pour les types non supportés, retourner une description
                    Ok(SafeVariant::BStr(format!("Type VARIANT VT_{} - Conversion en v0.1.4", vt.0)))
                }
            }
        }
    }

    /// Convertit vers un VARIANT Windows - VERSION CONSERVATRICE v0.1.3
    pub fn to_variant(&self) -> SageResult<VARIANT> {
        // Pour v0.1.3, on implémente seulement les conversions de base
        // Version complète avec unions ManuallyDrop dans v0.1.4
        
        match self {
            SafeVariant::Empty => {
                Ok(VARIANT::default()) // Par défaut VT_EMPTY
            },
            
            SafeVariant::BStr(_s) => {
                // TODO v0.1.4: Créer un VARIANT avec BSTR en utilisant ManuallyDrop
                // Pour l'instant, on retourne un VARIANT par défaut
                // L'important est que from_variant fonctionne pour les retours COM
                let variant = VARIANT::default();
                Ok(variant)
            },
            
            SafeVariant::Bool(_val) => {
                // TODO v0.1.4: Implémenter avec VARIANT_BOOL
                Ok(VARIANT::default())
            },
            
            SafeVariant::I4(_val) => {
                // TODO v0.1.4: Implémenter avec lVal
                Ok(VARIANT::default())
            },
            
            SafeVariant::R8(_val) => {
                // TODO v0.1.4: Implémenter avec dblVal
                Ok(VARIANT::default())
            },
            
            // Conversions automatiques vers les types de base
            SafeVariant::I2(val) => {
                SafeVariant::I4(*val as i32).to_variant()
            },
            
            SafeVariant::R4(val) => {
                SafeVariant::R8(*val as f64).to_variant()
            },
            
            _ => {
                // Pour les types complexes, on retourne un VARIANT vide pour l'instant
                // TODO v0.1.4: Conversion complète
                Ok(VARIANT::default())
            }
        }
    }

    /// Convertit vers String si possible - AMÉLIORÉE v0.1.3
    pub fn to_string(&self) -> SageResult<String> {
        match self {
            SafeVariant::BStr(s) => Ok(s.clone()),
            SafeVariant::I2(i) => Ok(i.to_string()),
            SafeVariant::I4(i) => Ok(i.to_string()),
            SafeVariant::I8(i) => Ok(i.to_string()),
            SafeVariant::UI2(i) => Ok(i.to_string()),
            SafeVariant::UI4(i) => Ok(i.to_string()),
            SafeVariant::UI8(i) => Ok(i.to_string()),
            SafeVariant::R4(f) => Ok(f.to_string()),
            SafeVariant::R8(f) => Ok(f.to_string()),
            SafeVariant::Bool(b) => Ok(b.to_string()),
            SafeVariant::Currency(c) => {
                // Conversion currency vers decimal (diviser par 10000)
                let decimal_val = *c as f64 / 10000.0;
                Ok(format!("{:.4}", decimal_val))
            },
            SafeVariant::Date(d) => {
                // Conversion DATE COM vers string lisible
                // DATE COM = nombre de jours depuis 30/12/1899
                let days_since_epoch = *d;
                if days_since_epoch == 0.0 {
                    Ok("1899-12-30T00:00:00".to_string())
                } else {
                    Ok(format!("DATE({})", days_since_epoch))
                }
            },
            SafeVariant::Empty => Ok(String::new()),
            SafeVariant::Null => Ok("NULL".to_string()),
            SafeVariant::Error(hr) => Ok(format!("ERROR(0x{:08X})", hr.0)),
            _ => Err(SageError::ConversionError {
                from_type: self.type_name().to_string(),
                to_type: "String".to_string(),
                value: format!("{:?}", self),
            }),
        }
    }

    /// Convertit vers i32 si possible - AMÉLIORÉE v0.1.3
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn to_i32(&self) -> SageResult<i32> {
        match self {
            SafeVariant::I2(i) => Ok(*i as i32),
            SafeVariant::I4(i) => Ok(*i),
            SafeVariant::I8(i) => {
                if *i >= i32::MIN as i64 && *i <= i32::MAX as i64 {
                    Ok(*i as i32)
                } else {
                    Err(SageError::ConversionError {
                        from_type: "I8".to_string(),
                        to_type: "i32".to_string(),
                        value: format!("Valeur {} hors limites i32", i),
                    })
                }
            },
            SafeVariant::UI2(i) => Ok(*i as i32),
            SafeVariant::UI4(i) => {
                if *i <= i32::MAX as u32 {
                    Ok(*i as i32)
                } else {
                    Err(SageError::ConversionError {
                        from_type: "UI4".to_string(),
                        to_type: "i32".to_string(),
                        value: format!("Valeur {} hors limites i32", i),
                    })
                }
            },
            SafeVariant::R4(f) => Ok(*f as i32),
            SafeVariant::R8(f) => Ok(*f as i32),
            SafeVariant::Currency(c) => Ok((*c / 10000) as i32),
            SafeVariant::Bool(b) => Ok(if *b { 1 } else { 0 }),
            SafeVariant::BStr(s) => s.parse::<i32>().map_err(|_| SageError::ConversionError {
                from_type: "String".to_string(),
                to_type: "i32".to_string(),
                value: s.clone(),
            }),
            _ => Err(SageError::ConversionError {
                from_type: self.type_name().to_string(),
                to_type: "i32".to_string(),
                value: format!("{:?}", self),
            }),
        }
    }

    /// Convertit vers f64 si possible - AMÉLIORÉE v0.1.3  
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn to_f64(&self) -> SageResult<f64> {
        match self {
            SafeVariant::R4(f) => Ok(*f as f64),
            SafeVariant::R8(f) => Ok(*f),
            SafeVariant::I2(i) => Ok(*i as f64),
            SafeVariant::I4(i) => Ok(*i as f64),
            SafeVariant::I8(i) => Ok(*i as f64),
            SafeVariant::UI2(i) => Ok(*i as f64),
            SafeVariant::UI4(i) => Ok(*i as f64),
            SafeVariant::UI8(i) => Ok(*i as f64),
            SafeVariant::Currency(c) => Ok(*c as f64 / 10000.0),
            SafeVariant::Date(d) => Ok(*d),
            SafeVariant::Bool(b) => Ok(if *b { 1.0 } else { 0.0 }),
            SafeVariant::BStr(s) => s.parse::<f64>().map_err(|_| SageError::ConversionError {
                from_type: "String".to_string(),
                to_type: "f64".to_string(),
                value: s.clone(),
            }),
            _ => Err(SageError::ConversionError {
                from_type: self.type_name().to_string(),
                to_type: "f64".to_string(),
                value: format!("{:?}", self),
            }),
        }
    }

    /// Convertit vers bool si possible - AMÉLIORÉE v0.1.3
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn to_bool(&self) -> SageResult<bool> {
        match self {
            SafeVariant::Bool(b) => Ok(*b),
            SafeVariant::I2(i) => Ok(*i != 0),
            SafeVariant::I4(i) => Ok(*i != 0),
            SafeVariant::I8(i) => Ok(*i != 0),
            SafeVariant::UI2(i) => Ok(*i != 0),
            SafeVariant::UI4(i) => Ok(*i != 0),
            SafeVariant::UI8(i) => Ok(*i != 0),
            SafeVariant::R4(f) => Ok(*f != 0.0),
            SafeVariant::R8(f) => Ok(*f != 0.0),
            SafeVariant::Currency(c) => Ok(*c != 0),
            SafeVariant::Empty => Ok(false),
            SafeVariant::Null => Ok(false),
            SafeVariant::BStr(s) => match s.to_lowercase().as_str() {
                "true" | "1" | "yes" | "oui" | "-1" => Ok(true),
                "false" | "0" | "no" | "non" | "" => Ok(false),
                _ => Err(SageError::ConversionError {
                    from_type: "String".to_string(),
                    to_type: "bool".to_string(),
                    value: s.clone(),
                }),
            },
            _ => Err(SageError::ConversionError {
                from_type: self.type_name().to_string(),
                to_type: "bool".to_string(),
                value: format!("{:?}", self),
            }),
        }
    }

    /// Convertit vers Currency (format Sage) - NOUVEAU v0.1.3
    pub fn to_currency(&self) -> SageResult<f64> {
        match self {
            SafeVariant::Currency(c) => Ok(*c as f64 / 10000.0),
            SafeVariant::R4(f) => Ok(*f as f64),
            SafeVariant::R8(f) => Ok(*f),
            SafeVariant::I2(i) => Ok(*i as f64),
            SafeVariant::I4(i) => Ok(*i as f64),
            SafeVariant::I8(i) => Ok(*i as f64),
            SafeVariant::BStr(s) => s.parse::<f64>().map_err(|_| SageError::ConversionError {
                from_type: "String".to_string(),
                to_type: "Currency".to_string(),
                value: s.clone(),
            }),
            _ => Err(SageError::ConversionError {
                from_type: self.type_name().to_string(),
                to_type: "Currency".to_string(),
                value: format!("{:?}", self),
            }),
        }
    }

    /// Retourne le nom du type
    pub fn type_name(&self) -> &'static str {
        match self {
            SafeVariant::Empty => "Empty",
            SafeVariant::Null => "Null",
            SafeVariant::Bool(_) => "Bool",
            SafeVariant::I2(_) => "I2",
            SafeVariant::I4(_) => "I4",
            SafeVariant::R4(_) => "R4",
            SafeVariant::R8(_) => "R8",
            SafeVariant::Currency(_) => "Currency",
            SafeVariant::Date(_) => "Date",
            SafeVariant::BStr(_) => "BStr",
            SafeVariant::Error(_) => "Error",
            SafeVariant::I1(_) => "I1",
            SafeVariant::UI1(_) => "UI1",
            SafeVariant::UI2(_) => "UI2",
            SafeVariant::UI4(_) => "UI4",
            SafeVariant::I8(_) => "I8",
            SafeVariant::UI8(_) => "UI8",
        }
    }

    /// Vérifie si la variante est vide ou nulle
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn is_empty_or_null(&self) -> bool {
        matches!(self, SafeVariant::Empty | SafeVariant::Null)
    }

    // === MÉTHODES DE CRÉATION v0.1.3 ===
    
    /// Crée un SafeVariant String à partir d'une chaîne
    pub fn from_string<S: AsRef<str>>(s: S) -> Self {
        SafeVariant::BStr(s.as_ref().to_string())
    }
    
    /// Crée un SafeVariant Integer 32 bits
    pub fn from_i32(val: i32) -> Self {
        SafeVariant::I4(val)
    }
    
    /// Crée un SafeVariant Float 64 bits
    pub fn from_f64(val: f64) -> Self {
        SafeVariant::R8(val)
    }
    
    /// Crée un SafeVariant Boolean
    pub fn from_bool(val: bool) -> Self {
        SafeVariant::Bool(val)
    }
    
    /// Crée un SafeVariant Currency (pour les montants Sage)
    pub fn from_currency(val: f64) -> Self {
        // Conversion en currency (10000 unités = 1.0)
        SafeVariant::Currency((val * 10000.0) as i64)
    }
    
    /// Crée un SafeVariant Date à partir d'un timestamp f64
    pub fn from_date(val: f64) -> Self {
        SafeVariant::Date(val)
    }
    
    /// Crée un SafeVariant vide
    pub fn empty() -> Self {
        SafeVariant::Empty
    }
    
    /// Crée un SafeVariant null
    pub fn null() -> Self {
        SafeVariant::Null
    }
}

// Implémentations pratiques pour créer des SafeVariant
impl From<String> for SafeVariant {
    fn from(s: String) -> Self {
        SafeVariant::BStr(s)
    }
}

impl From<&str> for SafeVariant {
    fn from(s: &str) -> Self {
        SafeVariant::BStr(s.to_string())
    }
}

impl From<i32> for SafeVariant {
    fn from(i: i32) -> Self {
        SafeVariant::I4(i)
    }
}

impl From<f64> for SafeVariant {
    fn from(f: f64) -> Self {
        SafeVariant::R8(f)
    }
}

impl From<bool> for SafeVariant {
    fn from(b: bool) -> Self {
        SafeVariant::Bool(b)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_string_conversion() {
        let variant = SafeVariant::from("test");
        assert_eq!(variant.to_string().unwrap(), "test");
    }

    #[test]
    fn test_number_conversion() {
        let variant = SafeVariant::from(42);
        assert_eq!(variant.to_i32().unwrap(), 42);
        assert_eq!(variant.to_f64().unwrap(), 42.0);
    }

    #[test]
    fn test_bool_conversion() {
        let variant = SafeVariant::from(true);
        assert_eq!(variant.to_bool().unwrap(), true);
    }
}
