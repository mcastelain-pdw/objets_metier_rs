use crate::errors::{SageError, SageResult};
use windows::{core::*, Win32::{System::{Com::IDispatch, Variant::*}}};

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
    Dispatch(IDispatch), // NOUVEAU: pour stocker les objets COM
    Unknown(String),     // NOUVEAU: pour les types non reconnus
}

impl SafeVariant {
    /// Crée un SafeVariant à partir d'un VARIANT Windows - PRAGMATIQUE v0.1.3
    pub fn from_variant(variant: VARIANT) -> SageResult<Self> {
        unsafe {
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
                
                VT_DISPATCH => {
                    // Gérer les objets COM IDispatch - CORRIGÉ
                    let dispatch_opt = &variant.Anonymous.Anonymous.Anonymous.pdispVal;
                    if let Some(dispatch) = dispatch_opt.as_ref() {
                        Ok(SafeVariant::Dispatch(dispatch.clone()))
                    } else {
                        Ok(SafeVariant::Unknown("IDispatch null".to_string()))
                    }
                },
                
                VT_UNKNOWN => {
                    // Gérer les objets COM IUnknown
                    Ok(SafeVariant::Unknown(format!("IUnknown object (VT_{})", vt.0)))
                },
                
                _ => {
                    // Pour les types non supportés, retourner une description
                    Ok(SafeVariant::Unknown(format!("Type VARIANT VT_{} - Conversion en v0.1.4", vt.0)))
                }
            }
        }
    }

    /// Convertit vers un VARIANT Windows - VERSION FONCTIONNELLE v0.1.3+
    pub fn to_variant(&self) -> SageResult<VARIANT> {
        // Solution ultime: Utiliser VariantInit qui retourne directement une VARIANT
        // puis utiliser la méthode brutale pour contourner ManuallyDrop
        
        match self {
            SafeVariant::Empty => {
                unsafe { Ok(windows::Win32::System::Variant::VariantInit()) }
            },
            
            SafeVariant::BStr(s) => {
                let bstr = BSTR::from(s.as_str());
                
                unsafe {
                    let mut variant = windows::Win32::System::Variant::VariantInit();
                    
                    // Approche directe avec transmute pour contourner ManuallyDrop
                    let variant_ptr = &mut variant as *mut VARIANT as *mut u8;
                    
                    // Offset vers vt field (généralement à l'offset 0)
                    let vt_ptr = variant_ptr as *mut u16;
                    *vt_ptr = VT_BSTR.0;
                    
                    // Offset vers le champ bstrVal (généralement après vt + wReserved fields)
                    let bstr_ptr = variant_ptr.add(8) as *mut BSTR; // 8 = 2 (vt) + 2 (wReserved1) + 2 (wReserved2) + 2 (wReserved3)
                    *bstr_ptr = bstr;
                    
                    Ok(variant)
                }
            },
            
            SafeVariant::Bool(val) => {
                unsafe {
                    let mut variant = windows::Win32::System::Variant::VariantInit();
                    
                    let variant_ptr = &mut variant as *mut VARIANT as *mut u8;
                    let vt_ptr = variant_ptr as *mut u16;
                    *vt_ptr = VT_BOOL.0;
                    
                    let bool_ptr = variant_ptr.add(8) as *mut i16;
                    *bool_ptr = if *val { -1 } else { 0 }; // VARIANT_BOOL convention
                    
                    Ok(variant)
                }
            },
            
            SafeVariant::I4(val) => {
                unsafe {
                    let mut variant = windows::Win32::System::Variant::VariantInit();
                    
                    let variant_ptr = &mut variant as *mut VARIANT as *mut u8;
                    let vt_ptr = variant_ptr as *mut u16;
                    *vt_ptr = VT_I4.0;
                    
                    let i4_ptr = variant_ptr.add(8) as *mut i32;
                    *i4_ptr = *val;
                    
                    Ok(variant)
                }
            },
            
            SafeVariant::R8(val) => {
                unsafe {
                    let mut variant = windows::Win32::System::Variant::VariantInit();
                    
                    let variant_ptr = &mut variant as *mut VARIANT as *mut u8;
                    let vt_ptr = variant_ptr as *mut u16;
                    *vt_ptr = VT_R8.0;
                    
                    let r8_ptr = variant_ptr.add(8) as *mut f64;
                    *r8_ptr = *val;
                    
                    Ok(variant)
                }
            },
            
            SafeVariant::Dispatch(dispatch) => {
                unsafe {
                    let mut variant = windows::Win32::System::Variant::VariantInit();
                    
                    let variant_ptr = &mut variant as *mut VARIANT as *mut u8;
                    let vt_ptr = variant_ptr as *mut u16;
                    *vt_ptr = VT_DISPATCH.0;
                    
                    let dispatch_ptr = variant_ptr.add(8) as *mut Option<IDispatch>;
                    *dispatch_ptr = Some(dispatch.clone());
                    
                    Ok(variant)
                }
            },
            
            // Conversions automatiques
            SafeVariant::I2(val) => {
                SafeVariant::I4(*val as i32).to_variant()
            },
            
            SafeVariant::R4(val) => {
                SafeVariant::R8(*val as f64).to_variant()
            },
            
            _ => {
                // Pour les autres types, retourner VT_EMPTY
                println!("VARIANT conversion: Type non implémenté: {:?}", self);
                unsafe { Ok(windows::Win32::System::Variant::VariantInit()) }
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
            SafeVariant::Dispatch(_) => Ok("IDispatch object".to_string()),
            SafeVariant::Unknown(desc) => Ok(desc.clone()),
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
            SafeVariant::Dispatch(_) => "IDispatch",
            SafeVariant::Unknown(_) => "Unknown",
        }
    }

    /// Vérifie si la variante est vide ou nulle
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn is_empty_or_null(&self) -> bool {
        matches!(self, SafeVariant::Empty | SafeVariant::Null)
    }

    /// Convertit le VARIANT en interface IDispatch si possible - CORRIGÉ v0.1.3
    pub fn to_dispatch(&self) -> SageResult<IDispatch> {
        match self {
            SafeVariant::Dispatch(dispatch) => {
                // Cloner l'interface pour éviter les problèmes de durée de vie
                Ok(dispatch.clone())
            },
            SafeVariant::Unknown(desc) if desc.contains("IDispatch") => {
                Err(SageError::ConversionError {
                    from_type: "Unknown".to_string(),
                    to_type: "IDispatch".to_string(),
                    value: "IDispatch null détecté".to_string(),
                })
            },
            _ => {
                Err(SageError::ConversionError {
                    from_type: self.type_name().to_string(),
                    to_type: "IDispatch".to_string(),
                    value: format!("Type {} ne peut pas être converti en IDispatch", self.type_name()),
                })
            }
        }
    }

    /// Vérifie si le VARIANT contient une interface COM - CORRIGÉ v0.1.3
    pub fn is_object(&self) -> bool {
        matches!(self, SafeVariant::Dispatch(_) | SafeVariant::Unknown(_))
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

    /// Crée un SafeVariant à partir d'un IDispatch
    pub fn from_dispatch(dispatch: IDispatch) -> Self {
        SafeVariant::Dispatch(dispatch)
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

impl From<IDispatch> for SafeVariant {
    fn from(dispatch: IDispatch) -> Self {
        SafeVariant::Dispatch(dispatch)
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

    #[test]
    fn test_dispatch_object() {
        // Ce test nécessiterait un vrai IDispatch, donc on teste juste le type
        let variant = SafeVariant::Unknown("IDispatch object".to_string());
        assert!(variant.is_object());
        assert_eq!(variant.type_name(), "Unknown");
    }
}
