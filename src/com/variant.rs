use crate::errors::{SageError, SageResult};
use windows::{Win32::System::Variant::*, core::*};

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
    /// Crée un SafeVariant à partir d'un VARIANT Windows
    pub fn from_variant(_variant: VARIANT) -> Self {
        // Implémentation simplifiée pour le moment
        // Dans une version complète, on utiliserait les fonctions de conversion Windows
        SafeVariant::BStr("Conversion VARIANT non implémentée".to_string())
    }

    /// Convertit vers un VARIANT Windows (implémentation basique)
    pub fn to_variant(&self) -> VARIANT {
        // Pour le moment, retourne un VARIANT par défaut
        // L'implémentation complète nécessiterait une gestion plus complexe
        // des unions ManuallyDrop de Windows
        VARIANT::default()
    }

    /// Convertit vers String si possible
    pub fn to_string(&self) -> SageResult<String> {
        match self {
            SafeVariant::BStr(s) => Ok(s.clone()),
            SafeVariant::I2(i) => Ok(i.to_string()),
            SafeVariant::I4(i) => Ok(i.to_string()),
            SafeVariant::R4(f) => Ok(f.to_string()),
            SafeVariant::R8(f) => Ok(f.to_string()),
            SafeVariant::Bool(b) => Ok(b.to_string()),
            SafeVariant::Empty => Ok(String::new()),
            SafeVariant::Null => Ok("NULL".to_string()),
            _ => Err(SageError::ConversionError {
                from_type: self.type_name().to_string(),
                to_type: "String".to_string(),
                value: format!("{:?}", self),
            }),
        }
    }

    /// Convertit vers i32 si possible
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn to_i32(&self) -> SageResult<i32> {
        match self {
            SafeVariant::I2(i) => Ok(*i as i32),
            SafeVariant::I4(i) => Ok(*i),
            SafeVariant::UI2(i) => Ok(*i as i32),
            SafeVariant::UI4(i) => Ok(*i as i32),
            SafeVariant::R4(f) => Ok(*f as i32),
            SafeVariant::R8(f) => Ok(*f as i32),
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

    /// Convertit vers f64 si possible
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn to_f64(&self) -> SageResult<f64> {
        match self {
            SafeVariant::R4(f) => Ok(*f as f64),
            SafeVariant::R8(f) => Ok(*f),
            SafeVariant::I2(i) => Ok(*i as f64),
            SafeVariant::I4(i) => Ok(*i as f64),
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

    /// Convertit vers bool si possible
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn to_bool(&self) -> SageResult<bool> {
        match self {
            SafeVariant::Bool(b) => Ok(*b),
            SafeVariant::I2(i) => Ok(*i != 0),
            SafeVariant::I4(i) => Ok(*i != 0),
            SafeVariant::BStr(s) => match s.to_lowercase().as_str() {
                "true" | "1" | "yes" | "oui" => Ok(true),
                "false" | "0" | "no" | "non" => Ok(false),
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
