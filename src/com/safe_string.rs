use crate::errors::{SageError, SageResult};
use windows::core::BSTR;

/// Wrapper sûr pour les chaînes BSTR COM
#[derive(Debug, Clone)]
#[allow(dead_code)] // Sera utilisé dans v0.2.0
pub struct SafeString {
    inner: String,
}

#[allow(dead_code)] // Toutes les méthodes seront utilisées dans v0.2.0
impl SafeString {
    /// Crée une nouvelle SafeString
    pub fn new(s: &str) -> Self {
        SafeString {
            inner: s.to_string(),
        }
    }

    /// Crée une SafeString à partir d'un BSTR Windows
    pub fn from_bstr(bstr: BSTR) -> Self {
        SafeString {
            inner: bstr.to_string(),
        }
    }

    /// Convertit vers BSTR Windows
    pub fn to_bstr(&self) -> BSTR {
        BSTR::from(self.inner.as_str())
    }

    /// Obtient la chaîne Rust
    pub fn as_str(&self) -> &str {
        &self.inner
    }

    /// Obtient la chaîne Rust
    pub fn into_string(self) -> String {
        self.inner
    }

    /// Vérifie si la chaîne est vide
    pub fn is_empty(&self) -> bool {
        self.inner.is_empty()
    }

    /// Obtient la longueur
    pub fn len(&self) -> usize {
        self.inner.len()
    }

    /// Valide que la chaîne respecte un format donné
    pub fn validate_format(&self, pattern: &str) -> SageResult<()> {
        // Validation basique pour les formats Sage communs
        match pattern {
            "numeric" => {
                if self
                    .inner
                    .chars()
                    .all(|c| c.is_ascii_digit() || c == '.' || c == '-')
                {
                    Ok(())
                } else {
                    Err(SageError::validation(
                        "string",
                        &self.inner,
                        "doit être numérique",
                    ))
                }
            }
            "alphanumeric" => {
                if self.inner.chars().all(|c| c.is_alphanumeric()) {
                    Ok(())
                } else {
                    Err(SageError::validation(
                        "string",
                        &self.inner,
                        "doit être alphanumérique",
                    ))
                }
            }
            "date" => {
                // Format DD/MM/YYYY ou YYYY-MM-DD
                if self.inner.len() == 10
                    && (self.inner.chars().nth(2) == Some('/')
                        && self.inner.chars().nth(5) == Some('/')
                        || self.inner.chars().nth(4) == Some('-')
                            && self.inner.chars().nth(7) == Some('-'))
                {
                    Ok(())
                } else {
                    Err(SageError::validation(
                        "string",
                        &self.inner,
                        "doit être au format DD/MM/YYYY ou YYYY-MM-DD",
                    ))
                }
            }
            "account_code" => {
                // Code de compte Sage (généralement 6-8 chiffres)
                if self.inner.len() >= 3
                    && self.inner.len() <= 8
                    && self.inner.chars().all(|c| c.is_ascii_digit())
                {
                    Ok(())
                } else {
                    Err(SageError::validation(
                        "string",
                        &self.inner,
                        "doit être un code de compte valide (3-8 chiffres)",
                    ))
                }
            }
            "journal_code" => {
                // Code de journal (généralement 2-3 caractères alphanumériques)
                if self.inner.len() >= 2
                    && self.inner.len() <= 3
                    && self.inner.chars().all(|c| c.is_alphanumeric())
                {
                    Ok(())
                } else {
                    Err(SageError::validation(
                        "string",
                        &self.inner,
                        "doit être un code de journal valide (2-3 caractères alphanumériques)",
                    ))
                }
            }
            _ => Ok(()), // Format non reconnu, pas de validation
        }
    }

    /// Normalise la chaîne selon les conventions Sage
    pub fn normalize_sage_format(&mut self, format_type: &str) {
        match format_type {
            "account_code" => {
                // Supprimer les espaces et convertir en majuscules
                self.inner = self.inner.trim().to_uppercase();
                // Pad avec des zéros à gauche si nécessaire (pour certains formats)
                if self.inner.len() < 6 && self.inner.chars().all(|c| c.is_ascii_digit()) {
                    self.inner = format!("{:0>6}", self.inner);
                }
            }
            "journal_code" => {
                // Convertir en majuscules et supprimer les espaces
                self.inner = self.inner.trim().to_uppercase();
            }
            "trim" => {
                self.inner = self.inner.trim().to_string();
            }
            "upper" => {
                self.inner = self.inner.to_uppercase();
            }
            "lower" => {
                self.inner = self.inner.to_lowercase();
            }
            _ => {} // Format non reconnu, pas de normalisation
        }
    }

    /// Encode pour être sûr dans les appels COM
    pub fn encode_for_com(&self) -> Vec<u16> {
        self.inner
            .encode_utf16()
            .chain(std::iter::once(0)) // Null terminator
            .collect()
    }
}

impl From<String> for SafeString {
    fn from(s: String) -> Self {
        SafeString { inner: s }
    }
}

impl From<&str> for SafeString {
    fn from(s: &str) -> Self {
        SafeString::new(s)
    }
}

impl From<BSTR> for SafeString {
    fn from(bstr: BSTR) -> Self {
        SafeString::from_bstr(bstr)
    }
}

impl std::fmt::Display for SafeString {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        write!(f, "{}", self.inner)
    }
}

impl PartialEq for SafeString {
    fn eq(&self, other: &Self) -> bool {
        self.inner == other.inner
    }
}

impl PartialEq<str> for SafeString {
    fn eq(&self, other: &str) -> bool {
        self.inner == other
    }
}

impl PartialEq<String> for SafeString {
    fn eq(&self, other: &String) -> bool {
        &self.inner == other
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_safe_string_creation() {
        let s = SafeString::new("test");
        assert_eq!(s.as_str(), "test");
    }

    #[test]
    fn test_format_validation() {
        let s = SafeString::new("123456");
        assert!(s.validate_format("numeric").is_ok());
        assert!(s.validate_format("account_code").is_ok());

        let s = SafeString::new("abc123");
        assert!(s.validate_format("alphanumeric").is_ok());
        assert!(s.validate_format("numeric").is_err());
    }

    #[test]
    fn test_normalization() {
        let mut s = SafeString::new("  vt  ");
        s.normalize_sage_format("journal_code");
        assert_eq!(s.as_str(), "VT");

        let mut s = SafeString::new("123");
        s.normalize_sage_format("account_code");
        assert_eq!(s.as_str(), "000123");
    }
}
