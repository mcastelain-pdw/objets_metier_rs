use std::fmt;
use windows::core::{Error as WindowsError, HRESULT};

/// Erreurs spécifiques à l'API Sage 100c
#[derive(Debug)]
#[allow(dead_code)] // Les variantes seront utilisées dans les futures versions
pub enum SageError {
    /// Erreur COM générique
    ComError { hresult: HRESULT, message: String },

    /// Erreur de connexion à la base de données
    ConnectionError {
        database_path: String,
        message: String,
    },

    /// Erreur d'authentification
    AuthenticationError { user: String, message: String },

    /// Erreur lors de l'appel d'une méthode
    MethodCallError {
        method_name: String,
        method_id: i32,
        message: String,
    },

    /// Erreur de conversion de données
    ConversionError {
        from_type: String,
        to_type: String,
        value: String,
    },

    /// Erreur de validation des données métier
    ValidationError {
        field: String,
        value: String,
        constraint: String,
    },

    /// Base de données non ouverte
    DatabaseNotOpen,

    /// CLSID non trouvé ou DLL non enregistrée
    ClassNotRegistered(String),

    /// Erreur de format de paramètre
    InvalidParameter {
        parameter: String,
        expected: String,
        received: String,
    },

    /// Opération non supportée
    UnsupportedOperation(String),

    /// Erreur interne inattendue
    InternalError(String),
}

impl fmt::Display for SageError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            SageError::ComError { hresult, message } => {
                write!(f, "Erreur COM (0x{:08X}): {}", hresult.0, message)
            }
            SageError::ConnectionError {
                database_path,
                message,
            } => {
                write!(f, "Erreur de connexion à '{}': {}", database_path, message)
            }
            SageError::AuthenticationError { user, message } => {
                write!(f, "Erreur d'authentification pour '{}': {}", user, message)
            }
            SageError::MethodCallError {
                method_name,
                method_id,
                message,
            } => {
                write!(
                    f,
                    "Erreur lors de l'appel de '{}' (ID: {}): {}",
                    method_name, method_id, message
                )
            }
            SageError::ConversionError {
                from_type,
                to_type,
                value,
            } => {
                write!(
                    f,
                    "Erreur de conversion de '{}' vers '{}' pour la valeur '{}'",
                    from_type, to_type, value
                )
            }
            SageError::ValidationError {
                field,
                value,
                constraint,
            } => {
                write!(
                    f,
                    "Erreur de validation du champ '{}' avec la valeur '{}': {}",
                    field, value, constraint
                )
            }
            SageError::DatabaseNotOpen => {
                write!(f, "Aucune base de données n'est ouverte")
            }
            SageError::ClassNotRegistered(clsid) => {
                write!(
                    f,
                    "Classe COM non enregistrée: {}. Exécutez 'regsvr32 objets100c.dll' en tant qu'administrateur",
                    clsid
                )
            }
            SageError::InvalidParameter {
                parameter,
                expected,
                received,
            } => {
                write!(
                    f,
                    "Paramètre '{}' invalide: attendu '{}', reçu '{}'",
                    parameter, expected, received
                )
            }
            SageError::UnsupportedOperation(op) => {
                write!(f, "Opération non supportée: {}", op)
            }
            SageError::InternalError(msg) => {
                write!(f, "Erreur interne: {}", msg)
            }
        }
    }
}

impl std::error::Error for SageError {}

impl From<WindowsError> for SageError {
    fn from(error: WindowsError) -> Self {
        let hresult = error.code();
        let message = error.message().to_string_lossy().to_owned();

        // Traiter les erreurs COM spécifiques
        match hresult.0 {
            val if val == 0x80040154u32 as i32 => {
                SageError::ClassNotRegistered("CLSID non trouvé".to_string())
            }
            val if val == 0x80070005u32 as i32 => SageError::ComError {
                hresult,
                message: "Accès refusé. Vérifiez les privilèges administrateur.".to_string(),
            },
            _ => SageError::ComError { hresult, message },
        }
    }
}

impl SageError {
    /// Crée une erreur de méthode COM
    pub fn method_call(method_name: &str, method_id: i32, message: &str) -> Self {
        SageError::MethodCallError {
            method_name: method_name.to_string(),
            method_id,
            message: message.to_string(),
        }
    }

    /// Crée une erreur de connexion
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn connection(database_path: &str, message: &str) -> Self {
        SageError::ConnectionError {
            database_path: database_path.to_string(),
            message: message.to_string(),
        }
    }

    /// Crée une erreur de validation
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn validation(field: &str, value: &str, constraint: &str) -> Self {
        SageError::ValidationError {
            field: field.to_string(),
            value: value.to_string(),
            constraint: constraint.to_string(),
        }
    }

    /// Crée une erreur de paramètre invalide
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn invalid_parameter(parameter: &str, expected: &str, received: &str) -> Self {
        SageError::InvalidParameter {
            parameter: parameter.to_string(),
            expected: expected.to_string(),
            received: received.to_string(),
        }
    }

    /// Vérifie si l'erreur est liée à une classe COM non enregistrée
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn is_class_not_registered(&self) -> bool {
        matches!(self, SageError::ClassNotRegistered(_))
    }

    /// Vérifie si l'erreur est liée à la connexion
    #[allow(dead_code)] // Sera utilisé dans v0.2.0
    pub fn is_connection_error(&self) -> bool {
        matches!(
            self,
            SageError::ConnectionError { .. } | SageError::DatabaseNotOpen
        )
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_error_display() {
        let error = SageError::validation("compte", "invalid", "doit être numérique");
        assert!(error.to_string().contains("Erreur de validation"));
    }

    #[test]
    fn test_error_classification() {
        let error = SageError::ClassNotRegistered("test".to_string());
        assert!(error.is_class_not_registered());

        let error = SageError::DatabaseNotOpen;
        assert!(error.is_connection_error());
    }
}
