pub mod sage_error;

pub use sage_error::SageError;
pub type SageResult<T> = Result<T, SageError>;
