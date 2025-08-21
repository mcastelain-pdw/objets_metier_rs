pub mod dispatch;
pub mod instance;
pub mod safe_string;
pub mod variant;

pub use dispatch::SafeDispatch;
pub use instance::ComInstance;
pub use variant::SafeVariant;
// SafeString sera utilisé dans les futures versions
#[allow(unused_imports)]
pub use safe_string::SafeString;
