pub mod instance;
pub mod dispatch;
pub mod variant;
pub mod safe_string;


pub use instance::{ComInstance, MemberInfo, MemberType};
pub use dispatch::{SafeDispatch, FromDispatch};
pub use variant::SafeVariant;
#[allow(unused_imports)] // Sera utilisé dans les futures versions
pub use safe_string::SafeString;

