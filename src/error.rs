use thiserror::Error;

/// The crate-wide result type used by RusDox.
pub type Result<T> = std::result::Result<T, DocxError>;

/// Error variants returned by RusDox operations.
#[derive(Debug, Error)]
pub enum DocxError {
    /// Wrapper for I/O failures.
    #[error("io error: {0}")]
    Io(#[from] std::io::Error),
    /// Wrapper for ZIP archive failures.
    #[error("zip error: {0}")]
    Zip(#[from] zip::result::ZipError),
    /// Wrapper for XML parser failures.
    #[error("xml error: {0}")]
    Xml(#[from] quick_xml::Error),
    /// Logical parsing or OOXML validation failures.
    #[error("parse error: {0}")]
    Parse(String),
    /// An input exceeded a documented resource ceiling.
    #[error("resource limit exceeded: {0}")]
    ResourceLimit(String),
    /// A completed verification whose generated outputs failed the parity contract.
    #[error("parity verification failed: {0}")]
    Parity(String),
    /// Rendering stopped after a cooperative cancellation request.
    #[error("render cancelled")]
    Cancelled,
}

impl DocxError {
    /// Creates a logical parse error with a custom message.
    pub(crate) fn parse(message: impl Into<String>) -> Self {
        Self::Parse(message.into())
    }

    /// Creates a resource-limit error with a custom explanation.
    pub(crate) fn resource_limit(message: impl Into<String>) -> Self {
        Self::ResourceLimit(message.into())
    }
}

impl From<quick_xml::escape::EscapeError> for DocxError {
    fn from(error: quick_xml::escape::EscapeError) -> Self {
        Self::Parse(error.to_string())
    }
}
