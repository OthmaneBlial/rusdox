//! Bounded concurrent rendering for service and batch workloads.

use std::collections::VecDeque;
use std::sync::{Arc, Mutex};
use std::thread;

use crate::{
    CancellationToken, DocxError, RenderRequest, RenderSource, RenderedDocument, Renderer, Result,
};

/// Capacity limits applied before a batch starts.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BatchLimits {
    /// Maximum requests accepted by one batch call.
    pub max_jobs: usize,
    /// Maximum worker threads rendering at the same time.
    pub max_concurrency: usize,
    /// Maximum source bytes accepted for one request.
    pub max_source_bytes_per_job: u64,
    /// Maximum aggregate source bytes accepted by one batch.
    pub max_total_source_bytes: u64,
}

impl Default for BatchLimits {
    fn default() -> Self {
        Self {
            max_jobs: 256,
            max_concurrency: 4,
            max_source_bytes_per_job: 8 * 1024 * 1024,
            max_total_source_bytes: 64 * 1024 * 1024,
        }
    }
}

impl BatchLimits {
    fn validate(self) -> Result<()> {
        if self.max_jobs == 0
            || self.max_concurrency == 0
            || self.max_source_bytes_per_job == 0
            || self.max_total_source_bytes == 0
        {
            return Err(DocxError::ResourceLimit(
                "batch limits must all be greater than zero".to_string(),
            ));
        }
        if self.max_source_bytes_per_job > self.max_total_source_bytes {
            return Err(DocxError::ResourceLimit(
                "max_source_bytes_per_job cannot exceed max_total_source_bytes".to_string(),
            ));
        }
        Ok(())
    }
}

/// One caller-identified rendering request in a batch.
#[derive(Debug, Clone)]
pub struct BatchRequest {
    /// Caller-selected identifier preserved in the result.
    pub id: String,
    /// Transport-neutral renderer request.
    pub request: RenderRequest,
}

/// Terminal state for one batch item.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum BatchStatus {
    /// Rendering completed and artifacts are available.
    Completed,
    /// Rendering failed with a validation, parsing, I/O, or layout error.
    Failed,
    /// Work did not complete because cancellation was requested.
    Cancelled,
}

/// Ordered result for one batch item.
#[derive(Debug)]
pub struct BatchItemResult {
    /// Identifier copied from [`BatchRequest::id`].
    pub id: String,
    /// Terminal item status.
    pub status: BatchStatus,
    /// In-memory artifacts for a completed item.
    pub output: Option<RenderedDocument>,
    /// Human-readable error for a failed or cancelled item.
    pub error: Option<String>,
}

/// Bounded worker pool over any thread-safe [`Renderer`] implementation.
#[derive(Clone)]
pub struct BatchRenderer {
    renderer: Arc<dyn Renderer>,
    limits: BatchLimits,
}

impl std::fmt::Debug for BatchRenderer {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("BatchRenderer")
            .field("limits", &self.limits)
            .finish_non_exhaustive()
    }
}

impl BatchRenderer {
    /// Creates a bounded batch renderer from a shared rendering backend.
    pub fn new(renderer: Arc<dyn Renderer>, limits: BatchLimits) -> Result<Self> {
        limits.validate()?;
        Ok(Self { renderer, limits })
    }

    /// Renders requests concurrently while preserving their original order.
    ///
    /// Capacity and source-byte limits are checked before any worker starts.
    /// Cancellation prevents queued work and is observed by native rendering
    /// between major stages. Already completed items remain available.
    pub fn render(
        &self,
        requests: Vec<BatchRequest>,
        cancellation: &CancellationToken,
    ) -> Result<Vec<BatchItemResult>> {
        self.preflight(&requests)?;
        if requests.is_empty() {
            return Ok(Vec::new());
        }

        let count = requests.len();
        let queue = Arc::new(Mutex::new(
            requests.into_iter().enumerate().collect::<VecDeque<_>>(),
        ));
        let results = Arc::new(Mutex::new(
            std::iter::repeat_with(|| None)
                .take(count)
                .collect::<Vec<Option<BatchItemResult>>>(),
        ));
        let workers = self.limits.max_concurrency.min(count);

        thread::scope(|scope| {
            for _ in 0..workers {
                let queue = Arc::clone(&queue);
                let results = Arc::clone(&results);
                let renderer = Arc::clone(&self.renderer);
                let cancellation = cancellation.clone();
                scope.spawn(move || loop {
                    let Some((index, item)) =
                        queue.lock().expect("batch queue poisoned").pop_front()
                    else {
                        break;
                    };
                    let result = if cancellation.is_cancelled() {
                        cancelled(item.id)
                    } else {
                        match renderer.render_cancellable(&item.request, &cancellation) {
                            Ok(output) => BatchItemResult {
                                id: item.id,
                                status: BatchStatus::Completed,
                                output: Some(output),
                                error: None,
                            },
                            Err(DocxError::Cancelled) => cancelled(item.id),
                            Err(error) => BatchItemResult {
                                id: item.id,
                                status: BatchStatus::Failed,
                                output: None,
                                error: Some(error.to_string()),
                            },
                        }
                    };
                    results.lock().expect("batch results poisoned")[index] = Some(result);
                });
            }
        });

        let mut locked = results.lock().expect("batch results poisoned");
        Ok(locked
            .iter_mut()
            .map(|result| result.take().expect("every queued item has a result"))
            .collect())
    }

    fn preflight(&self, requests: &[BatchRequest]) -> Result<()> {
        if requests.len() > self.limits.max_jobs {
            return Err(DocxError::ResourceLimit(format!(
                "batch contains {} jobs; limit is {}",
                requests.len(),
                self.limits.max_jobs
            )));
        }
        let mut total = 0_u64;
        for item in requests {
            let bytes = source_bytes(&item.request.source)?;
            if bytes > self.limits.max_source_bytes_per_job {
                return Err(DocxError::ResourceLimit(format!(
                    "batch item '{}' source is {bytes} bytes; per-job limit is {}",
                    item.id, self.limits.max_source_bytes_per_job
                )));
            }
            total = total.checked_add(bytes).ok_or_else(|| {
                DocxError::ResourceLimit("batch source byte total overflowed".to_string())
            })?;
        }
        if total > self.limits.max_total_source_bytes {
            return Err(DocxError::ResourceLimit(format!(
                "batch source total is {total} bytes; limit is {}",
                self.limits.max_total_source_bytes
            )));
        }
        Ok(())
    }
}

fn source_bytes(source: &RenderSource) -> Result<u64> {
    match source {
        RenderSource::Inline { content, .. } => Ok(content.len() as u64),
        RenderSource::Path { path } => Ok(std::fs::metadata(path)?.len()),
    }
}

fn cancelled(id: String) -> BatchItemResult {
    BatchItemResult {
        id,
        status: BatchStatus::Cancelled,
        output: None,
        error: Some(DocxError::Cancelled.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::{AtomicUsize, Ordering};
    use std::sync::Arc;
    use std::time::Duration;

    use super::{BatchLimits, BatchRenderer, BatchRequest, BatchStatus};
    use crate::config::RusdoxConfig;
    use crate::{
        CancellationToken, NativeRenderer, RenderRequest, RenderSource, RenderedDocument, Renderer,
        RendererValidation, SpecFormat, RENDERER_API_VERSION,
    };

    #[derive(Debug)]
    struct TrackingRenderer {
        active: AtomicUsize,
        maximum: AtomicUsize,
        delay: Duration,
    }

    impl Renderer for TrackingRenderer {
        fn validate(&self, _request: &RenderRequest) -> crate::Result<RendererValidation> {
            Ok(RendererValidation {
                valid: true,
                diagnostics: Vec::new(),
                parse_duration: Duration::ZERO,
                validation_duration: Duration::ZERO,
            })
        }

        fn render(&self, _request: &RenderRequest) -> crate::Result<RenderedDocument> {
            let active = self.active.fetch_add(1, Ordering::SeqCst) + 1;
            self.maximum.fetch_max(active, Ordering::SeqCst);
            std::thread::sleep(self.delay);
            self.active.fetch_sub(1, Ordering::SeqCst);
            Ok(RenderedDocument {
                docx: b"PK".to_vec(),
                pdf: None,
                diagnostics: Vec::new(),
                parse_duration: Duration::ZERO,
                validation_duration: Duration::ZERO,
                compose_duration: Duration::ZERO,
                docx_duration: Duration::ZERO,
                pdf_duration: Duration::ZERO,
            })
        }
    }

    fn request(id: usize, padding: usize) -> BatchRequest {
        BatchRequest {
            id: format!("job-{id}"),
            request: RenderRequest {
                renderer_api_version: RENDERER_API_VERSION,
                source: RenderSource::Inline {
                    format: SpecFormat::Yaml,
                    content: format!("version: 1\nblocks: []\n# {}", "x".repeat(padding)),
                },
                emit_pdf: false,
            },
        }
    }

    #[test]
    fn real_renderer_completes_a_bounded_load_batch() {
        let renderer = BatchRenderer::new(
            Arc::new(NativeRenderer::new(RusdoxConfig::default())),
            BatchLimits {
                max_jobs: 32,
                max_concurrency: 4,
                ..BatchLimits::default()
            },
        )
        .expect("batch renderer");
        let results = renderer
            .render(
                (0..16).map(|index| request(index, 0)).collect(),
                &CancellationToken::new(),
            )
            .expect("batch");
        assert_eq!(results.len(), 16);
        assert!(results.iter().all(|item| {
            item.status == BatchStatus::Completed
                && item
                    .output
                    .as_ref()
                    .is_some_and(|output| output.docx.starts_with(b"PK"))
        }));
    }

    #[test]
    fn concurrency_never_exceeds_the_worker_budget() {
        let backend = Arc::new(TrackingRenderer {
            active: AtomicUsize::new(0),
            maximum: AtomicUsize::new(0),
            delay: Duration::from_millis(15),
        });
        let renderer = BatchRenderer::new(
            backend.clone(),
            BatchLimits {
                max_concurrency: 3,
                ..BatchLimits::default()
            },
        )
        .expect("batch renderer");
        renderer
            .render(
                (0..12).map(|index| request(index, 0)).collect(),
                &CancellationToken::new(),
            )
            .expect("batch");
        let maximum = backend.maximum.load(Ordering::SeqCst);
        assert!((2..=3).contains(&maximum));
    }

    #[test]
    fn cancellation_stops_queued_work_without_losing_order() {
        let renderer = BatchRenderer::new(
            Arc::new(TrackingRenderer {
                active: AtomicUsize::new(0),
                maximum: AtomicUsize::new(0),
                delay: Duration::from_millis(25),
            }),
            BatchLimits {
                max_concurrency: 2,
                ..BatchLimits::default()
            },
        )
        .expect("batch renderer");
        let token = CancellationToken::new();
        let worker_token = token.clone();
        let handle = std::thread::spawn(move || {
            renderer.render(
                (0..20).map(|index| request(index, 0)).collect(),
                &worker_token,
            )
        });
        std::thread::sleep(Duration::from_millis(5));
        token.cancel();
        let results = handle.join().expect("thread").expect("batch");
        assert_eq!(results.len(), 20);
        assert_eq!(results[0].id, "job-0");
        assert_eq!(results[19].id, "job-19");
        assert!(results
            .iter()
            .any(|item| item.status == BatchStatus::Cancelled));
    }

    #[test]
    fn aggregate_source_memory_budget_fails_before_workers_start() {
        let renderer = BatchRenderer::new(
            Arc::new(NativeRenderer::new(RusdoxConfig::default())),
            BatchLimits {
                max_source_bytes_per_job: 100,
                max_total_source_bytes: 150,
                ..BatchLimits::default()
            },
        )
        .expect("batch renderer");
        let error = renderer
            .render(
                vec![request(0, 60), request(1, 60)],
                &CancellationToken::new(),
            )
            .expect_err("aggregate limit");
        assert!(error.to_string().contains("batch source total"));
    }
}
