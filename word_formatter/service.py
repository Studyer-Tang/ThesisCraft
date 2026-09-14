"""Batch execution and structured results independent of any user interface."""

from dataclasses import dataclass, field
from .conversion import LegacyConversionUnavailable


@dataclass
class BatchResult:
    outputs: list = field(default_factory=list)
    failures: list = field(default_factory=list)
    skipped: list = field(default_factory=list)
    cancelled: bool = False

    @property
    def exit_code(self):
        if self.failures:
            return 1
        if self.skipped or self.cancelled:
            return 2
        return 0


def run_jobs(processor, jobs, *, overwrite=False, cancel_event=None, on_result=None):
    result = BatchResult()
    for index, job in enumerate(jobs, 1):
        if cancel_event is not None and cancel_event.is_set():
            result.cancelled = True
            break
        try:
            processor.format_document(
                str(job.source), str(job.output), overwrite=overwrite
            )
            result.outputs.append(str(job.output))
            state, detail = "success", str(job.output)
        except LegacyConversionUnavailable as exc:
            state, detail = "skipped", str(exc)
            result.skipped.append({"source": str(job.source), "error": detail})
        except Exception as exc:
            state, detail = "failed", str(exc)
            result.failures.append({"source": str(job.source), "error": detail})
        if on_result:
            on_result(index, len(jobs), job, state, detail)
    return result
