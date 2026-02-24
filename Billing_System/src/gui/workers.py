"""Background worker for running service calls off the UI thread."""

import sys
import traceback

from PySide6.QtCore import QObject, QRunnable, Signal, Slot


class WorkerSignals(QObject):
    """Signals emitted by ServiceWorker."""

    finished = Signal(object)  # ServiceResult on success
    error = Signal(tuple)      # (exc_type, exc_value, traceback_str)


class ServiceWorker(QRunnable):
    """Run a service function in the thread pool.

    Parameters
    ----------
    fn : callable
        The service function to call (e.g. doc_service.generate_daily).
    *args, **kwargs
        Forwarded to *fn*.
    """

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.signals = WorkerSignals()
        self.setAutoDelete(True)

    @Slot()
    def run(self):
        # COM initialisation — needed when Word / Outlook COM objects are
        # created from a background thread (QThreadPool worker).
        import pythoncom
        pythoncom.CoInitialize()
        try:
            result = self.fn(*self.args, **self.kwargs)
            self.signals.finished.emit(result)
        except Exception:
            self.signals.error.emit(
                (type(sys.exc_info()[1]),
                 sys.exc_info()[1],
                 traceback.format_exc())
            )
        finally:
            pythoncom.CoUninitialize()
