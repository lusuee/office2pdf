import socket
import sys
import traceback

from main import (
    DEFAULT_HOST,
    DEFAULT_PORT,
    cleanup_office_applications,
    create_http_server,
    logger,
)

if sys.platform == "win32":
    import servicemanager
    import win32event
    import win32service
    import win32serviceutil

    ServiceFrameworkBase = win32serviceutil.ServiceFramework
else:
    servicemanager = None
    win32event = None
    win32service = None
    win32serviceutil = None
    ServiceFrameworkBase = object


class Office2PdfService(ServiceFrameworkBase):
    _svc_name_ = "office2pdf"
    _svc_display_name_ = "Office2PDF"
    _svc_description_ = "Convert Office documents to PDF over HTTP."

    def __init__(self, args):
        if sys.platform != "win32":
            raise RuntimeError("Windows service mode is only supported on Windows.")
        super().__init__(args)
        self.stop_event = win32event.CreateEvent(None, 0, 0, None)
        self.server = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        logger.info("Service stop requested.")

        if self.server is not None:
            try:
                self.server.close()
            except Exception as exc:
                logger.warning("Error stopping HTTP server: %s", exc)

        cleanup_office_applications()
        win32event.SetEvent(self.stop_event)

    def SvcDoRun(self):
        port = DEFAULT_PORT
        host = DEFAULT_HOST

        servicemanager.LogInfoMsg(f"{self._svc_name_} starting on {host}:{port}")
        logger.info("Windows service starting on %s:%s", host, port)

        try:
            socket.setdefaulttimeout(60)
            self.server = create_http_server(host=host, port=port)
            self.server.run()
        except Exception as exc:
            error_text = "".join(traceback.format_exception(exc))
            logger.exception("Windows service failed to start.")
            servicemanager.LogErrorMsg(error_text)
            raise
        finally:
            cleanup_office_applications()
            logger.info("Windows service stopped.")
            servicemanager.LogInfoMsg(f"{self._svc_name_} stopped")


if __name__ == "__main__":
    if sys.platform != "win32":
        raise RuntimeError("service.py can only run on Windows.")
    win32serviceutil.HandleCommandLine(Office2PdfService)
