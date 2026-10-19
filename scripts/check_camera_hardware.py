"""Read one camera frame in memory. Never saves, displays or logs image pixels."""

import json

from PySide6.QtCore import QCoreApplication, QTimer

from document_excel.infrastructure.camera import CameraService


def main():
    app = QCoreApplication([])
    service = CameraService(indices=range(2))
    report = {"camera_frame_saved": False, "stream_opened": False, "full_resolution_capture": False}

    def finish():
        service.close()
        if service.busy:
            QTimer.singleShot(100, finish)
        else:
            report["device_released"] = True
            app.quit()

    def found(devices):
        report["available_devices"] = len(devices)
        if devices:
            service.start(devices[0])
        else:
            finish()

    def ready(resolutions):
        report["stream_opened"] = True
        report["verified_resolutions"] = resolutions
        service.capture()

    def captured(array):
        report["captured_size"] = [array.shape[1], array.shape[0]]
        report["full_resolution_capture"] = True
        finish()

    def failed(message):
        report["technical_error"] = message
        finish()

    service.devices_found.connect(found)
    service.ready.connect(ready)
    service.photograph.connect(captured)
    service.error.connect(failed)
    QTimer.singleShot(0, service.discover)
    QTimer.singleShot(30000, finish)
    QTimer.singleShot(40000, app.quit)
    app.exec()
    print(json.dumps(report, ensure_ascii=True))
    return 0 if report["full_resolution_capture"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
