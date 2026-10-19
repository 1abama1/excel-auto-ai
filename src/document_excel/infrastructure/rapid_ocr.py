"""RapidOCR modules with local ONNX weights and CPUExecutionProvider only."""

from __future__ import annotations

import importlib.metadata
import time
from collections import OrderedDict
from pathlib import Path

from document_excel.domain.errors import AppError
from document_excel.infrastructure.ocr_contract import OcrDiagnostic, OcrResult
from document_excel.infrastructure.paddle_models import CYRILLIC, ENGLISH
from document_excel.infrastructure.paddle_ocr import PaddleOCREngine, _Backend, decode_prediction
from document_excel.infrastructure.rapid_models import (
    MODEL_FILES,
    ORT_VERSION,
    RAPID_VERSION,
    model_root,
    validate_bundle,
)


class _Predictor:
    def __init__(self, predictor):
        self.predictor = predictor

    def predict(self, pixels):
        from rapidocr.ch_ppocr_rec import TextRecInput

        output = self.predictor(TextRecInput(img=pixels))
        if not output.txts:
            return iter([{"rec_text": "", "rec_score": 0.0}])
        return iter([{"rec_text": output.txts[0], "rec_score": output.scores[0]}])


class _RapidBackend(_Backend):
    def __init__(self, root, device, threads):
        if device != "cpu":
            raise AppError("RapidOCR использует только CPUExecutionProvider.")
        import rapidocr
        from rapidocr.ch_ppocr_det import TextDetector
        from rapidocr.ch_ppocr_det.utils import DetPreProcess
        from rapidocr.ch_ppocr_rec import TextRecognizer
        from rapidocr.utils.log import logger
        from rapidocr.utils.parse_parameters import ParseParams
        from rapidocr.utils.typings import LangRec, OCRVersion

        logger.setLevel("ERROR")

        self.root, self.device, self.threads = root, device, threads
        self.cache = OrderedDict()
        self.model_loads = 0
        config = ParseParams.load(Path(rapidocr.__file__).with_name("config.yaml"))
        engine_config = config.EngineConfig.onnxruntime
        engine_config.intra_op_num_threads = threads
        engine_config.inter_op_num_threads = 1
        for flag in ("use_cuda", "use_dml", "use_cann", "use_coreml"):
            engine_config[flag] = False

        class FixedResolutionDetector(TextDetector):
            def get_preprocess(self, max_wh):
                return DetPreProcess(self.limit_side_len, self.limit_type, self.mean, self.std)

        detector = config.Det
        detector.ocr_version = OCRVersion.PPOCRV5
        detector.model_path = str(Path(root) / MODEL_FILES["det"][0])
        detector.model_root_dir = root
        detector.engine_cfg = engine_config
        detector.limit_side_len, detector.limit_type = 1600, "max"
        detector.mean, detector.std = [0.485, 0.456, 0.406], [0.229, 0.224, 0.225]
        detector.box_thresh, detector.unclip_ratio, detector.use_dilation = 0.6, 1.5, False
        self.detector = FixedResolutionDetector(detector)
        self._check_provider(self.detector)

        def create(model_name, **_kwargs):
            cfg = config.Rec.copy()
            cfg.ocr_version = OCRVersion.PPOCRV5
            cfg.lang_type = LangRec.EN if model_name == ENGLISH else LangRec.CYRILLIC
            cfg.model_path = str(
                Path(root) / MODEL_FILES["latin" if model_name == ENGLISH else "cyrillic"][0]
            )
            cfg.model_root_dir = root
            cfg.engine_cfg = engine_config
            cfg.rec_batch_num = 4
            cfg.font_path = None
            recognizer = TextRecognizer(cfg)
            self._check_provider(recognizer)
            alphabet = set(recognizer.postprocess_op.character)
            needed = set(
                "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789<"
                if model_name == ENGLISH
                else "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯҮӨҢүөң0123456789"
            )
            if not needed <= alphabet:
                raise AppError("Алфавит ONNX-модели не соответствует установленному профилю.")
            self.model_loads += 1
            return _Predictor(recognizer)

        self.recognition_class = create
        self.recognizers = {CYRILLIC: create(CYRILLIC)}

    @staticmethod
    def _check_provider(predictor):
        if predictor.session.session.get_providers() != ["CPUExecutionProvider"]:
            raise AppError("ONNX Runtime должен использовать только CPUExecutionProvider.")

    def line(self, pixels, model, image_id, variant):
        result = super().line(pixels, model, image_id, variant)
        result.engine_name = "rapidocr"
        result.model_version = f"PP-OCRv5 / RapidOCR {RAPID_VERSION} / ORT {ORT_VERSION} / {model}"
        return result

    def document(self, pixels, image_id, variant):
        from rapidocr.ch_ppocr_rec import TextRecInput
        from rapidocr.utils.process_img import get_rotate_crop_image

        started = time.perf_counter()
        detected = self.detector(pixels)
        detection_seconds = time.perf_counter() - started
        h, w = pixels.shape[:2]
        if detected.boxes is None or not len(detected.boxes):
            return OcrResult(
                [],
                image_id,
                variant,
                w,
                h,
                "rapidocr",
                f"PP-OCRv5 / RapidOCR {RAPID_VERSION}",
                timings={"text_detection_seconds": detection_seconds, "text_recognition_seconds": 0.0},
            )
        crops = [get_rotate_crop_image(pixels, box.copy()) for box in detected.boxes]
        started = time.perf_counter()
        read = self.recognizers[CYRILLIC].predictor(TextRecInput(img=crops))
        recognition_seconds = time.perf_counter() - started
        result = decode_prediction(
            {"rec_texts": read.txts, "rec_scores": read.scores, "rec_polys": detected.boxes},
            image_id,
            variant,
            w,
            h,
        )
        result.engine_name = "rapidocr"
        result.model_version = f"PP-OCRv5 / RapidOCR {RAPID_VERSION} / ORT {ORT_VERSION}"
        result.timings = {
            "text_detection_seconds": detection_seconds,
            "text_recognition_seconds": recognition_seconds,
        }
        return result


class RapidOCREngine(PaddleOCREngine):
    """Same geometry/fields/cache/cancellation contract; no Paddle import needed."""

    def __init__(self, settings):
        self.settings = settings.model_copy(deep=True)
        self.root = model_root(settings.rapid_model_directory)

    def diagnose(self):
        try:
            versions = {name: importlib.metadata.version(name) for name in ("rapidocr", "onnxruntime")}
            if versions != {"rapidocr": RAPID_VERSION, "onnxruntime": ORT_VERSION}:
                raise AppError(f"Требуются RapidOCR {RAPID_VERSION} и ONNX Runtime CPU {ORT_VERSION}.")
            if self.settings.paddle_device != "cpu":
                raise AppError("Выберите CPU; NVIDIA, CUDA и cuDNN не требуются.")
            validate_bundle(self.root)
            return OcrDiagnostic(
                str(self.root), f"RapidOCR {RAPID_VERSION} / PP-OCRv5 / CPU", ["ru", "ky", "en", "digits"]
            )
        except importlib.metadata.PackageNotFoundError:
            return OcrDiagnostic(
                str(self.root),
                "PP-OCRv5 / CPU",
                [],
                "Установите CPU-поставку RapidOCR и локальные ONNX-модели.",
            )
        except AppError as exc:
            return OcrDiagnostic(str(self.root), "PP-OCRv5 / CPU", [], str(exc))
