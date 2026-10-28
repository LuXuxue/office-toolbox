import threading

import sopdf
import numpy as np
import cv2


# PDFium requires process-wide exclusion, even for separate documents.
_PDFIUM_LOCK = threading.Lock()

def pdf_to_images(pdf_bytes: bytes, dpi: int = 150) -> list[np.ndarray]:
    """Convert PDF bytes to a list of OpenCV images (numpy arrays, BGR format)."""
    with _PDFIUM_LOCK:
        images = []
        with sopdf.open(stream=pdf_bytes) as doc:
            for page in doc.pages:
                png_bytes = page.render(dpi=dpi)
                nparr = np.frombuffer(png_bytes, np.uint8)
                cv_image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                images.append(cv_image)
        return images
