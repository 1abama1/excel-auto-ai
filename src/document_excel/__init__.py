"""Document Excel desktop application."""

import os

__version__ = "1.0.0"

# Apply before NumPy/OpenBLAS/OpenMP are imported, including spawned workers.
# The UI and an OCR worker otherwise each create a machine-sized thread pool.
for _variable in ("OMP_NUM_THREADS", "OPENBLAS_NUM_THREADS", "MKL_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
    try:
        _threads = min(2, max(1, int(os.environ.get(_variable, "2"))))
    except ValueError:
        _threads = 2
    os.environ[_variable] = str(_threads)
