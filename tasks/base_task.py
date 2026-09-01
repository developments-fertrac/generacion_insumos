from __future__ import annotations

import time
import traceback
from abc import ABC, abstractmethod
from pathlib import Path

from config.settings import Settings
from core.logger import get_logger
from core.email_notifier import EmailNotifier


class BaseTask(ABC):
    name: str = "base_task"

    def __init__(self, settings: Settings):
        self.settings = settings
        self.log = get_logger(self.name)
        self.notifier: EmailNotifier | None = None

    def run(self) -> bool:
        self.log.info("=" * 70)
        self.log.info("INICIO: %s", self.name)
        self.log.info("=" * 70)
        start = time.time()

        try:
            self.setup()
            self.execute()
            elapsed = int(time.time() - start)
            detail = f"Completado en {elapsed // 60}m {elapsed % 60}s"
            self.log.info("EXITO: %s", detail)
            self._notify_success(detail)
            return True
        except Exception as e:
            elapsed = int(time.time() - start)
            tb = traceback.format_exc()
            self.log.error("FALLO en %s (%ds): %s", self.name, elapsed, e)
            self.log.debug(tb)
            self._notify_failure(str(e))
            return False
        finally:
            self.teardown()
            self.log.info("FIN: %s (%ds total)", self.name, int(time.time() - start))

    def setup(self) -> None:
        self.log.info("Setup: %s", self.name)

    @abstractmethod
    def execute(self) -> None:
        ...

    def teardown(self) -> None:
        pass

    def _notify_success(self, detail: str) -> None:
        if self.notifier:
            self.notifier.notify_success(detail)

    def _notify_failure(self, error: str) -> None:
        if self.notifier:
            self.notifier.notify_failure(error)
