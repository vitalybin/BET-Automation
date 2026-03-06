# puralox/eln_client.py
"""
ElnClient — encapsulates all eLabFTW HTTP interactions.

Extracted from PuraloxApp to give the ELN integration a dedicated class,
making the class diagram accurate and the app layer cleaner.
"""

import os
import logging
import time
from io import BytesIO

import numpy as np
import matplotlib.pyplot as plt
import requests

import elabapi_python
from elabapi_python import ExperimentsApi
from elabapi_python.rest import ApiException

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Image as RLImage

logger = logging.getLogger(__name__)


class ElnClient:
    """
    Wraps all network calls to an eLabFTW instance:
      - API client configuration and connection test
      - Experiment creation / patching
      - BET plot PDF upload
      - Tag management
      - Template and experiment listing
    """

    def __init__(self, base_url: str, token: str, disable_ssl: bool = True):
        self.base_url = base_url
        self.token = token
        self.verify_ssl = not disable_ssl

        cfg = elabapi_python.Configuration()
        cfg.host = base_url
        cfg.verify_ssl = self.verify_ssl

        client = elabapi_python.ApiClient(cfg)
        client.set_default_header("Authorization", token)

        self.exp_api: ExperimentsApi = ExperimentsApi(client)

        self._test_connection()

    # ── Connection ──────────────────────────────────────────────────────────

    def _test_connection(self, max_attempts: int = 5) -> None:
        """Try up to `max_attempts` times to reach eLabFTW; warns on failure."""
        for attempt in range(1, max_attempts + 1):
            try:
                logger.debug("Testing eLabFTW connection (attempt %d)", attempt)
                self.exp_api.read_experiments(limit=1)
                logger.info("✅ eLabFTW API OK.")
                return
            except Exception as e:
                logger.warning("Connection attempt %d failed: %s", attempt, e)
                time.sleep(2 ** attempt)

    # ── Template / experiment listing ───────────────────────────────────────

    def fetch_templates(self) -> list:
        """Return a list of experiment templates from eLabFTW (empty list on error)."""
        url = f"{self.base_url}/experiments_templates?limit=100"
        logger.debug("GET %s", url)
        try:
            resp = requests.get(
                url,
                headers={"Authorization": self.token},
                verify=self.verify_ssl,
            )
        except Exception as e:
            logger.error("Template HTTP error: %s", e)
            return []

        if not resp.ok:
            logger.error("Template fetch failed [%s]: %s", resp.status_code, resp.text)
            return []

        try:
            return resp.json()
        except Exception:
            logger.exception("Failed to parse templates JSON")
            return []

    def fetch_experiments(self, limit: int = 10) -> list:
        """Return a list of recent experiments from eLabFTW."""
        exps = self.exp_api.read_experiments(limit=limit)
        return [e.to_dict() for e in exps]

    # ── Experiment lifecycle ─────────────────────────────────────────────────

    def create_experiment(self, title: str, body: str) -> str:
        """
        Create a new experiment in eLabFTW and immediately patch its title + body.

        Returns:
            exp_id (str): the numeric string ID of the created experiment.

        Raises:
            RuntimeError: if either the create or patch step fails.
        """
        _, status, headers = self.exp_api.post_experiment_with_http_info(body={})
        if status != 201:
            raise RuntimeError(f"Experiment creation failed (HTTP {status})")

        exp_id = headers["Location"].rstrip("/").split("/")[-1]

        self.exp_api.patch_experiment(exp_id, body={"title": title, "body": body})
        logger.info("Created eLabFTW experiment id=%s title=%r", exp_id, title)
        return exp_id

    def add_tag(self, exp_id: str, tag_name: str) -> None:
        """Add a tag to the experiment. Failures are logged but not re-raised."""
        try:
            resp = requests.post(
                f"{self.base_url}/experiments/{exp_id}/tags",
                headers={
                    "Authorization": self.token,
                    "Content-Type": "application/json",
                },
                json={"tag": tag_name},
                verify=self.verify_ssl,
            )
            logger.debug("Tag resp: %s %s", resp.status_code, resp.text)
        except Exception:
            logger.exception("Tagging failed (non-fatal)")

    # ── BET plot attachment ──────────────────────────────────────────────────

    def upload_plot(
        self,
        exp_id: str,
        file_id: int,
        pts: list[dict],
        x_min: float | None = None,
        x_max: float | None = None,
    ) -> None:
        """
        Build a BET scatter plot from `pts` (list of dicts with p_p0 / p_va_p0_p),
        wrap it in a PDF, and upload it as an attachment to the experiment.

        x_min / x_max: optional range coming from the UI sliders.
        Failures are logged but not re-raised.
        """
        if not pts:
            return

        try:
            x_all = np.array([r["p_p0"] for r in pts if r["p_p0"] is not None])
            y_all = np.array([r["p_va_p0_p"] for r in pts if r["p_va_p0_p"] is not None])

            if not len(x_all) or not len(y_all):
                return

            mask = np.ones_like(x_all, dtype=bool)
            if x_min is not None:
                mask &= x_all >= x_min
            if x_max is not None:
                mask &= x_all <= x_max

            x = x_all[mask] if mask.any() else x_all
            y = y_all[mask] if mask.any() else y_all

            fig, ax = plt.subplots()
            ax.scatter(x, y, s=20)
            ax.set_title("BET Plot")
            ax.set_xlabel("p/p0")
            ax.set_ylabel("p/va_p0_p")

            img_buf = BytesIO()
            fig.savefig(img_buf, format="PNG", bbox_inches="tight")
            plt.close(fig)
            img_buf.seek(0)

            pdf_buf = BytesIO()
            doc = SimpleDocTemplate(pdf_buf, pagesize=letter)
            doc.build([RLImage(img_buf, width=400, height=300)])
            pdf_buf.seek(0)

            files = {"file": (f"BET_Plot_{file_id}.pdf", pdf_buf.read(), "application/pdf")}
            up_url = f"{self.base_url}/experiments/{exp_id}/uploads"
            resp = requests.post(
                up_url,
                headers={"Authorization": self.token},
                files=files,
                verify=self.verify_ssl,
            )
            logger.debug(
                "Upload plot resp: %s %s (range: %s – %s)",
                resp.status_code, resp.text, x_min, x_max,
            )
        except Exception:
            logger.exception("Plot upload failed (non-fatal)")
