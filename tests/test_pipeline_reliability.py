"""Regression tests for import safety and data-output integrity."""

import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd

from src import data_import, data_join


class ConsumerQuarterlyTests(unittest.TestCase):
    def test_partial_missing_values_do_not_remove_quarter_rows(self) -> None:
        source = pd.DataFrame(
            {
                "indicator*": [1.0, None, None],
                "another indicator*": [None, 2.0, None],
                "monthly indicator": [4.0, 5.0, 6.0],
            },
            index=pd.date_range("2026-01-01", periods=3, freq="QS"),
        )

        monthly, quarterly = data_join._split_consumer_data(source)

        self.assertEqual(len(monthly), 3)
        self.assertEqual(quarterly.index.tolist(), source.index[:2].tolist())
        self.assertTrue(quarterly.isna().any().any())


class OutputIntegrityTests(unittest.TestCase):
    def test_duplicate_feature_names_are_disambiguated(self) -> None:
        frames = [
            pd.DataFrame({"duplicate": [1]}),
            pd.DataFrame({"duplicate": [2]}),
        ]
        combined = data_join._concat_unique_columns(
            frames,
            "monthly_data_0.xlsx",
        )
        self.assertEqual(
            combined.columns.tolist(),
            ["duplicate", "duplicate__duplicate_2"],
        )

    def test_empty_services_header_raises_value_error(self) -> None:
        with self.assertRaisesRegex(ValueError, "Empty header"):
            data_join._clean_services_header(None)

    def test_publish_removes_stale_managed_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            root = Path(temporary_directory)
            staging = root / "staging"
            output = root / "interim"
            staging.mkdir()
            output.mkdir()
            (staging / "monthly_data_0.xlsx").touch()
            (output / "monthly_data_3.xlsx").touch()
            (output / "quaterly_data_4.xlsx").touch()
            (output / "target.xlsx").touch()

            data_join._publish_output_files(staging, output)

            self.assertTrue((output / "monthly_data_0.xlsx").exists())
            self.assertFalse((output / "monthly_data_3.xlsx").exists())
            self.assertFalse((output / "quaterly_data_4.xlsx").exists())
            self.assertFalse((output / "target.xlsx").exists())


class ImporterReliabilityTests(unittest.TestCase):
    def test_failed_download_stage_does_not_archive_current_raw_files(self) -> None:
        with tempfile.TemporaryDirectory() as temporary_directory:
            data_directory = Path(temporary_directory) / "data"
            raw_directory = data_directory / "raw"
            raw_directory.mkdir(parents=True)
            existing_file = raw_directory / "current.xlsx"
            existing_file.write_bytes(b"existing")

            with (
                patch.object(data_import, "DATA_DIR", data_directory),
                patch.object(data_import, "RAW_DIR", raw_directory),
                patch.object(
                    data_import,
                    "_import_bulletin",
                    side_effect=RuntimeError("network failure"),
                ),
                patch.object(data_import, "_move_raw_files_to_bin") as archive,
            ):
                with self.assertRaisesRegex(RuntimeError, "network failure"):
                    data_import.importer()

            self.assertEqual(existing_file.read_bytes(), b"existing")
            archive.assert_not_called()

    def test_monthly_page_fetch_only_falls_back_on_404(self) -> None:
        original_date = data_import.date

        class FixedDate(original_date):
            @classmethod
            def today(cls):
                return cls(2026, 10, 1)

        error = data_import.requests.exceptions.HTTPError("not found")
        error.response = type("Response", (), {"status_code": 404})()
        soup = data_import.bs(
            "<ul class='files-list'><li><span class='filename'>"
            "publication</span><span class='file-url'>file.xlsx"
            "</span></li></ul>",
            "html.parser",
        )
        with patch.object(
            data_import, "_fetch_page", side_effect=[error, soup]
        ) as fetch, patch.object(data_import, "date", FixedDate):
            page_url, fetched_soup = data_import._fetch_latest_monthly_page(
                data_import.BUSINESS_SITUATION_URL,
                repeat_year=True,
            )

        self.assertIn("2026-september-2026", page_url)
        self.assertIs(fetched_soup, soup)
        self.assertEqual(fetch.call_count, 2)

    def test_monthly_page_fetch_propagates_non_404_http_errors(self) -> None:
        error = data_import.requests.exceptions.HTTPError("server error")
        error.response = type("Response", (), {"status_code": 500})()
        with patch.object(data_import, "_fetch_page", side_effect=error):
            with self.assertRaises(data_import.requests.exceptions.HTTPError):
                data_import._fetch_latest_monthly_page(
                    data_import.BUSINESS_SITUATION_URL,
                    repeat_year=True,
                )


if __name__ == "__main__":
    unittest.main()
