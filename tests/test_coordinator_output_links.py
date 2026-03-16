from __future__ import annotations

from pathlib import Path

from modules.coordinator import output_links


class _DummyModule:
    OUTPUT_LINK_NOT_AVAILABLE_KEY = "instructor.links.not_available"
    OUTPUT_LINK_OPEN_FILE_KEY = "instructor.links.open_file"
    OUTPUT_LINK_OPEN_FOLDER_KEY = "instructor.links.open_folder"
    OUTPUT_LINK_OPEN_FAILED_KEY = "instructor.links.open_failed"

    def __init__(self) -> None:
        self._files: list[Path] = []
        self._downloaded_outputs: list[Path] = []
        self.generated_outputs_view = None


class _DummyView:
    def __init__(self) -> None:
        self.html = ""

    def setHtml(self, html: str) -> None:  # noqa: N802 - Qt-style name
        self.html = html


def _ns() -> dict[str, object]:
    return {
        "t": lambda key, **kwargs: f"T({key})",
        "OUTPUT_LINK_MODE_FILE": "file",
        "OUTPUT_LINK_MODE_FOLDER": "folder",
        "OUTPUT_LINK_SEPARATOR": "::",
        "OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX": 10,
        "show_toast": lambda *args, **kwargs: None,
    }


def test_output_link_markup_without_path_uses_not_available_label() -> None:
    module = _DummyModule()
    html = output_links.output_link_markup(module, "Uploaded", None, ns=_ns())
    assert "<b>Uploaded</b>: T(instructor.links.not_available)" in html


def test_output_link_markup_with_path_renders_file_and_folder_links() -> None:
    module = _DummyModule()
    html = output_links.output_link_markup(module, "Uploaded", "C:/tmp/report.xlsx", ns=_ns())
    assert "file::C:/tmp/report.xlsx" in html
    assert "folder::C:/tmp/report.xlsx" in html
    assert "T(instructor.links.open_file)" in html
    assert "T(instructor.links.open_folder)" in html
    assert "report.xlsx" in html


def test_output_links_html_includes_uploaded_and_downloaded_placeholders_when_empty() -> None:
    module = _DummyModule()
    html = output_links.output_links_html(module, ns=_ns())
    assert html.count("margin-bottom:10px") == 2
    assert "T(coordinator.links.uploaded_report)" in html
    assert "T(coordinator.links.downloaded_output)" in html


def test_refresh_output_links_sets_generated_outputs_html() -> None:
    module = _DummyModule()
    module.generated_outputs_view = _DummyView()
    module._files = [Path("C:/tmp/a.xlsx")]

    output_links.refresh_output_links(module, ns=_ns())

    assert "a.xlsx" in module.generated_outputs_view.html


def test_on_output_link_activated_no_path_is_noop(monkeypatch) -> None:
    module = _DummyModule()
    calls: list[str] = []

    class _FakeQUrl:
        @staticmethod
        def fromLocalFile(path: str) -> str:
            return path

    monkeypatch.setattr(output_links, "QUrl", _FakeQUrl)
    monkeypatch.setattr(output_links.QDesktopServices, "openUrl", lambda url: calls.append(url) or True)

    output_links.on_output_link_activated(module, "file::   ", ns=_ns())

    assert calls == []


def test_on_output_link_activated_folder_mode_uses_parent_and_shows_error_toast_on_failure(monkeypatch) -> None:
    module = _DummyModule()
    opens: list[str] = []
    toasts: list[tuple[object, str, str, str]] = []

    class _FakeQUrl:
        @staticmethod
        def fromLocalFile(path: str) -> str:
            return path

    monkeypatch.setattr(output_links, "QUrl", _FakeQUrl)
    monkeypatch.setattr(output_links.QDesktopServices, "openUrl", lambda url: opens.append(url) or False)

    ns = _ns()
    ns["show_toast"] = lambda parent, message, *, title, level: toasts.append((parent, message, title, level))

    output_links.on_output_link_activated(module, "folder::C:/tmp/report.xlsx", ns=ns)

    assert opens == [str(Path("C:/tmp/report.xlsx").parent)]
    assert toasts == [
        (
            module,
            "T(instructor.links.open_failed)",
            "T(instructor.msg.error_title)",
            "error",
        )
    ]
