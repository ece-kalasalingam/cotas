"""Coordinator output-link rendering and activation helpers."""

from __future__ import annotations

from html import escape
from pathlib import Path
from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices


def output_link_markup(module: object, label: str, path: str | None, *, ns: dict[str, object]) -> str:
    if not path:
        return f"<b>{escape(label)}</b>: {ns['t'](module.OUTPUT_LINK_NOT_AVAILABLE_KEY)}"
    href_path = Path(path).as_posix()
    file_link = (
        f'<a href="{ns["OUTPUT_LINK_MODE_FILE"]}{ns["OUTPUT_LINK_SEPARATOR"]}{href_path}">'
        f"{ns['t'](module.OUTPUT_LINK_OPEN_FILE_KEY)}</a>"
    )
    folder_link = (
        f'<a href="{ns["OUTPUT_LINK_MODE_FOLDER"]}{ns["OUTPUT_LINK_SEPARATOR"]}{href_path}">'
        f"{ns['t'](module.OUTPUT_LINK_OPEN_FOLDER_KEY)}</a>"
    )
    name = escape(Path(path).name)
    full_path = escape(str(Path(path)))
    return (
        f"<b>{escape(label)}</b>: {name}<br>"
        f"<span>{full_path}</span><br>"
        f"{file_link} | {folder_link}"
    )


def output_links_html(module: object, *, ns: dict[str, object]) -> str:
    rows: list[str] = []
    uploaded_label = ns["t"]("coordinator.links.uploaded_report")
    downloaded_label = ns["t"]("coordinator.links.downloaded_output")
    row_spacing = ns["OUTPUT_LINK_ROW_MARGIN_BOTTOM_PX"]
    for path in module._files:
        rows.append(
            f"<div style='margin-bottom:{row_spacing}px'>{output_link_markup(module, uploaded_label, str(path), ns=ns)}</div>"
        )
    if not rows:
        rows.append(
            f"<div style='margin-bottom:{row_spacing}px'>{output_link_markup(module, uploaded_label, None, ns=ns)}</div>"
        )

    if module._downloaded_outputs:
        for path in module._downloaded_outputs:
            rows.append(
                f"<div style='margin-bottom:{row_spacing}px'>{output_link_markup(module, downloaded_label, str(path), ns=ns)}</div>"
            )
    else:
        rows.append(
            f"<div style='margin-bottom:{row_spacing}px'>{output_link_markup(module, downloaded_label, None, ns=ns)}</div>"
        )
    return "".join(rows)


def refresh_output_links(module: object, *, ns: dict[str, object]) -> None:
    module.generated_outputs_view.setHtml(output_links_html(module, ns=ns))


def on_output_link_activated(module: object, href: str, *, ns: dict[str, object]) -> None:
    mode, _, raw_path = href.partition(ns["OUTPUT_LINK_SEPARATOR"])
    path = raw_path.strip()
    if not path:
        return
    target = Path(path).parent if mode == ns["OUTPUT_LINK_MODE_FOLDER"] else Path(path)
    opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
    if opened:
        return
    ns["show_toast"](
        module,
        ns["t"](module.OUTPUT_LINK_OPEN_FAILED_KEY),
        title=ns["t"]("instructor.msg.error_title"),
        level="error",
    )
