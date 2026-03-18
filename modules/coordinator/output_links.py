"""Coordinator output-link rendering and activation helpers."""

from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Callable, Mapping, Protocol, TypedDict, cast

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices

from common.output_markup import render_labeled_value


class _GeneratedOutputsView(Protocol):
    def setHtml(self, html: str) -> None:  # noqa: N802 - Qt naming
        ...


class _OutputLinkModule(Protocol):
    OUTPUT_LINK_NOT_AVAILABLE_KEY: str
    OUTPUT_LINK_OPEN_FILE_KEY: str
    OUTPUT_LINK_OPEN_FOLDER_KEY: str
    OUTPUT_LINK_OPEN_FAILED_KEY: str
    _files: list[Path]
    _downloaded_outputs: list[Path]
    generated_outputs_view: _GeneratedOutputsView


class _OutputLinkNamespace(TypedDict):
    t: Callable[..., str]
    OUTPUT_LINK_MODE_FILE: str
    OUTPUT_LINK_MODE_FOLDER: str
    OUTPUT_LINK_SEPARATOR: str
    show_toast: Callable[..., None]


def output_link_markup(module: object, label: str, path: str | None, *, ns: Mapping[str, object]) -> str:
    typed_module = cast(_OutputLinkModule, module)
    typed_ns = cast(_OutputLinkNamespace, ns)
    t = typed_ns["t"]
    if not path:
        return f"{escape(label)}: {t(typed_module.OUTPUT_LINK_NOT_AVAILABLE_KEY)}"
    href_path = Path(path).as_posix()
    file_target = f'{typed_ns["OUTPUT_LINK_MODE_FILE"]}{typed_ns["OUTPUT_LINK_SEPARATOR"]}{href_path}'
    folder_target = f'{typed_ns["OUTPUT_LINK_MODE_FOLDER"]}{typed_ns["OUTPUT_LINK_SEPARATOR"]}{href_path}'
    name = escape(Path(path).name)
    full_path = escape(str(Path(path)))
    return (
        f"{render_labeled_value(label, name)}<br>"
        f"{full_path}<br>"
        f'<a href="{file_target}">{t(typed_module.OUTPUT_LINK_OPEN_FILE_KEY)}</a>'
        " | "
        f'<a href="{folder_target}">{t(typed_module.OUTPUT_LINK_OPEN_FOLDER_KEY)}</a>'
    )


def output_links_html(module: object, *, ns: Mapping[str, object]) -> str:
    typed_module = cast(_OutputLinkModule, module)
    typed_ns = cast(_OutputLinkNamespace, ns)
    t = typed_ns["t"]
    rows: list[str] = []
    uploaded_label = t("coordinator.links.uploaded_report")
    downloaded_label = t("coordinator.links.downloaded_output")
    for path in typed_module._files:
        rows.append(output_link_markup(typed_module, uploaded_label, str(path), ns=typed_ns))
    if not rows:
        rows.append(output_link_markup(typed_module, uploaded_label, None, ns=typed_ns))

    if typed_module._downloaded_outputs:
        for path in typed_module._downloaded_outputs:
            rows.append(output_link_markup(typed_module, downloaded_label, str(path), ns=typed_ns))
    else:
        rows.append(output_link_markup(typed_module, downloaded_label, None, ns=typed_ns))
    return "<br><br>".join(rows)


def refresh_output_links(module: object, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_OutputLinkModule, module)
    typed_ns = cast(_OutputLinkNamespace, ns)
    typed_module.generated_outputs_view.setHtml(output_links_html(typed_module, ns=typed_ns))


def on_output_link_activated(module: object, href: str, *, ns: Mapping[str, object]) -> None:
    typed_module = cast(_OutputLinkModule, module)
    typed_ns = cast(_OutputLinkNamespace, ns)
    t = typed_ns["t"]
    mode, _, raw_path = href.partition(typed_ns["OUTPUT_LINK_SEPARATOR"])
    path = raw_path.strip()
    if not path:
        return
    target = Path(path).parent if mode == typed_ns["OUTPUT_LINK_MODE_FOLDER"] else Path(path)
    opened = QDesktopServices.openUrl(QUrl.fromLocalFile(str(target)))
    if opened:
        return
    typed_ns["show_toast"](
        typed_module,
        t(typed_module.OUTPUT_LINK_OPEN_FAILED_KEY),
        title=t("instructor.msg.error_title"),
        level="error",
    )
