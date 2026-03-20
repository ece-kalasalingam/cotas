"""Shared generated-output panel model and rendering helpers."""

from __future__ import annotations

from dataclasses import dataclass
from html import escape
from pathlib import Path
from typing import Callable


@dataclass(frozen=True)
class OutputItem:
    label_key: str
    path: str


@dataclass(frozen=True)
class OutputPanelData:
    items: tuple[OutputItem, ...]
    empty_message_key: str = "outputs.none_generated"
    open_file_key: str = "instructor.links.open_file"
    open_folder_key: str = "instructor.links.open_folder"
    open_failed_key: str = "instructor.links.open_failed"


def render_output_panel_html(
    data: OutputPanelData,
    *,
    translate: Callable[..., str],
    output_link_mode_file: str,
    output_link_mode_folder: str,
    output_link_separator: str,
) -> str:
    if not data.items:
        return escape(translate(data.empty_message_key))
    rows: list[str] = []
    for item in data.items:
        label = escape(translate(item.label_key))
        path_obj = Path(item.path)
        file_name = escape(path_obj.name)
        full_path = escape(str(path_obj))
        href_path = path_obj.as_posix()
        file_target = f"{output_link_mode_file}{output_link_separator}{href_path}"
        folder_target = f"{output_link_mode_folder}{output_link_separator}{href_path}"
        rows.append(
            f"{label}: {file_name}<br>"
            f"{full_path}<br>"
            f'<a href="{file_target}">{escape(translate(data.open_file_key))}</a>'
            " | "
            f'<a href="{folder_target}">{escape(translate(data.open_folder_key))}</a>'
        )
    return "<br><br>".join(rows)


def open_output_link(
    href: str,
    *,
    output_link_mode_folder: str,
    output_link_separator: str,
    open_path: Callable[[Path], bool],
) -> bool:
    mode, _, raw_path = href.partition(output_link_separator)
    path = raw_path.strip()
    if not path:
        return False
    target = Path(path).parent if mode == output_link_mode_folder else Path(path)
    return bool(open_path(target))
