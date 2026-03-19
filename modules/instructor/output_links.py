"""Instructor output-link rendering and activation helpers."""

from __future__ import annotations

from html import escape
from pathlib import Path
from typing import Any, Callable, Mapping, Protocol, TypedDict

from common.output_markup import render_labeled_value


class _LabelTarget(Protocol):
    def setText(self, value: str) -> None:  # noqa: N802 - Qt naming
        ...


class _InstructorOutputModule(Protocol):
    @property
    def RAIL_LINK_NOT_AVAILABLE_KEY(self) -> str:
        ...

    @property
    def RAIL_LINK_OPEN_FILE_KEY(self) -> str:
        ...

    @property
    def RAIL_LINK_OPEN_FOLDER_KEY(self) -> str:
        ...

    @property
    def RAIL_LINK_OPEN_FAILED_KEY(self) -> str:
        ...

    def _quick_link_items(self) -> tuple[tuple[str, str | None], ...]:
        ...


class InstructorOutputNamespace(TypedDict):
    t: Callable[..., str]
    OUTPUT_LINK_MODE_FILE: str
    OUTPUT_LINK_MODE_FOLDER: str
    OUTPUT_LINK_SEPARATOR: str
    url_from_local_file: Callable[[str], Any]
    open_url: Callable[[Any], bool]
    show_toast: Callable[..., None]


def quick_link_markup(
    module: _InstructorOutputModule,
    label_key: str,
    path: str | None,
    *,
    ns: InstructorOutputNamespace,
) -> str:
    t = ns["t"]
    label = t(label_key)
    if not path:
        return f"{escape(label)}: {t(module.RAIL_LINK_NOT_AVAILABLE_KEY)}"
    name = escape(Path(path).name)
    full_path = escape(str(Path(path)))
    href_path = Path(path).as_posix()
    file_target = f'{ns["OUTPUT_LINK_MODE_FILE"]}{ns["OUTPUT_LINK_SEPARATOR"]}{href_path}'
    folder_target = f'{ns["OUTPUT_LINK_MODE_FOLDER"]}{ns["OUTPUT_LINK_SEPARATOR"]}{href_path}'
    return (
        f"{render_labeled_value(label, name)}<br>"
        f"{full_path}<br>"
        f'<a href="{file_target}">{t(module.RAIL_LINK_OPEN_FILE_KEY)}</a>'
        " | "
        f'<a href="{folder_target}">{t(module.RAIL_LINK_OPEN_FOLDER_KEY)}</a>'
    )


def quick_links_html(module: _InstructorOutputModule, *, ns: InstructorOutputNamespace) -> str:
    return "<br><br>".join(
        quick_link_markup(module, link_key, path, ns=ns)
        for link_key, path in module._quick_link_items()
    )


def refresh_quick_links(module: _InstructorOutputModule, *, ns: InstructorOutputNamespace) -> None:
    items = list(module._quick_link_items())
    generated_outputs_view = getattr(module, "generated_outputs_view", None)
    if generated_outputs_view is not None:
        generated_outputs_view.setHtml("<br><br>".join(quick_link_markup(module, key, path, ns=ns) for key, path in items))
    quick_link_labels: Mapping[str, _LabelTarget] = getattr(module, "quick_link_labels", {})
    for link_key, link_label in quick_link_labels.items():
        matching_paths = [path for key, path in items if key == link_key]
        if not matching_paths:
            link_label.setText(quick_link_markup(module, link_key, None, ns=ns))
            continue
        if len(matching_paths) == 1:
            link_label.setText(quick_link_markup(module, link_key, matching_paths[0], ns=ns))
            continue
        link_label.setText("<br><br>".join(quick_link_markup(module, link_key, path, ns=ns) for path in matching_paths))


def on_quick_link_activated(module: _InstructorOutputModule, href: str, *, ns: InstructorOutputNamespace) -> None:
    t = ns["t"]
    mode, _, raw_path = href.partition(ns["OUTPUT_LINK_SEPARATOR"])
    path = raw_path.strip()
    if not path:
        return
    target = Path(path).parent if mode == ns["OUTPUT_LINK_MODE_FOLDER"] else Path(path)
    opened = ns["open_url"](ns["url_from_local_file"](str(target)))
    if opened:
        return
    ns["show_toast"](
        module,
        t(module.RAIL_LINK_OPEN_FAILED_KEY),
        title=t("instructor.msg.error_title"),
        level="error",
    )
