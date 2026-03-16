"""Coordinator file-list action helpers."""

from __future__ import annotations

from pathlib import Path


def remove_file_by_path(module: object, file_path: str, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return
    target_key = ns["_path_key"](Path(file_path))
    before_count = len(module._files)
    module._files = [path for path in module._files if ns["_path_key"](path) != target_key]
    if len(module._files) == before_count:
        return

    for row in range(module.drop_list.count()):
        item = module.drop_list.item(row)
        path_value = str(item.data(ns["Qt"].ItemDataRole.UserRole) or "")
        if ns["_path_key"](Path(path_value)) == target_key:
            module.drop_list.takeItem(row)
            break

    module._refresh_ui()
    module._publish_status_key("coordinator.status.removed", count=1)
    ns["log_process_message"](
        "removing selected coordinator files",
        logger=module._logger,
        success_message="removing selected coordinator files completed successfully. removed=1",
        user_success_message=ns["build_i18n_log_message"](
            "coordinator.status.removed",
            kwargs={"count": 1},
            fallback=ns["t"]("coordinator.status.removed", count=1),
        ),
    )


def clear_all(module: object, *, ns: dict[str, object]) -> None:
    if module.state.busy:
        return
    if not module._files:
        return
    total = len(module._files)
    module._files.clear()
    module.drop_list.clear()
    module._refresh_ui()
    module._publish_status_key("coordinator.status.cleared", count=total)
    ns["log_process_message"](
        "clearing coordinator files",
        logger=module._logger,
        success_message=f"clearing coordinator files completed successfully. removed={total}",
        user_success_message=ns["build_i18n_log_message"](
            "coordinator.status.cleared",
            kwargs={"count": total},
            fallback=ns["t"]("coordinator.status.cleared", count=total),
        ),
    )
