"""
数据校验器 —— ID/SubID 唯一性校验等。
"""
from typing import Optional


class ValidationError(Exception):
    """校验错误"""
    def __init__(self, message: str):
        super().__init__(message)
        self.message = message


def validate_table(table) -> list[str]:
    """校验单张表的 ID/SubID 唯一性。
    返回警告信息列表（非致命），抛出 ValidationError 表示致命错误。
    """
    warnings: list[str] = []

    if not table.has_id:
        return warnings

    # 分离 Ignore 列（不在数据校验范围）
    data_fields = [f for f in table.fields if f.type_def.kind.name != 'IGNORE']

    if table.has_subid:
        # ID + SubID 模式
        seen: set[tuple] = set()
        for i, row in enumerate(table.rows):
            rid = row.get("ID")
            subid = row.get("SubID")
            if rid is None:
                raise ValidationError(
                    f"[{table.source_label}] 第{i+3}行: ID 不能为空"
                )
            if subid is None:
                raise ValidationError(
                    f"[{table.source_label}] 第{i+3}行: "
                    f"ID={rid} 存在重复，但缺少 SubID 值"
                )
            key = (rid, subid)
            if key in seen:
                raise ValidationError(
                    f"[{table.source_label}] 第{i+3}行: "
                    f"ID={rid}, SubID={subid} 组合重复"
                )
            seen.add(key)
    else:
        # 仅 ID 模式 —— 先检查是否有重复
        id_counts: dict = {}
        for i, row in enumerate(table.rows):
            rid = row.get("ID")
            if rid is not None:
                id_counts[rid] = id_counts.get(rid, 0) + 1

        has_duplicate = any(c > 1 for c in id_counts.values())

        if has_duplicate:
            raise ValidationError(
                f"[{table.source_label}] ID 存在重复值，但没有 SubID 字段。"
                f"请添加 SubID 列以支持 ID+SubID 组合唯一。"
                f"重复的 ID: {[k for k, v in id_counts.items() if v > 1]}"
            )

    return warnings
