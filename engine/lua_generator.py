"""
Lua 代码生成器 —— 将合并后的数据生成 .lua 文件。
"""
import os
from engine.merger import MergedTable
from engine.type_system import TypeKind


def generate_lua(merged: MergedTable) -> str:
    """根据合并表生成完整 Lua 代码字符串。"""
    lines = []

    # 文件头注释
    lines.append(f"-- 自动生成，请勿手动修改")
    lines.append(f"-- 来源: {', '.join(merged.sources)}")
    lines.append(f"")

    # 确定模式
    if merged.has_subid:
        return _generate_nested(merged, lines)
    else:
        return _generate_flat(merged, lines)


def _generate_flat(merged: MergedTable, lines: list[str]) -> str:
    """扁平模式: 每个字段一行"""
    table_name = merged.export_name
    lines.append(f"local {table_name} = {{")

    data_fields = _get_data_fields(merged.fields)

    for row in merged.rows:
        rid = row.get("ID")
        if rid is None:
            continue
        rid_str = _format_key(rid)
        field_lines = _format_data_lines(row, data_fields, indent=2)
        if field_lines:
            lines.append(f"\t{rid_str} = {{")
            lines.extend(field_lines)
            lines.append(f"\t}},")
        else:
            lines.append(f"\t{rid_str} = {{}},")

    lines.append(f"}}")
    lines.append(f"return {table_name}")
    return "\n".join(lines)


def _generate_nested(merged: MergedTable, lines: list[str]) -> str:
    """嵌套模式: 每个字段一行"""
    table_name = merged.export_name
    lines.append(f"local {table_name} = {{")

    groups: dict = {}
    order: list = []
    for row in merged.rows:
        rid = row.get("ID")
        subid = row.get("SubID")
        if rid is None or subid is None:
            continue
        if rid not in groups:
            groups[rid] = []
            order.append(rid)
        groups[rid].append((subid, row))

    data_fields = _get_data_fields(merged.fields)

    for rid in order:
        rid_str = _format_key(rid)
        lines.append(f"\t{rid_str} = {{")
        for subid, row in groups[rid]:
            subid_str = _format_key(subid)
            field_lines = _format_data_lines(row, data_fields, indent=3)
            if field_lines:
                lines.append(f"\t\t{subid_str} = {{")
                lines.extend(field_lines)
                lines.append(f"\t\t}},")
            else:
                lines.append(f"\t\t{subid_str} = {{}},")
        lines.append(f"\t}},")

    lines.append(f"}}")
    lines.append(f"return {table_name}")
    return "\n".join(lines)


def _get_data_fields(fields) -> list:
    """获取需要写入数据的字段（排除 ID、SubID、Ignore）。"""
    return [
        f for f in fields
        if f.name not in ("ID", "SubID")
        and f.type_def.kind != TypeKind.IGNORE
    ]


def _format_data_lines(row: dict, fields, indent: int = 1) -> list[str]:
    """格式化一行数据的字段列表，每个字段一行。跳过 nil 值。"""
    prefix = "\t" * indent
    result = []
    for fd in fields:
        val = row.get(fd.name)
        if val is None:
            continue
        lua_val = _to_lua_value(val, fd.type_def)
        result.append(f"{prefix}{fd.name} = {lua_val},")
    return result


def _format_key(value) -> str:
    """返回完整的 Lua 表键语法。
    Int    → [1001]
    String → ITEM_SWORD（合法标识符）或 ["含特殊字符"]（非标识符）
    """
    if isinstance(value, int):
        return f"[{value}]"
    if isinstance(value, float):
        v = int(value) if value == int(value) else value
        return f"[{v}]"
    if isinstance(value, str):
        if _is_valid_lua_identifier(value):
            return value
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        return f'["{escaped}"]'
    return f'["{value}"]'


def _to_lua_value(value, field_type=None) -> str:
    """将 Python 值转为 Lua 字面量字符串。"""
    if value is None:
        return "nil"

    if isinstance(value, bool):
        return "true" if value else "false"

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if value == int(value):
            return str(int(value))
        return str(value)

    if isinstance(value, str):
        # 转义特殊字符
        escaped = value.replace("\\", "\\\\").replace('"', '\\"')
        return f'"{escaped}"'

    if isinstance(value, list):
        items = []
        for item in value:
            items.append(_to_lua_value(item))
        return "{" + ", ".join(items) + "}"

    if isinstance(value, dict):
        pairs = []
        for k, v in value.items():
            k_str = _format_lua_dict_key(k)
            v_str = _to_lua_value(v)
            pairs.append(f"{k_str} = {v_str}")
        return "{" + ", ".join(pairs) + "}"

    return f'"{value}"'


def _format_lua_dict_key(key) -> str:
    """格式化 Lua 字典的键。字符串键若为合法标识符则不加引号。"""
    if isinstance(key, int):
        return f"[{key}]"
    if isinstance(key, float):
        if key == int(key):
            return f"[{int(key)}]"
        return f"[{key}]"
    if isinstance(key, str):
        if _is_valid_lua_identifier(key):
            return key
        return f'["{key}"]'
    return f'["{key}"]'


def _is_valid_lua_identifier(s: str) -> bool:
    """检查字符串是否为合法 Lua 标识符（仅 ASCII 字母/数字/下划线，不以数字开头）。"""
    if not s:
        return False
    # Lua 标识符仅支持 ASCII 字母
    first = s[0]
    if not (('a' <= first <= 'z') or ('A' <= first <= 'Z') or first == '_'):
        return False
    for c in s:
        if not (('a' <= c <= 'z') or ('A' <= c <= 'Z') or ('0' <= c <= '9') or c == '_'):
            return False
    return True


def write_lua_file(merged: MergedTable, output_dir: str) -> str:
    """将合并表写入 .lua 文件，返回写入的文件路径。"""
    os.makedirs(output_dir, exist_ok=True)
    lua_code = generate_lua(merged)
    file_name = f"{merged.export_name}.lua"
    file_path = os.path.join(output_dir, file_name)
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(lua_code)
    return file_path
