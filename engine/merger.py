"""
多表合并器 —— 将同名 export_name 的多张表按字段名匹配合并。
列取并集，缺失字段填 nil。
"""
from engine.excel_reader import TableData, FieldDef
from engine.type_system import FieldType, TypeKind


class MergedTable:
    """合并后的表"""
    def __init__(self):
        self.export_name: str = ""
        self.fields: list[FieldDef] = []   # 合并后的字段列表（并集）
        self.rows: list[dict] = []          # 合并后的数据行
        self.has_id: bool = False
        self.has_subid: bool = False
        self.sources: list[str] = []        # 来源标识列表

    @property
    def source_label(self) -> str:
        """兼容 validator 使用的 source_label 属性"""
        return ", ".join(self.sources)


def merge_tables(tables: list[TableData]) -> dict[str, MergedTable]:
    """将多个 TableData 按 export_name 分组合并。
    返回 {export_name: MergedTable}。
    """
    groups: dict[str, list[TableData]] = {}
    for t in tables:
        name = t.export_name
        if name not in groups:
            groups[name] = []
        groups[name].append(t)

    result: dict[str, MergedTable] = {}
    for name, group in groups.items():
        merged = _merge_group(name, group)
        result[name] = merged
    return result


def _merge_group(export_name: str, tables: list[TableData]) -> MergedTable:
    """合并同一 export_name 的一组表。"""
    merged = MergedTable()
    merged.export_name = export_name

    # 1. 收集所有字段（按字段名去重，保留首次出现的定义）
    field_map: dict[str, FieldDef] = {}  # name -> FieldDef
    for table in tables:
        merged.sources.append(table.source_label)
        for fd in table.fields:
            if fd.name not in field_map:
                field_map[fd.name] = fd

    # 按原始列顺序排序（保持首次出现表的列序，后续新增追加）
    merged.fields = list(field_map.values())

    # 2. 检测 ID / SubID
    merged.has_id = any(fd.name == "ID" for fd in merged.fields)
    merged.has_subid = any(fd.name == "SubID" for fd in merged.fields)

    # 3. 合并数据行
    for table in tables:
        for row in table.rows:
            merged_row: dict = {}
            for fd in merged.fields:
                merged_row[fd.name] = row.get(fd.name, None)
            merged.rows.append(merged_row)

    return merged
