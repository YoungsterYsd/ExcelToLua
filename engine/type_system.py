"""
类型系统 —— 解析 Row2 类型定义，校验并转换 Row3+ 的单元格值。
支持: Int / Float / String / Const(...) / Array / Dic(K:V) / List(T) / Ignore
"""
import re
from enum import Enum, auto
from typing import Any, Optional, Union


class TypeKind(Enum):
    INT = auto()
    FLOAT = auto()
    STRING = auto()
    CONST = auto()
    ARRAY = auto()
    DIC = auto()
    LIST = auto()
    IGNORE = auto()


class FieldType:
    """单个字段的类型描述"""

    def __init__(self, kind: TypeKind):
        self.kind = kind
        # Const: {name: value} 映射
        self.const_map: dict[str, int] = {}
        # Dic: (key_type, value_type) 两个 FieldType
        self.dic_key: Optional["FieldType"] = None
        self.dic_value: Optional["FieldType"] = None
        # List/Array: 元素类型 FieldType
        self.element_type: Optional["FieldType"] = None

    def __repr__(self):
        return f"FieldType({self.kind})"


# ---- 类型定义解析（Row2） ----

def parse_type_definition(raw: str) -> FieldType:
    """从 Row2 的类型字符串解析出 FieldType。
    例如 "Int" -> FieldType(INT)
        "Const(A=1,B=2)" -> FieldType(CONST) with const_map
        "Dic(String:Int)" -> FieldType(DIC) with key=STRING, value=INT
        "List(List(Int))" -> FieldType(LIST) with element=LIST(INT)
    """
    raw = raw.strip()
    if not raw:
        raise ValueError(f"类型定义不能为空")

    low = raw.lower()

    # Ignore
    if low == "ignore":
        return FieldType(TypeKind.IGNORE)

    # Int
    if low == "int":
        return FieldType(TypeKind.INT)

    # Float
    if low == "float":
        return FieldType(TypeKind.FLOAT)

    # String
    if low == "string":
        return FieldType(TypeKind.STRING)

    # Array
    if low == "array":
        ft = FieldType(TypeKind.ARRAY)
        ft.element_type = FieldType(TypeKind.INT)  # Array 元素默认 Int
        return ft

    # Const(Typename1=1,Typename2=2)
    const_match = re.match(r'^Const\((.*)\)$', raw.strip(), re.IGNORECASE)
    if const_match:
        ft = FieldType(TypeKind.CONST)
        inner = const_match.group(1)
        for pair in _split_top_level(inner, ','):
            pair = pair.strip()
            if '=' not in pair:
                raise ValueError(f"Const 定义格式错误: {pair}，应为 Name=Value")
            name, val = pair.split('=', 1)
            name = name.strip()
            try:
                ft.const_map[name] = int(val.strip())
            except ValueError:
                raise ValueError(f"Const 值必须为整数: {val.strip()}")
        return ft

    # Dic(String=Int) / Dic(Int=String) 等
    dic_match = re.match(r'^Dic\((.*)\)$', raw.strip(), re.IGNORECASE)
    if dic_match:
        inner = dic_match.group(1)
        if '=' not in inner:
            raise ValueError(f"Dic 定义格式错误: {raw}，应为 Dic(KeyType=ValueType)")
        key_str, val_str = inner.split('=', 1)
        ft = FieldType(TypeKind.DIC)
        ft.dic_key = _parse_simple_type(key_str.strip())
        ft.dic_value = parse_type_definition(val_str.strip())
        return ft

    # List(Int) / List(List(Int)) 等
    list_match = re.match(r'^List\((.*)\)$', raw.strip(), re.IGNORECASE)
    if list_match:
        inner = list_match.group(1)
        ft = FieldType(TypeKind.LIST)
        ft.element_type = parse_type_definition(inner.strip())
        return ft

    raise ValueError(f"无法识别的类型定义: {raw}")


def _parse_simple_type(raw: str) -> FieldType:
    """解析简单类型（仅 Int/String/Float，用于 Dic Key）"""
    low = raw.strip().lower()
    if low == "int":
        return FieldType(TypeKind.INT)
    if low == "string":
        return FieldType(TypeKind.STRING)
    if low == "float":
        return FieldType(TypeKind.FLOAT)
    raise ValueError(f"Dic Key 仅支持 Int/String/Float: {raw}")


# ---- 值解析与校验 ----

def parse_value(raw_value, field_type: FieldType) -> Any:
    """将单元格原始值按类型定义解析为 Python 对象。
    返回解析后的值；若无法解析则抛出 ValueError。
    """
    if raw_value is None:
        return None

    raw = str(raw_value).strip()
    if not raw:
        return None

    if field_type.kind == TypeKind.IGNORE:
        return None  # Ignore 列不返回数据

    if field_type.kind == TypeKind.INT:
        return _parse_int(raw)

    if field_type.kind == TypeKind.FLOAT:
        return _parse_float(raw)

    if field_type.kind == TypeKind.STRING:
        return raw

    if field_type.kind == TypeKind.CONST:
        return _parse_const(raw, field_type.const_map)

    if field_type.kind == TypeKind.ARRAY:
        return _parse_array(raw)

    if field_type.kind == TypeKind.DIC:
        return _parse_dic(raw, field_type)

    if field_type.kind == TypeKind.LIST:
        return _parse_list(raw, field_type)

    raise ValueError(f"未知类型: {field_type.kind}")


def _parse_int(raw: str) -> int:
    try:
        if '.' in raw:
            return int(float(raw))
        return int(raw)
    except ValueError:
        raise ValueError(f"无法转换为 Int: {raw}")


def _parse_float(raw: str) -> float:
    try:
        return float(raw)
    except ValueError:
        raise ValueError(f"无法转换为 Float: {raw}")


def _parse_const(raw: str, const_map: dict[str, int]) -> int:
    if raw not in const_map:
        valid = ', '.join(const_map.keys())
        raise ValueError(f"Const 值 \"{raw}\" 不在定义范围内，合法值: [{valid}]")
    return const_map[raw]


def _parse_array(raw: str) -> list:
    """解析 Array: [1,2,3] -> [1, 2, 3]"""
    raw = raw.strip()
    if not (raw.startswith('[') and raw.endswith(']')):
        raise ValueError(f"Array 必须用 [] 包裹: {raw}")
    inner = raw[1:-1].strip()
    if not inner:
        return []
    elements = _split_top_level(inner, ',')
    result = []
    for elem in elements:
        elem = elem.strip()
        result.append(_parse_loose_value(elem))
    return result


def _parse_dic(raw: str, field_type: FieldType) -> dict:
    """解析 Dic: {key1=val1, key2=val2} -> {key1: val1, key2: val2}
    使用 Lua 标准的 = 分隔符。
    """
    raw = raw.strip()
    if not (raw.startswith('{') and raw.endswith('}')):
        raise ValueError(f"Dic 必须用 {{}} 包裹: {raw}")
    inner = raw[1:-1].strip()
    if not inner:
        return {}
    pairs = _split_top_level(inner, ',')
    result = {}
    for pair in pairs:
        if '=' not in pair:
            raise ValueError(f"Dic 键值对必须用 = 分隔: {pair}")
        key, val = pair.split('=', 1)
        key = key.strip()
        val = val.strip()

        # 解析 key
        parsed_key = _parse_simple_value(key, field_type.dic_key)
        # 解析 value（可能是嵌套类型）
        if field_type.dic_value.kind in (TypeKind.INT, TypeKind.FLOAT, TypeKind.STRING):
            parsed_val = _parse_simple_value(val, field_type.dic_value)
        else:
            parsed_val = parse_value(val, field_type.dic_value)
        result[parsed_key] = parsed_val
    return result


def _parse_list(raw: str, field_type: FieldType) -> list:
    """解析 List: {1,2,3} -> [1, 2, 3]
    List(List(Int)): {{1,2},{3,4}} -> [[1,2],[3,4]]
    """
    raw = raw.strip()
    if not (raw.startswith('{') and raw.endswith('}')):
        raise ValueError(f"List 必须用 {{}} 包裹: {raw}")
    inner = raw[1:-1].strip()
    if not inner:
        return []
    elements = _split_top_level(inner, ',')
    result = []
    elem_type = field_type.element_type
    for elem in elements:
        elem = elem.strip()
        if elem_type.kind in (TypeKind.INT, TypeKind.FLOAT, TypeKind.STRING):
            result.append(_parse_simple_value(elem, elem_type))
        else:
            # 嵌套类型（List/Array/Dic）需要完整包裹
            result.append(parse_value(elem, elem_type))
    return result


def _parse_simple_value(raw: str, field_type: FieldType) -> Any:
    """解析简单标量值"""
    if field_type.kind == TypeKind.INT:
        return _parse_int(raw)
    if field_type.kind == TypeKind.FLOAT:
        return _parse_float(raw)
    if field_type.kind == TypeKind.STRING:
        return raw
    raise ValueError(f"不支持的简单类型: {field_type.kind}")


def _parse_loose_value(raw: str) -> Any:
    """宽松解析 —— 自动推断 Int/Float/String"""
    raw = raw.strip()
    # 尝试 Int
    try:
        return int(raw)
    except ValueError:
        pass
    # 尝试 Float
    try:
        return float(raw)
    except ValueError:
        pass
    # 默认 String
    return raw


def _split_top_level(text: str, delimiter: str = ',') -> list[str]:
    """在顶层按分隔符切分，忽略嵌套 {} 和 [] 内的分隔符。"""
    parts = []
    depth = 0
    current = []
    for ch in text:
        if ch in '{[':
            depth += 1
        elif ch in '}]':
            depth -= 1
        if ch == delimiter and depth == 0:
            parts.append(''.join(current))
            current = []
        else:
            current.append(ch)
    if current:
        parts.append(''.join(current))
    return parts
