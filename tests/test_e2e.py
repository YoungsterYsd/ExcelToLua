"""
端到端测试 —— 读取测试 Excel，验证完整的 解析→合并→校验→生成 流程。
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from engine.excel_scanner import scan_xlsx_files, filter_by_name
from engine.excel_reader import read_excel
from engine.validator import validate_table, ValidationError
from engine.merger import merge_tables
from engine.lua_generator import generate_lua, write_lua_file


def main():
    test_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_data")
    output_dir = os.path.join(test_dir, "output")

    print("=" * 60)
    print("Lua 配置转换器 —— 端到端测试")
    print("=" * 60)

    # 1. 扫描文件
    files = scan_xlsx_files(test_dir)
    print(f"\n[1] 扫描到 {len(files)} 个 Excel 文件:")
    for f in files:
        print(f"    - {os.path.basename(f)}")

    # 2. 读取所有表
    all_tables = []
    errors = []
    for fpath in files:
        try:
            tables = read_excel(fpath)
            for t in tables:
                print(f"\n[2] 读取 [{t.source_label}] → 导出名={t.export_name}, "
                      f"字段数={len(t.fields)}, 数据行={len(t.rows)}")
                for fd in t.fields:
                    print(f"      字段: {fd.note} | {fd.name} ({fd.type_def})")
            all_tables.extend(tables)
        except Exception as e:
            errors.append(f"读取失败: {os.path.basename(fpath)}\n  {e}")
            print(f"\n[2] ❌ 读取失败 [{os.path.basename(fpath)}]: {e}")

    # 3. 合并
    merged_map = merge_tables(all_tables)
    print(f"\n[3] 合并后共 {len(merged_map)} 张表:")
    for name, merged in merged_map.items():
        print(f"    {name}: {len(merged.fields)} 字段, {len(merged.rows)} 行, "
              f"来源={merged.sources}")

    # 4. 校验 + 生成 Lua
    print(f"\n[4] 校验与生成:")
    for name, merged in merged_map.items():
        try:
            validate_table(merged)
            lua_code = generate_lua(merged)
            file_path = write_lua_file(merged, output_dir)
            print(f"\n  ✅ {name}.lua → {file_path}")
            print(f"  {'─' * 50}")
            print(lua_code)
            print(f"  {'─' * 50}")
        except ValidationError as e:
            print(f"\n  ❌ {name} 校验失败: {e.message}")

    # 5. 汇总
    if errors:
        print(f"\n{'=' * 60}")
        print(f"❌ 读取阶段有 {len(errors)} 个错误:")
        for e in errors:
            print(f"  {e}")

    print(f"\n{'=' * 60}")
    print("测试完成")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
