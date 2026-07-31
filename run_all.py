#!/usr/bin/env python3
"""
停车票核销报表 v3.0 - 一键生成所有报表
用法: python3 run_all.py <核销记录表.xlsx> <车辆进出报表.xlsx> [输出目录]

输出:
  1. 主报表 (31个子表) - xxx_停车票核销报表.xlsx
  2. 商家计费文件 (9个) - 含专业排版
  3. 集团/九号车场电子券报表
  4. 车场报表(物业、集团、哥弟来访条)
"""

import sys, os

# Add v1 skill path for build_report
SKILL_V1 = os.path.expanduser("~/.claude/skills/parking-report")
SKILL_V2 = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SKILL_V1)
sys.path.insert(0, SKILL_V2)


def main():
    if len(sys.argv) < 3:
        print("用法: python3 run_all.py <核销记录表.xlsx> <车辆进出报表.xlsx> [输出目录]")
        print("示例: python3 run_all.py ~/Desktop/核销记录表2026-7-5.xlsx ~/Desktop/车辆进出报表2026-7-5.xlsx ~/Desktop/输出/")
        sys.exit(1)

    main_file = os.path.expanduser(sys.argv[1])
    vehicle_file = os.path.expanduser(sys.argv[2])
    output_dir = os.path.expanduser(sys.argv[3]) if len(sys.argv) >= 4 else os.path.dirname(main_file)

    os.makedirs(output_dir, exist_ok=True)

    # Validate inputs
    for f in [main_file, vehicle_file]:
        if not os.path.exists(f):
            print(f"❌ 文件不存在: {f}")
            sys.exit(1)

    # Step 1: Generate main report using build_report.py
    print("=" * 60)
    print("🅿️  停车票核销报表生成器 v2.0")
    print("=" * 60)
    print(f"  核销记录表: {main_file}")
    print(f"  车辆进出报表: {vehicle_file}")
    print(f"  输出目录: {output_dir}")
    print("=" * 60)

    base_name = os.path.splitext(os.path.basename(main_file))[0]
    main_output = os.path.join(output_dir, f'{base_name}_停车票核销报表.xlsx')

    import subprocess
    build_script = os.path.join(SKILL_V1, 'build_report.py')

    print("\n📋 第一步: 生成主报表 (31个子表)...")
    result = subprocess.run(
        ['python3', build_script, main_file, main_output],
        capture_output=True, text=True
    )
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr)
        print("❌ 主报表生成失败!")
        sys.exit(1)

    # Steps 2-4: Generate branch files
    from generate_branch import step2_merchant_files, step3_jituan_report, step4_chechang_report

    print("\n📄 第二步: 生成商家计费文件...")
    period_text = step2_merchant_files(main_output, output_dir)

    print("\n🏢 第三步: 生成集团/九号报表...")
    step3_jituan_report(main_output, output_dir, period_text)

    print("\n🚗 第四步: 生成车场报表...")
    step4_chechang_report(main_output, vehicle_file, output_dir)

    # Summary
    print("\n" + "=" * 60)
    print("✅ 全部完成!")
    print(f"  所有文件已保存到: {output_dir}")
    print("=" * 60)

    # List output files
    for f in sorted(os.listdir(output_dir)):
        fpath = os.path.join(output_dir, f)
        if os.path.isfile(fpath) and f.endswith('.xlsx'):
            size_kb = os.path.getsize(fpath) / 1024
            print(f"  📄 {f} ({size_kb:.0f} KB)")


if __name__ == '__main__':
    main()
