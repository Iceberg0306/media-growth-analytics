"""
一键运行全部分析 Pipeline
用法: python run_all.py
依次执行 4 个分析脚本，输出到 reports/ 目录
"""

import subprocess
import sys
import os

SCRIPTS = [
    ("Phase 1: 基础统计 + TOP10", "src/01_basic_stats.py"),
    ("Phase 2: TOP10 vs 倒数10 对比", "src/02_top10_comparison.py"),
    ("Phase 3: 全量特征分析", "src/03_feature_analysis.py"),
    ("Phase 4: 全维度 + 混淆变量控制", "src/04_full_pipeline.py"),
]

def main():
    project_root = os.path.dirname(os.path.abspath(__file__))
    os.chdir(project_root)
    print(f"Working directory: {project_root}\n")

    for label, script in SCRIPTS:
        print(f"\n{'='*60}")
        print(f"  {label}")
        print(f"{'='*60}\n")
        result = subprocess.run(
            [sys.executable, script],
            capture_output=True, text=True, encoding="utf-8", errors="replace"
        )
        if result.stdout:
            print(result.stdout[-500:])  # last 500 chars
        if result.returncode != 0:
            print(f"[WARN] {script} exited with code {result.returncode}")
            print(result.stderr[-300:] if result.stderr else "")

    print("\n\n全部分析完成。报告位于 reports/ 目录。")


if __name__ == "__main__":
    main()
