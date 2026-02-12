#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
台灣大樂透分析工具 - 快速示範

這是一個簡化的示範腳本，展示如何使用 lotto_analyzer.py
"""

from lotto_analyzer import TaiwanLotto649Analyzer


def quick_demo():
    """快速示範：執行完整分析流程"""
    
    print("=== 台灣大樂透分析工具 - 快速示範 ===\n")
    
    # 1. 建立分析器
    analyzer = TaiwanLotto649Analyzer()
    
    # 2. 獲取歷史數據（使用模擬數據）
    print("步驟 1: 獲取歷史開獎數據")
    analyzer.fetch_historical_data(max_draws=300)
    
    # 3. 分析號碼頻率
    print("\n步驟 2: 分析號碼出現頻率")
    analyzer.analyze_frequency()
    
    # 4. 顯示前10名熱門號碼
    print("\n步驟 3: 顯示熱門號碼")
    print("\n【TOP 10 熱門號碼】")
    top_10 = analyzer.get_top_numbers(10)
    for rank, (number, count) in enumerate(top_10, 1):
        print(f"  {rank}. 號碼 {number:02d} - 出現 {count} 次")
    
    # 5. 生成推薦組合
    print("\n步驟 4: 生成推薦號碼組合")
    combinations = analyzer.generate_recommended_combinations(3)
    
    print("\n【推薦號碼組合】")
    for idx, (strategy, numbers) in enumerate(combinations, 1):
        numbers_str = " - ".join(f"{n:02d}" for n in numbers)
        print(f"  組合 {idx} ({strategy}): {numbers_str}")
    
    # 6. 匯出結果
    print("\n步驟 5: 匯出分析結果")
    analyzer.export_results("demo_results.txt")
    
    print("\n✅ 示範完成！")
    print("💡 提示：執行 python lotto_analyzer.py 可進行完整分析")


def custom_analysis_example():
    """自訂分析示範"""
    
    print("\n=== 自訂分析示範 ===\n")
    
    analyzer = TaiwanLotto649Analyzer()
    analyzer.fetch_historical_data(max_draws=200)
    frequency = analyzer.analyze_frequency()
    
    # 自訂分析：找出特定範圍的熱門號碼
    print("分析不同區間的熱門號碼：")
    
    zones = [
        (1, 16, "小號區 (01-16)"),
        (17, 33, "中號區 (17-33)"),
        (34, 49, "大號區 (34-49)")
    ]
    
    for start, end, name in zones:
        zone_numbers = {n: count for n, count in frequency.items() if start <= n <= end}
        top_in_zone = sorted(zone_numbers.items(), key=lambda x: x[1], reverse=True)[:5]
        
        print(f"\n{name}:")
        for number, count in top_in_zone:
            print(f"  號碼 {number:02d}: {count} 次")


if __name__ == "__main__":
    # 執行快速示範
    quick_demo()
    
    # 執行自訂分析
    custom_analysis_example()
    
    print("\n" + "="*60)
    print("⚠️  重要提醒：")
    print("   本工具僅供統計分析與學術研究使用")
    print("   樂透為隨機遊戲，歷史數據不代表未來結果")
    print("   請理性購彩，量力而為")
    print("="*60)
