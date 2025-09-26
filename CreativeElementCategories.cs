using System;
using System.Collections.Generic;
using System.Linq;

// クリエイティブ要素カテゴリを一元管理するヘルパ
// 追加・名称変更はここを書き換えるだけで他ロジックへ反映されるよう SimpleBot などから参照する
public static class CreativeElementCategories
{
    // 利用順序も保持したいので配列（変更時は5要素制約ロジックも見直すこと）
    public static readonly string[] All = new[]
    {
        "状況",
        "課題・欲求",
        "感情",
        "温度感",
        "場・舞台・空間"
    };
}
