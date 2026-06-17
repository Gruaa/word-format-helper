using System;
using System.Collections.Generic;

namespace WordFormatHelper.AddIn.Models
{
    /// <summary>
    /// 中文字号与磅值的映射工具。
    /// Word 中文字号：初号/小初/一号/小一/二号/小二/三号/小三/四号/小四/五号/小五/六号/小六/七号/八号
    /// </summary>
    internal static class ChineseFontSize
    {
        // 中文字号名称 -> 磅值（按 Word 标准）
        private static readonly KeyValuePair<string, float>[] _entries =
        {
            new KeyValuePair<string, float>("初号", 42f),
            new KeyValuePair<string, float>("小初", 36f),
            new KeyValuePair<string, float>("一号", 26f),
            new KeyValuePair<string, float>("小一", 24f),
            new KeyValuePair<string, float>("二号", 22f),
            new KeyValuePair<string, float>("小二", 18f),
            new KeyValuePair<string, float>("三号", 16f),
            new KeyValuePair<string, float>("小三", 15f),
            new KeyValuePair<string, float>("四号", 14f),
            new KeyValuePair<string, float>("小四", 12f),
            new KeyValuePair<string, float>("五号", 10.5f),
            new KeyValuePair<string, float>("小五", 9f),
            new KeyValuePair<string, float>("六号", 7.5f),
            new KeyValuePair<string, float>("小六", 6.5f),
            new KeyValuePair<string, float>("七号", 5.5f),
            new KeyValuePair<string, float>("八号", 5f)
        };

        /// <summary>所有中文字号名称（从大到小），用于下拉框。</summary>
        public static string[] AllNames
        {
            get
            {
                var names = new string[_entries.Length];
                for (int i = 0; i < _entries.Length; i++) names[i] = _entries[i].Key;
                return names;
            }
        }

        /// <summary>根据磅值获取中文字号名称；找不到时返回最接近的名称。</summary>
        public static string GetName(float size)
        {
            for (int i = 0; i < _entries.Length; i++)
            {
                if (Math.Abs(_entries[i].Value - size) < 0.05f) return _entries[i].Key;
            }
            // 找最接近的
            float minDiff = float.MaxValue;
            int minIdx = 0;
            for (int i = 0; i < _entries.Length; i++)
            {
                float d = Math.Abs(_entries[i].Value - size);
                if (d < minDiff) { minDiff = d; minIdx = i; }
            }
            return _entries[minIdx].Key;
        }

        /// <summary>根据中文字号名称获取磅值；找不到时返回三号(16)。</summary>
        public static float GetSize(string name)
        {
            if (string.IsNullOrEmpty(name)) return 16f;
            for (int i = 0; i < _entries.Length; i++)
            {
                if (_entries[i].Key == name) return _entries[i].Value;
            }
            return 16f;
        }
    }
}
