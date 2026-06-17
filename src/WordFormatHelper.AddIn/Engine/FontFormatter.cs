using System;
using Word = Microsoft.Office.Interop.Word;
using WordFormatHelper.AddIn.Models;

namespace WordFormatHelper.AddIn.Engine
{
    internal static class FontFormatter
    {
        public static void ApplyFont(Word.Range range, ElementFormat fmt)
        {
            // 只设置 UI 中显示的字段：中文字体、英文字体、字号、加粗
            var font = range.Font;
            font.NameFarEast = ResolveChineseFont(fmt.ChineseFont);
            font.NameAscii = fmt.EnglishFont;
            font.NameOther = fmt.EnglishFont;
            font.Name = fmt.EnglishFont;
            font.Size = fmt.FontSize;
            font.Bold = fmt.Bold ? 1 : 0;
        }

        public static void ApplyParagraphFormat(Word.Range range, ElementFormat fmt)
        {
            // 只设置 UI 中显示的字段：对齐、行距规则、行距、段前、段后、首行缩进
            var pf = range.ParagraphFormat;
            pf.Alignment = ToWordAlignment(fmt.Alignment);

            switch (fmt.LineSpacingRule)
            {
                case LineSpacingRule.Single:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    break;
                case LineSpacingRule.OnePointFive:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpace1pt5;
                    break;
                case LineSpacingRule.Double:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceDouble;
                    break;
                case LineSpacingRule.AtLeast:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast;
                    pf.LineSpacing = fmt.LineSpacing;
                    break;
                case LineSpacingRule.Exactly:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly;
                    pf.LineSpacing = fmt.LineSpacing;
                    break;
                case LineSpacingRule.Multiple:
                    pf.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple;
                    pf.LineSpacing = fmt.LineSpacing;
                    break;
            }

            pf.SpaceBefore = fmt.SpaceBefore;
            pf.SpaceAfter = fmt.SpaceAfter;

            if (fmt.FirstLineIndentChars > 0)
            {
                pf.CharacterUnitFirstLineIndent = fmt.FirstLineIndentChars;
            }
            else
            {
                pf.CharacterUnitFirstLineIndent = 0f;
                pf.FirstLineIndent = 0f;
            }
        }

        public static Word.WdParagraphAlignment ToWordAlignment(AlignmentType a)
        {
            switch (a)
            {
                case AlignmentType.Left: return Word.WdParagraphAlignment.wdAlignParagraphLeft;
                case AlignmentType.Center: return Word.WdParagraphAlignment.wdAlignParagraphCenter;
                case AlignmentType.Right: return Word.WdParagraphAlignment.wdAlignParagraphRight;
                case AlignmentType.Justify: return Word.WdParagraphAlignment.wdAlignParagraphJustify;
                default: return Word.WdParagraphAlignment.wdAlignParagraphJustify;
            }
        }

        public static string ResolveChineseFont(string fontName)
        {
            string[] fallbacks;
            if (fontName == "方正小标宋简体")
                fallbacks = new[] { "方正小标宋简体", "方正小标宋_GBK", "华文中宋", "黑体", "宋体" };
            else if (fontName == "仿宋_GB2312")
                fallbacks = new[] { "仿宋_GB2312", "仿宋_GB2312", "仿宋", "FangSong", "宋体" };
            else if (fontName == "楷体_GB2312")
                fallbacks = new[] { "楷体_GB2312", "楷体", "KaiTi", "宋体" };
            else
                fallbacks = new[] { fontName, "宋体" };

            foreach (var f in fallbacks)
            {
                if (IsFontInstalled(f)) return f;
            }
            return "宋体";
        }

        public static bool IsFontInstalled(string fontName)
        {
            try
            {
                using (var family = new System.Drawing.FontFamily(fontName))
                {
                    return family.Name == fontName || string.Equals(family.Name, fontName, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }
        }

        public static int ParseColor(string color)
        {
            if (string.IsNullOrEmpty(color)) return -16777216;
            if (color.ToLowerInvariant() == "auto") return -16777216;
            try
            {
                if (color.StartsWith("#"))
                {
                    var hex = color.Substring(1);
                    int r = Convert.ToInt32(hex.Substring(0, 2), 16);
                    int g = Convert.ToInt32(hex.Substring(2, 2), 16);
                    int b = Convert.ToInt32(hex.Substring(4, 2), 16);
                    return r + (g << 8) + (b << 16);
                }
                return Convert.ToInt32(color);
            }
            catch
            {
                return -16777216;
            }
        }
    }
}
