using System;
using Word = Microsoft.Office.Interop.Word;
using WordFormatHelper.AddIn.Models;

namespace WordFormatHelper.AddIn.Engine
{
    internal static class ParagraphFormatter
    {
        public static void FormatDocument(Word.Document doc, FormatSettings settings)
        {
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                // 跳过表格内的段落（表格单元格格式由 TableFormatter 单独处理，
                // 避免给单元格应用正文的首行缩进等设置）
                if (IsInTable(para)) continue;

                var fmt = DetectElementFormat(para, settings);
                if (fmt == null) continue;

                var range = para.Range;
                FontFormatter.ApplyFont(range, fmt);
                FontFormatter.ApplyParagraphFormat(range, fmt);
            }
        }

        /// <summary>判断段落是否位于表格单元格内。用多种方式检测，确保可靠。</summary>
        private static bool IsInTable(Word.Paragraph para)
        {
            try
            {
                // 方式1：用 Information[wdWithInTable]
                object result = para.Range.Information[Word.WdInformation.wdWithInTable];
                if (result != null)
                {
                    if (result is bool b) return b;
                    if (result is int i) return i != 0;
                    if (result is short s) return s != 0;
                    try { return Convert.ToBoolean(result); } catch { }
                }
            }
            catch { }

            try
            {
                // 方式2：检查段落范围内是否包含表格
                return para.Range.Tables.Count > 0;
            }
            catch { }

            return false;
        }

        private static ElementFormat DetectElementFormat(Word.Paragraph para, FormatSettings settings)
        {
            int level = 0;
            try { level = (int)para.OutlineLevel; } catch { }
            if (level >= 1 && level <= 4)
            {
                switch (level)
                {
                    case 1: return settings.Heading1;
                    case 2: return settings.Heading2;
                    case 3: return settings.Heading3;
                    case 4: return settings.Heading4;
                }
            }

            string styleName = "";
            try { styleName = ((Word.Style)para.get_Style()).NameLocal ?? ""; } catch { }

            if (styleName.Contains("标题 1") || styleName.Contains("Heading 1")) return settings.Heading1;
            if (styleName.Contains("标题 2") || styleName.Contains("Heading 2")) return settings.Heading2;
            if (styleName.Contains("标题 3") || styleName.Contains("Heading 3")) return settings.Heading3;
            if (styleName.Contains("标题 4") || styleName.Contains("Heading 4")) return settings.Heading4;
            if (styleName.Contains("标题") || styleName.Contains("Heading")) return settings.Heading4;

            return settings.BodyText;
        }
    }
}
