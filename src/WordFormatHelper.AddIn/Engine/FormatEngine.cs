using System;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using WordFormatHelper.AddIn.Models;

namespace WordFormatHelper.AddIn.Engine
{
    internal static class FormatEngine
    {
        public static void ApplyToDocument(Word.Application app, Word.Document doc, FormatSettings settings)
        {
            bool prevScreenUpdating = app.ScreenUpdating;
            app.ScreenUpdating = false;

            object undoName = "应用中文报告格式";
            Word.UndoRecord undoRecord = null;
            try
            {
                try
                {
                    undoRecord = app.UndoRecord;
                    undoRecord.StartCustomRecord((string)undoName);
                }
                catch
                {
                }

                // 引擎调用链与格式设置 UI 完全对应，无隐藏调整：
                // - ParagraphFormatter：标题/正文格式（跳过表格）
                // - TableFormatter：表格格式
                // - ApplyPunctuation：标点转换（跳过表格）
                try { ParagraphFormatter.FormatDocument(doc, settings); } catch { }
                try { TableFormatter.Format(doc, settings.TableSettings); } catch { }
                try { ApplyPunctuation(doc, settings.PunctuationMode); } catch { }
            }
            finally
            {
                try { undoRecord?.EndCustomRecord(); } catch { }
                app.ScreenUpdating = prevScreenUpdating;
            }
        }

        // 标点映射：半角 → 全角
        private static readonly string[] HalfPuncts = { ",", ".", ";", ":", "!", "?" };
        private static readonly string[] FullPuncts = { "，", "。", "；", "：", "！", "？" };

        private static void ApplyPunctuation(Word.Document doc, PunctuationMode mode)
        {
            if (mode == PunctuationMode.Keep) return;

            if (mode == PunctuationMode.FullWidth || mode == PunctuationMode.HalfWidth)
            {
                // 遍历非表格段落，逐个替换标点（避免影响表格内数字格式）
                string[] fromPuncts = mode == PunctuationMode.FullWidth ? HalfPuncts : FullPuncts;
                string[] toPuncts = mode == PunctuationMode.FullWidth ? FullPuncts : HalfPuncts;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    if (IsInTable(para)) continue;
                    try
                    {
                        var range = para.Range;
                        string text = range.Text;
                        if (string.IsNullOrEmpty(text)) continue;

                        string newText = text;
                        for (int i = 0; i < fromPuncts.Length; i++)
                        {
                            newText = newText.Replace(fromPuncts[i], toPuncts[i]);
                        }

                        if (newText != text)
                        {
                            try { range.Text = newText; } catch { }
                        }
                    }
                    catch { }
                }
                return;
            }

            // Smart 模式：中文上下文用中文标点，英文/数字上下文用英文标点
            ApplySmartPunctuation(doc);
        }

        /// <summary>智能标点：遍历每个段落，根据标点前后字符是中文还是英文/数字决定用全角还是半角。
        /// 跳过表格内段落，避免破坏表格数字格式（表格格式由 TableFormatter 单独处理）。</summary>
        private static void ApplySmartPunctuation(Word.Document doc)
        {
            foreach (Word.Paragraph para in doc.Paragraphs)
            {
                // 跳过表格内段落，避免破坏表格数字格式
                if (IsInTable(para)) continue;

                try
                {
                    var range = para.Range;
                    string text = range.Text;
                    if (string.IsNullOrEmpty(text)) continue;

                    string newText = SmartConvertPunctuation(text);
                    if (newText != text)
                    {
                        try { range.Text = newText; } catch { }
                    }
                }
                catch { }
            }
        }

        /// <summary>判断段落是否位于表格单元格内。</summary>
        private static bool IsInTable(Word.Paragraph para)
        {
            try
            {
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
            try { return para.Range.Tables.Count > 0; } catch { }
            return false;
        }

        /// <summary>智能标点转换：中文上下文→全角，英文/数字上下文→半角。</summary>
        private static string SmartConvertPunctuation(string text)
        {
            var sb = new StringBuilder(text);
            for (int i = 0; i < sb.Length; i++)
            {
                char c = sb[i];

                // Smart 模式不处理 ./。 这一对，因为小数点和句号歧义太大，
                // 容易把数字后的句号错误转成小数点（如 "1000。" → "1000."）
                if (c == '.' || c == '。') continue;

                int halfIdx = Array.IndexOf(HalfPuncts, c.ToString());
                int fullIdx = Array.IndexOf(FullPuncts, c.ToString());

                if (halfIdx < 0 && fullIdx < 0) continue;

                // 判断前一个非空格字符是否中文
                char? prevChar = GetPrevNonSpaceChar(sb, i);
                char? nextChar = GetNextNonSpaceChar(sb, i);

                bool isChineseContext = IsChineseChar(prevChar) || IsChineseChar(nextChar);
                bool isEnglishContext = IsEnglishOrDigitChar(prevChar) || IsEnglishOrDigitChar(nextChar);

                // 中文上下文 → 全角；英文/数字上下文 → 半角
                if (isChineseContext && !isEnglishContext)
                {
                    // 转全角
                    if (halfIdx >= 0) sb[i] = FullPuncts[halfIdx][0];
                }
                else if (isEnglishContext && !isChineseContext)
                {
                    // 转半角
                    if (fullIdx >= 0) sb[i] = HalfPuncts[fullIdx][0];
                }
                // 两者都有或都没有，保持原样
            }
            return sb.ToString();
        }

        private static char? GetPrevNonSpaceChar(StringBuilder sb, int index)
        {
            for (int i = index - 1; i >= 0; i--)
            {
                if (!char.IsWhiteSpace(sb[i])) return sb[i];
            }
            return null;
        }

        private static char? GetNextNonSpaceChar(StringBuilder sb, int index)
        {
            for (int i = index + 1; i < sb.Length; i++)
            {
                if (!char.IsWhiteSpace(sb[i])) return sb[i];
            }
            return null;
        }

        private static bool IsChineseChar(char? c)
        {
            if (!c.HasValue) return false;
            char ch = c.Value;
            // CJK 统一汉字、扩展A、兼容汉字
            return (ch >= '\u4e00' && ch <= '\u9fff') ||
                   (ch >= '\u3400' && ch <= '\u4dbf') ||
                   (ch >= '\uf900' && ch <= '\ufaff');
        }

        private static bool IsEnglishOrDigitChar(char? c)
        {
            if (!c.HasValue) return false;
            char ch = c.Value;
            return (ch >= 'a' && ch <= 'z') ||
                   (ch >= 'A' && ch <= 'Z') ||
                   (ch >= '0' && ch <= '9');
        }
    }
}

