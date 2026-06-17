using WordFormatHelper.AddIn.Models;

namespace WordFormatHelper.AddIn.Presets
{
    /// <summary>
    /// 预设默认值：
    /// - 公文报告：严格按 GB/T 9704-2012《党政机关公文格式》
    /// - 审计报告：按审计行业惯例
    /// </summary>
    internal static class DefaultPresets
    {
        /// <summary>公文报告（GB/T 9704-2012）。</summary>
        public static FormatSettings CreateStandard()
        {
            var s = new FormatSettings();

            // 一级标题：3号黑体，不加粗
            s.Heading1 = new ElementFormat
            {
                ChineseFont = "黑体",
                EnglishFont = "Times New Roman",
                FontSize = 16f, // 三号
                Bold = false,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.Exactly,
                LineSpacing = 28f,
                SpaceBefore = 6f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 2f
            };

            // 二级标题：3号楷体，加粗
            s.Heading2 = new ElementFormat
            {
                ChineseFont = "楷体_GB2312",
                EnglishFont = "Times New Roman",
                FontSize = 16f, // 三号
                Bold = true,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.Exactly,
                LineSpacing = 28f,
                SpaceBefore = 6f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 2f
            };

            // 三级标题：3号仿宋，加粗
            s.Heading3 = new ElementFormat
            {
                ChineseFont = "仿宋_GB2312",
                EnglishFont = "Times New Roman",
                FontSize = 16f, // 三号
                Bold = true,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.Exactly,
                LineSpacing = 28f,
                SpaceBefore = 6f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 2f
            };

            // 四级标题：3号仿宋，不加粗
            s.Heading4 = new ElementFormat
            {
                ChineseFont = "仿宋_GB2312",
                EnglishFont = "Times New Roman",
                FontSize = 16f, // 三号
                Bold = false,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.Exactly,
                LineSpacing = 28f,
                SpaceBefore = 6f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 2f
            };

            // 正文：3号仿宋_GB2312，固定行距28磅，首行缩进2字符
            s.BodyText = new ElementFormat
            {
                ChineseFont = "仿宋_GB2312",
                EnglishFont = "Times New Roman",
                FontSize = 16f, // 三号
                Bold = false,
                Alignment = AlignmentType.Justify,
                LineSpacingRule = LineSpacingRule.Exactly,
                LineSpacing = 28f,
                SpaceBefore = 0f,
                SpaceAfter = 0f,
                FirstLineIndentChars = 2f
            };

            // 表格：公文表格用三线表风格
            s.TableSettings = new TableSettings
            {
                HeaderFont = "黑体",
                HeaderFontSize = 14f, // 四号
                HeaderBold = false,
                HeaderAlignment = AlignmentType.Center,
                HeaderBackColor = "auto",
                HeaderRepeat = true,
                CellFont = "仿宋_GB2312",
                CellFontSize = 14f, // 四号
                CellBold = false,
                CellAlignment = AlignmentType.Center,
                CellVerticalAlign = 1,
                TableAlignment = AlignmentType.Center,
                AutoFitMode = 0,
                // 三线表：上下粗线，无左右边框，无内部竖线
                BorderTop = new BorderEdge { Enabled = true, Style = 0, Width = 1.5f, Color = "auto" },
                BorderBottom = new BorderEdge { Enabled = true, Style = 0, Width = 1.5f, Color = "auto" },
                BorderLeft = new BorderEdge { Enabled = false, Style = 0, Width = 0.5f, Color = "auto" },
                BorderRight = new BorderEdge { Enabled = false, Style = 0, Width = 0.5f, Color = "auto" },
                BorderHorizontal = new BorderEdge { Enabled = true, Style = 0, Width = 0.5f, Color = "auto" },
                BorderVertical = new BorderEdge { Enabled = false, Style = 0, Width = 0.5f, Color = "auto" }
            };

            s.PunctuationMode = PunctuationMode.Smart;
            return s;
        }

        /// <summary>审计报告（行业惯例）。</summary>
        public static FormatSettings CreateAudit()
        {
            var s = new FormatSettings();

            // 一级标题：黑体三号
            s.Heading1 = new ElementFormat
            {
                ChineseFont = "黑体",
                EnglishFont = "Times New Roman",
                FontSize = 16f, // 三号
                Bold = false,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.OnePointFive,
                LineSpacing = 0f,
                SpaceBefore = 12f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 0f
            };

            // 二级标题：楷体四号加粗
            s.Heading2 = new ElementFormat
            {
                ChineseFont = "楷体",
                EnglishFont = "Times New Roman",
                FontSize = 14f, // 四号
                Bold = true,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.OnePointFive,
                LineSpacing = 0f,
                SpaceBefore = 6f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 0f
            };

            // 三级标题：仿宋四号
            s.Heading3 = new ElementFormat
            {
                ChineseFont = "仿宋",
                EnglishFont = "Times New Roman",
                FontSize = 14f, // 四号
                Bold = false,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.OnePointFive,
                LineSpacing = 0f,
                SpaceBefore = 6f,
                SpaceAfter = 6f,
                FirstLineIndentChars = 0f
            };

            // 四级标题：仿宋小四
            s.Heading4 = new ElementFormat
            {
                ChineseFont = "仿宋",
                EnglishFont = "Times New Roman",
                FontSize = 12f, // 小四
                Bold = true,
                Alignment = AlignmentType.Left,
                LineSpacingRule = LineSpacingRule.OnePointFive,
                LineSpacing = 0f,
                SpaceBefore = 3f,
                SpaceAfter = 3f,
                FirstLineIndentChars = 0f
            };

            // 正文：宋体小四，1.5倍行距，首行缩进2字符
            s.BodyText = new ElementFormat
            {
                ChineseFont = "宋体",
                EnglishFont = "Times New Roman",
                FontSize = 12f, // 小四
                Bold = false,
                Alignment = AlignmentType.Justify,
                LineSpacingRule = LineSpacingRule.OnePointFive,
                LineSpacing = 0f,
                SpaceBefore = 0f,
                SpaceAfter = 0f,
                FirstLineIndentChars = 2f
            };

            // 表格：审计报告用网格表
            s.TableSettings = new TableSettings
            {
                HeaderFont = "黑体",
                HeaderFontSize = 12f, // 小四
                HeaderBold = true,
                HeaderAlignment = AlignmentType.Center,
                HeaderBackColor = "#D9D9D9",
                HeaderRepeat = true,
                CellFont = "宋体",
                CellFontSize = 12f, // 小四
                CellBold = false,
                CellAlignment = AlignmentType.Left,
                CellVerticalAlign = 1,
                TableAlignment = AlignmentType.Center,
                AutoFitMode = 0,
                BorderTop = new BorderEdge { Enabled = true, Style = 0, Width = 0.75f, Color = "auto" },
                BorderBottom = new BorderEdge { Enabled = true, Style = 0, Width = 0.75f, Color = "auto" },
                BorderLeft = new BorderEdge { Enabled = true, Style = 0, Width = 0.75f, Color = "auto" },
                BorderRight = new BorderEdge { Enabled = true, Style = 0, Width = 0.75f, Color = "auto" },
                BorderHorizontal = new BorderEdge { Enabled = true, Style = 0, Width = 0.5f, Color = "auto" },
                BorderVertical = new BorderEdge { Enabled = true, Style = 0, Width = 0.5f, Color = "auto" }
            };

            s.PunctuationMode = PunctuationMode.Smart;
            return s;
        }

        public static FormatSettings Create(PresetType type)
        {
            switch (type)
            {
                case PresetType.Standard: return CreateStandard();
                case PresetType.Audit: return CreateAudit();
                case PresetType.Custom: return CreateStandard();
                default: return CreateStandard();
            }
        }
    }
}
