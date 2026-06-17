using Word = Microsoft.Office.Interop.Word;
using WordFormatHelper.AddIn.Models;

namespace WordFormatHelper.AddIn.Engine
{
    internal static class TableFormatter
    {
        public static void Format(Word.Document doc, TableSettings settings)
        {
            foreach (Word.Table table in doc.Tables)
            {
                FormatTable(table, settings);
            }
        }

        private static void FormatTable(Word.Table table, TableSettings settings)
        {
            try
            {
                // 表格对齐
                table.Rows.Alignment = ToRowAlignment(settings.TableAlignment);

                // 自动调整
                try
                {
                    if (settings.AutoFitMode == 1)
                        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
                    else if (settings.AutoFitMode == 2)
                        table.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow);
                }
                catch { }

                // 6 条独立边框
                ApplyBorder(table.Borders[Word.WdBorderType.wdBorderTop], settings.BorderTop);
                ApplyBorder(table.Borders[Word.WdBorderType.wdBorderBottom], settings.BorderBottom);
                ApplyBorder(table.Borders[Word.WdBorderType.wdBorderLeft], settings.BorderLeft);
                ApplyBorder(table.Borders[Word.WdBorderType.wdBorderRight], settings.BorderRight);
                ApplyBorder(table.Borders[Word.WdBorderType.wdBorderHorizontal], settings.BorderHorizontal);
                ApplyBorder(table.Borders[Word.WdBorderType.wdBorderVertical], settings.BorderVertical);

                // 表头重复
                if (settings.HeaderRepeat && table.Rows.Count > 0)
                {
                    try { table.Rows[1].HeadingFormat = -1; } catch { }
                }

                // 行格式化
                for (int r = 1; r <= table.Rows.Count; r++)
                {
                    var row = table.Rows[r];
                    bool isHeader = (r == 1);
                    var font = isHeader ? settings.HeaderFont : settings.CellFont;
                    var fontSize = isHeader ? settings.HeaderFontSize : settings.CellFontSize;
                    var bold = isHeader ? settings.HeaderBold : settings.CellBold;
                    var align = isHeader ? settings.HeaderAlignment : settings.CellAlignment;
                    var backColor = isHeader ? settings.HeaderBackColor : null;
                    var vAlign = settings.CellVerticalAlign;

                    foreach (Word.Cell cell in row.Cells)
                    {
                        try
                        {
                            // 垂直对齐
                            cell.VerticalAlignment = ToCellVAlign(vAlign);

                            var range = cell.Range;
                            var f = range.Font;
                            // 只设置 UI 中显示的字段：字体、字号、加粗
                            f.NameFarEast = FontFormatter.ResolveChineseFont(font);
                            f.Size = fontSize;
                            f.Bold = bold ? 1 : 0;

                            range.ParagraphFormat.Alignment = FontFormatter.ToWordAlignment(align);

                            // 表头底色
                            try
                            {
                                if (isHeader && !string.IsNullOrEmpty(backColor) && backColor != "auto")
                                {
                                    cell.Shading.BackgroundPatternColor = (Word.WdColor)FontFormatter.ParseColor(backColor);
                                }
                            }
                            catch { }
                        }
                        catch { }
                    }
                }
            }
            catch
            {
            }
        }

        private static void ApplyBorder(Word.Border border, BorderEdge edge)
        {
            try
            {
                if (edge.Enabled)
                {
                    border.LineStyle = ToLineStyle(edge.Style);
                    border.LineWidth = ToLineWidth(edge.Width);
                    border.Color = (Word.WdColor)FontFormatter.ParseColor(edge.Color);
                }
                else
                {
                    border.LineStyle = Word.WdLineStyle.wdLineStyleNone;
                }
            }
            catch { }
        }

        private static Word.WdLineStyle ToLineStyle(int style)
        {
            switch (style)
            {
                case 0: return Word.WdLineStyle.wdLineStyleSingle;
                case 1: return Word.WdLineStyle.wdLineStyleDashSmallGap;
                case 2: return Word.WdLineStyle.wdLineStyleDot;
                case 3: return Word.WdLineStyle.wdLineStyleDashDot;
                case 4: return Word.WdLineStyle.wdLineStyleDouble;
                case 5: return Word.WdLineStyle.wdLineStyleTriple;
                case 6: return Word.WdLineStyle.wdLineStyleSingleWavy;
                default: return Word.WdLineStyle.wdLineStyleSingle;
            }
        }

        private static Word.WdLineWidth ToLineWidth(float width)
        {
            if (width <= 0.25f) return Word.WdLineWidth.wdLineWidth025pt;
            if (width <= 0.5f) return Word.WdLineWidth.wdLineWidth050pt;
            if (width <= 0.75f) return Word.WdLineWidth.wdLineWidth075pt;
            if (width <= 1f) return Word.WdLineWidth.wdLineWidth100pt;
            if (width <= 1.5f) return Word.WdLineWidth.wdLineWidth150pt;
            if (width <= 2.25f) return Word.WdLineWidth.wdLineWidth225pt;
            if (width <= 3f) return Word.WdLineWidth.wdLineWidth300pt;
            return Word.WdLineWidth.wdLineWidth450pt;
        }

        private static Word.WdRowAlignment ToRowAlignment(AlignmentType align)
        {
            switch (align)
            {
                case AlignmentType.Left: return Word.WdRowAlignment.wdAlignRowLeft;
                case AlignmentType.Center: return Word.WdRowAlignment.wdAlignRowCenter;
                case AlignmentType.Right: return Word.WdRowAlignment.wdAlignRowRight;
                default: return Word.WdRowAlignment.wdAlignRowCenter;
            }
        }

        private static Word.WdCellVerticalAlignment ToCellVAlign(int vAlign)
        {
            switch (vAlign)
            {
                case 0: return Word.WdCellVerticalAlignment.wdCellAlignVerticalTop;
                case 2: return Word.WdCellVerticalAlignment.wdCellAlignVerticalBottom;
                default: return Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
            }
        }
    }
}
