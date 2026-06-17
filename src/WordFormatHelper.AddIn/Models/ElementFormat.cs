using System;
using System.Xml.Serialization;

namespace WordFormatHelper.AddIn.Models
{
    public enum AlignmentType
    {
        Left,
        Center,
        Right,
        Justify,
        Distribute
    }

    public enum LineSpacingRule
    {
        Single,
        OnePointFive,
        Double,
        AtLeast,
        Exactly,
        Multiple
    }

    public enum PunctuationMode
    {
        Keep,
        FullWidth,
        HalfWidth,
        Smart // 中文上下文用中文标点，英文/数字上下文用英文标点
    }

    /// <summary>标题/正文等段落元素格式（精简版：10 个核心字段）。</summary>
    [Serializable]
    public class ElementFormat
    {
        [XmlElement("ChineseFont")] public string ChineseFont { get; set; } = "仿宋";
        [XmlElement("EnglishFont")] public string EnglishFont { get; set; } = "Times New Roman";
        [XmlElement("FontSize")] public float FontSize { get; set; } = 16f;
        [XmlElement("Bold")] public bool Bold { get; set; } = false;
        [XmlElement("Alignment")] public AlignmentType Alignment { get; set; } = AlignmentType.Justify;
        [XmlElement("LineSpacingRule")] public LineSpacingRule LineSpacingRule { get; set; } = LineSpacingRule.Exactly;
        [XmlElement("LineSpacing")] public float LineSpacing { get; set; } = 28f;
        [XmlElement("SpaceBefore")] public float SpaceBefore { get; set; } = 0f;
        [XmlElement("SpaceAfter")] public float SpaceAfter { get; set; } = 0f;
        [XmlElement("FirstLineIndentChars")] public float FirstLineIndentChars { get; set; } = 0f;

        public ElementFormat Clone()
        {
            return (ElementFormat)MemberwiseClone();
        }
    }

    /// <summary>单条边框的设置。</summary>
    [Serializable]
    public class BorderEdge
    {
        [XmlElement("Enabled")] public bool Enabled { get; set; } = true;
        [XmlElement("Style")] public int Style { get; set; } = 0; // 0=单实线 1=虚线 2=点线 3=点划线 4=双线 5=三线 6=波浪线
        [XmlElement("Width")] public float Width { get; set; } = 0.5f;
        [XmlElement("Color")] public string Color { get; set; } = "auto";

        public BorderEdge Clone()
        {
            return (BorderEdge)MemberwiseClone();
        }
    }

    /// <summary>表格设置（精简版：12 个核心字段 + 6 条独立边框）。</summary>
    [Serializable]
    public class TableSettings
    {
        // 表头（6 个）
        [XmlElement("HeaderFont")] public string HeaderFont { get; set; } = "黑体";
        [XmlElement("HeaderFontSize")] public float HeaderFontSize { get; set; } = 10.5f;
        [XmlElement("HeaderBold")] public bool HeaderBold { get; set; } = true;
        [XmlElement("HeaderAlignment")] public AlignmentType HeaderAlignment { get; set; } = AlignmentType.Center;
        [XmlElement("HeaderBackColor")] public string HeaderBackColor { get; set; } = "#D9D9D9";
        [XmlElement("HeaderRepeat")] public bool HeaderRepeat { get; set; } = true;

        // 单元格（5 个）
        [XmlElement("CellFont")] public string CellFont { get; set; } = "仿宋";
        [XmlElement("CellFontSize")] public float CellFontSize { get; set; } = 10.5f;
        [XmlElement("CellBold")] public bool CellBold { get; set; } = false;
        [XmlElement("CellAlignment")] public AlignmentType CellAlignment { get; set; } = AlignmentType.Center;
        [XmlElement("CellVerticalAlign")] public int CellVerticalAlign { get; set; } = 1; // 0=顶 1=中 2=底

        // 表格整体（2 个）
        [XmlElement("TableAlignment")] public AlignmentType TableAlignment { get; set; } = AlignmentType.Center;
        [XmlElement("AutoFitMode")] public int AutoFitMode { get; set; } = 0; // 0=无 1=根据内容 2=根据窗口

        // 6 条独立边框
        [XmlElement("BorderTop")] public BorderEdge BorderTop { get; set; } = new BorderEdge();
        [XmlElement("BorderBottom")] public BorderEdge BorderBottom { get; set; } = new BorderEdge();
        [XmlElement("BorderLeft")] public BorderEdge BorderLeft { get; set; } = new BorderEdge();
        [XmlElement("BorderRight")] public BorderEdge BorderRight { get; set; } = new BorderEdge();
        [XmlElement("BorderHorizontal")] public BorderEdge BorderHorizontal { get; set; } = new BorderEdge();
        [XmlElement("BorderVertical")] public BorderEdge BorderVertical { get; set; } = new BorderEdge();

        public TableSettings Clone()
        {
            var c = (TableSettings)MemberwiseClone();
            c.BorderTop = BorderTop.Clone();
            c.BorderBottom = BorderBottom.Clone();
            c.BorderLeft = BorderLeft.Clone();
            c.BorderRight = BorderRight.Clone();
            c.BorderHorizontal = BorderHorizontal.Clone();
            c.BorderVertical = BorderVertical.Clone();
            return c;
        }
    }
}
