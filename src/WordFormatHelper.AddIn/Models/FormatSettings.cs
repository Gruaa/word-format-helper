using System;
using System.Xml.Serialization;

namespace WordFormatHelper.AddIn.Models
{
    [Serializable]
    [XmlRoot("FormatSettings")]
    public class FormatSettings
    {
        [XmlElement("Heading1")] public ElementFormat Heading1 { get; set; } = new ElementFormat();
        [XmlElement("Heading2")] public ElementFormat Heading2 { get; set; } = new ElementFormat();
        [XmlElement("Heading3")] public ElementFormat Heading3 { get; set; } = new ElementFormat();
        [XmlElement("Heading4")] public ElementFormat Heading4 { get; set; } = new ElementFormat();
        [XmlElement("BodyText")] public ElementFormat BodyText { get; set; } = new ElementFormat();

        [XmlElement("TableSettings")] public TableSettings TableSettings { get; set; } = new TableSettings();

        [XmlElement("PunctuationMode")] public PunctuationMode PunctuationMode { get; set; } = PunctuationMode.FullWidth;

        public FormatSettings Clone()
        {
            var copy = new FormatSettings
            {
                Heading1 = Heading1.Clone(),
                Heading2 = Heading2.Clone(),
                Heading3 = Heading3.Clone(),
                Heading4 = Heading4.Clone(),
                BodyText = BodyText.Clone(),
                TableSettings = TableSettings.Clone(),
                PunctuationMode = PunctuationMode
            };
            return copy;
        }

        public ElementFormat GetElement(string tag)
        {
            switch (tag)
            {
                case "h1": return Heading1;
                case "h2": return Heading2;
                case "h3": return Heading3;
                case "h4": return Heading4;
                case "body": return BodyText;
                default: return BodyText;
            }
        }
    }
}
