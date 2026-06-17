namespace WordFormatHelper.AddIn.Models
{
    public enum PresetType
    {
        Standard,
        Audit,
        Custom
    }

    public static class PresetTypeExtensions
    {
        public static string GetDisplayName(this PresetType type)
        {
            switch (type)
            {
                case PresetType.Standard: return "公文报告";
                case PresetType.Audit: return "审计报告";
                case PresetType.Custom: return "自定义格式";
                default: return type.ToString();
            }
        }

        public static string GetTag(this PresetType type)
        {
            switch (type)
            {
                case PresetType.Standard: return "standard";
                case PresetType.Audit: return "audit";
                case PresetType.Custom: return "custom";
                default: return type.ToString().ToLowerInvariant();
            }
        }

        public static PresetType FromTag(string tag)
        {
            switch ((tag ?? "").ToLowerInvariant())
            {
                case "standard": return PresetType.Standard;
                case "audit": return PresetType.Audit;
                case "custom": return PresetType.Custom;
                default: return PresetType.Standard;
            }
        }
    }
}
