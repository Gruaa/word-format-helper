using System;
using System.IO;
using System.Xml.Serialization;
using WordFormatHelper.AddIn.Models;

namespace WordFormatHelper.AddIn.Presets
{
    internal static class PresetManager
    {
        private static readonly string _appDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordFormatHelper");

        private static readonly string _presetsFile = Path.Combine(_appDir, "presets.xml");

        private static FormatSettingsContainer _container;

        private static FormatSettingsContainer Container
        {
            get
            {
                if (_container == null)
                {
                    _container = LoadContainer();
                }
                return _container;
            }
        }

        private static FormatSettingsContainer LoadContainer()
        {
            try
            {
                if (File.Exists(_presetsFile))
                {
                    var serializer = new XmlSerializer(typeof(FormatSettingsContainer));
                    using (var fs = File.OpenRead(_presetsFile))
                    {
                        var loaded = (FormatSettingsContainer)serializer.Deserialize(fs);
                        // 版本不匹配时重置为新默认值，确保升级后用户能看到新预设
                        if (loaded == null || loaded.Version != FormatSettingsContainer.CurrentVersion)
                        {
                            return CreateFreshContainer();
                        }
                        return loaded;
                    }
                }
            }
            catch
            {
            }
            return CreateFreshContainer();
        }

        private static FormatSettingsContainer CreateFreshContainer()
        {
            return new FormatSettingsContainer
            {
                Version = FormatSettingsContainer.CurrentVersion,
                Standard = DefaultPresets.CreateStandard(),
                Audit = DefaultPresets.CreateAudit(),
                Custom = DefaultPresets.CreateStandard()
            };
        }

        private static void SaveContainer(FormatSettingsContainer container)
        {
            try
            {
                Directory.CreateDirectory(_appDir);
                var serializer = new XmlSerializer(typeof(FormatSettingsContainer));
                using (var fs = File.Create(_presetsFile))
                {
                    serializer.Serialize(fs, container);
                }
            }
            catch
            {
            }
        }

        public static FormatSettings GetPreset(PresetType type)
        {
            switch (type)
            {
                case PresetType.Standard: return Container.Standard;
                case PresetType.Audit: return Container.Audit;
                case PresetType.Custom: return Container.Custom;
                default: return Container.Standard;
            }
        }

        public static void SavePreset(PresetType type, FormatSettings settings)
        {
            var c = Container;
            switch (type)
            {
                case PresetType.Standard: c.Standard = settings; break;
                case PresetType.Audit: c.Audit = settings; break;
                case PresetType.Custom: c.Custom = settings; break;
            }
            _container = c;
            SaveContainer(c);
        }

        public static void ResetToDefault(PresetType type)
        {
            SavePreset(type, DefaultPresets.Create(type));
        }
    }

    [Serializable]
    [XmlRoot("Presets")]
    public class FormatSettingsContainer
    {
        // 预设格式版本号：修改默认预设时递增，确保升级后重置为新默认值
        public const string CurrentVersion = "4";

        [XmlAttribute("version")] public string Version { get; set; } = CurrentVersion;
        [XmlElement("Standard")] public FormatSettings Standard { get; set; }
        [XmlElement("Audit")] public FormatSettings Audit { get; set; }
        [XmlElement("Custom")] public FormatSettings Custom { get; set; }
    }
}
