using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Extensibility;
using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using WordFormatHelper.AddIn.Engine;
using WordFormatHelper.AddIn.Models;
using WordFormatHelper.AddIn.Presets;
using WordFormatHelper.AddIn.Forms;

namespace WordFormatHelper.AddIn
{
    [ComVisible(true)]
    [Guid("8A3F2B1C-4D5E-4A6B-9C7D-8E2F1A3B5C4D")]
    [ProgId("WordFormatHelper.AddIn")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Connect : IDTExtensibility2, IRibbonExtensibility
    {
        private Word.Application _wordApp;
        private IRibbonUI _ribbonUI;

        static Connect()
        {
            Log("Add-in assembly loaded");
        }

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                _wordApp = (Word.Application)application;
                Log("OnConnection: Word " + _wordApp.Version);
            }
            catch (Exception ex)
            {
                Log("OnConnection FAILED: " + ex);
                throw;
            }
        }

        private static void Log(string msg)
        {
            try
            {
                string path = System.IO.Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData),
                    "WordFormatHelper", "addin.log");
                System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(path));
                System.IO.File.AppendAllText(path,
                    System.DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + " " + msg + System.Environment.NewLine);
            }
            catch { }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _wordApp = null;
            _ribbonUI = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var resourceName = "WordFormatHelper.AddIn.Ribbon.xml";
                using (var stream = asm.GetManifestResourceStream(resourceName))
                {
                    if (stream == null) return string.Empty;
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch
            {
                return string.Empty;
            }
        }

        public void OnRibbonLoad(IRibbonUI ribbonUI)
        {
            _ribbonUI = ribbonUI;
            Log("OnRibbonLoad called");
        }

        public void OnApplyPreset(IRibbonControl control)
        {
            Log("OnApplyPreset called, tag=" + (control.Tag ?? "null"));
            try
            {
                if (_wordApp == null || _wordApp.ActiveDocument == null)
                {
                    MessageBox.Show("请先打开一个 Word 文档。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var presetType = PresetTypeExtensions.FromTag(control.Tag);
                var settings = PresetManager.GetPreset(presetType);
                FormatEngine.ApplyToDocument(_wordApp, _wordApp.ActiveDocument, settings);
            }
            catch (Exception ex)
            {
                Log("OnApplyPreset FAILED: " + ex);
                MessageBox.Show("应用格式时出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnOpenSettings(IRibbonControl control)
        {
            Log("OnOpenSettings called");
            try
            {
                using (var dlg = new FormatSettingsDialog())
                {
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Log("OnOpenSettings FAILED: " + ex);
                MessageBox.Show("打开格式设置时出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnApplyStyle(IRibbonControl control)
        {
            Log("OnApplyStyle called, tag=" + (control.Tag ?? "null"));
            try
            {
                if (_wordApp == null || _wordApp.ActiveDocument == null)
                {
                    MessageBox.Show("请先打开一个 Word 文档。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var tag = control.Tag ?? "body";
                int outlineLevel = GetOutlineLevel(tag);

                // 段落级别按钮：只修改大纲级别，不修改字体字号等其他格式
                var sel = _wordApp.Selection;
                if (sel != null && sel.Range != null)
                {
                    try
                    {
                        foreach (Word.Paragraph para in sel.Range.Paragraphs)
                        {
                            para.OutlineLevel = (Word.WdOutlineLevel)outlineLevel;
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Log("OnApplyStyle FAILED: " + ex);
                MessageBox.Show("应用样式时出错：" + ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>根据按钮 tag 返回 Word 大纲级别。1-9 为标题级别，10 为正文级别(wdOutlineLevelBodyText)。</summary>
        private static int GetOutlineLevel(string tag)
        {
            switch (tag)
            {
                case "h1": return 1;
                case "h2": return 2;
                case "h3": return 3;
                case "h4": return 4;
                case "body": return 10; // wdOutlineLevelBodyText，清除标题级别，从导航中移除
                default: return 10;
            }
        }
    }
}
