using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Win32;
using WordFormatHelper.AddIn.Models;
using WordFormatHelper.AddIn.Presets;

namespace WordFormatHelper.AddIn.Forms
{
    internal class FormatSettingsDialog : Form
    {
        private ComboBox _presetCombo;
        private TabControl _tabs;
        private Button _btnOk, _btnCancel, _btnReset;

        private PresetType _currentPreset = PresetType.Standard;
        private FormatSettings _settings;

        // 标题/正文编辑器
        private ElementFormatEditor _h1, _h2, _h3, _h4, _body;
        // 标点
        private RadioButton _punctKeep, _punctFull, _punctHalf, _punctSmart;
        // 表格（13 个字段）
        private ComboBox _tblHeaderFont, _tblCellFont, _tblHeaderSize, _tblCellSize;
        private CheckBox _tblHeaderBold, _tblCellBold, _tblHeaderRepeat;
        private ComboBox _tblHeaderAlign, _tblHeaderBackColor, _tblCellAlign, _tblCellVAlign, _tblTableAlign, _tblAutoFit;
        // 6 条独立边框
        private BorderEdgeEditor _borderTop, _borderBottom, _borderLeft, _borderRight, _borderH, _borderV;

        private static string[] _installedFonts;

        /// <summary>获取系统已安装字体列表（从注册表读取中文名），与 Word 字体下拉框一致。</summary>
        internal static string[] InstalledFonts
        {
            get
            {
                if (_installedFonts == null)
                    _installedFonts = LoadInstalledFonts();
                return _installedFonts;
            }
        }

        /// <summary>从注册表读取已安装字体列表，返回中文显示名。</summary>
        private static string[] LoadInstalledFonts()
        {
            var fonts = new List<string>();
            try
            {
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"))
                {
                    if (key != null)
                    {
                        foreach (var name in key.GetValueNames())
                        {
                            var fontName = name;
                            // 去掉后缀如 " (TrueType)" " (OpenType)"
                            var idx = fontName.LastIndexOf(" (");
                            if (idx > 0) fontName = fontName.Substring(0, idx);
                            // 过滤旋转字体（以@开头）
                            if (fontName.StartsWith("@")) continue;
                            if (!fonts.Contains(fontName)) fonts.Add(fontName);
                        }
                    }
                }
            }
            catch { }

            // 确保常用中文字体在列表中
            var commonFonts = new[] { "宋体", "黑体", "仿宋", "仿宋_GB2312", "楷体", "楷体_GB2312", "方正小标宋简体", "Times New Roman" };
            foreach (var f in commonFonts)
            {
                if (!fonts.Contains(f)) fonts.Add(f);
            }

            fonts.Sort();
            return fonts.ToArray();
        }

        /// <summary>创建字体下拉框（DropDown 样式，可输入搜索，与 Word 字体框一致）。</summary>
        private static ComboBox CreateFontComboBox()
        {
            var combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDown,
                Width = 140,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            combo.Items.AddRange(InstalledFonts);
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
            return combo;
        }

        private static readonly string[] ColorNames = { "自动", "黑色", "白色", "红色", "蓝色", "绿色", "灰色", "浅灰" };
        private static readonly string[] AlignNames = { "左对齐", "居中", "右对齐", "两端对齐", "分散对齐" };

        public FormatSettingsDialog()
        {
            _settings = PresetManager.GetPreset(_currentPreset);
            BuildUi();
            LoadSettings();
        }

        private void BuildUi()
        {
            Text = "格式设置";
            Width = 860;
            Height = 680;
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;

            // 顶部预设选择
            var topPanel = new Panel { Dock = DockStyle.Top, Height = 56, Padding = new Padding(16, 14, 16, 8) };
            var lblPreset = new Label { Text = "预设方案：", AutoSize = true, Location = new Point(16, 20) };
            _presetCombo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Location = new Point(100, 17), Width = 200 };
            _presetCombo.Items.AddRange(new object[] { "公文报告", "审计报告", "自定义格式" });
            _presetCombo.SelectedIndex = 0;
            _presetCombo.SelectedIndexChanged += (s, e) =>
            {
                SaveCurrentToMemory();
                _currentPreset = (PresetType)_presetCombo.SelectedIndex;
                _settings = PresetManager.GetPreset(_currentPreset);
                LoadSettings();
            };
            topPanel.Controls.Add(lblPreset);
            topPanel.Controls.Add(_presetCombo);

            // 底部按钮
            var bottomPanel = new TableLayoutPanel { Dock = DockStyle.Bottom, Height = 56, ColumnCount = 3, RowCount = 1 };
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            bottomPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110));
            _btnReset = new Button { Text = "恢复默认值", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Anchor = AnchorStyles.Left, Margin = new Padding(8, 14, 8, 8) };
            _btnReset.Click += (s, e) =>
            {
                if (MessageBox.Show("确定要将当前预设恢复为内置默认值吗？", "确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    _settings = DefaultPresets.Create(_currentPreset);
                    LoadSettings();
                }
            };
            _btnCancel = new Button { Text = "取消", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, DialogResult = DialogResult.Cancel, Anchor = AnchorStyles.Right, Margin = new Padding(8, 14, 8, 8) };
            _btnOk = new Button { Text = "确定", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, DialogResult = DialogResult.OK, Anchor = AnchorStyles.Right, Margin = new Padding(8, 14, 8, 8) };
            _btnOk.Click += (s, e) => SaveCurrentToMemory();
            bottomPanel.Controls.Add(_btnReset, 0, 0);
            bottomPanel.Controls.Add(_btnCancel, 1, 0);
            bottomPanel.Controls.Add(_btnOk, 2, 0);

            // 标签页（顺序：标题→正文→表格）
            _tabs = new TabControl { Dock = DockStyle.Fill, Padding = new Point(12, 8) };
            BuildHeadingsTab();
            BuildBodyTab();
            BuildTableTab();

            // 添加顺序：先 Fill，再 Top/Bottom，确保 Dock 布局正确
            Controls.Add(_tabs);
            Controls.Add(topPanel);
            Controls.Add(bottomPanel);

            AcceptButton = _btnOk;
            CancelButton = _btnCancel;
        }

        // ==================== 标题标签页 ====================
        private void BuildHeadingsTab()
        {
            var page = new TabPage("标题");
            var scroll = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(10)
            };

            _h1 = new ElementFormatEditor("一级标题");
            _h2 = new ElementFormatEditor("二级标题");
            _h3 = new ElementFormatEditor("三级标题");
            _h4 = new ElementFormatEditor("四级标题");
            scroll.Controls.AddRange(new Control[] { _h1, _h2, _h3, _h4 });
            page.Controls.Add(scroll);
            _tabs.TabPages.Add(page);
        }

        // ==================== 正文标签页 ====================
        private void BuildBodyTab()
        {
            var page = new TabPage("正文");
            var scroll = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(10)
            };

            _body = new ElementFormatEditor("正文格式");
            scroll.Controls.Add(_body);

            // 标点设置
            var punctBox = new GroupBox
            {
                Text = "标点设置",
                Width = 760,
                Height = 140,
                Padding = new Padding(16, 28, 16, 12)
            };
            _punctKeep = new RadioButton { Text = "保持原样", AutoSize = true, Location = new Point(20, 36) };
            _punctFull = new RadioButton { Text = "全部转全角中文标点", AutoSize = true, Location = new Point(180, 36) };
            _punctHalf = new RadioButton { Text = "全部转半角英文标点", AutoSize = true, Location = new Point(400, 36) };
            _punctSmart = new RadioButton { Text = "智能转换（推荐）：中文用中文标点，英文/数字用英文标点", AutoSize = true, Location = new Point(20, 70) };
            var tip = new Label
            {
                Text = "智能转换：根据标点前后的字符类型自动选择全角或半角标点",
                AutoSize = true,
                ForeColor = Color.Gray,
                Location = new Point(20, 100)
            };
            punctBox.Controls.AddRange(new Control[] { _punctKeep, _punctFull, _punctHalf, _punctSmart, tip });
            scroll.Controls.Add(punctBox);

            page.Controls.Add(scroll);
            _tabs.TabPages.Add(page);
        }

        // ==================== 表格标签页 ====================
        private void BuildTableTab()
        {
            var page = new TabPage("表格");
            var flow = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(10)
            };

            // 表头（两排显示，2行3列）
            flow.Controls.Add(CreateSectionLabel("—— 表头设置 ——"));
            flow.Controls.Add(BuildTableHeaderPanel());

            // 单元格（两排显示，2行3列）
            flow.Controls.Add(CreateSectionLabel("—— 单元格设置 ——"));
            flow.Controls.Add(BuildTableCellPanel());

            // 表格整体
            flow.Controls.Add(CreateSectionLabel("—— 表格整体 ——"));
            _tblTableAlign = CreateLabeledCombo("表格对齐:", AlignNames, out var l12); flow.Controls.Add(CreateRowPanel(l12, _tblTableAlign));
            _tblAutoFit = CreateLabeledCombo("自动调整:", new[] { "无", "根据内容自动调整", "根据窗口自动调整" }, out var l13); flow.Controls.Add(CreateRowPanel(l13, _tblAutoFit));

            // 6 条独立边框
            flow.Controls.Add(CreateSectionLabel("—— 上边框 ——"));
            _borderTop = new BorderEdgeEditor(); flow.Controls.Add(_borderTop);
            flow.Controls.Add(CreateSectionLabel("—— 下边框 ——"));
            _borderBottom = new BorderEdgeEditor(); flow.Controls.Add(_borderBottom);
            flow.Controls.Add(CreateSectionLabel("—— 左边框 ——"));
            _borderLeft = new BorderEdgeEditor(); flow.Controls.Add(_borderLeft);
            flow.Controls.Add(CreateSectionLabel("—— 右边框 ——"));
            _borderRight = new BorderEdgeEditor(); flow.Controls.Add(_borderRight);
            flow.Controls.Add(CreateSectionLabel("—— 内部横线 ——"));
            _borderH = new BorderEdgeEditor(); flow.Controls.Add(_borderH);
            flow.Controls.Add(CreateSectionLabel("—— 内部竖线 ——"));
            _borderV = new BorderEdgeEditor(); flow.Controls.Add(_borderV);

            page.Controls.Add(flow);
            _tabs.TabPages.Add(page);
        }

        // 表头两排面板（2行3列）
        private Panel BuildTableHeaderPanel()
        {
            _tblHeaderFont = CreateLabeledFontCombo("表头字体:", out var l1);
            _tblHeaderSize = CreateLabeledCombo("表头字号:", ChineseFontSize.AllNames, out var l2);
            _tblHeaderBold = CreateLabeledCheck("表头加粗:", out var l3);
            _tblHeaderAlign = CreateLabeledCombo("表头对齐:", AlignNames, out var l4);
            _tblHeaderBackColor = CreateLabeledCombo("表头底色:", ColorNames, out var l5);
            _tblHeaderRepeat = CreateLabeledCheck("表头重复:", out var l6);
            return BuildTwoRowPanel(l1, _tblHeaderFont, l2, _tblHeaderSize, l3, _tblHeaderBold,
                                     l4, _tblHeaderAlign, l5, _tblHeaderBackColor, l6, _tblHeaderRepeat);
        }

        // 单元格两排面板（2行3列，第6格留空）
        private Panel BuildTableCellPanel()
        {
            _tblCellFont = CreateLabeledFontCombo("单元格字体:", out var l1);
            _tblCellSize = CreateLabeledCombo("单元格字号:", ChineseFontSize.AllNames, out var l2);
            _tblCellBold = CreateLabeledCheck("单元格加粗:", out var l3);
            _tblCellAlign = CreateLabeledCombo("单元格对齐:", AlignNames, out var l4);
            _tblCellVAlign = CreateLabeledCombo("垂直对齐:", new[] { "顶端对齐", "垂直居中", "底端对齐" }, out var l5);
            return BuildTwoRowPanel(l1, _tblCellFont, l2, _tblCellSize, l3, _tblCellBold,
                                     l4, _tblCellAlign, l5, _tblCellVAlign, null, null);
        }

        // ==================== 通用辅助方法 ====================
        private static Label CreateSectionLabel(string text)
        {
            return new Label
            {
                Text = text,
                AutoSize = true,
                ForeColor = Color.SteelBlue,
                Font = new Font(SystemFonts.DefaultFont, FontStyle.Bold),
                Margin = new Padding(0, 10, 0, 4),
                Width = 760
            };
        }

        private static Panel CreateRowPanel(Control label, Control value)
        {
            var p = new Panel { Width = 760, Height = 30, Margin = new Padding(0, 2, 0, 2) };
            label.Location = new Point(0, 5);
            label.AutoSize = true;
            value.Location = new Point(160, 3);
            p.Controls.Add(label);
            p.Controls.Add(value);
            return p;
        }

        /// <summary>构建两排三列面板（6 个字段分两行显示）。</summary>
        private static Panel BuildTwoRowPanel(
            Control l1, Control v1, Control l2, Control v2, Control l3, Control v3,
            Control l4, Control v4, Control l5, Control v5, Control l6, Control v6)
        {
            var p = new Panel { Width = 760, Height = 95, Margin = new Padding(0, 2, 0, 2) };
            // 第一排
            PlaceField(p, l1, v1, 0, 0);
            PlaceField(p, l2, v2, 260, 0);
            PlaceField(p, l3, v3, 520, 0);
            // 第二排
            PlaceField(p, l4, v4, 0, 40);
            PlaceField(p, l5, v5, 260, 40);
            if (l6 != null && v6 != null) PlaceField(p, l6, v6, 520, 40);
            return p;
        }

        private static void PlaceField(Panel parent, Control label, Control value, int x, int y)
        {
            label.Location = new Point(x, y + 5);
            label.AutoSize = true;
            value.Location = new Point(x + 90, y + 3);
            parent.Controls.Add(label);
            parent.Controls.Add(value);
        }

        private static ComboBox CreateLabeledCombo(string labelText, string[] items, out Label label)
        {
            label = new Label { Text = labelText, AutoSize = true };
            var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 140 };
            combo.Items.AddRange(items);
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
            return combo;
        }

        private static CheckBox CreateLabeledCheck(string labelText, out Label label)
        {
            label = new Label { Text = labelText, AutoSize = true };
            return new CheckBox { AutoSize = true };
        }

        /// <summary>创建带标签的字体下拉框（DropDown 样式，可输入搜索）。</summary>
        private static ComboBox CreateLabeledFontCombo(string labelText, out Label label)
        {
            label = new Label { Text = labelText, AutoSize = true };
            return CreateFontComboBox();
        }

        private static NumericUpDown CreateLabeledNum(string labelText, float value, out Label label)
        {
            label = new Label { Text = labelText, AutoSize = true };
            return new NumericUpDown { Width = 100, DecimalPlaces = 1, Minimum = 0, Maximum = 5000, Value = (decimal)value };
        }

        // ==================== 加载/保存 ====================
        private void LoadSettings()
        {
            _h1.LoadFrom(_settings.Heading1);
            _h2.LoadFrom(_settings.Heading2);
            _h3.LoadFrom(_settings.Heading3);
            _h4.LoadFrom(_settings.Heading4);
            _body.LoadFrom(_settings.BodyText);

            switch (_settings.PunctuationMode)
            {
                case PunctuationMode.Keep: _punctKeep.Checked = true; break;
                case PunctuationMode.FullWidth: _punctFull.Checked = true; break;
                case PunctuationMode.HalfWidth: _punctHalf.Checked = true; break;
                case PunctuationMode.Smart: _punctSmart.Checked = true; break;
            }

            // 表格
            SelectComboByText(_tblHeaderFont, _settings.TableSettings.HeaderFont);
            SelectSizeCombo(_tblHeaderSize, _settings.TableSettings.HeaderFontSize);
            _tblHeaderBold.Checked = _settings.TableSettings.HeaderBold;
            _tblHeaderAlign.SelectedIndex = (int)_settings.TableSettings.HeaderAlignment;
            SelectComboByText(_tblHeaderBackColor, ColorNameToDisplay(_settings.TableSettings.HeaderBackColor));
            _tblHeaderRepeat.Checked = _settings.TableSettings.HeaderRepeat;

            SelectComboByText(_tblCellFont, _settings.TableSettings.CellFont);
            SelectSizeCombo(_tblCellSize, _settings.TableSettings.CellFontSize);
            _tblCellBold.Checked = _settings.TableSettings.CellBold;
            _tblCellAlign.SelectedIndex = (int)_settings.TableSettings.CellAlignment;
            _tblCellVAlign.SelectedIndex = _settings.TableSettings.CellVerticalAlign;

            _tblTableAlign.SelectedIndex = (int)_settings.TableSettings.TableAlignment;
            _tblAutoFit.SelectedIndex = _settings.TableSettings.AutoFitMode;

            _borderTop.LoadFrom(_settings.TableSettings.BorderTop);
            _borderBottom.LoadFrom(_settings.TableSettings.BorderBottom);
            _borderLeft.LoadFrom(_settings.TableSettings.BorderLeft);
            _borderRight.LoadFrom(_settings.TableSettings.BorderRight);
            _borderH.LoadFrom(_settings.TableSettings.BorderHorizontal);
            _borderV.LoadFrom(_settings.TableSettings.BorderVertical);
        }

        private void SaveCurrentToMemory()
        {
            _h1.SaveTo(_settings.Heading1);
            _h2.SaveTo(_settings.Heading2);
            _h3.SaveTo(_settings.Heading3);
            _h4.SaveTo(_settings.Heading4);
            _body.SaveTo(_settings.BodyText);

            if (_punctKeep.Checked) _settings.PunctuationMode = PunctuationMode.Keep;
            else if (_punctFull.Checked) _settings.PunctuationMode = PunctuationMode.FullWidth;
            else if (_punctHalf.Checked) _settings.PunctuationMode = PunctuationMode.HalfWidth;
            else if (_punctSmart.Checked) _settings.PunctuationMode = PunctuationMode.Smart;

            // 表格
            _settings.TableSettings.HeaderFont = string.IsNullOrEmpty(_tblHeaderFont.Text) ? "黑体" : _tblHeaderFont.Text;
            _settings.TableSettings.HeaderFontSize = _tblHeaderSize.SelectedItem != null
                ? ChineseFontSize.GetSize(_tblHeaderSize.SelectedItem.ToString()) : 10.5f;
            _settings.TableSettings.HeaderBold = _tblHeaderBold.Checked;
            _settings.TableSettings.HeaderAlignment = (AlignmentType)_tblHeaderAlign.SelectedIndex;
            _settings.TableSettings.HeaderBackColor = ColorComboToName(_tblHeaderBackColor.SelectedItem?.ToString() ?? "浅灰");
            _settings.TableSettings.HeaderRepeat = _tblHeaderRepeat.Checked;

            _settings.TableSettings.CellFont = string.IsNullOrEmpty(_tblCellFont.Text) ? "仿宋" : _tblCellFont.Text;
            _settings.TableSettings.CellFontSize = _tblCellSize.SelectedItem != null
                ? ChineseFontSize.GetSize(_tblCellSize.SelectedItem.ToString()) : 10.5f;
            _settings.TableSettings.CellBold = _tblCellBold.Checked;
            _settings.TableSettings.CellAlignment = (AlignmentType)_tblCellAlign.SelectedIndex;
            _settings.TableSettings.CellVerticalAlign = _tblCellVAlign.SelectedIndex;

            _settings.TableSettings.TableAlignment = (AlignmentType)_tblTableAlign.SelectedIndex;
            _settings.TableSettings.AutoFitMode = _tblAutoFit.SelectedIndex;

            _borderTop.SaveTo(_settings.TableSettings.BorderTop);
            _borderBottom.SaveTo(_settings.TableSettings.BorderBottom);
            _borderLeft.SaveTo(_settings.TableSettings.BorderLeft);
            _borderRight.SaveTo(_settings.TableSettings.BorderRight);
            _borderH.SaveTo(_settings.TableSettings.BorderHorizontal);
            _borderV.SaveTo(_settings.TableSettings.BorderVertical);

            PresetManager.SavePreset(_currentPreset, _settings);
        }

        private static void SelectComboByText(ComboBox combo, string text)
        {
            if (combo == null || string.IsNullOrEmpty(text)) return;
            int idx = combo.FindStringExact(text);
            if (idx >= 0) combo.SelectedIndex = idx;
            else combo.Text = text; // 对于 DropDown 样式（字体），直接设置文本
        }

        private static void SelectSizeCombo(ComboBox combo, float size)
        {
            if (combo == null) return;
            string name = ChineseFontSize.GetName(size);
            int idx = combo.FindStringExact(name);
            if (idx >= 0) combo.SelectedIndex = idx;
            else if (combo.Items.Count > 0) combo.SelectedIndex = 0;
        }

        public static string ColorNameToDisplay(string color)
        {
            switch (color?.ToUpperInvariant())
            {
                case "#000000": return "黑色";
                case "#FFFFFF": return "白色";
                case "#FF0000": return "红色";
                case "#0000FF": return "蓝色";
                case "#008000": return "绿色";
                case "#808080": return "灰色";
                case "#D9D9D9": return "浅灰";
                default: return "自动";
            }
        }

        public static string ColorComboToName(string displayName)
        {
            switch (displayName)
            {
                case "自动": return "auto";
                case "黑色": return "#000000";
                case "白色": return "#FFFFFF";
                case "红色": return "#FF0000";
                case "蓝色": return "#0000FF";
                case "绿色": return "#008000";
                case "灰色": return "#808080";
                case "浅灰": return "#D9D9D9";
                default: return "auto";
            }
        }
    }

    /// <summary>元素格式编辑器（标题、正文通用，10 个核心字段，5 行 2 列）。</summary>
    internal class ElementFormatEditor : GroupBox
    {
        private ComboBox _cnFont, _enFont, _size, _align, _lineSpacingRule;
        private CheckBox _bold;
        private NumericUpDown _lineSpacing, _spaceBefore, _spaceAfter, _firstIndent;

        public ElementFormatEditor(string title)
        {
            Text = title;
            Width = 760;
            Height = 210;

            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 5,
                Padding = new Padding(12, 20, 12, 8)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 120));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 38));
            for (int i = 0; i < 5; i++)
                panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            _cnFont = CreateFontCombo();
            _enFont = CreateFontCombo();
            _size = CreateSizeCombo();
            _bold = new CheckBox { AutoSize = true };
            _align = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
            _align.Items.AddRange(new object[] { "左对齐", "居中", "右对齐", "两端对齐", "分散对齐" });
            _align.SelectedIndex = 3;
            _lineSpacingRule = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
            _lineSpacingRule.Items.AddRange(new object[] { "单倍", "1.5倍", "2倍", "最小值", "固定值", "多倍" });
            _lineSpacingRule.SelectedIndex = 4;
            // 行距规则联动：单倍/1.5倍/2倍时禁用行距磅数
            _lineSpacingRule.SelectedIndexChanged += (s, e) => UpdateLineSpacingEnabled();
            _lineSpacing = new NumericUpDown { Width = 80, DecimalPlaces = 1, Minimum = 0, Maximum = 200, Value = 28 };
            _spaceBefore = new NumericUpDown { Width = 80, DecimalPlaces = 1, Minimum = 0, Maximum = 200, Value = 0 };
            _spaceAfter = new NumericUpDown { Width = 80, DecimalPlaces = 1, Minimum = 0, Maximum = 200, Value = 0 };
            _firstIndent = new NumericUpDown { Width = 80, DecimalPlaces = 1, Minimum = 0, Maximum = 20, Value = 0 };

            int row = 0;
            AddRow(panel, row++, "中文字体:", _cnFont, "英文字体:", _enFont);
            AddRow(panel, row++, "字号:", _size, "加粗:", _bold);
            AddRow(panel, row++, "对齐:", _align, "行距规则:", _lineSpacingRule);
            AddRow(panel, row++, "行距(磅):", _lineSpacing, "段前(磅):", _spaceBefore);
            AddRow(panel, row++, "段后(磅):", _spaceAfter, "首行缩进(字符):", _firstIndent);

            Controls.Add(panel);
        }

        /// <summary>行距规则联动：单倍(0)/1.5倍(1)/2倍(2)不需要磅数，禁用输入框。</summary>
        private void UpdateLineSpacingEnabled()
        {
            bool needPoints = _lineSpacingRule.SelectedIndex >= 3; // 最小值/固定值/多倍
            _lineSpacing.Enabled = needPoints;
            if (!needPoints) _lineSpacing.Value = 0;
        }

        private static ComboBox CreateFontCombo()
        {
            var combo = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDown,
                Width = 140,
                AutoCompleteMode = AutoCompleteMode.SuggestAppend,
                AutoCompleteSource = AutoCompleteSource.ListItems
            };
            combo.Items.AddRange(FormatSettingsDialog.InstalledFonts);
            if (combo.Items.Count > 0) combo.SelectedIndex = 0;
            return combo;
        }

        private static ComboBox CreateSizeCombo()
        {
            var combo = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 90 };
            combo.Items.AddRange(ChineseFontSize.AllNames);
            int idx = combo.FindStringExact("三号");
            combo.SelectedIndex = idx >= 0 ? idx : 0;
            return combo;
        }

        private static void AddRow(TableLayoutPanel panel, int row, string l1, Control c1, string l2, Control c2)
        {
            panel.Controls.Add(new Label { Text = l1, AutoSize = true, Anchor = AnchorStyles.Left }, 0, row);
            if (c1 != null) panel.Controls.Add(c1, 1, row);
            if (!string.IsNullOrEmpty(l2))
            {
                panel.Controls.Add(new Label { Text = l2, AutoSize = true, Anchor = AnchorStyles.Left }, 2, row);
                if (c2 != null) panel.Controls.Add(c2, 3, row);
            }
        }

        public void LoadFrom(ElementFormat fmt)
        {
            SelectByText(_cnFont, fmt.ChineseFont);
            SelectByText(_enFont, fmt.EnglishFont);
            SelectSize(_size, fmt.FontSize);
            _bold.Checked = fmt.Bold;
            _align.SelectedIndex = (int)fmt.Alignment;
            _lineSpacingRule.SelectedIndex = (int)fmt.LineSpacingRule;
            _lineSpacing.Value = (decimal)fmt.LineSpacing;
            _spaceBefore.Value = (decimal)fmt.SpaceBefore;
            _spaceAfter.Value = (decimal)fmt.SpaceAfter;
            _firstIndent.Value = (decimal)fmt.FirstLineIndentChars;
            UpdateLineSpacingEnabled();
        }

        public void SaveTo(ElementFormat fmt)
        {
            fmt.ChineseFont = string.IsNullOrEmpty(_cnFont.Text) ? "仿宋" : _cnFont.Text;
            fmt.EnglishFont = string.IsNullOrEmpty(_enFont.Text) ? "Times New Roman" : _enFont.Text;
            fmt.FontSize = _size.SelectedItem != null ? ChineseFontSize.GetSize(_size.SelectedItem.ToString()) : 16f;
            fmt.Bold = _bold.Checked;
            fmt.Alignment = (AlignmentType)_align.SelectedIndex;
            fmt.LineSpacingRule = (LineSpacingRule)_lineSpacingRule.SelectedIndex;
            fmt.LineSpacing = (float)_lineSpacing.Value;
            fmt.SpaceBefore = (float)_spaceBefore.Value;
            fmt.SpaceAfter = (float)_spaceAfter.Value;
            fmt.FirstLineIndentChars = (float)_firstIndent.Value;
        }

        private static void SelectByText(ComboBox combo, string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            int idx = combo.FindStringExact(text);
            if (idx >= 0) combo.SelectedIndex = idx;
            else combo.Text = text; // 字体不在列表中，直接设置文本（DropDown 样式）
        }

        private static void SelectSize(ComboBox combo, float size)
        {
            string name = ChineseFontSize.GetName(size);
            int idx = combo.FindStringExact(name);
            if (idx >= 0) combo.SelectedIndex = idx;
            else if (combo.Items.Count > 0) combo.SelectedIndex = 0;
        }
    }

    /// <summary>单条边框编辑器。</summary>
    internal class BorderEdgeEditor : GroupBox
    {
        private CheckBox _enabled;
        private ComboBox _style, _color;
        private NumericUpDown _width;

        public BorderEdgeEditor()
        {
            Width = 760;
            Height = 115;
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 2,
                Padding = new Padding(12, 20, 12, 8)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 100));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 30));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 30));

            _enabled = new CheckBox { AutoSize = true };
            _style = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120 };
            _style.Items.AddRange(new object[] { "单实线", "虚线", "点线", "点划线", "双线", "三线", "波浪线" });
            _style.SelectedIndex = 0;
            _width = new NumericUpDown { Width = 80, DecimalPlaces = 2, Minimum = 0.25m, Maximum = 6, Value = 0.5m, Increment = 0.25m };
            _color = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 110 };
            _color.Items.AddRange(new object[] { "自动", "黑色", "白色", "红色", "蓝色", "绿色", "灰色", "浅灰" });
            _color.SelectedIndex = 0;

            panel.Controls.Add(new Label { Text = "启用:", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0);
            panel.Controls.Add(_enabled, 1, 0);
            panel.Controls.Add(new Label { Text = "线型:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 0);
            panel.Controls.Add(_style, 3, 0);
            panel.Controls.Add(new Label { Text = "粗细(磅):", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 1);
            panel.Controls.Add(_width, 1, 1);
            panel.Controls.Add(new Label { Text = "颜色:", AutoSize = true, Anchor = AnchorStyles.Left }, 2, 1);
            panel.Controls.Add(_color, 3, 1);

            Controls.Add(panel);
        }

        public void LoadFrom(BorderEdge edge)
        {
            _enabled.Checked = edge.Enabled;
            _style.SelectedIndex = edge.Style;
            _width.Value = (decimal)edge.Width;
            SelectByText(_color, FormatSettingsDialog.ColorNameToDisplay(edge.Color));
        }

        public void SaveTo(BorderEdge edge)
        {
            edge.Enabled = _enabled.Checked;
            edge.Style = _style.SelectedIndex;
            edge.Width = (float)_width.Value;
            edge.Color = FormatSettingsDialog.ColorComboToName(_color.SelectedItem?.ToString() ?? "自动");
        }

        private static void SelectByText(ComboBox combo, string text)
        {
            int idx = combo.FindStringExact(text);
            if (idx >= 0) combo.SelectedIndex = idx;
            else combo.SelectedIndex = 0;
        }
    }
}
