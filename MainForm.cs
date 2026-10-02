using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace NssfPayroll
{
    public partial class MainForm : Form
    {
        private readonly AppData _data;
        private Establishment _est;            // المؤسسة المختارة بتبويب الأجور
        private PeriodData _current;
        private readonly PeriodContext _ctx = new PeriodContext();   // المؤسسة/الفترة/طريقة العرض الحالية
        private LineRowList _rows;
        private AutoCompleteStringCollection _nameSource = new AutoCompleteStringCollection();
        private bool _loading;
        private bool _dirty;
        private Font _boldFont;

        // System.Windows.Forms.Timer بالاسم الكامل لأن System.Threading.Timer بيتلخبط معو
        private readonly System.Windows.Forms.Timer _saveTimer = new System.Windows.Forms.Timer();

        private readonly TabControl _tabs = new TabControl();

        // ---- تبويب المؤسسات والنسب ----
        private readonly ListBox _lstEst = new ListBox();
        private readonly TextBox _txtEstName = new TextBox();
        private readonly TextBox _txtEstNssf = new TextBox();
        private readonly TextBox _txtOwner = new TextBox();
        private readonly TextBox _txtPhone = new TextBox();
        private readonly Label _lblEstStatus = new Label();
        private Button _btnSaveEst;
        private Button _btnEmployees;
        private Button _btnDeleteEst;
        private Establishment _editing;        // المؤسسة المعروضة بالحقول (ممكن تكون جديدة ولسا ما انحفظت)
        private bool _editingIsNew;
        private readonly NumericUpDown _nudSm = new NumericUpDown();
        private readonly NumericUpDown _nudEos = new NumericUpDown();
        private readonly NumericUpDown _nudFa = new NumericUpDown();

        // ---- تبويب أجور الفترة والطباعة ----
        private readonly ComboBox _cmbEst = new ComboBox();
        private readonly Label _lblNssf = new Label();
        private readonly NumericUpDown _nudYear = new NumericUpDown();
        private readonly ComboBox _cmbPeriod = new ComboBox();
        private readonly DateTimePicker _dtpDate = new DateTimePicker();
        private readonly DataGridView _grid = new DataGridView();
        private readonly DataGridView _grid70 = new DataGridView();   // "الفترة 70": تجميع سنوي 21+22+23+24 لكل موظف
        private readonly DataGridView _summary = new DataGridView();
        private DataGridViewTextBoxColumn _colName;
        private DataGridViewTextBoxColumn _colEos;
        private DataGridViewTextBoxColumn _colSm;
        private DataGridViewTextBoxColumn _colFa;
        private DataGridViewTextBoxColumn _colDue;    // اشتراكات مستحقة (محسوبة، بسقف 28 مليون شهري)
        private DataGridViewTextBoxColumn _colPaid;
        private DataGridViewTextBoxColumn _colNet;    // الاشتراكات الصافية = اشتراكات مستحقة − التعويض المدفوع
        private readonly ComboBox _cmbView = new ComboBox();     // الفصل كامل / شهر 1 / شهر 2 / شهر 3
        private readonly Label _lblGridHint = new Label();

        private Button _btnCopyPrev;
        private Button _btnInsertAll;
        private Button _btnOpen;
        private Button _btnPrint;
        private Button _btnPdf;

        // ---- تبويب الطباعة ----
        private readonly ComboBox _cmbPrintEmp = new ComboBox();
        private readonly NumericUpDown _nudPrintYear = new NumericUpDown();
        private readonly ComboBox _cmbPrintGranularity = new ComboBox();
        private readonly DataGridView _gridPrintPreview = new DataGridView();
        private Button _btnPrintEmpExcel;
        private Button _btnPrintEmpPrint;
        private Button _btnPrintEmpPdf;

        public MainForm()
        {
            _data = DataStore.Load();

            Text = "جداول الاشتراكات - الصندوق الوطني للضمان الاجتماعي (CNSS 190A)";
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Font = new Font("Segoe UI", 11F);
            ClientSize = new Size(1100, 720);
            MinimumSize = new Size(900, 600);
            StartPosition = FormStartPosition.CenterScreen;

            _boldFont = new Font(Font, FontStyle.Bold);
            _rows = new LineRowList(new List<EmployeeLine>(), _ctx);

            _saveTimer.Interval = 800;
            _saveTimer.Tick += delegate { _saveTimer.Stop(); SaveIfDirty(); };
            FormClosing += delegate { SaveIfDirty(); };

            BuildUi();

            _loading = true;
            LoadRatesToUi();
            int year = _data.LastYear > 0 ? _data.LastYear : DateTime.Today.Year;
            year = Math.Min(Math.Max(year, (int)_nudYear.Minimum), (int)_nudYear.Maximum);
            _nudYear.Value = year;
            _cmbPeriod.SelectedIndex = (_data.LastPeriod >= 1 && _data.LastPeriod <= 4) ? _data.LastPeriod - 1 : 0;
            _loading = false;

            RefreshEstList(_data.LastEstablishmentId);
            RefreshEstCombo(_data.LastEstablishmentId);
            RebuildNameSource();
            LoadPeriod();

            // أول مرة (ما في مؤسسات) بنفتح على تبويب المؤسسات
            _tabs.SelectedIndex = _data.Establishments.Count == 0 ? 0 : 1;
        }

        // =====================================================================
        //  بناء الواجهة
        // =====================================================================
        private void BuildUi()
        {
            _tabs.Dock = DockStyle.Fill;
            _tabs.RightToLeftLayout = true;

            TabPage tabCompany = new TabPage("المعلومات الأساسية");
            TabPage tabPeriod = new TabPage("احتساب الجداول");
            TabPage tabPrint = new TabPage("الطباعة");
            TabPage tabReports = new TabPage("المتابعة والتقارير");
            _tabs.TabPages.Add(tabCompany);
            _tabs.TabPages.Add(tabPeriod);
            _tabs.TabPages.Add(tabPrint);
            _tabs.TabPages.Add(tabReports);
            Controls.Add(_tabs);

            BuildCompanyTab(tabCompany);
            BuildPeriodTab(tabPeriod);
            BuildPrintTab(tabPrint);
            BuildReportsTab(tabReports);

            // كل مرة تفتح تبويب المتابعة والتقارير أو الطباعة، بيتحدّث (يلتقط أي تعديل صار بتبويب تاني)
            _tabs.SelectedIndexChanged += delegate
            {
                if (_tabs.SelectedTab == tabReports) RefreshReportsTab();
                else if (_tabs.SelectedTab == tabPrint) RefreshPrintTab();
            };
        }

        private static void AddRow(TableLayoutPanel table, string label, Control input)
        {
            Label l = new Label();
            l.Text = label;
            l.AutoSize = true;
            l.Margin = new Padding(6, 12, 12, 6);
            input.Margin = new Padding(6);
            table.Controls.Add(l);
            table.Controls.Add(input);
        }

        private TableLayoutPanel MakeFieldsTable()
        {
            TableLayoutPanel table = new TableLayoutPanel();
            table.Dock = DockStyle.Top;
            table.AutoSize = true;
            table.ColumnCount = 2;
            table.Padding = new Padding(10, 0, 10, 0);
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            return table;
        }

        private Label MakeHeader(string text, int topPadding)
        {
            Label l = new Label();
            l.Text = text;
            l.Font = _boldFont;
            l.AutoSize = true;
            l.Dock = DockStyle.Top;
            l.Padding = new Padding(10, topPadding, 10, 6);
            return l;
        }

        // ------------------------- تبويب المؤسسات والنسب -------------------------
        private void BuildCompanyTab(TabPage page)
        {
            // ---------- يمين: قائمة المؤسسات ----------
            Label lblList = new Label();
            lblList.Text = "المؤسسات";
            lblList.Font = _boldFont;
            lblList.AutoSize = true;
            lblList.Dock = DockStyle.Top;
            lblList.Padding = new Padding(4, 4, 4, 6);

            _lstEst.Dock = DockStyle.Fill;
            _lstEst.IntegralHeight = false;
            _lstEst.SelectedIndexChanged += delegate
            {
                if (_loading) return;
                Establishment sel = _lstEst.SelectedItem as Establishment;
                if (sel != null) ShowEstablishment(sel);
            };

            FlowLayoutPanel listButtons = new FlowLayoutPanel();
            listButtons.Dock = DockStyle.Bottom;
            listButtons.AutoSize = true;
            listButtons.Controls.Add(UiHelpers.MakeButton("مؤسسة جديدة", delegate { NewEstablishment(); }));
            _btnDeleteEst = UiHelpers.MakeButton("حذف المؤسسة", delegate { DeleteEstablishment(); });
            listButtons.Controls.Add(_btnDeleteEst);

            Panel listPanel = new Panel();
            listPanel.Dock = DockStyle.Fill;
            listPanel.Padding = new Padding(6);
            listPanel.Controls.Add(_lstEst);         // Fill أول
            listPanel.Controls.Add(listButtons);     // وبعدين Bottom
            listPanel.Controls.Add(lblList);         // وبعدين Top

            // ---------- بيانات المؤسسة ----------
            TableLayoutPanel fields = MakeFieldsTable();
            _txtEstName.Width = 420;
            _txtEstNssf.Width = 240;
            _txtOwner.Width = 420;
            _txtPhone.Width = 240;
            AddRow(fields, "اسم المؤسسة", _txtEstName);
            AddRow(fields, "رقم المؤسسة في الضمان", _txtEstNssf);
            AddRow(fields, "اسم صاحب المؤسسة", _txtOwner);
            AddRow(fields, "رقم تلفون صاحب المؤسسة", _txtPhone);

            FlowLayoutPanel estButtons = new FlowLayoutPanel();
            estButtons.Dock = DockStyle.Top;
            estButtons.AutoSize = true;
            estButtons.Padding = new Padding(10, 6, 10, 0);
            _btnSaveEst = UiHelpers.MakeButton("حفظ المؤسسة", delegate { SaveEstablishment(); });
            _btnEmployees = UiHelpers.MakeButton("موظفو المؤسسة (إضافة / تعديل)", delegate { OpenEmployees(); });
            estButtons.Controls.Add(_btnSaveEst);
            estButtons.Controls.Add(_btnEmployees);

            _lblEstStatus.Dock = DockStyle.Top;
            _lblEstStatus.AutoSize = true;
            _lblEstStatus.ForeColor = Color.DimGray;
            _lblEstStatus.Padding = new Padding(14, 6, 10, 6);

            EventHandler fieldChanged = delegate
            {
                if (_loading || _editing == null || _editingIsNew) return;
                SetEstStatus("فيه تعديلات لسا ما انحفظت، اضغط «حفظ المؤسسة».", true);
            };
            _txtEstName.TextChanged += fieldChanged;
            _txtEstNssf.TextChanged += fieldChanged;
            _txtOwner.TextChanged += fieldChanged;
            _txtPhone.TextChanged += fieldChanged;

            // ---------- النسب (واحدة لكل المؤسسات) ----------
            TableLayoutPanel rates = MakeFieldsTable();
            foreach (NumericUpDown n in new NumericUpDown[] { _nudSm, _nudEos, _nudFa })
            {
                n.DecimalPlaces = 2;
                n.Minimum = 0;
                n.Maximum = 100;
                n.Width = 110;
            }
            AddRow(rates, "نسبة المرض والأمومة (%)", _nudSm);
            AddRow(rates, "نسبة تعويض نهاية الخدمة (%)", _nudEos);
            AddRow(rates, "نسبة التعويضات العائلية (%)", _nudFa);

            _nudSm.ValueChanged += delegate
            {
                if (_loading) return;
                _data.Rates.SickMaternity = _nudSm.Value / 100m;
                MarkDirty();
                RefreshSummary();
            };
            _nudEos.ValueChanged += delegate
            {
                if (_loading) return;
                _data.Rates.EndOfService = _nudEos.Value / 100m;
                MarkDirty();
                RefreshSummary();
            };
            _nudFa.ValueChanged += delegate
            {
                if (_loading) return;
                _data.Rates.FamilyAllowance = _nudFa.Value / 100m;
                MarkDirty();
                RefreshSummary();
            };

            // ---------- تجميع القسم الأيسر (الأخير بينضاف بيصير فوق لأنو Dock=Top) ----------
            Panel details = new Panel();
            details.Dock = DockStyle.Fill;
            details.AutoScroll = true;
            details.Controls.Add(rates);
            details.Controls.Add(MakeHeader("النسب (بتنطبق على كل المؤسسات)", 18));
            details.Controls.Add(_lblEstStatus);
            details.Controls.Add(estButtons);
            details.Controls.Add(fields);
            details.Controls.Add(MakeHeader("بيانات المؤسسة", 8));

            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 2;
            root.RowCount = 1;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 300F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.Controls.Add(listPanel, 0, 0);
            root.Controls.Add(details, 1, 0);
            page.Controls.Add(root);
        }

        private void BuildPeriodTab(TabPage page)
        {
            // ---------- الشريط العلوي: المؤسسة / السنة / الفترة / التاريخ / طريقة الإدخال ----------
            _cmbEst.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbEst.Width = 300;

            _lblNssf.AutoSize = true;
            _lblNssf.ForeColor = Color.DarkGreen;
            _lblNssf.Margin = new Padding(10, 9, 10, 0);

            _nudYear.Minimum = 2000;
            _nudYear.Maximum = 2100;
            _nudYear.Width = 90;

            _cmbPeriod.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbPeriod.Width = 80;
            _cmbPeriod.Items.AddRange(new object[] { "21", "22", "23", "24", "70" });

            _dtpDate.Format = DateTimePickerFormat.Short;
            _dtpDate.Width = 140;

            _cmbView.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbView.Width = 240;

            _btnCopyPrev = UiHelpers.MakeButton("نسخ الموظفين من الفترة السابقة", delegate { CopyFromPrevious(); });
            _btnInsertAll = UiHelpers.MakeButton("إدراج موظفي المؤسسة لهذه الفترة", delegate { InsertEstablishmentEmployees(); });

            FlowLayoutPanel top = new FlowLayoutPanel();
            top.Dock = DockStyle.Fill;
            top.AutoSize = true;
            top.Padding = new Padding(6);
            top.Controls.Add(UiHelpers.MakeLabel("المؤسسة"));
            top.Controls.Add(_cmbEst);
            top.Controls.Add(_lblNssf);
            top.Controls.Add(UiHelpers.MakeLabel("السنة"));
            top.Controls.Add(_nudYear);
            top.Controls.Add(UiHelpers.MakeLabel("الفترة (النموذج)"));
            top.Controls.Add(_cmbPeriod);
            top.Controls.Add(UiHelpers.MakeLabel("تاريخ التوقيع"));
            top.Controls.Add(_dtpDate);
            top.Controls.Add(UiHelpers.MakeLabel("طريقة الإدخال"));
            top.Controls.Add(_cmbView);
            top.Controls.Add(_btnCopyPrev);
            top.Controls.Add(_btnInsertAll);

            // ---------- سطر الشرح / التنبيه ----------
            _lblGridHint.AutoSize = false;
            _lblGridHint.Dock = DockStyle.Fill;
            _lblGridHint.Height = 46;
            _lblGridHint.Padding = new Padding(10, 2, 10, 4);
            _lblGridHint.Font = new Font(Font.FontFamily, 9.5F);
            _lblGridHint.ForeColor = Color.DimGray;

            // ---------- جدول إدخال الأجور ----------
            _grid.Dock = DockStyle.Fill;
            _grid.AutoGenerateColumns = false;
            _grid.AllowUserToAddRows = true;
            _grid.AllowUserToDeleteRows = true;
            _grid.BackgroundColor = SystemColors.Window;
            _grid.RowHeadersWidth = 40;
            _grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            _colName = new DataGridViewTextBoxColumn();
            _colName.HeaderText = "الاسم";
            _colName.DataPropertyName = "Name";
            _colName.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            _colName.FillWeight = 160;
            _grid.Columns.Add(_colName);

            _colEos = NumberColumn("نهاية خدمة", "Eos");
            _colSm = NumberColumn("مرض وأمومة", "Sm");
            _colFa = NumberColumn("تعويضات عائلية", "Fa");
            _colDue = ComputedColumn("اشتراكات مستحقة", "DueContributions");
            _colPaid = NumberColumn("التعويض العائلي المدفوع", "Paid");
            _colNet = ComputedColumn("الاشتراكات", "NetContributions");
            _grid.Columns.Add(_colEos);
            _grid.Columns.Add(_colSm);
            _grid.Columns.Add(_colFa);
            _grid.Columns.Add(_colDue);
            _grid.Columns.Add(_colPaid);
            _grid.Columns.Add(_colNet);

            _grid.CellParsing += Grid_CellParsing;
            _grid.DataError += Grid_DataError;
            _grid.EditingControlShowing += Grid_EditingControlShowing;
            _grid.CellBeginEdit += Grid_CellBeginEdit;
            _grid.CellFormatting += Grid_CellFormatting;
            _grid.CellEndEdit += Grid_CellEndEdit;
            _grid.UserDeletedRow += delegate { AfterEdit(); };

            // ---------- جدول "الفترة 70" (تجميع سنوي، للعرض بس، بدون تعديل مباشر) ----------
            _grid70.Dock = DockStyle.Fill;
            _grid70.ReadOnly = true;
            _grid70.AllowUserToAddRows = false;
            _grid70.AllowUserToDeleteRows = false;
            _grid70.RowHeadersVisible = false;
            _grid70.TabStop = false;
            _grid70.BackgroundColor = SystemColors.Window;
            _grid70.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _grid70.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _grid70.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True;
            _grid70.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            _grid70.Columns.Add("name", "الاسم");
            _grid70.Columns.Add("eos", "نهاية الخدمة");
            _grid70.Columns.Add("paid", "الاشتراكات المدفوعة");
            _grid70.Columns.Add("fa", "التعويضات العائلية");
            _grid70.Columns.Add("sm", "المرض والأمومة");
            _grid70.Columns.Add("due", "المتوجب");
            foreach (DataGridViewColumn c in _grid70.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            _grid70.Columns["name"].FillWeight = 160;
            _grid70.Visible = false;

            // الجدول العادي (_grid) وجدول الفترة 70 (_grid70) بيشغلوا نفس المكان، وحدة بس ظاهرة كل مرة
            Panel gridHost = new Panel();
            gridHost.Dock = DockStyle.Fill;
            gridHost.Controls.Add(_grid70);
            gridHost.Controls.Add(_grid);

            // ---------- ملخص النموذج (نفس جدول الحسابات بشيتات 21-24) ----------
            _summary.Dock = DockStyle.Fill;
            _summary.ReadOnly = true;
            _summary.AllowUserToAddRows = false;
            _summary.AllowUserToDeleteRows = false;
            _summary.RowHeadersVisible = false;
            _summary.TabStop = false;
            _summary.BackgroundColor = SystemColors.Control;
            _summary.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _summary.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _summary.DefaultCellStyle.BackColor = Color.White;
            _summary.DefaultCellStyle.SelectionBackColor = Color.White;
            _summary.DefaultCellStyle.SelectionForeColor = Color.Black;
            _summary.Columns.Add("item", "البند");
            _summary.Columns.Add("wages", "الأجور ولواحقها");
            _summary.Columns.Add("rate", "المعدل");
            _summary.Columns.Add("count", "عدد الأجراء");
            _summary.Columns.Add("due", "الاشتراكات المستحقة");
            foreach (DataGridViewColumn c in _summary.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            _summary.Columns[0].FillWeight = 220;

            // ---------- أزرار الطباعة ----------
            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Bottom;
            buttons.AutoSize = true;
            buttons.Padding = new Padding(6);
            _btnOpen = UiHelpers.MakeButton("فتح النموذج في Excel (معاينة وطباعة)", delegate { OpenInExcel(); });
            _btnPrint = UiHelpers.MakeButton("طباعة مباشرة", delegate { PrintDirect(); });
            _btnPdf = UiHelpers.MakeButton("حفظ PDF", delegate { SavePdf(); });
            buttons.Controls.Add(_btnOpen);
            buttons.Controls.Add(_btnPrint);
            buttons.Controls.Add(_btnPdf);

            Panel bottom = new Panel();
            bottom.Dock = DockStyle.Fill;
            bottom.Controls.Add(_summary);   // Fill أول
            bottom.Controls.Add(buttons);    // وبعدين Bottom

            // ---------- التجميع ----------
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 4;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, 300F));
            root.Controls.Add(top, 0, 0);
            root.Controls.Add(_lblGridHint, 0, 1);
            root.Controls.Add(gridHost, 0, 2);
            root.Controls.Add(bottom, 0, 3);
            page.Controls.Add(root);

            // ---------- أحداث ----------
            _cmbEst.SelectedIndexChanged += delegate
            {
                if (_loading) return;
                CommitGridEdit();
                _est = _cmbEst.SelectedItem as Establishment;
                _data.LastEstablishmentId = _est != null ? (Guid?)_est.Id : null;
                UpdateEstUi();
                LoadPeriod();
            };
            _nudYear.ValueChanged += delegate { if (!_loading) LoadPeriod(); };
            _cmbPeriod.SelectedIndexChanged += delegate { if (!_loading) LoadPeriod(); };
            _cmbView.SelectedIndexChanged += delegate
            {
                if (_loading) return;
                CommitGridEdit();
                ApplyView();
            };
            _dtpDate.ValueChanged += delegate
            {
                if (_loading || _current == null) return;
                _current.SigningDate = _dtpDate.Value.Date;
                MarkDirty();
            };
        }

        // =====================================================================
        //  تبويب الطباعة: طباعة النموذج الحالي + طباعة تقرير موظف محدد
        // =====================================================================
        private void BuildPrintTab(TabPage page)
        {
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 4;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // عنوان: للموظف
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // اختيار الموظف/السنة/فصلي-شهري
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));   // جدول المعاينة
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));   // أزرار تقرير الموظف

            // ---------- للموظف (طباعة النموذج الأساسي رجعت لتبويب «احتساب الجداول») ----------
            Label lblEmpHeader = MakeHeader("طباعة تقرير أجور لموظف محدد", 8);

            _cmbPrintEmp.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbPrintEmp.DisplayMember = "DisplayName";
            _cmbPrintEmp.Width = 260;
            _cmbPrintEmp.SelectedIndexChanged += delegate { if (!_loading) RefreshPrintPreview(); };

            _nudPrintYear.Minimum = 2000;
            _nudPrintYear.Maximum = 2100;
            _nudPrintYear.Width = 90;
            _nudPrintYear.ValueChanged += delegate { if (!_loading) RefreshPrintPreview(); };

            _cmbPrintGranularity.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbPrintGranularity.Width = 100;
            _cmbPrintGranularity.Items.AddRange(new object[] { "فصلي", "شهري" });
            _cmbPrintGranularity.SelectedIndex = 0;
            _cmbPrintGranularity.SelectedIndexChanged += delegate { if (!_loading) RefreshPrintPreview(); };

            FlowLayoutPanel empControls = new FlowLayoutPanel();
            empControls.AutoSize = true;
            empControls.Padding = new Padding(6);
            empControls.Controls.Add(UiHelpers.MakeLabel("الموظف"));
            empControls.Controls.Add(_cmbPrintEmp);
            empControls.Controls.Add(UiHelpers.MakeLabel("السنة"));
            empControls.Controls.Add(_nudPrintYear);
            empControls.Controls.Add(_cmbPrintGranularity);

            _gridPrintPreview.Dock = DockStyle.Fill;
            _gridPrintPreview.ReadOnly = true;
            _gridPrintPreview.AllowUserToAddRows = false;
            _gridPrintPreview.AllowUserToDeleteRows = false;
            _gridPrintPreview.RowHeadersVisible = false;
            _gridPrintPreview.TabStop = false;
            _gridPrintPreview.BackgroundColor = SystemColors.Window;
            _gridPrintPreview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _gridPrintPreview.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            BuildPrintPreviewColumns(false);

            _btnPrintEmpExcel = UiHelpers.MakeButton("فتح تقرير الموظف في Excel", delegate { RunEmployeePrint(ExcelAction.OpenInExcel, null); });
            _btnPrintEmpPrint = UiHelpers.MakeButton("طباعة تقرير الموظف", delegate { RunEmployeePrint(ExcelAction.Print, null); });
            _btnPrintEmpPdf = UiHelpers.MakeButton("حفظ PDF", delegate { SaveEmployeePrintPdf(); });
            _btnPrintEmpExcel.Enabled = false;
            _btnPrintEmpPrint.Enabled = false;
            _btnPrintEmpPdf.Enabled = false;

            FlowLayoutPanel empButtons = new FlowLayoutPanel();
            empButtons.AutoSize = true;
            empButtons.Padding = new Padding(6);
            empButtons.Controls.Add(_btnPrintEmpExcel);
            empButtons.Controls.Add(_btnPrintEmpPrint);
            empButtons.Controls.Add(_btnPrintEmpPdf);

            root.Controls.Add(lblEmpHeader, 0, 0);
            root.Controls.Add(empControls, 0, 1);
            root.Controls.Add(_gridPrintPreview, 0, 2);
            root.Controls.Add(empButtons, 0, 3);
            page.Controls.Add(root);
        }

        // بيستدعيها تبديل التبويب لتبويب الطباعة، وكل ما تتبدّل المؤسسة أو الفترة، حتى يضل محدّث
        private void RefreshPrintTab()
        {
            bool prev = _loading;
            _loading = true;
            try
            {
                Employee keepSelected = _cmbPrintEmp.SelectedItem as Employee;
                _cmbPrintEmp.Items.Clear();
                if (_est != null)
                    foreach (Employee e in _est.Employees) _cmbPrintEmp.Items.Add(e);

                if (keepSelected != null && _cmbPrintEmp.Items.Contains(keepSelected))
                    _cmbPrintEmp.SelectedItem = keepSelected;
                else if (_cmbPrintEmp.Items.Count > 0)
                    _cmbPrintEmp.SelectedIndex = 0;
                else
                    _cmbPrintEmp.SelectedIndex = -1;

                int y = (int)_nudYear.Value;
                _nudPrintYear.Value = Math.Min(Math.Max(y, (int)_nudPrintYear.Minimum), (int)_nudPrintYear.Maximum);
            }
            finally
            {
                _loading = prev;
            }

            RefreshPrintPreview();
        }

        private void BuildPrintPreviewColumns(bool monthly)
        {
            _gridPrintPreview.Columns.Clear();
            _gridPrintPreview.Columns.Add("year", "السنة");
            _gridPrintPreview.Columns.Add("period", "الفترة");
            if (monthly) _gridPrintPreview.Columns.Add("month", "الشهر");
            _gridPrintPreview.Columns.Add("sm", "مرض وأمومة");
            _gridPrintPreview.Columns.Add("eos", "نهاية خدمة");
            _gridPrintPreview.Columns.Add("fa", "تعويضات عائلية");
            _gridPrintPreview.Columns.Add("paid", "التعويض المدفوع");
            foreach (DataGridViewColumn c in _gridPrintPreview.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void RefreshPrintPreview()
        {
            _gridPrintPreview.Rows.Clear();
            Employee emp = _cmbPrintEmp.SelectedItem as Employee;
            bool monthly = _cmbPrintGranularity.SelectedIndex == 1;
            BuildPrintPreviewColumns(monthly);

            if (_est == null || emp == null)
            {
                _btnPrintEmpExcel.Enabled = false;
                _btnPrintEmpPrint.Enabled = false;
                _btnPrintEmpPdf.Enabled = false;
                return;
            }

            int year = (int)_nudPrintYear.Value;
            List<EmployeeWageReportRow> rows = WageRules.BuildWageReport(_est, emp, monthly).Where(r => r.Year == year).ToList();

            foreach (EmployeeWageReportRow r in rows)
            {
                if (monthly)
                    _gridPrintPreview.Rows.Add(r.Year.ToString(CultureInfo.InvariantCulture), r.Period, r.Month,
                        NumberInput.Format(r.SickMaternity), NumberInput.Format(r.EndOfService), NumberInput.Format(r.FamilyAllowance), NumberInput.Format(r.FamilyAllowancePaid));
                else
                    _gridPrintPreview.Rows.Add(r.Year.ToString(CultureInfo.InvariantCulture), r.Period,
                        NumberInput.Format(r.SickMaternity), NumberInput.Format(r.EndOfService), NumberInput.Format(r.FamilyAllowance), NumberInput.Format(r.FamilyAllowancePaid));
            }

            if (rows.Count > 0)
            {
                int idx = monthly
                    ? _gridPrintPreview.Rows.Add("", "", "المجموع", NumberInput.Format(rows.Sum(r => r.SickMaternity)), NumberInput.Format(rows.Sum(r => r.EndOfService)), NumberInput.Format(rows.Sum(r => r.FamilyAllowance)), NumberInput.Format(rows.Sum(r => r.FamilyAllowancePaid)))
                    : _gridPrintPreview.Rows.Add("", "المجموع", NumberInput.Format(rows.Sum(r => r.SickMaternity)), NumberInput.Format(rows.Sum(r => r.EndOfService)), NumberInput.Format(rows.Sum(r => r.FamilyAllowance)), NumberInput.Format(rows.Sum(r => r.FamilyAllowancePaid)));
                _gridPrintPreview.Rows[idx].DefaultCellStyle.Font = _boldFont;
                _gridPrintPreview.Rows[idx].DefaultCellStyle.BackColor = Color.Gainsboro;
            }

            _btnPrintEmpExcel.Enabled = true;
            _btnPrintEmpPrint.Enabled = true;
            _btnPrintEmpPdf.Enabled = true;
        }

        private bool RunEmployeePrint(ExcelAction action, string pdfPath)
        {
            Employee emp = _cmbPrintEmp.SelectedItem as Employee;
            if (_est == null || emp == null)
            {
                UiHelpers.Warn(this, "اختار موظف أولاً.");
                return false;
            }

            bool monthly = _cmbPrintGranularity.SelectedIndex == 1;
            int year = (int)_nudPrintYear.Value;
            List<EmployeeWageReportRow> rows = WageRules.BuildWageReport(_est, emp, monthly).Where(r => r.Year == year).ToList();

            try
            {
                UseWaitCursor = true;
                Application.DoEvents();
                EmployeeReportPrinter.Run(_est, emp, monthly, rows, action, pdfPath);
                return true;
            }
            catch (Exception ex)
            {
                UiHelpers.Error(this, "صار خطأ مع Excel:" + Environment.NewLine + ex.Message);
                return false;
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        private void SaveEmployeePrintPdf()
        {
            Employee emp = _cmbPrintEmp.SelectedItem as Employee;
            if (_est == null || emp == null)
            {
                UiHelpers.Warn(this, "اختار موظف أولاً.");
                return;
            }

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "PDF (*.pdf)|*.pdf";
                string safeName = string.IsNullOrWhiteSpace(emp.FullName) ? "موظف" : emp.FullName;
                dlg.FileName = "تقرير_أجور_" + safeName + "_" + (int)_nudPrintYear.Value + ".pdf";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                if (RunEmployeePrint(ExcelAction.ExportPdf, dlg.FileName))
                {
                    try { Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true }); }
                    catch (Exception) { }
                }
            }
        }

        private static DataGridViewTextBoxColumn NumberColumn(string header, string property)
        {
            DataGridViewTextBoxColumn c = new DataGridViewTextBoxColumn();
            c.HeaderText = header;
            c.DataPropertyName = property;
            c.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            c.FillWeight = 100;
            c.DefaultCellStyle.Format = "N0";
            c.DefaultCellStyle.FormatProvider = CultureInfo.InvariantCulture;
            c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            c.DefaultCellStyle.BackColor = Color.LightYellow;   // أصفر = خانة إدخال (متل الإكسل)
            return c;
        }

        // خانة محسوبة تلقائياً (اشتراكات مستحقة / الاشتراكات الصافية) — مش خانة إدخال، فمالها نفس لون الأصفر
        private static DataGridViewTextBoxColumn ComputedColumn(string header, string property)
        {
            DataGridViewTextBoxColumn c = NumberColumn(header, property);
            c.ReadOnly = true;
            c.DefaultCellStyle.BackColor = Color.WhiteSmoke;
            return c;
        }

        // =====================================================================
        //  المؤسسات (تبويب بيانات المؤسسة)
        // =====================================================================
        private void LoadRatesToUi()
        {
            _nudSm.Value = ClampPercent(_data.Rates.SickMaternity * 100m);
            _nudEos.Value = ClampPercent(_data.Rates.EndOfService * 100m);
            _nudFa.Value = ClampPercent(_data.Rates.FamilyAllowance * 100m);
        }

        private static decimal ClampPercent(decimal v)
        {
            return Math.Max(0m, Math.Min(100m, v));
        }

        private void SetEstStatus(string text, bool warn)
        {
            _lblEstStatus.Text = text;
            _lblEstStatus.ForeColor = warn ? Color.Firebrick : Color.DimGray;
        }

        // بيعبّي قائمة المؤسسات وبيعرض المختارة (أو أول وحدة). إذا ما في ولا وحدة بيفتح مؤسسة جديدة.
        private void RefreshEstList(Guid? select)
        {
            bool prev = _loading;
            _loading = true;
            try
            {
                _lstEst.Items.Clear();
                foreach (Establishment e in _data.Establishments) _lstEst.Items.Add(e);

                int idx = -1;
                if (select.HasValue) idx = _data.Establishments.FindIndex(x => x.Id == select.Value);
                if (idx < 0 && _lstEst.Items.Count > 0) idx = 0;
                _lstEst.SelectedIndex = idx;
            }
            finally
            {
                _loading = prev;
            }

            Establishment sel = _lstEst.SelectedItem as Establishment;
            if (sel != null) ShowEstablishment(sel);
            else NewEstablishment();
        }

        private void ShowEstablishment(Establishment est)
        {
            _editing = est;
            _editingIsNew = false;

            bool prev = _loading;
            _loading = true;
            try
            {
                _txtEstName.Text = est.Name;
                _txtEstNssf.Text = est.NssfNumber;
                _txtOwner.Text = est.OwnerName;
                _txtPhone.Text = est.OwnerPhone;
            }
            finally
            {
                _loading = prev;
            }

            _btnEmployees.Enabled = true;
            _btnDeleteEst.Enabled = true;
            SetEstStatus("عدد الموظفين المسجّلين: " + est.Employees.Count, false);
        }

        private void NewEstablishment()
        {
            _editing = new Establishment();
            _editingIsNew = true;

            bool prev = _loading;
            _loading = true;
            try
            {
                _lstEst.SelectedIndex = -1;
                _txtEstName.Text = "";
                _txtEstNssf.Text = "";
                _txtOwner.Text = "";
                _txtPhone.Text = "";
            }
            finally
            {
                _loading = prev;
            }

            _btnEmployees.Enabled = false;
            _btnDeleteEst.Enabled = false;
            SetEstStatus("مؤسسة جديدة: عبّي البيانات واضغط «حفظ المؤسسة»، وبعدها بتقدر تضيف موظفينها.", false);
            _txtEstName.Focus();
        }

        private void SaveEstablishment()
        {
            if (_editing == null) return;

            string name = _txtEstName.Text.Trim();
            if (name.Length == 0)
            {
                UiHelpers.Warn(this, "اكتب اسم المؤسسة أول شي.");
                _txtEstName.Focus();
                return;
            }

            _editing.Name = name;
            _editing.NssfNumber = _txtEstNssf.Text.Trim();
            _editing.OwnerName = _txtOwner.Text.Trim();
            _editing.OwnerPhone = _txtPhone.Text.Trim();

            bool wasNew = _editingIsNew;
            if (wasNew)
            {
                _data.Establishments.Add(_editing);
                _editingIsNew = false;
            }
            MarkDirty();

            // بنحدّث القائمتين (تبويب المؤسسات + القائمة المنسدلة بتبويب الأجور)
            Guid? keepInCombo = _est != null ? (Guid?)_est.Id : _editing.Id;
            RefreshEstList(_editing.Id);
            RefreshEstCombo(keepInCombo);
            RebuildNameSource();
            LoadPeriod();

            SetEstStatus("✔ تم حفظ المؤسسة. عدد الموظفين المسجّلين: " + _editing.Employees.Count, false);

            if (wasNew &&
                UiHelpers.Ask(this, "تم حفظ المؤسسة. بدك تضيف موظفين لها هلأ؟") == DialogResult.Yes)
            {
                OpenEmployees();
            }
        }

        private void DeleteEstablishment()
        {
            if (_editing == null || _editingIsNew) return;

            string label = _editing.ToString();
            if (UiHelpers.Ask(this, "بدك تحذف المؤسسة \"" + label + "\" مع كل موظفينها وكل بيانات فتراتها؟ ما في رجعة.") != DialogResult.Yes)
                return;

            Guid removedId = _editing.Id;
            _data.Establishments.Remove(_editing);
            if (_data.LastEstablishmentId.HasValue && _data.LastEstablishmentId.Value == removedId)
                _data.LastEstablishmentId = null;
            MarkDirty();

            RefreshEstList(null);
            RefreshEstCombo(_data.LastEstablishmentId);
            RebuildNameSource();
            LoadPeriod();
        }

        private void OpenEmployees()
        {
            if (_editing == null || _editingIsNew)
            {
                UiHelpers.Warn(this, "احفظ بيانات المؤسسة أولاً، وبعدين ضيف الموظفين.");
                return;
            }

            using (EmployeesForm f = new EmployeesForm(_editing, (int)_nudYear.Value, delegate { MarkDirty(); }))
            {
                f.ShowDialog(this);
            }

            SaveIfDirty();
            if (_est == _editing)
            {
                // نفس المؤسسة معروضة بتبويب الأجور: منعيد بناء الجدول من الصفر، حتى لو انحذف موظف
                // كان إلو سطر بالفترة الحالية (حذف الموظف بيشيل أسطرو من كل الفترات، بس الجدول
                // المفتوح هون ما بيعرف بهيك تلقائياً إلا لو رجعنا بنينا _rows من جديد).
                LoadPeriod();
            }
            else
            {
                RebuildNameSource();      // ممكن تغيّرت أسماء أو تواريخ ترك
                _grid.Invalidate();
                RefreshSummary();
            }
            SetEstStatus("عدد الموظفين المسجّلين: " + _editing.Employees.Count, false);
        }

        // =====================================================================
        //  القائمة المنسدلة للمؤسسات (تبويب الأجور)
        // =====================================================================
        private void RefreshEstCombo(Guid? select)
        {
            bool prev = _loading;
            _loading = true;
            try
            {
                _cmbEst.Items.Clear();
                foreach (Establishment e in _data.Establishments) _cmbEst.Items.Add(e);

                int idx = -1;
                if (select.HasValue) idx = _data.Establishments.FindIndex(x => x.Id == select.Value);
                if (idx < 0 && _cmbEst.Items.Count > 0) idx = 0;
                _cmbEst.SelectedIndex = idx;
            }
            finally
            {
                _loading = prev;
            }

            _est = _cmbEst.SelectedItem as Establishment;
            _data.LastEstablishmentId = _est != null ? (Guid?)_est.Id : null;
            UpdateEstUi();
        }

        // رقم الضمان بيظهر لحالو أول ما تختار المؤسسة (وهو يلي بينكتب بالنموذج)
        private void UpdateEstUi()
        {
            bool has = _est != null;
            bool annual = IsAnnualSummaryPeriod();

            if (has)
                _lblNssf.Text = "رقم الضمان: " + (string.IsNullOrWhiteSpace(_est.NssfNumber) ? "—" : _est.NssfNumber);
            else
                _lblNssf.Text = "ما في مؤسسات. أضف مؤسسة من تبويب «بيانات المؤسسة والنسب».";

            _grid.Enabled = has && !annual;
            _cmbView.Enabled = has && !annual;
            _btnCopyPrev.Enabled = has && !annual;
            _btnInsertAll.Enabled = has && !annual;
            _btnOpen.Enabled = has && !annual;
            _btnPrint.Enabled = has && !annual;
            _btnPdf.Enabled = has && !annual;

            _grid.Visible = !annual;
            _grid70.Visible = annual;
            _summary.Visible = !annual;
        }

        // "الفترة 70" مش فترة حقيقية مخزّنة، هي عرض تجميعي سنوي محسوب من الفصول 21-24
        private bool IsAnnualSummaryPeriod()
        {
            return _cmbPeriod.SelectedIndex == 4;
        }

        // بيبني لائحة الأسماء يلي بتكمّل لحالها (تعبئة تلقائية) من موظفي المؤسسة المختارة.
        // الموظف يلي ترك قبل هالفصل ما بيظهر (بينشال من الأجراء).
        private void RebuildNameSource()
        {
            AutoCompleteStringCollection src = new AutoCompleteStringCollection();
            if (_est != null && !IsAnnualSummaryPeriod())
            {
                int year = (int)_nudYear.Value;
                int quarter = _cmbPeriod.SelectedIndex + 1;
                if (quarter < 1) quarter = 1;

                foreach (Employee e in _est.Employees)
                {
                    if (string.IsNullOrWhiteSpace(e.FullName) || src.Contains(e.FullName)) continue;
                    if (!e.IsActiveInPeriod(year, quarter)) continue;
                    src.Add(e.FullName);
                }
            }
            _nameSource = src;
        }

        private void LoadPeriod()
        {
            bool prev = _loading;
            bool annual = IsAnnualSummaryPeriod();
            _loading = true;
            try
            {
                int year = (int)_nudYear.Value;
                int number = annual ? 0 : _cmbPeriod.SelectedIndex + 1;

                _ctx.Establishment = _est;
                _ctx.Year = year;
                _ctx.Quarter = number < 1 ? 1 : number;
                _ctx.Rates = _data.Rates;

                if (annual)
                {
                    _current = null;
                    _rows = new LineRowList(new List<EmployeeLine>(), _ctx);
                    _data.LastYear = year;
                    RefreshAnnualSummary();
                }
                else if (_est == null || number < 1)
                {
                    _current = null;
                    _rows = new LineRowList(new List<EmployeeLine>(), _ctx);
                }
                else
                {
                    _current = _est.GetPeriod(year, number);
                    _rows = new LineRowList(_current.Lines, _ctx);   // بتلفّ نفس لائحة الفترة، فأي تعديل بينحفظ

                    DateTime d = _current.SigningDate.HasValue ? _current.SigningDate.Value : DateTime.Today;
                    _dtpDate.Value = d.Date;

                    _data.LastYear = year;
                    _data.LastPeriod = number;
                }

                _grid.DataSource = _rows;
                RebuildViewItems();
                ApplyView();
                RebuildNameSource();
                UpdateEstUi();   // بيبدّل الظهور بين الجدول العادي وجدول الفترة 70، وبيعطّل الأزرار يلي ما بتنطبق عليها
            }
            finally
            {
                _loading = prev;
            }

            if (!annual) RefreshSummary();
            MarkDirty();
        }

        // ---------- طريقة الإدخال: الفصل كامل أو شهر بشهر ----------
        private void RebuildViewItems()
        {
            int quarter = _ctx.Quarter < 1 ? 1 : _ctx.Quarter;
            int sel = _cmbView.SelectedIndex;

            _cmbView.Items.Clear();
            _cmbView.Items.Add("الفصل كامل");
            for (int i = 0; i < 3; i++)
                _cmbView.Items.Add("الشهر " + (i + 1) + " (" + WageRules.MonthNames[(quarter - 1) * 3 + i] + ")");

            _cmbView.SelectedIndex = (sel >= 0 && sel < 4) ? sel : 0;
        }

        private void ApplyView()
        {
            if (IsAnnualSummaryPeriod())
            {
                _lblGridHint.ForeColor = Color.DimGray;
                _lblGridHint.Text = "الفترة 70: تجميع سنوي تلقائي (نهاية الخدمة، الاشتراكات المدفوعة، التعويضات العائلية، المرض والأمومة) "
                    + "من الفصول 21+22+23+24 لكل موظف — مش فترة تدخل عليها بيانات لحالها. عدّل من الفصل الأصلي وبينعكس هون تلقائياً. "
                    + "ما فيها طباعة ولا نسخ/إدراج موظفين.";
                return;
            }

            int v = _cmbView.SelectedIndex < 0 ? 0 : _cmbView.SelectedIndex;
            _ctx.ViewMonth = v;

            string label = v == 0 ? "الفصل كامل" : "الشهر " + v;
            _colEos.HeaderText = "نهاية خدمة" + Environment.NewLine + label;
            _colSm.HeaderText = "مرض وأمومة" + Environment.NewLine + label;
            _colFa.HeaderText = "تعويضات عائلية" + Environment.NewLine + label;
            _colDue.HeaderText = "اشتراكات مستحقة" + Environment.NewLine + label;
            _colPaid.HeaderText = "التعويض العائلي المدفوع" + Environment.NewLine + label;
            _colNet.HeaderText = "الاشتراكات" + Environment.NewLine + label;

            _grid.Invalidate();
            UpdateGridHint();
        }

        private void UpdateGridHint()
        {
            _lblGridHint.ForeColor = Color.DimGray;
            if (_ctx.ViewMonth == 0)
                _lblGridHint.Text = "عرض الفصل كامل: الرقم يلي بتكتبو بيتوزّع بالتساوي على أشهر الفصل. لتدخل أجور كل شهر لحالو اختار الشهر من «طريقة الإدخال». "
                    + "الصفر = الموظف مو محسوب بهالفرع (بينقص من عدد الأجراء). الرمادي = ما بينحسب (موظف ترك).";
            else
                _lblGridHint.Text = "عم تدخل أجور " + _cmbView.Text + " بس، والأشهر التانية ما بتتغير. "
                    + "الصفر = الموظف مو محسوب بهالفرع (بينقص من عدد الأجراء). الرمادي = ما بينحسب (موظف ترك).";
        }

        private void SetGridHint(string text)
        {
            _lblGridHint.ForeColor = Color.Firebrick;
            _lblGridHint.Text = text;
        }

        private void RefreshSummary()
        {
            if (_current == null)
            {
                _summary.Rows.Clear();
                return;
            }

            PeriodResult r = Calculator.Calculate(_current, _data.Rates, _est);

            _summary.Rows.Clear();
            AddSummaryRow(r.SickMaternity);
            AddSummaryRow(r.EndOfService);
            AddSummaryRow(r.FamilyAllowance);
            _summary.Rows.Add("مجموع الاشتراكات المستحقة", "", "", "", NumberInput.Format(r.TotalDue));
            _summary.Rows.Add("التعويضات العائلية المدفوعة", "", "", "", NumberInput.Format(r.FamilyAllowancePaid));
            _summary.Rows.Add("الزيادة المتوجبة على الصندوق / الباقي المتوجب للصندوق", "", "", "", NumberInput.Format(r.Net));

            for (int i = 3; i < _summary.Rows.Count; i++)
                _summary.Rows[i].DefaultCellStyle.Font = _boldFont;

            _summary.ClearSelection();
        }

        // بيعبّي جدول "الفترة 70" (تجميع سنوي 21+22+23+24 لكل موظف) مع سطر مجموع بالآخر
        private void RefreshAnnualSummary()
        {
            _grid70.Rows.Clear();
            if (_est == null) return;

            List<AnnualSummaryRow> rows = WageRules.BuildAnnualSummary(_est, (int)_nudYear.Value, _data.Rates.EndOfService);
            foreach (AnnualSummaryRow r in rows)
                _grid70.Rows.Add(r.EmployeeName, NumberInput.Format(r.EndOfService), NumberInput.Format(r.PaidContribution),
                    NumberInput.Format(r.FamilyAllowance), NumberInput.Format(r.SickMaternity), NumberInput.Format(r.Due));

            if (rows.Count > 0)
            {
                int idx = _grid70.Rows.Add("المجموع",
                    NumberInput.Format(rows.Sum(x => x.EndOfService)), NumberInput.Format(rows.Sum(x => x.PaidContribution)),
                    NumberInput.Format(rows.Sum(x => x.FamilyAllowance)), NumberInput.Format(rows.Sum(x => x.SickMaternity)),
                    NumberInput.Format(rows.Sum(x => x.Due)));
                _grid70.Rows[idx].DefaultCellStyle.Font = _boldFont;
                _grid70.Rows[idx].DefaultCellStyle.BackColor = Color.Gainsboro;
            }
        }

        private void AddSummaryRow(BranchResult b)
        {
            _summary.Rows.Add(
                b.Name,
                NumberInput.Format(b.Wages),
                (b.Rate * 100m).ToString("0.##", CultureInfo.InvariantCulture) + "%",
                b.Employees.ToString(CultureInfo.InvariantCulture),
                NumberInput.Format(b.Contribution));
        }

        // =====================================================================
        //  أحداث الجدول
        // =====================================================================
        private void AfterEdit()
        {
            if (_loading) return;
            _grid.Invalidate();
            RefreshSummary();
            MarkDirty();
        }

        // تعبئة تلقائية للأسماء (متل الإكسل): بس تكتب أول حروف الاسم بيقترحو ويكمّلو
        private void Grid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox tb = e.Control as TextBox;
            if (tb == null) return;

            bool isNameCell = _grid.CurrentCell != null && _grid.CurrentCell.ColumnIndex == _colName.Index;
            if (isNameCell)
            {
                tb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                tb.AutoCompleteSource = AutoCompleteSource.CustomSource;
                tb.AutoCompleteCustomSource = _nameSource;
            }
            else
            {
                tb.AutoCompleteMode = AutoCompleteMode.None;   // نفس خانة التعديل بتنعاد استعمالها للأعمدة التانية
            }
        }

        // الأعمدة الرقمية: بترجّع نوع الأجر (null = مو عمود أجر)
        private WageKind? KindOfColumn(int col)
        {
            if (col == _colSm.Index) return WageKind.SickMaternity;
            if (col == _colEos.Index) return WageKind.EndOfService;
            if (col == _colFa.Index) return WageKind.FamilyAllowance;
            if (col == _colPaid.Index) return WageKind.FamilyAllowancePaid;
            return null;
        }

        // بيمنع تعديل خانة ما بينفع تنعدّل (موظف ترك، أو أشهر مختلفة بعرض الفصل) وبيقلّك ليش
        private void Grid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            UpdateGridHint();
            if (e.RowIndex < 0 || e.RowIndex >= _rows.Count) return;   // السطر الجديد الفاضي: مسموح

            WageKind? kind = KindOfColumn(e.ColumnIndex);
            if (!kind.HasValue) return;

            string reason;
            if (!_rows[e.RowIndex].CanEdit(kind.Value, out reason))
            {
                e.Cancel = true;
                SetGridHint(reason);
            }
        }

        // الخانات يلي ما بتنحسب أو ما بتنعدّل بتصير رمادي
        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _rows.Count) return;

            WageKind? kind = KindOfColumn(e.ColumnIndex);
            if (!kind.HasValue) return;

            string reason;
            if (!_rows[e.RowIndex].CanEdit(kind.Value, out reason))
            {
                e.CellStyle.BackColor = Color.Gainsboro;
                e.CellStyle.ForeColor = Color.DimGray;
            }
        }

        private void Grid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (!_loading && e.RowIndex >= 0 && e.RowIndex < _rows.Count && e.ColumnIndex == _colName.Index)
                AutoFillFromEmployee(_rows[e.RowIndex]);
            AfterEdit();
        }

        // إذا الاسم المكتوب لموظف مسجّل بالمؤسسة: منوحّد كتابتو ومنربطو ببطاقتو، ومنعبّي رواتبو الشهرية
        // من بطاقتو لكل شهر فعّال (بس إذا رواتب السطر فاضية). إذا الموظف تارك قبل هالفصل بيتنبّه.
        private void AutoFillFromEmployee(LineRow row)
        {
            EmployeeLine line = row.Line;
            if (_est == null || string.IsNullOrWhiteSpace(line.Name)) return;

            string typed = line.Name.Trim();
            Employee card = _est.Employees.FirstOrDefault(x =>
                string.Equals((x.FullName ?? "").Trim(), typed, StringComparison.CurrentCultureIgnoreCase));
            if (card == null)
            {
                line.EmployeeId = null;
                return;
            }

            line.EmployeeId = card.Id;
            line.Name = card.FullName;

            if (!card.IsActiveInPeriod(_ctx.Year, _ctx.Quarter))
            {
                SetGridHint("الموظف «" + card.FullName + "» ترك بتاريخ "
                    + card.EndDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                    + ": ما بينحسب بهالفصل.");
                return;
            }

            if (WageRules.IsEmpty(line))
                WageRules.FillFromCard(line, card, _ctx.Year, _ctx.Quarter);
        }

        // بيقبل 28000000 أو 28,000,000 أو أرقام عربية ٢٨٠٠٠٠٠٠
        private void Grid_CellParsing(object sender, DataGridViewCellParsingEventArgs e)
        {
            if (e.DesiredType != typeof(decimal)) return;

            decimal d;
            if (NumberInput.TryParseDecimal(Convert.ToString(e.Value), out d))
            {
                e.Value = d;
                e.ParsingApplied = true;
            }
        }

        private void Grid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if ((e.Context & DataGridViewDataErrorContexts.Parsing) != 0)
                UiHelpers.Warn(this, "القيمة يلي دخلتها مو رقم صحيح. اكتب أرقام بس (أو اضغط Esc للتراجع).");
            e.ThrowException = false;
        }

        private void CopyFromPrevious()
        {
            if (_est == null || _current == null) return;

            int year = (int)_nudYear.Value;
            int number = _cmbPeriod.SelectedIndex + 1;
            int prevYear = number == 1 ? year - 1 : year;
            int prevNumber = number == 1 ? 4 : number - 1;

            PeriodData prev = _est.FindPeriod(prevYear, prevNumber);
            if (prev == null || prev.Lines.Count == 0)
            {
                UiHelpers.Info(this, "ما في بيانات بالفترة السابقة (2" + prevNumber + "-" + prevYear + ").");
                return;
            }

            CommitGridEdit();
            if (_rows.Count > 0 &&
                UiHelpers.Ask(this, "الجدول الحالي فيه بيانات. بدك تستبدلها ببيانات الفترة السابقة؟") != DialogResult.Yes)
                return;

            int copied = 0;
            int skipped = 0;
            _loading = true;
            try
            {
                _current.Lines.Clear();      // منعدّل اللائحة مباشرة (أسرع وأسلم من إضافة سطر سطر على جدول مربوط)
                foreach (EmployeeLine l in prev.Lines)
                {
                    Employee card = WageRules.FindEmployee(_est, l);
                    if (card != null && !card.IsActiveInPeriod(year, number))
                    {
                        skipped++;      // ترك قبل هالفصل: بينشال من الأجراء
                        continue;
                    }

                    // الراتب الحالي = أجور آخر شهر فعّال بالفصل السابق (ممكن تكون الزيادة صارت بنص الفصل)
                    MonthWages basis = LastActiveMonth(l, card, prevYear, prevNumber);

                    EmployeeLine copy = new EmployeeLine();
                    copy.EmployeeId = l.EmployeeId;
                    copy.Name = l.Name;
                    for (int i = 0; i < 3; i++)
                    {
                        bool active = WageRules.IsMonthActive(card, year, number, i);
                        copy.Months[i].SickMaternity = active ? basis.SickMaternity : 0m;
                        copy.Months[i].EndOfService = active ? basis.EndOfService : 0m;
                        copy.Months[i].FamilyAllowance = active ? basis.FamilyAllowance : 0m;
                    }
                    // ملاحظة: ما منعبّي التعويض العائلي المدفوع من الفترة السابقة — بيضل صفر (القيمة
                    // الافتراضية لسطر جديد)، لأنو بيتحدّد من جديد لكل فترة حسب المبلغ يلي فعلاً انصرف.
                    _current.Lines.Add(copy);
                    copied++;
                }
            }
            finally
            {
                _rows = new LineRowList(_current.Lines, _ctx);   // ربط الجدول من جديد مرة وحدة
                _grid.DataSource = _rows;
                _loading = false;
            }

            AfterEdit();
            if (skipped > 0)
                UiHelpers.Info(this, "اننسخ " + copied + " موظف. ما اننسخ " + skipped + " لأنهم تاركين قبل هالفصل.");
        }

        private static MonthWages LastActiveMonth(EmployeeLine l, Employee card, int year, int quarter)
        {
            l.EnsureMonths();
            for (int i = 2; i >= 0; i--)
            {
                if (!WageRules.IsMonthActive(card, year, quarter, i)) continue;
                MonthWages m = l.Months[i];
                if (m.SickMaternity != 0m || m.EndOfService != 0m || m.FamilyAllowance != 0m) return m;
            }
            return new MonthWages();
        }

        // بيضيف كل موظفي المؤسسة (يلي مؤهّلين) لجدول الفترة الحالية دفعة وحدة. بيمتنع إذا الجدول فيه
        // أسماء أصلاً (لتجنّب التكرار)، وبيتجاهل موظف لسا ما مرق شهر كامل على استلامو (WageRules.FillFromCardForBulkInsert).
        private void InsertEstablishmentEmployees()
        {
            if (_est == null || _current == null) return;

            CommitGridEdit();
            if (_rows.Any(r => !string.IsNullOrWhiteSpace(r.Name)))
            {
                UiHelpers.Warn(this, "الجدول عندو أسماء موظفين مسجّلة أصلاً. فضّي الجدول (أو احذف الأسطر) قبل ما تستعمل هالزر، حتى ما ينصير تكرار.");
                return;
            }

            int inserted = 0;
            int skipped = 0;
            _loading = true;
            try
            {
                foreach (Employee emp in _est.Employees)
                {
                    if (string.IsNullOrWhiteSpace(emp.FullName)) continue;   // بطاقة بلا اسم: تجاهلها

                    EmployeeLine line = new EmployeeLine();
                    if (WageRules.FillFromCardForBulkInsert(line, emp, _ctx.Year, _ctx.Quarter))
                    {
                        _current.Lines.Add(line);
                        inserted++;
                    }
                    else
                    {
                        skipped++;
                    }
                }
            }
            finally
            {
                _rows = new LineRowList(_current.Lines, _ctx);   // ربط الجدول من جديد مرة وحدة
                _grid.DataSource = _rows;
                _loading = false;
            }

            AfterEdit();
            if (inserted == 0)
                UiHelpers.Info(this, "ما انضاف حدا: يا المؤسسة بلا موظفين مسجّلين، يا كلن تاركين، يا لسا ما مرق شهر كامل على استلامن.");
            else if (skipped > 0)
                UiHelpers.Info(this, "انضاف " + inserted + " موظف. " + skipped + " ما انضافوا (تاركين قبل هالفترة، أو لسا ما مرق شهر على استلامن).");
        }

        // =====================================================================
        //  Excel: معاينة / طباعة / PDF
        // =====================================================================
        private string FormLabel()
        {
            return "2" + (_cmbPeriod.SelectedIndex + 1) + "-" + (int)_nudYear.Value;
        }

        private void CommitGridEdit()
        {
            if (_grid.IsCurrentCellInEditMode) _grid.EndEdit();
            _grid.CurrentCell = null;   // بيثبّت السطر الجديد إذا كان لسا بالتعديل
        }

        private bool CheckReadyToPrint()
        {
            if (_est == null || _current == null)
            {
                UiHelpers.Warn(this, "اختار مؤسسة أولاً. إذا ما في مؤسسات، أضف وحدة من تبويب «بيانات المؤسسة والنسب».");
                return false;
            }

            if (string.IsNullOrWhiteSpace(_est.NssfNumber) &&
                UiHelpers.Ask(this, "رقم الضمان للمؤسسة فاضي وبيطلع فاضي بالنموذج. بدك تكمّل؟") != DialogResult.Yes)
                return false;

            return true;
        }

        private bool RunExcel(ExcelAction action, string pdfPath)
        {
            if (_est == null || _current == null) return false;

            CommitGridEdit();
            RefreshSummary();
            SaveIfDirty();

            PeriodResult result = Calculator.Calculate(_current, _data.Rates, _est);
            DateTime date = _dtpDate.Value.Date;

            try
            {
                UseWaitCursor = true;
                Application.DoEvents();
                ExcelPrinter.Run(_est, result, date, action, pdfPath);
                LogWorkAction(_est.Id, _current.Year, _current.Number, action.ToString());
                return true;
            }
            catch (Exception ex)
            {
                UiHelpers.Error(this, "صار خطأ مع Excel:" + Environment.NewLine + ex.Message);
                return false;
            }
            finally
            {
                UseWaitCursor = false;
            }
        }

        // بيسجّل حركة إصدار الأجور (فتح بالإكسل/طباعة/PDF) بسجلّ سير العمل، لتبويب «المتابعة والتقارير»
        private void LogWorkAction(Guid establishmentId, int year, int quarter, string action)
        {
            _data.WorkLog.Add(new WorkLogEntry
            {
                Timestamp = DateTime.Now,
                EstablishmentId = establishmentId,
                Year = year,
                Quarter = quarter,
                Action = action
            });
            MarkDirty();
            RefreshWorkLogGrid();   // موجودة بتبويب المتابعة والتقارير؛ بتحدّث نفسها إذا مفتوح على نفس المؤسسة
        }

        private void OpenInExcel()
        {
            if (!CheckReadyToPrint()) return;
            RunExcel(ExcelAction.OpenInExcel, null);
        }

        private void PrintDirect()
        {
            if (!CheckReadyToPrint()) return;
            if (UiHelpers.Ask(this, "بدك تطبع النموذج " + FormLabel() + " (" + _est.ToString() + ") على الطابعة الافتراضية؟") != DialogResult.Yes) return;
            RunExcel(ExcelAction.Print, null);
        }

        private void SavePdf()
        {
            if (!CheckReadyToPrint()) return;

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "PDF (*.pdf)|*.pdf";
                dlg.FileName = "CNSS190A_" + FormLabel() + ".pdf";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                if (RunExcel(ExcelAction.ExportPdf, dlg.FileName))
                {
                    try
                    {
                        Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true });
                    }
                    catch (Exception) { }
                }
            }
        }

        // =====================================================================
        //  حفظ تلقائي
        // =====================================================================
        private void MarkDirty()
        {
            _dirty = true;
            _saveTimer.Stop();
            _saveTimer.Start();
        }

        private void SaveIfDirty()
        {
            if (!_dirty) return;
            try
            {
                DataStore.Save(_data);
                _dirty = false;
            }
            catch (Exception ex)
            {
                UiHelpers.Error(this, "ما قدرت احفظ البيانات:" + Environment.NewLine + ex.Message);
            }
        }
    }
}
