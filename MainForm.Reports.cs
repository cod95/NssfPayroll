using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace NssfPayroll
{
    /// <summary>
    /// تبويب "المتابعة والتقارير": تصفّح كل المؤسسات وموظفيها (مع فلترة وتواريخ استلام/ترك)، سجلّ
    /// حركة إصدار الأجور (فتح بالإكسل/طباعة/PDF) لكل مؤسسة مع تصدير تقرير عنو، وأجور أي موظف لحالو
    /// فصلي أو شهري عبر كل الفترات المسجّلة.
    /// </summary>
    public partial class MainForm
    {
        private const string RepAllEstablishments = "— كل المؤسسات —";

        private readonly ComboBox _cmbRepEst = new ComboBox();
        private readonly Label _lblRepInfo = new Label();

        private readonly TextBox _txtRepSearch = new TextBox();
        private readonly ComboBox _cmbRepStatus = new ComboBox();
        private readonly DataGridView _gridRepEmployees = new DataGridView();
        private List<KeyValuePair<Establishment, Employee>> _repEmployeeRows = new List<KeyValuePair<Establishment, Employee>>();

        private readonly Label _lblRepEmpTitle = new Label();
        private readonly ComboBox _cmbRepGranularity = new ComboBox();
        private readonly DataGridView _gridRepWages = new DataGridView();
        private Button _btnRepEmpExcel;
        private Button _btnRepEmpPrint;
        private Button _btnRepEmpPdf;

        private readonly DataGridView _gridRepLog = new DataGridView();
        private Button _btnRepExportLog;

        // =====================================================================
        //  بناء الواجهة
        // =====================================================================
        private void BuildReportsTab(TabPage page)
        {
            // ---------- الشريط العلوي: اختيار المؤسسة وبياناتها ----------
            _cmbRepEst.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbRepEst.Width = 260;
            _cmbRepEst.SelectedIndexChanged += delegate
            {
                RefreshEstInfo();
                RefreshEmployeesGrid();
                RefreshWorkLogGrid();
            };

            _lblRepInfo.AutoSize = true;
            _lblRepInfo.ForeColor = System.Drawing.Color.DimGray;
            _lblRepInfo.Margin = new Padding(14, 9, 4, 0);

            FlowLayoutPanel top = new FlowLayoutPanel();
            top.Dock = DockStyle.Fill;
            top.AutoSize = true;
            top.Padding = new Padding(6);
            top.Controls.Add(UiHelpers.MakeLabel("المؤسسة"));
            top.Controls.Add(_cmbRepEst);
            top.Controls.Add(_lblRepInfo);

            // ---------- يسار: الموظفون (بحث + فلتر + جدول) ----------
            Label lblEmpHeader = new Label();
            lblEmpHeader.Text = "الموظفون";
            lblEmpHeader.Font = _boldFont;
            lblEmpHeader.AutoSize = true;
            lblEmpHeader.Dock = DockStyle.Top;
            lblEmpHeader.Padding = new Padding(4, 6, 4, 6);

            _txtRepSearch.Width = 160;
            _txtRepSearch.TextChanged += delegate { RefreshEmployeesGrid(); };

            _cmbRepStatus.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbRepStatus.Width = 110;
            _cmbRepStatus.Items.AddRange(new object[] { "الكل", "نشيطين", "تاركين" });
            _cmbRepStatus.SelectedIndex = 0;
            _cmbRepStatus.SelectedIndexChanged += delegate { RefreshEmployeesGrid(); };

            FlowLayoutPanel empFilter = new FlowLayoutPanel();
            empFilter.Dock = DockStyle.Top;
            empFilter.AutoSize = true;
            empFilter.Padding = new Padding(4, 0, 4, 4);
            empFilter.Controls.Add(UiHelpers.MakeLabel("بحث بالاسم"));
            empFilter.Controls.Add(_txtRepSearch);
            empFilter.Controls.Add(UiHelpers.MakeLabel("الحالة"));
            empFilter.Controls.Add(_cmbRepStatus);

            _gridRepEmployees.Dock = DockStyle.Fill;
            _gridRepEmployees.ReadOnly = true;
            _gridRepEmployees.AllowUserToAddRows = false;
            _gridRepEmployees.AllowUserToDeleteRows = false;
            _gridRepEmployees.RowHeadersVisible = false;
            _gridRepEmployees.MultiSelect = false;
            _gridRepEmployees.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            _gridRepEmployees.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _gridRepEmployees.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _gridRepEmployees.Columns.Add("est", "المؤسسة");     // بيظهر بس بوضع "كل المؤسسات"
            _gridRepEmployees.Columns.Add("name", "الاسم");
            _gridRepEmployees.Columns.Add("nssf", "رقم الضمان");
            _gridRepEmployees.Columns.Add("marital", "الحالة");
            _gridRepEmployees.Columns.Add("start", "تاريخ الاستخدام");
            _gridRepEmployees.Columns.Add("end", "تاريخ الترك");
            foreach (DataGridViewColumn c in _gridRepEmployees.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            _gridRepEmployees.Columns["name"].FillWeight = 160;
            _gridRepEmployees.SelectionChanged += delegate { RefreshWagesGrid(); };

            TableLayoutPanel leftPanel = new TableLayoutPanel();
            leftPanel.Dock = DockStyle.Fill;
            leftPanel.ColumnCount = 1;
            leftPanel.RowCount = 3;
            leftPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            leftPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            leftPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            leftPanel.Controls.Add(lblEmpHeader, 0, 0);
            leftPanel.Controls.Add(empFilter, 0, 1);
            leftPanel.Controls.Add(_gridRepEmployees, 0, 2);

            // ---------- يمين فوق: أجور الموظف المختار (فصلي/شهري) ----------
            _lblRepEmpTitle.Text = "أجور الموظف: (اختار موظف من القائمة)";
            _lblRepEmpTitle.Font = _boldFont;
            _lblRepEmpTitle.AutoSize = true;
            _lblRepEmpTitle.Margin = new Padding(4, 6, 4, 6);

            _cmbRepGranularity.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbRepGranularity.Width = 90;
            _cmbRepGranularity.Items.AddRange(new object[] { "فصلي", "شهري" });
            _cmbRepGranularity.SelectedIndex = 0;
            _cmbRepGranularity.SelectedIndexChanged += delegate { RefreshWagesGrid(); };

            FlowLayoutPanel wagesHeader = new FlowLayoutPanel();
            wagesHeader.Dock = DockStyle.Top;
            wagesHeader.AutoSize = true;
            wagesHeader.Padding = new Padding(4, 0, 4, 4);
            wagesHeader.Controls.Add(_lblRepEmpTitle);
            wagesHeader.Controls.Add(_cmbRepGranularity);

            _gridRepWages.Dock = DockStyle.Fill;
            _gridRepWages.ReadOnly = true;
            _gridRepWages.AllowUserToAddRows = false;
            _gridRepWages.AllowUserToDeleteRows = false;
            _gridRepWages.RowHeadersVisible = false;
            _gridRepWages.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _gridRepWages.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            BuildWagesColumns(false);

            _btnRepEmpExcel = UiHelpers.MakeButton("فتح تقرير الموظف في Excel", delegate { RunEmployeeReport(ExcelAction.OpenInExcel, null); });
            _btnRepEmpPrint = UiHelpers.MakeButton("طباعة تقرير الموظف", delegate { RunEmployeeReport(ExcelAction.Print, null); });
            _btnRepEmpPdf = UiHelpers.MakeButton("حفظ PDF", delegate { SaveEmployeeReportPdf(); });
            _btnRepEmpExcel.Enabled = false;
            _btnRepEmpPrint.Enabled = false;
            _btnRepEmpPdf.Enabled = false;

            FlowLayoutPanel wagesButtons = new FlowLayoutPanel();
            wagesButtons.Dock = DockStyle.Bottom;
            wagesButtons.AutoSize = true;
            wagesButtons.Padding = new Padding(4);
            wagesButtons.Controls.Add(_btnRepEmpExcel);
            wagesButtons.Controls.Add(_btnRepEmpPrint);
            wagesButtons.Controls.Add(_btnRepEmpPdf);

            TableLayoutPanel wagesPanel = new TableLayoutPanel();
            wagesPanel.Dock = DockStyle.Fill;
            wagesPanel.ColumnCount = 1;
            wagesPanel.RowCount = 3;
            wagesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            wagesPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            wagesPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            wagesPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            wagesPanel.Controls.Add(wagesHeader, 0, 0);
            wagesPanel.Controls.Add(_gridRepWages, 0, 1);
            wagesPanel.Controls.Add(wagesButtons, 0, 2);

            // ---------- يمين تحت: سجلّ حركة إصدار الأجور + تصدير تقرير ----------
            Label lblLogHeader = new Label();
            lblLogHeader.Text = "سجلّ إصدار الأجور";
            lblLogHeader.Font = _boldFont;
            lblLogHeader.AutoSize = true;
            lblLogHeader.Dock = DockStyle.Top;
            lblLogHeader.Padding = new Padding(4, 6, 4, 6);

            _gridRepLog.Dock = DockStyle.Fill;
            _gridRepLog.ReadOnly = true;
            _gridRepLog.AllowUserToAddRows = false;
            _gridRepLog.AllowUserToDeleteRows = false;
            _gridRepLog.RowHeadersVisible = false;
            _gridRepLog.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            _gridRepLog.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            _gridRepLog.Columns.Add("est", "المؤسسة");
            _gridRepLog.Columns.Add("period", "الفترة");
            _gridRepLog.Columns.Add("action", "الإجراء");
            _gridRepLog.Columns.Add("time", "التاريخ والوقت");
            foreach (DataGridViewColumn c in _gridRepLog.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }

            _btnRepExportLog = UiHelpers.MakeButton("تصدير سجل العمل (CSV)", delegate { ExportWorkLog(); });
            FlowLayoutPanel logButtons = new FlowLayoutPanel();
            logButtons.Dock = DockStyle.Bottom;
            logButtons.AutoSize = true;
            logButtons.Padding = new Padding(4);
            logButtons.Controls.Add(_btnRepExportLog);

            TableLayoutPanel logPanel = new TableLayoutPanel();
            logPanel.Dock = DockStyle.Fill;
            logPanel.ColumnCount = 1;
            logPanel.RowCount = 3;
            logPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            logPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            logPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            logPanel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            logPanel.Controls.Add(lblLogHeader, 0, 0);
            logPanel.Controls.Add(_gridRepLog, 0, 1);
            logPanel.Controls.Add(logButtons, 0, 2);

            // ---------- يمين: أجور الموظف فوق + سجلّ العمل تحت ----------
            TableLayoutPanel rightPanel = new TableLayoutPanel();
            rightPanel.Dock = DockStyle.Fill;
            rightPanel.ColumnCount = 1;
            rightPanel.RowCount = 2;
            rightPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 55F));
            rightPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 45F));
            rightPanel.Controls.Add(wagesPanel, 0, 0);
            rightPanel.Controls.Add(logPanel, 0, 1);

            // ---------- التجميع: يسار (موظفون) + يمين (أجور وسجلّ) ----------
            TableLayoutPanel mainSplit = new TableLayoutPanel();
            mainSplit.Dock = DockStyle.Fill;
            mainSplit.ColumnCount = 2;
            mainSplit.RowCount = 1;
            mainSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42F));
            mainSplit.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58F));
            mainSplit.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            mainSplit.Controls.Add(leftPanel, 0, 0);
            mainSplit.Controls.Add(rightPanel, 1, 0);

            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 1;
            root.RowCount = 2;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.Controls.Add(top, 0, 0);
            root.Controls.Add(mainSplit, 0, 1);
            page.Controls.Add(root);

            RefreshReportsTab();
        }

        private void BuildWagesColumns(bool monthly)
        {
            _gridRepWages.Columns.Clear();
            _gridRepWages.Columns.Add("year", "السنة");
            _gridRepWages.Columns.Add("period", "الفترة");
            if (monthly) _gridRepWages.Columns.Add("month", "الشهر");
            _gridRepWages.Columns.Add("sm", "مرض وأمومة");
            _gridRepWages.Columns.Add("eos", "نهاية خدمة");
            _gridRepWages.Columns.Add("fa", "تعويضات عائلية");
            _gridRepWages.Columns.Add("paid", "التعويض المدفوع");
            foreach (DataGridViewColumn c in _gridRepWages.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
                c.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        // بيضيف سطر لجدول الأجور، بيراعي وجود عمود "الشهر" من عدمو حسب monthly. بيرجّع رقم السطر.
        private int AddWageRow(bool monthly, string year, string period, string month, string sm, string eos, string fa, string paid)
        {
            if (monthly)
                return _gridRepWages.Rows.Add(year, period, month, sm, eos, fa, paid);
            return _gridRepWages.Rows.Add(year, period, sm, eos, fa, paid);
        }

        // =====================================================================
        //  التحديث
        // =====================================================================

        // بينعمل مرة أول ما يتبنى التبويب، وكل مرة المستخدم يفتحو (لحتى يلتقط تعديلات صارت بتبويب تاني)
        private void RefreshReportsTab()
        {
            RefreshRepEstCombo();
            RefreshEstInfo();
            RefreshEmployeesGrid();
            RefreshWorkLogGrid();
        }

        private void RefreshRepEstCombo()
        {
            Establishment keep = _cmbRepEst.SelectedItem as Establishment;
            Guid? keepId = keep != null ? (Guid?)keep.Id : null;

            _cmbRepEst.BeginUpdate();
            _cmbRepEst.Items.Clear();
            _cmbRepEst.Items.Add(RepAllEstablishments);
            foreach (Establishment est in _data.Establishments) _cmbRepEst.Items.Add(est);
            _cmbRepEst.EndUpdate();

            int idx = 0;
            if (keepId.HasValue)
            {
                for (int i = 1; i < _cmbRepEst.Items.Count; i++)
                {
                    Establishment e = _cmbRepEst.Items[i] as Establishment;
                    if (e != null && e.Id == keepId.Value) { idx = i; break; }
                }
            }
            if (_cmbRepEst.Items.Count > 0) _cmbRepEst.SelectedIndex = idx;
        }

        private void RefreshEstInfo()
        {
            Establishment est = _cmbRepEst.SelectedItem as Establishment;
            if (est == null)
            {
                int totalEmp = _data.Establishments.Sum(e => e.Employees.Count);
                _lblRepInfo.Text = "كل المؤسسات: " + _data.Establishments.Count + " مؤسسة، " + totalEmp + " موظف بالمجموع.";
                return;
            }

            _lblRepInfo.Text =
                "رقم الضمان: " + (string.IsNullOrWhiteSpace(est.NssfNumber) ? "—" : est.NssfNumber)
                + "     صاحب المؤسسة: " + (string.IsNullOrWhiteSpace(est.OwnerName) ? "—" : est.OwnerName)
                + "     الهاتف: " + (string.IsNullOrWhiteSpace(est.OwnerPhone) ? "—" : est.OwnerPhone)
                + "     عدد الموظفين: " + est.Employees.Count;
        }

        private void RefreshEmployeesGrid()
        {
            bool allMode = !(_cmbRepEst.SelectedItem is Establishment);
            _gridRepEmployees.Columns["est"].Visible = allMode;

            _gridRepEmployees.Rows.Clear();
            _repEmployeeRows.Clear();

            IEnumerable<Establishment> ests = allMode
                ? (IEnumerable<Establishment>)_data.Establishments
                : new Establishment[] { (Establishment)_cmbRepEst.SelectedItem };

            string search = (_txtRepSearch.Text ?? "").Trim();
            int status = _cmbRepStatus.SelectedIndex;   // 0 الكل / 1 نشيطين / 2 تاركين

            foreach (Establishment est in ests)
            {
                foreach (Employee emp in est.Employees)
                {
                    if (search.Length > 0 &&
                        (emp.FullName ?? "").IndexOf(search, StringComparison.CurrentCultureIgnoreCase) < 0)
                        continue;
                    if (status == 1 && emp.EndDate.HasValue) continue;
                    if (status == 2 && !emp.EndDate.HasValue) continue;

                    _gridRepEmployees.Rows.Add(
                        est.Name,
                        string.IsNullOrWhiteSpace(emp.FullName) ? "(بدون اسم)" : emp.FullName,
                        emp.NssfNumber,
                        emp.IsMarried ? "متأهل" : "أعزب",
                        FormatRepDate(emp.StartDate),
                        FormatRepDate(emp.EndDate));
                    _repEmployeeRows.Add(new KeyValuePair<Establishment, Employee>(est, emp));
                }
            }

            RefreshWagesGrid();   // القائمة تغيّرت، فالسطر المختار (إذا في) ممكن يكون تغيّر مكانو
        }

        private void RefreshWagesGrid()
        {
            bool monthly = _cmbRepGranularity.SelectedIndex == 1;
            BuildWagesColumns(monthly);
            _gridRepWages.Rows.Clear();

            int sel = _gridRepEmployees.SelectedRows.Count > 0 ? _gridRepEmployees.SelectedRows[0].Index : -1;
            if (sel < 0 || sel >= _repEmployeeRows.Count)
            {
                _lblRepEmpTitle.Text = "أجور الموظف: (اختار موظف من القائمة يسار)";
                _btnRepEmpExcel.Enabled = false;
                _btnRepEmpPrint.Enabled = false;
                _btnRepEmpPdf.Enabled = false;
                return;
            }
            _btnRepEmpExcel.Enabled = true;
            _btnRepEmpPrint.Enabled = true;
            _btnRepEmpPdf.Enabled = true;

            Establishment est = _repEmployeeRows[sel].Key;
            Employee emp = _repEmployeeRows[sel].Value;
            string empName = string.IsNullOrWhiteSpace(emp.FullName) ? "(بدون اسم)" : emp.FullName;

            List<EmployeeWageReportRow> rows = WageRules.BuildWageReport(est, emp, monthly);
            foreach (EmployeeWageReportRow r in rows)
                AddWageRow(monthly, r.Year.ToString(CultureInfo.InvariantCulture), r.Period, r.Month,
                    NumberInput.Format(r.SickMaternity), NumberInput.Format(r.EndOfService),
                    NumberInput.Format(r.FamilyAllowance), NumberInput.Format(r.FamilyAllowancePaid));

            if (rows.Count > 0)
            {
                decimal totalSm = rows.Sum(r => r.SickMaternity);
                decimal totalEos = rows.Sum(r => r.EndOfService);
                decimal totalFa = rows.Sum(r => r.FamilyAllowance);
                decimal totalPaid = rows.Sum(r => r.FamilyAllowancePaid);
                int idx = monthly
                    ? AddWageRow(true, "", "", "المجموع", NumberInput.Format(totalSm), NumberInput.Format(totalEos), NumberInput.Format(totalFa), NumberInput.Format(totalPaid))
                    : AddWageRow(false, "", "المجموع", null, NumberInput.Format(totalSm), NumberInput.Format(totalEos), NumberInput.Format(totalFa), NumberInput.Format(totalPaid));
                _gridRepWages.Rows[idx].DefaultCellStyle.Font = _boldFont;
                _gridRepWages.Rows[idx].DefaultCellStyle.BackColor = System.Drawing.Color.Gainsboro;
            }

            _lblRepEmpTitle.Text = "أجور الموظف: " + empName + "  —  " + est.Name
                + (_gridRepWages.Rows.Count == 0 ? "  (ما في أجور مسجّلة)" : "");
        }

        // =====================================================================
        //  طباعة تقرير الموظف (معلومات المؤسسة + الموظف + جدول الأجور + المجموع)
        // =====================================================================

        // المؤسسة والموظف المختارين حالياً بجدول الموظفين. false إذا ما في اختيار.
        private bool GetSelectedReportEmployee(out Establishment est, out Employee emp)
        {
            est = null;
            emp = null;
            int sel = _gridRepEmployees.SelectedRows.Count > 0 ? _gridRepEmployees.SelectedRows[0].Index : -1;
            if (sel < 0 || sel >= _repEmployeeRows.Count) return false;
            est = _repEmployeeRows[sel].Key;
            emp = _repEmployeeRows[sel].Value;
            return true;
        }

        private bool RunEmployeeReport(ExcelAction action, string pdfPath)
        {
            Establishment est;
            Employee emp;
            if (!GetSelectedReportEmployee(out est, out emp))
            {
                UiHelpers.Warn(this, "اختار موظف من القائمة أولاً.");
                return false;
            }

            bool monthly = _cmbRepGranularity.SelectedIndex == 1;
            List<EmployeeWageReportRow> rows = WageRules.BuildWageReport(est, emp, monthly);

            try
            {
                UseWaitCursor = true;
                Application.DoEvents();
                EmployeeReportPrinter.Run(est, emp, monthly, rows, action, pdfPath);
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

        private void SaveEmployeeReportPdf()
        {
            Establishment est;
            Employee emp;
            if (!GetSelectedReportEmployee(out est, out emp))
            {
                UiHelpers.Warn(this, "اختار موظف من القائمة أولاً.");
                return;
            }

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "PDF (*.pdf)|*.pdf";
                string safeName = string.IsNullOrWhiteSpace(emp.FullName) ? "موظف" : emp.FullName;
                dlg.FileName = "تقرير_أجور_" + safeName + ".pdf";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                if (RunEmployeeReport(ExcelAction.ExportPdf, dlg.FileName))
                {
                    try { Process.Start(new ProcessStartInfo(dlg.FileName) { UseShellExecute = true }); }
                    catch (Exception) { }
                }
            }
        }

        // بيستدعيها RunExcel بتبويب الأجور كل مرة الأجور تنفتح/تنطبع/تصدّر بنجاح
        private void RefreshWorkLogGrid()
        {
            _gridRepLog.Rows.Clear();
            foreach (WorkLogEntry e in CurrentWorkLogRows())
            {
                Establishment est = _data.FindEstablishment(e.EstablishmentId);
                _gridRepLog.Rows.Add(
                    est != null ? est.Name : "(مؤسسة محذوفة)",
                    e.Year.ToString(CultureInfo.InvariantCulture) + "-2" + e.Quarter,
                    ActionText(e.Action),
                    e.Timestamp.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture));
            }
        }

        private List<WorkLogEntry> CurrentWorkLogRows()
        {
            Establishment sel = _cmbRepEst.SelectedItem as Establishment;
            IEnumerable<WorkLogEntry> src = _data.WorkLog;
            if (sel != null)
                src = src.Where(e => e.EstablishmentId == sel.Id);
            return src.OrderByDescending(e => e.Timestamp).ToList();
        }

        private static string FormatRepDate(DateTime? d)
        {
            return d.HasValue ? d.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : "—";
        }

        private static string ActionText(string action)
        {
            switch (action)
            {
                case "OpenInExcel": return "فتح بالإكسل";
                case "Print": return "طباعة";
                case "ExportPdf": return "تصدير PDF";
                default: return action ?? "";
            }
        }

        // =====================================================================
        //  تصدير تقرير سجل العمل (CSV)
        // =====================================================================
        private void ExportWorkLog()
        {
            List<WorkLogEntry> rows = CurrentWorkLogRows();
            if (rows.Count == 0)
            {
                UiHelpers.Info(this, "ما في حركة مسجّلة لتصدّرها.");
                return;
            }

            using (SaveFileDialog dlg = new SaveFileDialog())
            {
                dlg.Filter = "CSV (*.csv)|*.csv";
                dlg.FileName = "سجل_العمل_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".csv";
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                try
                {
                    StringBuilder sb = new StringBuilder();
                    sb.AppendLine(CsvRow(new string[] { "المؤسسة", "السنة", "الفترة", "الإجراء", "التاريخ والوقت" }));
                    foreach (WorkLogEntry e in rows)
                    {
                        Establishment est = _data.FindEstablishment(e.EstablishmentId);
                        sb.AppendLine(CsvRow(new string[] {
                            est != null ? est.Name : "(مؤسسة محذوفة)",
                            e.Year.ToString(CultureInfo.InvariantCulture),
                            "2" + e.Quarter,
                            ActionText(e.Action),
                            e.Timestamp.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture)
                        }));
                    }
                    // BOM حتى Excel يفتح العربي صح
                    File.WriteAllText(dlg.FileName, sb.ToString(), new UTF8Encoding(true));
                    UiHelpers.Info(this, "تصدّر الملف:" + Environment.NewLine + dlg.FileName);
                }
                catch (Exception ex)
                {
                    UiHelpers.Error(this, "ما قدرت اصدّر الملف:" + Environment.NewLine + ex.Message);
                }
            }
        }

        private static string CsvRow(string[] cells)
        {
            string[] escaped = new string[cells.Length];
            for (int i = 0; i < cells.Length; i++) escaped[i] = CsvCell(cells[i]);
            return string.Join(",", escaped);
        }

        private static string CsvCell(string s)
        {
            s = s ?? "";
            if (s.IndexOf(',') >= 0 || s.IndexOf('"') >= 0 || s.IndexOf('\n') >= 0)
                return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }
}
