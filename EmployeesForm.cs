using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;

namespace NssfPayroll
{
    /// <summary>
    /// نافذة موظفي مؤسسة وحدة: قائمة الموظفين + بطاقة الموظف.
    /// كل الحقول اختيارية. لما تختار "متأهل" بيظهر حقل عدد الأولاد فوراً.
    /// كل تعديل بينحفظ لحالو (ما في زر حفظ).
    /// </summary>
    public class EmployeesForm : Form
    {
        private readonly Establishment _est;
        private readonly Action _onChanged;
        private readonly int _referenceYear;      // سنة العمل: يلي تركوا قبلها بيختفوا من القائمة
        private BindingList<Employee> _list;
        private bool _loading;
        private Employee _shown;      // الموظف المعروض بالبطاقة هلأ

        private readonly ListBox _lst = new ListBox();
        private readonly CheckBox _chkLeft = new CheckBox();
        private readonly TableLayoutPanel _card = new TableLayoutPanel();
        private readonly TextBox _txtName = new TextBox();
        private readonly TextBox _txtNssf = new TextBox();
        private readonly ComboBox _cmbMarital = new ComboBox();
        private readonly Label _lblChildren = new Label();
        private readonly TextBox _txtChildren = new TextBox();
        private readonly DateTimePicker _dtpStart = new DateTimePicker();
        private readonly DateTimePicker _dtpEnd = new DateTimePicker();
        private readonly TextBox _txtSm = new TextBox();
        private readonly TextBox _txtEos = new TextBox();
        private readonly TextBox _txtFa = new TextBox();
        private readonly DateTimePicker _dtpSm = new DateTimePicker();    // اعتباراً من (مرض وأمومة)
        private readonly DateTimePicker _dtpEos = new DateTimePicker();   // اعتباراً من (نهاية خدمة)
        private readonly DateTimePicker _dtpFa = new DateTimePicker();   // اعتباراً من (تعويضات عائلية)
        private readonly Label _lblSmPrev = new Label();      // "وقبلو" - بيظهر بس إذا في تاريخ "اعتباراً من"
        private readonly Label _lblEosPrev = new Label();
        private readonly Label _lblFaPrev = new Label();
        private readonly TextBox _txtSmPrev = new TextBox();  // الراتب القديم (قبل تاريخ "اعتباراً من")
        private readonly TextBox _txtEosPrev = new TextBox();
        private readonly TextBox _txtFaPrev = new TextBox();

        public EmployeesForm(Establishment est, int referenceYear, Action onChanged)
        {
            _est = est;
            _referenceYear = referenceYear;
            _onChanged = onChanged;

            Text = "موظفو المؤسسة: " + _est.ToString();
            RightToLeft = RightToLeft.Yes;
            RightToLeftLayout = true;
            Font = new Font("Segoe UI", 11F);
            ClientSize = new Size(1180, 660);
            MinimumSize = new Size(1000, 600);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;

            BuildUi();
            RebuildList(null);
        }

        // =====================================================================
        //  الواجهة
        // =====================================================================
        private void BuildUi()
        {
            // ---------- القائمة ----------
            Label lblList = new Label();
            lblList.Text = "الموظفين";
            lblList.Font = new Font(Font, FontStyle.Bold);
            lblList.AutoSize = true;
            lblList.Dock = DockStyle.Top;
            lblList.Padding = new Padding(4, 4, 4, 6);

            _chkLeft.Text = "إظهار المتروكين من سنوات سابقة";
            _chkLeft.AutoSize = true;
            _chkLeft.Dock = DockStyle.Top;
            _chkLeft.Padding = new Padding(4, 0, 4, 6);
            _chkLeft.CheckedChanged += delegate { RebuildList(_shown); };

            _lst.Dock = DockStyle.Fill;
            _lst.IntegralHeight = false;
            _lst.SelectedIndexChanged += delegate
            {
                if (_loading) return;
                SyncDates(_shown);     // نثبّت تواريخ الموظف يلي كنا عارضينو قبل ما نغيّر البطاقة
                LoadSelected();
            };
            FormClosing += delegate { SyncDates(_shown); };

            FlowLayoutPanel listButtons = new FlowLayoutPanel();
            listButtons.Dock = DockStyle.Bottom;
            listButtons.AutoSize = true;
            listButtons.Controls.Add(UiHelpers.MakeButton("موظف جديد", delegate { AddEmployee(); }));
            listButtons.Controls.Add(UiHelpers.MakeButton("حذف الموظف", delegate { DeleteEmployee(); }));

            Panel listPanel = new Panel();
            listPanel.Dock = DockStyle.Fill;
            listPanel.Padding = new Padding(6);
            listPanel.Controls.Add(_lst);          // Fill أول
            listPanel.Controls.Add(listButtons);   // وبعدين Bottom
            listPanel.Controls.Add(_chkLeft);      // وبعدين Top (تحت العنوان)
            listPanel.Controls.Add(lblList);       // وبعدين Top

            // ---------- البطاقة ----------
            _card.Dock = DockStyle.Fill;
            _card.ColumnCount = 2;
            _card.Padding = new Padding(16);
            _card.AutoScroll = true;
            _card.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

            _txtName.Width = 380;
            _txtNssf.Width = 240;

            _cmbMarital.DropDownStyle = ComboBoxStyle.DropDownList;
            _cmbMarital.Width = 160;
            _cmbMarital.Items.AddRange(new object[] { "أعزب", "متأهل" });
            _cmbMarital.SelectedIndex = 0;

            _lblChildren.Text = "عدد الأولاد";
            _txtChildren.Width = 100;
            SetChildrenVisible(false);

            foreach (DateTimePicker dtp in new DateTimePicker[] { _dtpStart, _dtpEnd, _dtpSm, _dtpEos, _dtpFa })
            {
                dtp.Format = DateTimePickerFormat.Short;
                dtp.ShowCheckBox = true;     // إذا الخانة مو معلّمة = ما في تاريخ
                dtp.Width = 170;
                dtp.Checked = false;
            }
            foreach (DateTimePicker dtp in new DateTimePicker[] { _dtpSm, _dtpEos, _dtpFa })
                dtp.Width = 140;             // أضيق شوي لأنها جنب خانة الراتب بنفس السطر

            foreach (TextBox t in new TextBox[] { _txtSm, _txtEos, _txtFa, _txtSmPrev, _txtEosPrev, _txtFaPrev })
            {
                t.Width = 200;
                t.TextAlign = HorizontalAlignment.Center;
            }
            foreach (TextBox t in new TextBox[] { _txtSmPrev, _txtEosPrev, _txtFaPrev })
                t.Width = 140;   // أضيق شوي لأنها جنب راتب+تاريخ بنفس السطر

            AddCardRow("الاسم الثلاثي", _txtName);
            AddCardRow("رقم الضمان", _txtNssf);
            AddCardRow("الحالة الاجتماعية", _cmbMarital);
            AddCardRow(_lblChildren, _txtChildren);
            AddCardRow("تاريخ الاستخدام", _dtpStart);
            AddCardRow("تاريخ الترك (اختياري)", _dtpEnd);
            AddCardRow("راتب المرض والأمومة (شهري)", WageRow(_txtSm, _dtpSm, _lblSmPrev, _txtSmPrev));
            AddCardRow("راتب نهاية الخدمة (شهري)", WageRow(_txtEos, _dtpEos, _lblEosPrev, _txtEosPrev));
            AddCardRow("راتب التعويض العائلي (شهري)", WageRow(_txtFa, _dtpFa, _lblFaPrev, _txtFaPrev));

            // ---------- الأحداث ----------
            _txtName.TextChanged += delegate { Apply(e => e.FullName = _txtName.Text.Trim(), true); };
            _txtNssf.TextChanged += delegate { Apply(e => e.NssfNumber = _txtNssf.Text.Trim(), false); };

            _cmbMarital.SelectedIndexChanged += delegate { OnMaritalChanged(); };

            _txtChildren.KeyPress += delegate(object s, KeyPressEventArgs ev) { if (!char.IsControl(ev.KeyChar) && !char.IsDigit(ev.KeyChar)) ev.Handled = true; };
            _txtChildren.TextChanged += delegate
            {
                int? n;
                if (NumberInput.TryParseInt(_txtChildren.Text, out n))
                {
                    Apply(e =>
                    {
                        e.ChildrenCount = n;
                        e.FamilyAllowanceWage = WageRules.ComputeFamilyAllowance(e.IsMarried, e.ChildrenCount);
                    }, false);
                    RefreshFamilyAllowanceDisplay();
                }
            };

            // التواريخ: منقرا حالة الخانة والتاريخ مباشرة من الـ DateTimePicker (عند التغيير، وعند الطلوع منو،
            // وقبل ما نبدّل الموظف أو نسكّر النافذة) لحتى ما يضيع أي تعديل
            foreach (DateTimePicker dtp in new DateTimePicker[] { _dtpStart, _dtpEnd, _dtpSm, _dtpEos, _dtpFa })
            {
                dtp.ValueChanged += delegate { SyncDates(_shown, true); UpdatePrevVisibility(); };
                dtp.Leave += delegate { SyncDates(_shown, true); UpdatePrevVisibility(); };
                // تفعيل/إلغاء خانة "ما في تاريخ" (Checked) ما بيغيّر Value ولا بيطلّع من الخانة (Leave)،
                // وDateTimePicker ما عندو حدث CheckedChanged جاهز، فمنعتمد MouseUp/KeyUp لنمسك أي
                // تبديل عليها (بالماوس أو الكيبورد) وما ننتظر لحد ما تتبدّل الموظف أو تسكّر النافذة
                dtp.MouseUp += delegate { SyncDates(_shown, true); UpdatePrevVisibility(); };
                dtp.KeyUp += delegate { SyncDates(_shown, true); UpdatePrevVisibility(); };
            }

            WireWage(_txtSm, delegate(Employee e, decimal v) { e.SickMaternityWage = v; }, delegate(Employee e) { return e.SickMaternityWage; });
            WireWage(_txtEos, delegate(Employee e, decimal v) { e.EndOfServiceWage = v; }, delegate(Employee e) { return e.EndOfServiceWage; });
            WireWage(_txtFa, delegate(Employee e, decimal v) { e.FamilyAllowanceWage = v; }, delegate(Employee e) { return e.FamilyAllowanceWage; });
            WireWage(_txtSmPrev, delegate(Employee e, decimal v) { e.SickMaternityPreviousWage = v; }, delegate(Employee e) { return e.SickMaternityPreviousWage; });
            WireWage(_txtEosPrev, delegate(Employee e, decimal v) { e.EndOfServicePreviousWage = v; }, delegate(Employee e) { return e.EndOfServicePreviousWage; });
            WireWage(_txtFaPrev, delegate(Employee e, decimal v) { e.FamilyAllowancePreviousWage = v; }, delegate(Employee e) { return e.FamilyAllowancePreviousWage; });

            // ---------- التجميع ----------
            TableLayoutPanel root = new TableLayoutPanel();
            root.Dock = DockStyle.Fill;
            root.ColumnCount = 2;
            root.RowCount = 1;
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 280F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.Controls.Add(listPanel, 0, 0);
            root.Controls.Add(_card, 1, 0);

            Label hint = new Label();
            hint.Text = "كل الحقول اختيارية والتعديلات بتنحفظ لحالها. الرواتب شهرية (بتتعبّى تلقائياً بجدول الفصل لكل شهر). "
                + "لو صارت زيادة: حط تاريخها بـ«اعتباراً من» واكتب الراتب الجديد بالخانة الأساسية، وبيظهر حقل «وقبلو» "
                + "تحط فيه الراتب القديم. الأشهر قبل التاريخ بتاخد القديم، ومن هيك تاريخ وبعدو بتاخد الجديد (بلا تاريخ = الراتب الحالي بس، كل الأشهر). "
                + "تاريخ الترك: الموظف بينحسب لحد شهر الترك، وما بينحسب بالفصول يلي بعدو، وبالسنة يلي بعدها بيختفي من القائمة "
                + "(بتقدر تظهرو بالخانة فوق القائمة). تاريخ الاستخدام للمعلومات بس.";
            hint.AutoSize = false;
            hint.Height = 98;
            hint.ForeColor = Color.DimGray;
            hint.Dock = DockStyle.Bottom;
            hint.Padding = new Padding(10, 6, 10, 8);

            Controls.Add(root);   // Fill أول
            Controls.Add(hint);   // وبعدين Bottom
        }

        private void AddCardRow(string label, Control input)
        {
            Label l = new Label();
            l.Text = label;
            l.AutoSize = true;
            l.Margin = new Padding(6, 10, 14, 6);
            AddCardRow(l, input);
        }

        private void AddCardRow(Label label, Control input)
        {
            label.AutoSize = true;
            label.Margin = new Padding(6, 10, 14, 6);
            input.Margin = new Padding(6, 6, 6, 6);
            _card.Controls.Add(label);
            _card.Controls.Add(input);
        }

        // صف راتب + تاريخ "اعتباراً من" + الراتب القديم (يظهر بس إذا في تاريخ)، كلهن بنفس السطر
        private static FlowLayoutPanel WageRow(TextBox box, DateTimePicker dtp, Label prevLabel, TextBox prevBox)
        {
            FlowLayoutPanel p = new FlowLayoutPanel();
            p.AutoSize = true;
            p.WrapContents = false;
            p.Margin = new Padding(0);

            Label l = new Label();
            l.Text = "اعتباراً من";
            l.AutoSize = true;
            l.Margin = new Padding(12, 8, 4, 0);

            prevLabel.Text = "وقبلو";
            prevLabel.AutoSize = true;
            prevLabel.Margin = new Padding(10, 8, 4, 0);

            box.Margin = new Padding(0);
            dtp.Margin = new Padding(0);
            prevBox.Margin = new Padding(0);

            p.Controls.Add(box);
            p.Controls.Add(l);
            p.Controls.Add(dtp);
            p.Controls.Add(prevLabel);
            p.Controls.Add(prevBox);
            return p;
        }

        // بيظهر/يخبّي حقول "الراتب القديم" الثلاثة، حسب إذا كل تاريخ "اعتباراً من" معلّم أو لأ
        private void UpdatePrevVisibility()
        {
            _lblSmPrev.Visible = _dtpSm.Checked;
            _txtSmPrev.Visible = _dtpSm.Checked;
            _lblEosPrev.Visible = _dtpEos.Checked;
            _txtEosPrev.Visible = _dtpEos.Checked;
            _lblFaPrev.Visible = _dtpFa.Checked;
            _txtFaPrev.Visible = _dtpFa.Checked;
        }

        private void SetChildrenVisible(bool visible)
        {
            _lblChildren.Visible = visible;
            _txtChildren.Visible = visible;
        }

        // =====================================================================
        //  ربط الحقول بالموظف الحالي
        // =====================================================================
        private Employee Current()
        {
            return _lst.SelectedItem as Employee;
        }

        // بيطبّق تعديل على الموظف المختار. refreshList = لتتحدّث القائمة (لما بيتغير الاسم)
        private void Apply(Action<Employee> change, bool refreshList)
        {
            if (_loading) return;
            Employee emp = Current();
            if (emp == null) return;

            change(emp);

            if (refreshList) RefreshListItem(emp);
            if (_onChanged != null) _onChanged();
        }

        // بيحدّث نص الموظف بالقائمة (الاسم + "ترك ...") من دون ما تتغير الخانة المختارة
        private void RefreshListItem(Employee emp)
        {
            int idx = _list.IndexOf(emp);
            if (idx < 0) return;

            int sel = _lst.SelectedIndex;
            _loading = true;
            try
            {
                _list.ResetItem(idx);
                if (_lst.SelectedIndex != sel) _lst.SelectedIndex = sel;
            }
            finally
            {
                _loading = false;
            }
        }

        // موظف ترك قبل سنة العمل: بيختفي من القائمة (إلا إذا معلّم "إظهار المتروكين")
        private bool IsOldLeft(Employee e)
        {
            return e.EndDate.HasValue && e.EndDate.Value.Year < _referenceYear;
        }

        // بيبني القائمة المعروضة من موظفي المؤسسة (من دون المتروكين من سنوات سابقة إلا إذا طلبت تشوفن)
        private void RebuildList(Employee keepSelected)
        {
            List<Employee> visible = new List<Employee>();
            int oldCount = 0;
            foreach (Employee e in _est.Employees)
            {
                if (IsOldLeft(e))
                {
                    oldCount++;
                    if (!_chkLeft.Checked) continue;
                }
                visible.Add(e);
            }

            _loading = true;
            try
            {
                _list = new BindingList<Employee>(visible);
                _lst.DataSource = _list;
                _lst.DisplayMember = "DisplayName";
                if (keepSelected != null)
                {
                    int i = _list.IndexOf(keepSelected);
                    if (i >= 0) _lst.SelectedIndex = i;
                }
                _chkLeft.Text = "إظهار المتروكين من سنوات سابقة (" + oldCount + ")";
            }
            finally
            {
                _loading = false;
            }

            LoadSelected();
        }

        // بيكتب تاريخ الاستخدام والترك من الـ DateTimePicker للموظف emp (الخانة غير معلّمة = ما في تاريخ)
        private void SyncDates(Employee emp)
        {
            SyncDates(emp, false);
        }

        private void SyncDates(Employee emp, bool refreshList)
        {
            if (emp == null || _loading) return;

            DateTime? start = _dtpStart.Checked ? (DateTime?)_dtpStart.Value.Date : null;
            DateTime? end = _dtpEnd.Checked ? (DateTime?)_dtpEnd.Value.Date : null;
            DateTime? smDate = _dtpSm.Checked ? (DateTime?)_dtpSm.Value.Date : null;
            DateTime? eosDate = _dtpEos.Checked ? (DateTime?)_dtpEos.Value.Date : null;
            DateTime? faDate = _dtpFa.Checked ? (DateTime?)_dtpFa.Value.Date : null;

            bool changed = emp.StartDate != start || emp.EndDate != end
                || emp.SickMaternityEffectiveDate != smDate
                || emp.EndOfServiceEffectiveDate != eosDate
                || emp.FamilyAllowanceEffectiveDate != faDate;

            if (changed)
            {
                emp.StartDate = start;
                emp.EndDate = end;
                emp.SickMaternityEffectiveDate = smDate;
                emp.EndOfServiceEffectiveDate = eosDate;
                emp.FamilyAllowanceEffectiveDate = faDate;
                if (refreshList) RefreshListItem(emp);
                if (_onChanged != null) _onChanged();
            }
        }

        private void OnMaritalChanged()
        {
            if (_loading) return;

            bool married = _cmbMarital.SelectedIndex == 1;
            SetChildrenVisible(married);     // متأهل: بيظهر حقل عدد الأولاد فوراً

            if (!married)
            {
                _loading = true;
                _txtChildren.Text = "";
                _loading = false;
            }

            Apply(e =>
            {
                e.IsMarried = married;
                if (!married) e.ChildrenCount = null;
                // احتساب تلقائي للتعويض العائلي (2,100,000 عن الزوجة + 1,155,000 عن كل ولد، شهرياً).
                // المستخدم لسا فيه يعدّل الناتج يدوياً بعدها لو بدو؛ التعديل بيضل لحد ما الحالة
                // الاجتماعية أو عدد الأولاد يتغيّروا من جديد (عندها بينحتسب من جديد).
                e.FamilyAllowanceWage = WageRules.ComputeFamilyAllowance(e.IsMarried, e.ChildrenCount);
            }, false);

            RefreshFamilyAllowanceDisplay();
            if (married) _txtChildren.Focus();
        }

        // بيحدّث خانة راتب التعويض العائلي المعروضة بعد احتساب تلقائي (متل ما بيصير بـ WireWage.Leave)
        private void RefreshFamilyAllowanceDisplay()
        {
            Employee emp = Current();
            if (emp == null) return;
            _loading = true;
            try { _txtFa.Text = NumberInput.FormatOrBlank(emp.FamilyAllowanceWage); }
            finally { _loading = false; }
        }

        private void WireWage(TextBox box, Action<Employee, decimal> set, Func<Employee, decimal> get)
        {
            box.KeyPress += delegate(object s, KeyPressEventArgs ev)
            {
                if (!NumberInput.IsNumericKey(ev.KeyChar)) ev.Handled = true;
            };

            box.TextChanged += delegate
            {
                decimal v;
                if (NumberInput.TryParseDecimal(box.Text, out v))
                    Apply(e => set(e, v), false);
            };

            // لما تطلع من الخانة بيتنسّق الرقم (28,000,000)
            box.Leave += delegate
            {
                Employee emp = Current();
                if (emp == null) return;
                _loading = true;
                try { box.Text = NumberInput.FormatOrBlank(get(emp)); }
                finally { _loading = false; }
            };
        }

        // =====================================================================
        //  عرض الموظف المختار
        // =====================================================================
        private void LoadSelected()
        {
            Employee emp = Current();
            _shown = emp;

            _loading = true;
            try
            {
                _card.Enabled = emp != null;

                if (emp == null)
                {
                    _txtName.Text = "";
                    _txtNssf.Text = "";
                    _cmbMarital.SelectedIndex = 0;
                    _txtChildren.Text = "";
                    SetChildrenVisible(false);
                    SetDate(_dtpStart, null);
                    SetDate(_dtpEnd, null);
                    _txtSm.Text = "";
                    _txtEos.Text = "";
                    _txtFa.Text = "";
                    _txtSmPrev.Text = "";
                    _txtEosPrev.Text = "";
                    _txtFaPrev.Text = "";
                    SetDate(_dtpSm, null);
                    SetDate(_dtpEos, null);
                    SetDate(_dtpFa, null);
                    UpdatePrevVisibility();
                    return;
                }

                _txtName.Text = emp.FullName;
                _txtNssf.Text = emp.NssfNumber;
                _cmbMarital.SelectedIndex = emp.IsMarried ? 1 : 0;
                _txtChildren.Text = emp.ChildrenCount.HasValue
                    ? emp.ChildrenCount.Value.ToString(CultureInfo.InvariantCulture) : "";
                SetChildrenVisible(emp.IsMarried);
                SetDate(_dtpStart, emp.StartDate);
                SetDate(_dtpEnd, emp.EndDate);
                _txtSm.Text = NumberInput.FormatOrBlank(emp.SickMaternityWage);
                _txtEos.Text = NumberInput.FormatOrBlank(emp.EndOfServiceWage);
                _txtFa.Text = NumberInput.FormatOrBlank(emp.FamilyAllowanceWage);
                _txtSmPrev.Text = NumberInput.FormatOrBlank(emp.SickMaternityPreviousWage);
                _txtEosPrev.Text = NumberInput.FormatOrBlank(emp.EndOfServicePreviousWage);
                _txtFaPrev.Text = NumberInput.FormatOrBlank(emp.FamilyAllowancePreviousWage);
                SetDate(_dtpSm, emp.SickMaternityEffectiveDate);
                SetDate(_dtpEos, emp.EndOfServiceEffectiveDate);
                SetDate(_dtpFa, emp.FamilyAllowanceEffectiveDate);
                UpdatePrevVisibility();
            }
            finally
            {
                _loading = false;
            }
        }

        private static void SetDate(DateTimePicker dtp, DateTime? value)
        {
            if (value.HasValue)
            {
                DateTime v = value.Value;
                if (v < dtp.MinDate) v = dtp.MinDate;
                if (v > dtp.MaxDate) v = dtp.MaxDate;
                dtp.Value = v;
                dtp.Checked = true;
            }
            else
            {
                dtp.Value = DateTime.Today;   // لازم قبل Checked لأنو تعيين Value بيعلّم الخانة
                dtp.Checked = false;
            }
        }

        private void AddEmployee()
        {
            SyncDates(_shown);
            Employee emp = new Employee();
            _est.Employees.Add(emp);
            _list.Add(emp);
            _lst.SelectedIndex = _list.Count - 1;
            LoadSelected();
            if (_onChanged != null) _onChanged();
            _txtName.Focus();
        }

        private void DeleteEmployee()
        {
            Employee emp = Current();
            if (emp == null) return;

            int lineCount = WageRules.CountEmployeeLines(_est, emp);
            string msg = "بدك تحذف الموظف \"" + emp.DisplayName + "\"؟";
            if (lineCount > 0)
                msg += Environment.NewLine + Environment.NewLine
                     + "عندو أجور مسجّلة بـ " + lineCount + " " + (lineCount == 1 ? "فترة" : "فترات")
                     + ". حذف الموظف رح يمسح هالأجور معه من كل الفترات. ما في رجعة.";

            if (UiHelpers.Ask(this, msg) != DialogResult.Yes) return;

            WageRules.RemoveEmployeeLines(_est, emp);

            int idx = _lst.SelectedIndex;
            _est.Employees.Remove(emp);
            _list.RemoveAt(idx);
            if (_list.Count > 0) _lst.SelectedIndex = Math.Min(idx, _list.Count - 1);
            LoadSelected();
            if (_onChanged != null) _onChanged();
        }
    }
}
