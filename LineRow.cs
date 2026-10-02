using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;

namespace NssfPayroll
{
    /// <summary>سياق الجدول الحالي: المؤسسة والفترة وطريقة العرض.</summary>
    public class PeriodContext
    {
        public Establishment Establishment;
        public int Year;
        public int Quarter = 1;     // 1..4  (النموذج 21..24)
        public int ViewMonth;       // 0 = الفصل كامل، 1..3 = الشهر الأول/التاني/التالت
        public Rates Rates;         // النسب الحالية (لحساب "اشتراكات مستحقة" و"الاشتراكات" الصافية لكل سطر)
    }

    /// <summary>
    /// سطر بجدول الأجور. بيعرض ويعدّل الأجور حسب طريقة العرض:
    ///  - الفصل كامل: مجموع الأشهر الفعّالة (والتعديل بيتوزّع بالتساوي على الأشهر الفعّالة)
    ///  - شهر: أجور هالشهر بس
    /// الأشهر بعد تاريخ الترك ما بتنعرض ولا بتنعدّل ولا بتنحسب.
    /// </summary>
    public class LineRow
    {
        private readonly EmployeeLine _line;
        private readonly PeriodContext _ctx;

        public LineRow(EmployeeLine line, PeriodContext ctx)
        {
            _line = line;
            _ctx = ctx;
            _line.EnsureMonths();
        }

        public EmployeeLine Line { get { return _line; } }

        /// <summary>بطاقة الموظف (ممكن تكون null إذا الاسم مو مسجّل).</summary>
        public Employee Card { get { return WageRules.FindEmployee(_ctx.Establishment, _line); } }

        // ---- الخصائص يلي بيربطها الجدول ----
        public string Name
        {
            get { return _line.Name; }
            set { _line.Name = value ?? ""; }
        }

        public decimal Sm
        {
            get { return GetWage(WageKind.SickMaternity); }
            set { SetWage(WageKind.SickMaternity, value); }
        }

        public decimal Eos
        {
            get { return GetWage(WageKind.EndOfService); }
            set { SetWage(WageKind.EndOfService, value); }
        }

        public decimal Fa
        {
            get { return GetWage(WageKind.FamilyAllowance); }
            set { SetWage(WageKind.FamilyAllowance, value); }
        }

        public decimal Paid
        {
            get { return GetWage(WageKind.FamilyAllowancePaid); }
            set { SetWage(WageKind.FamilyAllowancePaid, value); }
        }

        /// <summary>الاشتراكات المستحقة لهالموظف (مجموع الفروع التلاتة، بالسقف الشهري 28 مليون
        /// للمرض والأمومة وللتعويضات العائلية، بلا سقف لنهاية الخدمة) — نفس منطق NssfContributionCalculator.</summary>
        public decimal DueContributions { get { return ComputeContributions(); } }

        /// <summary>الاشتراكات الصافية الواجب سدادها = الاشتراكات المستحقة − التعويض العائلي المدفوع.</summary>
        public decimal NetContributions { get { return DueContributions - Paid; } }

        private decimal ComputeContributions()
        {
            if (_ctx.Rates == null) return 0m;
            Employee emp = Card;

            if (_ctx.ViewMonth == 0)
            {
                decimal smC = WageRules.CappedQuarterTotal(_line, emp, _ctx.Year, _ctx.Quarter, WageKind.SickMaternity, NssfContributionCalculator.HealthMaxCeiling) * _ctx.Rates.SickMaternity;
                decimal faC = WageRules.CappedQuarterTotal(_line, emp, _ctx.Year, _ctx.Quarter, WageKind.FamilyAllowance, NssfContributionCalculator.FamilyMaxCeiling) * _ctx.Rates.FamilyAllowance;
                decimal eosC = WageRules.QuarterTotal(_line, emp, _ctx.Year, _ctx.Quarter, WageKind.EndOfService) * _ctx.Rates.EndOfService;
                return smC + faC + eosC;
            }

            int i = _ctx.ViewMonth - 1;
            if (!WageRules.IsMonthActive(emp, _ctx.Year, _ctx.Quarter, i)) return 0m;
            decimal monthSmC = Math.Min(WageRules.Get(_line.Months[i], WageKind.SickMaternity), NssfContributionCalculator.HealthMaxCeiling) * _ctx.Rates.SickMaternity;
            decimal monthFaC = Math.Min(WageRules.Get(_line.Months[i], WageKind.FamilyAllowance), NssfContributionCalculator.FamilyMaxCeiling) * _ctx.Rates.FamilyAllowance;
            decimal monthEosC = WageRules.Get(_line.Months[i], WageKind.EndOfService) * _ctx.Rates.EndOfService;
            return monthSmC + monthFaC + monthEosC;
        }

        // ---- المنطق ----
        public decimal GetWage(WageKind k)
        {
            Employee emp = Card;
            if (_ctx.ViewMonth == 0)
                return WageRules.QuarterTotal(_line, emp, _ctx.Year, _ctx.Quarter, k);

            int i = _ctx.ViewMonth - 1;
            if (!WageRules.IsMonthActive(emp, _ctx.Year, _ctx.Quarter, i)) return 0m;
            return WageRules.Get(_line.Months[i], k);
        }

        public void SetWage(WageKind k, decimal value)
        {
            Employee emp = Card;
            if (_ctx.ViewMonth == 0)
            {
                if (!WageRules.IsEvenSplit(_line, emp, _ctx.Year, _ctx.Quarter, k)) return;   // الجدول بيمنع هالتعديل قبل
                WageRules.SetQuarterTotal(_line, emp, _ctx.Year, _ctx.Quarter, k, value);
                return;
            }

            int i = _ctx.ViewMonth - 1;
            if (!WageRules.IsMonthActive(emp, _ctx.Year, _ctx.Quarter, i)) return;
            WageRules.Set(_line.Months[i], k, value);
        }

        /// <summary>هل بيجوز تعديل هالخانة؟ وإذا لأ، ليش.</summary>
        public bool CanEdit(WageKind k, out string reason)
        {
            reason = null;
            Employee emp = Card;

            if (_ctx.ViewMonth == 0)
            {
                if (WageRules.ActiveMonths(emp, _ctx.Year, _ctx.Quarter).Count == 0)
                {
                    reason = LeftReason(emp);
                    return false;
                }
                if (!WageRules.IsEvenSplit(_line, emp, _ctx.Year, _ctx.Quarter, k))
                {
                    reason = "أجور هالفصل مختلفة بين الأشهر، فما بينفع تعدّل الإجمالي. اختار شهر من «طريقة الإدخال» وعدّل أجرو.";
                    return false;
                }
                return true;
            }

            if (!WageRules.IsMonthActive(emp, _ctx.Year, _ctx.Quarter, _ctx.ViewMonth - 1))
            {
                reason = LeftReason(emp);
                return false;
            }
            return true;
        }

        private static string LeftReason(Employee emp)
        {
            string when = (emp != null && emp.EndDate.HasValue)
                ? emp.EndDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : "";
            return "الموظف ترك بتاريخ " + when + ": بينحسب لحد شهر الترك بس، وما بيتعدّل أجرو بالأشهر يلي بعدو.";
        }
    }

    /// <summary>
    /// لائحة أسطر الجدول. بتلفّ لائحة أسطر الفترة (PeriodData.Lines) وبتضل متزامنة معها
    /// (إضافة/حذف بالجدول = إضافة/حذف بالفترة).
    /// </summary>
    public class LineRowList : BindingList<LineRow>
    {
        private readonly List<EmployeeLine> _lines;
        private readonly PeriodContext _ctx;

        public LineRowList(List<EmployeeLine> lines, PeriodContext ctx)
        {
            _lines = lines;
            _ctx = ctx;

            // مباشرة على Items (بدون InsertItem) لأنو الأسطر موجودة أصلاً بالفترة
            foreach (EmployeeLine l in lines)
                Items.Add(new LineRow(l, ctx));

            AllowNew = true;
            AddingNew += delegate(object sender, AddingNewEventArgs e)
            {
                e.NewObject = new LineRow(new EmployeeLine(), _ctx);
            };
        }

        protected override void InsertItem(int index, LineRow item)
        {
            _lines.Insert(index, item.Line);
            base.InsertItem(index, item);
        }

        protected override void RemoveItem(int index)
        {
            _lines.RemoveAt(index);
            base.RemoveItem(index);
        }

        protected override void SetItem(int index, LineRow item)
        {
            _lines[index] = item.Line;
            base.SetItem(index, item);
        }

        protected override void ClearItems()
        {
            _lines.Clear();
            base.ClearItems();
        }
    }
}
