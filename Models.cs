using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace NssfPayroll
{
    /// <summary>كل بيانات البرنامج. بتنحفظ بملف JSON واحد.</summary>
    public class AppData
    {
        public List<Establishment> Establishments { get; set; } = new List<Establishment>();

        /// <summary>النسب واحدة لكل المؤسسات (نسب الضمان).</summary>
        public Rates Rates { get; set; } = new Rates();

        /// <summary>سجلّ حركة سير العمل: كل مرة الأجور تنفتح بالإكسل أو تنطبع أو تصدّر PDF لأي مؤسسة/فترة.</summary>
        public List<WorkLogEntry> WorkLog { get; set; } = new List<WorkLogEntry>();

        // آخر مؤسسة/سنة/فترة كنت شغّال عليها (لتنفتح عليها مباشرة)
        public Guid? LastEstablishmentId { get; set; }
        public int LastYear { get; set; }
        public int LastPeriod { get; set; }

        // ---- من النسخة الأولى (مؤسسة وحدة): بتنقرا مرة وحدة وبتتحوّل لمؤسسة، وبعدها بتصير null ----
        public CompanyInfo Company { get; set; }
        public List<PeriodData> Periods { get; set; }

        public Establishment FindEstablishment(Guid? id)
        {
            if (!id.HasValue) return null;
            return Establishments.FirstOrDefault(e => e.Id == id.Value);
        }

        /// <summary>
        /// بيحوّل ملفات النسخ القديمة للشكل الحالي. بيرجّع true إذا صار تحويل (لنحفظ الملف).
        ///  - النسخة 1 (مؤسسة وحدة): بتتحوّل لمؤسسة.
        ///  - النسخة 2 (أجور الفصل كاملة): بتتوزّع بالتساوي على 3 أشهر (المجموع ما بيتغير).
        /// </summary>
        public bool MigrateLegacy()
        {
            bool changed = false;

            // ---- النسخة 1 -> مؤسسة ----
            if (Company != null || Periods != null)
            {
                bool hasCompany = Company != null &&
                    (!string.IsNullOrWhiteSpace(Company.Name) || !string.IsNullOrWhiteSpace(Company.NssfNumber));
                bool hasLines = Periods != null &&
                    Periods.Any(p => p != null && p.Lines != null && p.Lines.Count > 0);

                if (Establishments.Count == 0 && (hasCompany || hasLines))
                {
                    Establishment est = new Establishment();
                    est.Name = (Company != null && !string.IsNullOrWhiteSpace(Company.Name)) ? Company.Name.Trim() : "مؤسسة";
                    est.NssfNumber = Company != null ? (Company.NssfNumber ?? "") : "";

                    if (Periods != null)
                    {
                        foreach (PeriodData p in Periods)
                        {
                            if (p == null) continue;
                            if (p.Lines == null) p.Lines = new List<EmployeeLine>();
                            foreach (EmployeeLine l in p.Lines)
                            {
                                // بالنسخة الأولى "مرض وأمومة" كان دايماً = "نهاية خدمة"
                                l.SickMaternityWage = l.EndOfServiceWage ?? 0m;
                                AddCardFromLine(est, l);
                            }
                            est.Periods.Add(p);
                        }
                    }

                    Establishments.Add(est);
                    LastEstablishmentId = est.Id;
                }

                Company = null;
                Periods = null;
                changed = true;
            }

            // ---- النسخة 2 -> أجور شهرية (وتأكيد إنو كل سطر عندو 3 أشهر) ----
            foreach (Establishment est in Establishments)
            {
                if (est.Employees == null) est.Employees = new List<Employee>();
                if (est.Periods == null) est.Periods = new List<PeriodData>();
                foreach (PeriodData p in est.Periods)
                {
                    if (p.Lines == null) p.Lines = new List<EmployeeLine>();
                    foreach (EmployeeLine l in p.Lines)
                    {
                        l.EnsureMonths();
                        if (l.MigrateWagesToMonths()) changed = true;
                    }
                }
            }

            if (WorkLog == null) WorkLog = new List<WorkLogEntry>();

            return changed;
        }

        // بيعمل بطاقة موظف من الأسماء يلي كانت مكتوبة بجداول الفترات (لتشتغل التعبئة التلقائية عليها)
        private static void AddCardFromLine(Establishment est, EmployeeLine l)
        {
            string name = (l.Name ?? "").Trim();
            if (name.Length == 0) return;
            bool exists = est.Employees.Any(e =>
                string.Equals((e.FullName ?? "").Trim(), name, StringComparison.CurrentCultureIgnoreCase));
            if (exists) return;

            Employee emp = new Employee();
            emp.FullName = name;
            emp.SickMaternityWage = l.SickMaternityWage ?? 0m;
            emp.EndOfServiceWage = l.EndOfServiceWage ?? 0m;
            emp.FamilyAllowanceWage = l.FamilyAllowanceWage ?? 0m;
            est.Employees.Add(emp);
        }
    }

    /// <summary>النسب (كانت مكتوبة داخل كل شيت بالإكسل: O10 و O11 و O12).</summary>
    public class Rates
    {
        public decimal SickMaternity { get; set; } = 0.11m;    // المرض والأمومة
        public decimal EndOfService { get; set; } = 0.085m;    // تعويض نهاية الخدمة
        public decimal FamilyAllowance { get; set; } = 0.06m;  // التعويضات العائلية
    }

    /// <summary>
    /// سطر بسجلّ حركة سير العمل: تسجيل تلقائي كل مرة الأجور تنفتح بالإكسل أو تنطبع أو تصدّر PDF.
    /// Action نص ثابت (مو enum) حتى ما يرتبط هالملف بـ ExcelPrinter: "OpenInExcel"، "Print"، "ExportPdf".
    /// </summary>
    public class WorkLogEntry
    {
        public DateTime Timestamp { get; set; }
        public Guid EstablishmentId { get; set; }
        public int Year { get; set; }
        public int Quarter { get; set; }     // 1..4 (النموذج = 2 + الرقم، يعني 21..24)
        public string Action { get; set; } = "";
    }

    /// <summary>بيانات المؤسسة القديمة (النسخة الأولى). للتحويل بس.</summary>
    public class CompanyInfo
    {
        public string Name { get; set; } = "";
        public string NssfNumber { get; set; } = "";
    }

    /// <summary>مؤسسة: إلها موظفينها وفتراتها (21..24 لكل سنة).</summary>
    public class Establishment
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string Name { get; set; } = "";
        public string NssfNumber { get; set; } = "";
        public string OwnerName { get; set; } = "";
        public string OwnerPhone { get; set; } = "";
        public List<Employee> Employees { get; set; } = new List<Employee>();
        public List<PeriodData> Periods { get; set; } = new List<PeriodData>();

        public PeriodData FindPeriod(int year, int number)
        {
            return Periods.FirstOrDefault(p => p.Year == year && p.Number == number);
        }

        public PeriodData GetPeriod(int year, int number)
        {
            PeriodData p = FindPeriod(year, number);
            if (p == null)
            {
                p = new PeriodData { Year = year, Number = number };
                Periods.Add(p);
            }
            return p;
        }

        // بتظهر بالقوائم المنسدلة
        public override string ToString()
        {
            return string.IsNullOrWhiteSpace(Name) ? "(مؤسسة بدون اسم)" : Name;
        }
    }

    /// <summary>بطاقة موظف. كل الحقول اختيارية. الرواتب هون شهرية (بتنعبّى تلقائياً بالجدول لكل شهر فعّال).</summary>
    public class Employee
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public string FullName { get; set; } = "";              // الاسم الثلاثي
        public string NssfNumber { get; set; } = "";            // رقم الضمان
        public bool IsMarried { get; set; }                     // متأهل / أعزب
        public int? ChildrenCount { get; set; }                 // عدد الأولاد (بس إذا متأهل)
        public DateTime? StartDate { get; set; }                // تاريخ الاستخدام (للمعلومات بس، ما بيأثر عالحساب)
        public DateTime? EndDate { get; set; }                  // تاريخ الترك (اختياري): بيوقّف الحساب بعد شهر الترك
        public decimal SickMaternityWage { get; set; }          // راتب المرض والأمومة (شهري، الحالي)
        public decimal EndOfServiceWage { get; set; }           // راتب نهاية الخدمة (شهري، الحالي)
        public decimal FamilyAllowanceWage { get; set; }        // راتب التعويض العائلي (شهري، الحالي)

        // الراتب "القديم" قبل تاريخ "اعتباراً من" (لو في زيادة مثلاً). ما بينستعمل أبداً إذا ما في
        // تاريخ "اعتباراً من" محدّد لهالفرع — عندها الراتب الحالي بس هو يلي بينحسب، كل الأشهر.
        public decimal SickMaternityPreviousWage { get; set; }
        public decimal EndOfServicePreviousWage { get; set; }
        public decimal FamilyAllowancePreviousWage { get; set; }

        // "اعتباراً من" لكل راتب (اختياري): إذا محدّد، الراتب الحالي بيصير محسوب بس بالأشهر يلي بعد
        // (أو بنفس) هالتاريخ، والأشهر يلي قبلو بينحسبلها الراتب "القديم" (مش صفر). بلا تاريخ = الراتب
        // الحالي بس، كل الأشهر (نفس السلوك القديم). كل راتب وتاريخو، مستقلين عن بعض.
        public DateTime? SickMaternityEffectiveDate { get; set; }
        public DateTime? EndOfServiceEffectiveDate { get; set; }
        public DateTime? FamilyAllowanceEffectiveDate { get; set; }

        // بتظهر بقائمة الموظفين (ما بتنحفظ بالملف)
        public string DisplayName
        {
            get
            {
                string name = string.IsNullOrWhiteSpace(FullName) ? "(موظف بدون اسم)" : FullName;
                if (EndDate.HasValue)
                    name += "  — ترك " + EndDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
                return name;
            }
        }

        /// <summary>
        /// هل الموظف بينحسب بهالشهر؟ بعد تاريخ الترك لأ. شهر الترك نفسو بينحسب (الموظف اشتغل فيه)،
        /// والأشهر يلي بعدو لأ، وبالسنة الجاية كل الأشهر لأ.
        /// </summary>
        public bool IsActiveInMonth(int year, int month)
        {
            if (!EndDate.HasValue) return true;
            return EndDate.Value.Date >= new DateTime(year, month, 1);
        }

        /// <summary>هل الموظف بينحسب بشهر واحد على الأقل من الفصل (1..4)؟</summary>
        public bool IsActiveInPeriod(int year, int quarter)
        {
            return IsActiveInMonth(year, (quarter - 1) * 3 + 1);
        }
    }

    /// <summary>فترة وحدة = جدول واحد = نموذج واحد (21 أو 22 أو 23 أو 24 = الفصول الأربعة) لسنة معينة.</summary>
    public class PeriodData
    {
        public int Year { get; set; }
        public int Number { get; set; }          // 1..4  (النموذج 2 + الرقم)
        public DateTime? SigningDate { get; set; }
        public List<EmployeeLine> Lines { get; set; } = new List<EmployeeLine>();
    }

    /// <summary>أجور شهر واحد لموظف (الصفر = ما عندو هالفرع بهالشهر).</summary>
    public class MonthWages
    {
        public decimal SickMaternity { get; set; }
        public decimal EndOfService { get; set; }
        public decimal FamilyAllowance { get; set; }
        public decimal FamilyAllowancePaid { get; set; }   // التعويض العائلي المدفوع، لهالشهر بالذات
    }

    /// <summary>سطر موظف بجدول الفصل: أجور كل شهر من الأشهر التلاتة، بما فيهم التعويض العائلي المدفوع.</summary>
    public class EmployeeLine
    {
        public Guid? EmployeeId { get; set; }     // ربط ببطاقة الموظف (لتطبيق تاريخ الترك)
        public string Name { get; set; } = "";
        public List<MonthWages> Months { get; set; } = NewMonths();   // 3 أشهر: 0 و1 و2

        // ---- من نسخ قديمة (أجور الفصل كاملة، بما فيهم التعويض المدفوع): بتنقرا مرة وحدة
        //      وبتتوزّع على الأشهر بالتساوي، وبعدها بتصير null ----
        public decimal? SickMaternityWage { get; set; }
        public decimal? EndOfServiceWage { get; set; }
        public decimal? FamilyAllowanceWage { get; set; }
        public decimal? FamilyAllowancePaid { get; set; }

        private static List<MonthWages> NewMonths()
        {
            return new List<MonthWages> { new MonthWages(), new MonthWages(), new MonthWages() };
        }

        /// <summary>بيتأكد إنو في 3 أشهر بالظبط.</summary>
        public void EnsureMonths()
        {
            if (Months == null) Months = new List<MonthWages>();
            while (Months.Count < 3) Months.Add(new MonthWages());
            while (Months.Count > 3) Months.RemoveAt(Months.Count - 1);
            for (int i = 0; i < 3; i++)
                if (Months[i] == null) Months[i] = new MonthWages();
        }

        /// <summary>بيوزّع أجور الفصل القديمة (بما فيهم التعويض المدفوع) بالتساوي على 3 أشهر (المجموع بيضل نفسو). بيرجّع true إذا صار تحويل.</summary>
        public bool MigrateWagesToMonths()
        {
            if (!SickMaternityWage.HasValue && !EndOfServiceWage.HasValue && !FamilyAllowanceWage.HasValue && !FamilyAllowancePaid.HasValue)
                return false;

            EnsureMonths();
            decimal[] sm = WageRules.SplitEvenly(SickMaternityWage ?? 0m, 3);
            decimal[] eos = WageRules.SplitEvenly(EndOfServiceWage ?? 0m, 3);
            decimal[] fa = WageRules.SplitEvenly(FamilyAllowanceWage ?? 0m, 3);
            decimal[] paid = WageRules.SplitEvenly(FamilyAllowancePaid ?? 0m, 3);
            for (int i = 0; i < 3; i++)
            {
                Months[i].SickMaternity = sm[i];
                Months[i].EndOfService = eos[i];
                Months[i].FamilyAllowance = fa[i];
                Months[i].FamilyAllowancePaid = paid[i];
            }

            SickMaternityWage = null;
            EndOfServiceWage = null;
            FamilyAllowanceWage = null;
            FamilyAllowancePaid = null;
            return true;
        }
    }
}
