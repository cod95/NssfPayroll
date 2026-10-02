using System;
using System.Collections.Generic;
using System.Linq;

namespace NssfPayroll
{
    public enum WageKind
    {
        SickMaternity = 0,        // مرض وأمومة
        EndOfService = 1,         // نهاية خدمة
        FamilyAllowance = 2,      // تعويضات عائلية
        FamilyAllowancePaid = 3   // التعويض العائلي المدفوع (مو فرع اشتراك، بس بيستعمل نفس آلية الشهر/الفصل)
    }

    /// <summary>سطر من جدول النموذج: الأجور (Wages، الحقيقية بلا سقف) × النسبة، بعد تطبيق السقف
    /// الشهري (إذا الفرع عندو سقف) = الاشتراك المستحق. Contribution بينحسب ويتحدّد مسبقاً بـ
    /// Calculator.Calculate (بيطبّق السقف شهر شهر قبل الجمع)، مش P × O بسيطة متل قبل.</summary>
    public class BranchResult
    {
        public string Name { get; set; }
        public decimal Wages { get; set; }
        public decimal Rate { get; set; }
        public int Employees { get; set; }
        public decimal Contribution { get; set; }
    }

    /// <summary>
    /// حساب اشتراكات الضمان الاجتماعي اللبناني — دالة نقية (pure function) مطابقة بالضبط لمنهجية
    /// الحساب الرسمية: سقف شهري 28,000,000 للمرض والأمومة وللتعويضات العائلية (بدون سقف لنهاية
    /// الخدمة)، والنسب القابلة للتعديل من تبويب «المعلومات الأساسية» (Rates). قابلة للاختبار لحالها
    /// بمعزل عن باقي البرنامج.
    /// </summary>
    public static class NssfContributionCalculator
    {
        /// <summary>سقف المرض والأمومة شهرياً (سقف مختلف عن التعويضات العائلية — قانونياً)</summary>
        public const decimal HealthMaxCeiling = 120000000m;
        /// <summary>سقف التعويضات العائلية شهرياً</summary>
        public const decimal FamilyMaxCeiling = 28000000m;

        public class ContributionResult
        {
            public decimal FamilyAllowanceDue { get; set; }         // التعويضات العائلية المستحقة (يمرَّر كما هو، معلوماتي)
            public decimal TotalDueContributions { get; set; }      // الاشتراكات المستحقة الإجمالية
            public decimal FamilyAllowancePaid { get; set; }        // التعويض العائلي المدفوع (يمرَّر كما هو)
            public decimal NetPayableContributions { get; set; }    // صافي الاشتراكات الواجب سدادها
        }

        /// <summary>
        /// بيحسب اشتراكات موظف واحد (لشهر أو فصل، حسب الأجور المعطاة). healthWage/familyWage/
        /// endOfServiceWage هني الأجور الخاضعة للاشتراك لكل فرع على حدة (ممكن تكون نفس الرقم لو
        /// الراتب موحّد بين الفروع، أو أرقام مختلفة حسب بطاقة الموظف — البرنامج بيسجّل كل فرع
        /// لحالو). familyAllowanceDue بيتمرّر كما هو بالنتيجة، ما بيدخل بحساب الاشتراكات.
        /// </summary>
        public static ContributionResult ComputeContributions(
            decimal healthWage, decimal familyWage, decimal endOfServiceWage,
            decimal familyAllowanceDue, decimal familyAllowancePaid,
            decimal healthRate, decimal familyRate, decimal endOfServiceRate)
        {
            decimal healthContribution = Math.Min(healthWage, HealthMaxCeiling) * healthRate;
            decimal familyContribution = Math.Min(familyWage, FamilyMaxCeiling) * familyRate;
            decimal endOfServiceContribution = endOfServiceWage * endOfServiceRate;   // بدون سقف

            decimal totalDue = healthContribution + familyContribution + endOfServiceContribution;
            decimal net = totalDue - familyAllowancePaid;

            return new ContributionResult
            {
                FamilyAllowanceDue = familyAllowanceDue,
                TotalDueContributions = totalDue,
                FamilyAllowancePaid = familyAllowancePaid,
                NetPayableContributions = net
            };
        }
    }

    public class PeriodResult
    {
        public int Year { get; set; }
        public int Number { get; set; }

        /// <summary>رمز النموذج: 21 أو 22 أو 23 أو 24 (وهو اسم الشيت بالقالب).</summary>
        public string FormCode { get { return "2" + Number; } }

        public BranchResult SickMaternity { get; set; }      // الصف 10
        public BranchResult EndOfService { get; set; }       // الصف 11
        public BranchResult FamilyAllowance { get; set; }    // الصف 12

        public decimal FamilyAllowancePaid { get; set; }     // N14

        /// <summary>N13 = SUM(N10:N12)</summary>
        public decimal TotalDue
        {
            get { return SickMaternity.Contribution + EndOfService.Contribution + FamilyAllowance.Contribution; }
        }

        /// <summary>N15 = N13 - N14</summary>
        public decimal Net { get { return TotalDue - FamilyAllowancePaid; } }
    }

    /// <summary>سطر بتقرير أجور موظف: فصل كامل (Month = null) أو شهر لحالو.</summary>
    public class EmployeeWageReportRow
    {
        public int Year { get; set; }
        public string Period { get; set; }       // "21".."24"
        public string Month { get; set; }        // اسم الشهر بالعربي، أو null لسطر الفصل الكامل
        public decimal SickMaternity { get; set; }
        public decimal EndOfService { get; set; }
        public decimal FamilyAllowance { get; set; }
        public decimal FamilyAllowancePaid { get; set; }
    }

    /// <summary>
    /// سطر موظف بـ"الفترة 70" (تجميعية سنوية: 21+22+23+24 مع بعض بسطر واحد لكل موظف).
    /// </summary>
    public class AnnualSummaryRow
    {
        public string EmployeeName { get; set; }
        public decimal EndOfService { get; set; }         // نهاية الخدمة = مجموع 21+22+23+24
        public decimal PaidContribution { get; set; }      // الاشتراكات المدفوعة = نهاية الخدمة × نسبة نهاية الخدمة
        public decimal FamilyAllowance { get; set; }        // التعويضات العائلية = مجموع 21+22+23+24
        public decimal SickMaternity { get; set; }           // المرض والأمومة = مجموع 21+22+23+24

        /// <summary>المتوجب = الاشتراكات المدفوعة − نهاية الخدمة.</summary>
        public decimal Due { get { return PaidContribution - EndOfService; } }
    }

    /// <summary>
    /// قواعد الأجور: أجور كل شهر، تاريخ الترك، وتوزيع أجور الفصل على الأشهر.
    /// (بيستعملها الحساب والجدول، فالرقم المعروض هو نفسو المطبوع.)
    /// </summary>
    public static class WageRules
    {
        /// <summary>أسماء الأشهر بالعربي (0 = كانون الثاني ... 11 = كانون الأول). بلا اعتماد على WinForms
        /// حتى يضل هالملف قابل للاختبار بمشروع .NET عادي.</summary>
        public static readonly string[] MonthNames =
        {
            "كانون الثاني", "شباط", "آذار", "نيسان", "أيار", "حزيران",
            "تموز", "آب", "أيلول", "تشرين الأول", "تشرين الثاني", "كانون الأول"
        };

        // مبالغ التعويضات العائلية الشهرية الرسمية (بيستعملها الاحتساب التلقائي ببطاقة الموظف)
        public const decimal FamilyAllowanceWifeAmount = 2100000m;    // عن الزوجة (لو متأهل)
        public const decimal FamilyAllowanceChildAmount = 1155000m;   // عن كل ولد

        /// <summary>
        /// التعويض العائلي المحتسب تلقائياً (شهري) حسب الحالة الاجتماعية وعدد الأولاد: 2,100,000 عن
        /// الزوجة لو متأهل + 1,155,000 عن كل ولد مسجّل. بيستعملو EmployeesForm كل مرة تتغيّر الحالة
        /// الاجتماعية أو عدد الأولاد؛ المستخدم لسا فيه يعدّل الناتج يدوياً بعدها لو بدو.
        /// </summary>
        public static decimal ComputeFamilyAllowance(bool isMarried, int? childrenCount)
        {
            decimal total = isMarried ? FamilyAllowanceWifeAmount : 0m;
            total += (childrenCount ?? 0) * FamilyAllowanceChildAmount;
            return total;
        }

        public static decimal Get(MonthWages m, WageKind k)
        {
            switch (k)
            {
                case WageKind.SickMaternity: return m.SickMaternity;
                case WageKind.EndOfService: return m.EndOfService;
                case WageKind.FamilyAllowance: return m.FamilyAllowance;
                default: return m.FamilyAllowancePaid;
            }
        }

        public static void Set(MonthWages m, WageKind k, decimal v)
        {
            switch (k)
            {
                case WageKind.SickMaternity: m.SickMaternity = v; break;
                case WageKind.EndOfService: m.EndOfService = v; break;
                case WageKind.FamilyAllowance: m.FamilyAllowance = v; break;
                default: m.FamilyAllowancePaid = v; break;
            }
        }

        /// <summary>بطاقة الموظف يلي إلو هالسطر: بالرقم التعريفي، وإذا ما في بالاسم.</summary>
        public static Employee FindEmployee(Establishment est, EmployeeLine line)
        {
            if (est == null || line == null || est.Employees == null) return null;

            if (line.EmployeeId.HasValue)
            {
                Employee byId = est.Employees.FirstOrDefault(x => x.Id == line.EmployeeId.Value);
                if (byId != null) return byId;
            }

            string name = (line.Name ?? "").Trim();
            if (name.Length == 0) return null;
            return est.Employees.FirstOrDefault(x =>
                string.Equals((x.FullName ?? "").Trim(), name, StringComparison.CurrentCultureIgnoreCase));
        }

        /// <summary>عكس FindEmployee: سطر الموظف هيدا بفترة معيّنة (بالرقم التعريفي، وإذا ما في بالاسم). لتقرير أجور موظف.</summary>
        public static EmployeeLine FindLineForEmployee(PeriodData period, Employee emp)
        {
            if (period == null || emp == null || period.Lines == null) return null;

            EmployeeLine byId = period.Lines.FirstOrDefault(l => l.EmployeeId.HasValue && l.EmployeeId.Value == emp.Id);
            if (byId != null) return byId;

            string name = (emp.FullName ?? "").Trim();
            if (name.Length == 0) return null;
            return period.Lines.FirstOrDefault(l =>
                string.Equals((l.Name ?? "").Trim(), name, StringComparison.CurrentCultureIgnoreCase));
        }

        private static bool LineMatchesEmployee(EmployeeLine l, Employee emp, string name)
        {
            return (l.EmployeeId.HasValue && l.EmployeeId.Value == emp.Id)
                || (name.Length > 0 && string.Equals((l.Name ?? "").Trim(), name, StringComparison.CurrentCultureIgnoreCase));
        }

        /// <summary>كم سطر أجور مسجّل لهالموظف عبر كل فترات المؤسسة (بلا ما يشيل شي). لرسالة التأكيد قبل الحذف.</summary>
        public static int CountEmployeeLines(Establishment est, Employee emp)
        {
            if (est == null || emp == null) return 0;
            string name = (emp.FullName ?? "").Trim();
            int count = 0;
            foreach (PeriodData p in est.Periods)
                count += p.Lines.Count(l => LineMatchesEmployee(l, emp, name));
            return count;
        }

        /// <summary>بيشيل كل أسطر أجور الموظف هيدا من كل فترات المؤسسة. بيرجّع عدد الأسطر يلي انشالت.</summary>
        public static int RemoveEmployeeLines(Establishment est, Employee emp)
        {
            if (est == null || emp == null) return 0;
            string name = (emp.FullName ?? "").Trim();
            int removed = 0;
            foreach (PeriodData p in est.Periods)
                removed += p.Lines.RemoveAll(l => LineMatchesEmployee(l, emp, name));
            return removed;
        }

        /// <summary>هل الشهر (0..2 من الفصل 1..4) بينحسب للموظف؟ بلا بطاقة = دايماً.</summary>
        public static bool IsMonthActive(Employee emp, int year, int quarter, int monthIndex)
        {
            return emp == null || emp.IsActiveInMonth(year, (quarter - 1) * 3 + monthIndex + 1);
        }

        /// <summary>تاريخ "اعتباراً من" المسجّل ببطاقة الموظف لفرع معيّن (بلا بطاقة = ما في تاريخ).</summary>
        private static DateTime? EffectiveDate(Employee card, WageKind k)
        {
            if (card == null) return null;
            switch (k)
            {
                case WageKind.SickMaternity: return card.SickMaternityEffectiveDate;
                case WageKind.EndOfService: return card.EndOfServiceEffectiveDate;
                case WageKind.FamilyAllowance: return card.FamilyAllowanceEffectiveDate;
                default: return null;   // FamilyAllowancePaid: ما في تاريخ "اعتباراً من" لهالحقل (مو مأخوذ من البطاقة أصلاً)
            }
        }

        private static decimal CardWage(Employee card, WageKind k)
        {
            switch (k)
            {
                case WageKind.SickMaternity: return card.SickMaternityWage;
                case WageKind.EndOfService: return card.EndOfServiceWage;
                case WageKind.FamilyAllowance: return card.FamilyAllowanceWage;
                default: return 0m;
            }
        }

        private static decimal PreviousCardWage(Employee card, WageKind k)
        {
            switch (k)
            {
                case WageKind.SickMaternity: return card.SickMaternityPreviousWage;
                case WageKind.EndOfService: return card.EndOfServicePreviousWage;
                case WageKind.FamilyAllowance: return card.FamilyAllowancePreviousWage;
                default: return 0m;   // FamilyAllowancePaid: ما في "راتب قديم" لهالحقل، مو مأخوذ من البطاقة أصلاً
            }
        }

        /// <summary>
        /// راتب البطاقة الصحيح لهالشهر: صفر إذا ترك (بيغلب كل شي)، وإلا الراتب "القديم" إذا الشهر قبل
        /// تاريخ "اعتباراً من"، وإلا الراتب الحالي (بعد التاريخ، أو بلا تاريخ أصلاً). هيدا يلي بيخلّي
        /// زيادة الراتب تنعكس صح: الأشهر يلي قبل الزيادة بتاخد الراتب القديم مش صفر.
        /// </summary>
        public static decimal ResolveCardWage(Employee card, int year, int quarter, int monthIndex, WageKind k)
        {
            if (card == null) return 0m;
            if (!IsMonthActive(card, year, quarter, monthIndex)) return 0m;

            int month = (quarter - 1) * 3 + monthIndex + 1;
            bool started = WageStartedByMonth(EffectiveDate(card, k), year, month);
            return started ? CardWage(card, k) : PreviousCardWage(card, k);
        }

        /// <summary>
        /// هل راتب الفرع هيدا صار "شغّال" بهالشهر (year/month)؟ بلا تاريخ = دايماً شغّال (نفس السلوك
        /// القديم). الشهر يلي فيه تاريخ "اعتباراً من" (أو بعدو) بيصير الراتب محسوب فيه كامل، تماماً
        /// متل ما شهر الترك نفسو بينحسب كامل (بلا تجزئة الشهر).
        /// </summary>
        private static bool WageStartedByMonth(DateTime? effectiveDate, int year, int month)
        {
            if (!effectiveDate.HasValue) return true;
            DateTime firstOfNextMonth = new DateTime(year, month, 1).AddMonths(1);
            return effectiveDate.Value.Date < firstOfNextMonth;
        }

        /// <summary>هل راتب الفرع هيدا شغّال بهالشهر (0..2 من الفصل) — الترك + تاريخ "اعتباراً من" مع بعض.</summary>
        public static bool IsWageActive(Employee card, int year, int quarter, int monthIndex, WageKind k)
        {
            if (!IsMonthActive(card, year, quarter, monthIndex)) return false;
            int month = (quarter - 1) * 3 + monthIndex + 1;
            return WageStartedByMonth(EffectiveDate(card, k), year, month);
        }

        public static List<int> ActiveMonths(Employee emp, int year, int quarter)
        {
            List<int> r = new List<int>();
            for (int i = 0; i < 3; i++)
                if (IsMonthActive(emp, year, quarter, i)) r.Add(i);
            return r;
        }

        /// <summary>مجموع أجور الفصل: الأشهر الفعّالة بس (يلي بعد الترك ما بتنحسب مهما كان مكتوب فيها).</summary>
        public static decimal QuarterTotal(EmployeeLine line, Employee emp, int year, int quarter, WageKind k)
        {
            line.EnsureMonths();
            decimal sum = 0m;
            for (int i = 0; i < 3; i++)
                if (IsMonthActive(emp, year, quarter, i)) sum += Get(line.Months[i], k);
            return sum;
        }

        /// <summary>متل QuarterTotal، بس كل شهر بينسقّف (Math.Min) لحالو قبل ما ينضاف للمجموع —
        /// حتى ما يصير فرق لو الراتب زاد بنص الفصل (تسقيف شهري صحيح، مش تسقيف المجموع كاملاً).
        /// ceiling = decimal.MaxValue يعني بلا سقف أصلاً (نفس نتيجة QuarterTotal بالضبط).</summary>
        public static decimal CappedQuarterTotal(EmployeeLine line, Employee emp, int year, int quarter, WageKind k, decimal ceiling)
        {
            line.EnsureMonths();
            decimal sum = 0m;
            for (int i = 0; i < 3; i++)
                if (IsMonthActive(emp, year, quarter, i)) sum += Math.Min(Get(line.Months[i], k), ceiling);
            return sum;
        }

        /// <summary>بيقسم المبلغ على n بالتساوي (الأولى بالأرقام الصحيحة والباقي بالأخيرة، فالمجموع ما بيتغير).</summary>
        public static decimal[] SplitEvenly(decimal total, int n)
        {
            decimal[] r = new decimal[n];
            if (n <= 0) return r;

            decimal share = Math.Floor(total / n);
            decimal used = 0m;
            for (int i = 0; i < n - 1; i++)
            {
                r[i] = share;
                used += share;
            }
            r[n - 1] = total - used;
            return r;
        }

        /// <summary>
        /// هل أجور الأشهر الفعّالة موزّعة بالتساوي؟ إذا إي، بيجوز تعدّل إجمالي الفصل بالجدول.
        /// إذا الأشهر مختلفة (مثلاً زيادة براتب بشهر معيّن) لازم تعدّل من عرض الشهر.
        /// </summary>
        public static bool IsEvenSplit(EmployeeLine line, Employee emp, int year, int quarter, WageKind k)
        {
            line.EnsureMonths();
            List<int> act = ActiveMonths(emp, year, quarter);
            if (act.Count <= 1) return true;

            decimal sum = 0m;
            foreach (int i in act) sum += Get(line.Months[i], k);

            decimal[] expected = SplitEvenly(sum, act.Count);
            for (int j = 0; j < act.Count; j++)
                if (Get(line.Months[act[j]], k) != expected[j]) return false;
            return true;
        }

        /// <summary>بيحط إجمالي الفصل موزّع بالتساوي على الأشهر الفعّالة (والأشهر بعد الترك بتصير صفر).</summary>
        public static void SetQuarterTotal(EmployeeLine line, Employee emp, int year, int quarter, WageKind k, decimal total)
        {
            line.EnsureMonths();
            List<int> act = ActiveMonths(emp, year, quarter);
            if (act.Count == 0) return;

            decimal[] parts = SplitEvenly(total, act.Count);
            for (int i = 0; i < 3; i++) Set(line.Months[i], k, 0m);
            for (int j = 0; j < act.Count; j++) Set(line.Months[act[j]], k, parts[j]);
        }

        /// <summary>كل الأجور (كل الأشهر وكل الفروع) صفر؟</summary>
        public static bool IsEmpty(EmployeeLine line)
        {
            line.EnsureMonths();
            foreach (MonthWages m in line.Months)
                if (m.SickMaternity != 0m || m.EndOfService != 0m || m.FamilyAllowance != 0m) return false;
            return true;
        }

        /// <summary>
        /// بيعبّي السطر براتب البطاقة الشهري لكل شهر فعّال من الفصل. كل فرع (مرض وأمومة/نهاية خدمة/
        /// تعويضات عائلية) بينحسب لحالو حسب تاريخ "اعتباراً من" المسجّل إلو بالبطاقة (وصفر قبلو، أو
        /// بعد شهر الترك، أو بلا بطاقة أصلاً).
        /// </summary>
        public static void FillFromCard(EmployeeLine line, Employee card, int year, int quarter)
        {
            line.EnsureMonths();
            line.EmployeeId = card.Id;
            line.Name = card.FullName;
            for (int i = 0; i < 3; i++)
            {
                line.Months[i].SickMaternity = ResolveCardWage(card, year, quarter, i, WageKind.SickMaternity);
                line.Months[i].EndOfService = ResolveCardWage(card, year, quarter, i, WageKind.EndOfService);
                line.Months[i].FamilyAllowance = ResolveCardWage(card, year, quarter, i, WageKind.FamilyAllowance);
            }
        }

        /// <summary>
        /// هل الموظف كان مستخدم ولو يوم واحد بالشهر المعطى (year/month)، حسب تاريخ استخدامو؟ قاعدة
        /// "كسر الشهر" الرسمية: أي جزء من الشهر (حتى يوم واحد) بيتحسب شهر كامل بالراتب الكامل، بلا أي
        /// تجزئة أو تقسيط يومي — نفس مبدأ "شهر الترك بيتحسب كامل" (IsActiveInMonth) بس بالاتجاه
        /// المعاكس (شهر الاستخدام). بلا تاريخ استخدام مسجّل = بلا قيد (نفس سلوك التواريخ الاختيارية بهالبرنامج).
        /// </summary>
        private static bool WasEmployedDuringMonth(DateTime? startDate, int year, int month)
        {
            if (!startDate.HasValue) return true;
            DateTime hireMonth = new DateTime(startDate.Value.Year, startDate.Value.Month, 1);
            DateTime targetMonth = new DateTime(year, month, 1);
            return hireMonth <= targetMonth;
        }

        /// <summary>
        /// متل FillFromCard، بس كمان بتراعي إنو ما بيتحسب شي قبل شهر الاستخدام (وشهر الاستخدام نفسو
        /// بينحسب شهر كامل بالراتب الكامل — قاعدة "كسر الشهر"، يستعملها زر "إدراج موظفي المؤسسة لهذه
        /// الفترة"). كل شهر من الفصل بينحسب لحالو. بترجّع false (وما بتعبّي شي) إذا الموظف مو مؤهّل
        /// ولا شهر واحد من الفصل (مثلاً استلم بعد الفصل)، حتى ما تنضاف لو سطر فاضي بلا فايدة.
        /// </summary>
        public static bool FillFromCardForBulkInsert(EmployeeLine line, Employee card, int year, int quarter)
        {
            line.EnsureMonths();
            bool anyEligible = false;
            for (int i = 0; i < 3; i++)
            {
                int month = (quarter - 1) * 3 + i + 1;
                bool hireOk = WasEmployedDuringMonth(card.StartDate, year, month);

                line.Months[i].SickMaternity = hireOk ? ResolveCardWage(card, year, quarter, i, WageKind.SickMaternity) : 0m;
                line.Months[i].EndOfService = hireOk ? ResolveCardWage(card, year, quarter, i, WageKind.EndOfService) : 0m;
                line.Months[i].FamilyAllowance = hireOk ? ResolveCardWage(card, year, quarter, i, WageKind.FamilyAllowance) : 0m;

                if (hireOk && IsMonthActive(card, year, quarter, i)) anyEligible = true;
            }

            if (!anyEligible) return false;

            line.EmployeeId = card.Id;
            line.Name = card.FullName;
            return true;
        }

        /// <summary>
        /// كل أجور موظف معيّن عبر كل فترات المؤسسة يلي إلو سطر فيها — فصل كامل أو شهر شهر حسب
        /// monthly. بيستعملها معاينة تبويب "المتابعة والتقارير" وتقرير الطباعة، حتى الرقمين يطلعوا نفس الشي دايماً.
        /// </summary>
        public static List<EmployeeWageReportRow> BuildWageReport(Establishment est, Employee emp, bool monthly)
        {
            List<EmployeeWageReportRow> rows = new List<EmployeeWageReportRow>();
            if (est == null || emp == null) return rows;

            foreach (PeriodData p in est.Periods.OrderBy(x => x.Year).ThenBy(x => x.Number))
            {
                EmployeeLine line = FindLineForEmployee(p, emp);
                if (line == null) continue;
                line.EnsureMonths();

                if (!monthly)
                {
                    rows.Add(new EmployeeWageReportRow
                    {
                        Year = p.Year,
                        Period = "2" + p.Number,
                        SickMaternity = QuarterTotal(line, emp, p.Year, p.Number, WageKind.SickMaternity),
                        EndOfService = QuarterTotal(line, emp, p.Year, p.Number, WageKind.EndOfService),
                        FamilyAllowance = QuarterTotal(line, emp, p.Year, p.Number, WageKind.FamilyAllowance),
                        FamilyAllowancePaid = QuarterTotal(line, emp, p.Year, p.Number, WageKind.FamilyAllowancePaid)
                    });
                }
                else
                {
                    for (int i = 0; i < 3; i++)
                    {
                        // نفس منطق LineRow.GetWage بالضبط: بعد الترك، الرقم المعروض صفر حتى لو
                        // ضل شي مخزّن قديم بالشهر (مثلاً تاريخ الترك انضاف بعد ما انكتبت الأجور).
                        bool active = IsMonthActive(emp, p.Year, p.Number, i);
                        MonthWages m = line.Months[i];
                        int monthNum = (p.Number - 1) * 3 + i + 1;
                        rows.Add(new EmployeeWageReportRow
                        {
                            Year = p.Year,
                            Period = "2" + p.Number,
                            Month = MonthNames[monthNum - 1],
                            SickMaternity = active ? m.SickMaternity : 0m,
                            EndOfService = active ? m.EndOfService : 0m,
                            FamilyAllowance = active ? m.FamilyAllowance : 0m,
                            FamilyAllowancePaid = active ? m.FamilyAllowancePaid : 0m
                        });
                    }
                }
            }

            return rows;
        }

        /// <summary>
        /// "الفترة 70": سطر واحد لكل موظف بمجموع أرباعه (21+22+23+24) لسنة معيّنة — مش فترة حقيقية
        /// مخزّنة، محسوبة كل مرة من فترات 21-24 الفعلية (بما فيها أي تعديل يدوي عليهن). موظف بلا ولا
        /// سطر بكل الأرباع الأربعة ما بيظهر (ما في شي يتجمّع).
        /// </summary>
        public static List<AnnualSummaryRow> BuildAnnualSummary(Establishment est, int year, decimal endOfServiceRate)
        {
            List<AnnualSummaryRow> rows = new List<AnnualSummaryRow>();
            if (est == null) return rows;

            foreach (Employee emp in est.Employees)
            {
                decimal eos = 0m, fa = 0m, sm = 0m;
                bool hasAny = false;

                for (int q = 1; q <= 4; q++)
                {
                    PeriodData p = est.FindPeriod(year, q);   // FindPeriod (مش GetPeriod) حتى ما ننشئ فترة فاضية بمجرد ما نعاين الفترة 70
                    if (p == null) continue;

                    EmployeeLine line = FindLineForEmployee(p, emp);
                    if (line == null) continue;

                    hasAny = true;
                    eos += QuarterTotal(line, emp, year, q, WageKind.EndOfService);
                    fa += QuarterTotal(line, emp, year, q, WageKind.FamilyAllowance);
                    sm += QuarterTotal(line, emp, year, q, WageKind.SickMaternity);
                }

                if (!hasAny) continue;

                rows.Add(new AnnualSummaryRow
                {
                    EmployeeName = string.IsNullOrWhiteSpace(emp.FullName) ? "(بدون اسم)" : emp.FullName,
                    EndOfService = eos,
                    PaidContribution = eos * endOfServiceRate,
                    FamilyAllowance = fa,
                    SickMaternity = sm
                });
            }

            return rows;
        }
    }

    public static class Calculator
    {
        /// <summary>
        /// بيحسب نموذج الفصل. الأجر الصفر بفرع = الموظف مو محسوب بهالفرع (ما بيدخل بالمجموع وبيتنقص من عدد الأجراء).
        /// est = المؤسسة (لتطبيق تاريخ الترك من بطاقات الموظفين).
        /// </summary>
        public static PeriodResult Calculate(PeriodData period, Rates rates, Establishment est)
        {
            decimal sm = 0m, eos = 0m, fa = 0m, paid = 0m;
            decimal smContribution = 0m, eosContribution = 0m, faContribution = 0m;
            int nSm = 0, nEos = 0, nFa = 0;

            foreach (EmployeeLine line in period.Lines)
            {
                Employee emp = WageRules.FindEmployee(est, line);

                decimal lsm = WageRules.QuarterTotal(line, emp, period.Year, period.Number, WageKind.SickMaternity);
                decimal leos = WageRules.QuarterTotal(line, emp, period.Year, period.Number, WageKind.EndOfService);
                decimal lfa = WageRules.QuarterTotal(line, emp, period.Year, period.Number, WageKind.FamilyAllowance);

                sm += lsm;
                eos += leos;
                fa += lfa;

                // الاشتراك المستحق: السقف الشهري (28 مليون) بينطبّق شهر شهر قبل الجمع، مش على
                // مجموع الفصل كامل — حتى ما يفرق شكل توزيع الراتب عبر الأشهر بالنتيجة.
                smContribution += WageRules.CappedQuarterTotal(line, emp, period.Year, period.Number, WageKind.SickMaternity, NssfContributionCalculator.HealthMaxCeiling) * rates.SickMaternity;
                faContribution += WageRules.CappedQuarterTotal(line, emp, period.Year, period.Number, WageKind.FamilyAllowance, NssfContributionCalculator.FamilyMaxCeiling) * rates.FamilyAllowance;
                eosContribution += leos * rates.EndOfService;   // نهاية الخدمة بدون سقف

                if (lsm > 0m) nSm++;
                if (leos > 0m) nEos++;
                if (lfa > 0m) nFa++;

                paid += WageRules.QuarterTotal(line, emp, period.Year, period.Number, WageKind.FamilyAllowancePaid);
            }

            return new PeriodResult
            {
                Year = period.Year,
                Number = period.Number,
                SickMaternity = new BranchResult { Name = "المرض والأمومة", Wages = sm, Rate = rates.SickMaternity, Employees = nSm, Contribution = smContribution },
                EndOfService = new BranchResult { Name = "تعويض نهاية الخدمة", Wages = eos, Rate = rates.EndOfService, Employees = nEos, Contribution = eosContribution },
                FamilyAllowance = new BranchResult { Name = "التعويضات العائلية", Wages = fa, Rate = rates.FamilyAllowance, Employees = nFa, Contribution = faContribution },
                FamilyAllowancePaid = paid
            };
        }
    }
}
