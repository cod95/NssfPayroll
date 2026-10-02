using System;
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace NssfPayroll
{
    public enum ExcelAction
    {
        OpenInExcel,   // بيفتح النموذج معبّى بـ Excel (منه بتعاين وبتطبع متل ما بدك)
        Print,         // طباعة مباشرة على الطابعة الافتراضية
        ExportPdf      // حفظ PDF
    }

    /// <summary>
    /// بياخد قالب الإكسل (ملفك الأصلي) وبينسخو لملف مؤقت، وبيعبّي الخلايا المتغيرة بس
    /// بشيت النموذج المطلوب (21..24)، وبعدين بيفتحو/بيطبعو/بيصدّرو PDF.
    /// القالب الأصلي ما بيتعدّل أبداً. بنستعمل Excel بالـ late binding (dynamic)
    /// فما منحتاج ولا مكتبة خارجية.
    /// </summary>
    public static class ExcelPrinter
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        private static string TemplatePath
        {
            get { return Path.Combine(AppContext.BaseDirectory, "Template", "NssfTemplate.xlsx"); }
        }

        public static void Run(Establishment establishment, PeriodResult result, DateTime signingDate,
                               ExcelAction action, string pdfPath)
        {
            if (!File.Exists(TemplatePath))
                throw new FileNotFoundException("ما لقيت ملف القالب: " + TemplatePath);

            Type excelType = ComHelper.GetExcelType();

            string temp = Path.Combine(Path.GetTempPath(),
                "nssf_" + result.FormCode + "_" + result.Year + "_" + Guid.NewGuid().ToString("N") + ".xlsx");
            File.Copy(TemplatePath, temp, true);
            File.SetAttributes(temp, FileAttributes.Normal);

            bool leaveOpen = false;
            try
            {
                leaveOpen = RunCore(excelType, temp, establishment, result, signingDate, action, pdfPath);
            }
            finally
            {
                // بعد ما يخلص RunCore كل متغيراتو بتصير برّا النطاق. منجبر الـ GC يحرّر أي كائن COM
                // ضايل (Range, Sheets...) لحتى ما يضل EXCEL.EXE مخفي شغّال بالخلفية.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();

                if (!leaveOpen)
                {
                    try { File.Delete(temp); } catch (Exception) { }   // إذا ما انمسح، بينمسح بالتشغيل الجاي
                }
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static bool RunCore(Type excelType, string temp, Establishment establishment, PeriodResult result,
                                    DateTime signingDate, ExcelAction action, string pdfPath)
        {
            string sheetName = result.FormCode;   // "21" .. "24"

            dynamic app = null;
            dynamic wb = null;
            dynamic ws = null;
            bool leaveOpen = false;

            try
            {
                app = Activator.CreateInstance(excelType);
                // أول اتصال بـ Excel لسا عم يبلّش (خصوصاً أول مرة عالكمبيوتر): بيرفض الأوامر لثانية أو
                // تنتين بخطأ "Call was rejected by callee". منعيد المحاولة تلقائياً بدل ما نفشل فوراً.
                ComHelper.Retry(() => { app.DisplayAlerts = false; });
                // Excel ظاهر لفتح النموذج وللطباعة المباشرة كمان (حتى يشوف المستخدم الجدول قبل ما
                // يطبع، مش يصير الطبع بالخفاء بلا ما يبين شي)؛ مخفي بس لتصدير PDF.
                ComHelper.Retry(() => { app.Visible = (action == ExcelAction.OpenInExcel || action == ExcelAction.Print); });

                wb = ComHelper.Retry<object>(() => app.Workbooks.Open(temp));
                ws = ComHelper.Retry<object>(() => wb.Worksheets[sheetName]);

                ComHelper.Retry(() => { Fill(ws, establishment, result, signingDate); });
                ComHelper.Retry(() => { KeepOnlySheet(wb, sheetName); });
                ComHelper.Retry(() => { ws.Activate(); });

                switch (action)
                {
                    case ExcelAction.Print:
                        ComHelper.Retry(() => { ws.PrintOut(); });
                        Thread.Sleep(1500);      // ندّي وقت للطابعة/الـ spooler قبل ما نطلّع الشباك
                        leaveOpen = ShowAndKeepOpen(app);   // بيضل الجدول ظاهر بعد الطباعة، يشوفو ويعيد الطباعة إذا لزم
                        break;

                    case ExcelAction.ExportPdf:
                        ComHelper.Retry(() => { ws.ExportAsFixedFormat(0, pdfPath); });   // 0 = xlTypePDF
                        break;

                    case ExcelAction.OpenInExcel:
                        leaveOpen = ShowAndKeepOpen(app);
                        break;
                }

                return leaveOpen;
            }
            finally
            {
                if (!leaveOpen)
                {
                    try { if ((object)wb != null) ComHelper.Retry(() => { wb.Close(false); }); } catch (Exception) { }
                    try { if ((object)app != null) ComHelper.Retry(() => { app.Quit(); }); } catch (Exception) { }
                }

                ComHelper.Release((object)ws);
                ComHelper.Release((object)wb);
                ComHelper.Release((object)app);
            }
        }

        // بيخلّي Excel ظاهر ومفتوح بعد ما يسكر البرنامج، وبيطلّعو لقدّام المستخدم. مستعمل لـ"فتح
        // بالإكسل" ولـ"طباعة مباشرة" مع بعض، حتى يشوف المستخدم الجدول ويقدر يعيد الطباعة إذا لزم.
        private static bool ShowAndKeepOpen(dynamic app)
        {
            ComHelper.Retry(() => { app.DisplayAlerts = true; });
            ComHelper.Retry(() => { app.UserControl = true; });        // بيضل Excel مفتوح بعد ما يسكر البرنامج
            try { SetForegroundWindow(new IntPtr(ComHelper.Retry<int>(() => (int)app.Hwnd))); } catch (Exception) { }
            return true;
        }

        // ------------------------------------------------------------------
        //  خريطة الخلايا: هي الخلايا الوحيدة يلي بتتعبّى بشيت النموذج
        // ------------------------------------------------------------------
        private static void Fill(dynamic ws, Establishment establishment, PeriodResult r, DateTime signingDate)
        {
            SetText(ws, "H1", r.FormCode + "-" + r.Year);   // مثال: 21-2026
            SetText(ws, "N6", establishment.Name);                // اسم المؤسسة
            SetText(ws, "N7", establishment.NssfNumber);          // رقم المؤسسة

            SetBranch(ws, 10, r.SickMaternity);             // المرض والأمومة
            SetBranch(ws, 11, r.EndOfService);              // تعويض نهاية الخدمة
            SetBranch(ws, 12, r.FamilyAllowance);           // التعويضات العائلية

            SetNumber(ws, "N13", r.TotalDue);               // مجموع الاشتراكات المستحقة
            SetNumber(ws, "N14", r.FamilyAllowancePaid);    // التعويضات العائلية المدفوعة
            SetNumber(ws, "N15", r.Net);                    // الباقي المتوجب للصندوق

            ws.Range["N18"].Value2 = signingDate.Date.ToOADate();   // التاريخ
        }

        private static void SetBranch(dynamic ws, int row, BranchResult b)
        {
            ws.Range["O" + row].Value2 = (double)b.Rate;          // المعدل
            ws.Range["P" + row].Value2 = (double)b.Wages;         // الأجور ولواحقها
            ws.Range["Q" + row].Value2 = b.Employees;             // عدد الأجراء
            ws.Range["N" + row].Value2 = (double)b.Contribution;  // الاشتراك المستحق
        }

        private static void SetNumber(dynamic ws, string address, decimal value)
        {
            ws.Range[address].Value2 = (double)value;
        }

        // منكتب النص بخلية بصيغة "نص" لحتى Excel ما يفسّر "21-2026" أو رقم المؤسسة على إنو تاريخ
        private static void SetText(dynamic ws, string address, string text)
        {
            dynamic cell = ws.Range[address];
            cell.NumberFormat = "@";
            cell.Value2 = text ?? "";
        }

        // منشيل باقي الشيتات (بما فيهم "ضمان") من النسخة المؤقتة، فبيضل النموذج المطلوب لحالو
        private static void KeepOnlySheet(dynamic wb, string keepName)
        {
            int count = wb.Worksheets.Count;
            for (int i = count; i >= 1; i--)
            {
                dynamic s = wb.Worksheets[i];
                string name = s.Name;
                if (!string.Equals(name, keepName, StringComparison.Ordinal))
                {
                    s.Delete();
                    ComHelper.Release((object)s);
                }
                // الشيت يلي بدنا ياه: ما منعمللو Release هون! لأنو نفس الـ RCW تبع ws،
                // وإذا انحرّر بيطلع الخطأ "COM object that has been separated from its underlying RCW".
                // بينحرّر مرة وحدة بالـ finally تبع RunCore.
            }
        }

        /// <summary>بيمسح ملفات مؤقتة قديمة من تشغيلات سابقة.</summary>
        public static void CleanOldTempFiles()
        {
            try
            {
                foreach (string f in Directory.GetFiles(Path.GetTempPath(), "nssf_*.xlsx"))
                {
                    try
                    {
                        if (File.GetLastWriteTime(f) < DateTime.Now.AddDays(-1))
                            File.Delete(f);
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }
        }
    }
}
