using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Threading;

namespace NssfPayroll
{
    /// <summary>
    /// بيبني تقرير أجور موظف رسمي (ترويسة فيها اسم الموظف ورقمو بالضمان ونوع التقرير، بيانات
    /// المؤسسة والموظف كاملة، جدول الأجور فصلي أو شهري مع حدود وتظليل، وسطر مجموع + تاريخ الطباعة)
    /// بملف إكسل جديد كامل من اليمين لليسار (مو من قالب، لأنو طول الجدول بيختلف حسب عدد الفترات)،
    /// وبعدين بيفتحو بالإكسل أو بيطبعو أو بيصدّرو PDF متل باقي البرنامج.
    /// </summary>
    public static class EmployeeReportPrinter
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // ثوابت Excel المستعملة (بلا مرجع مكتبة، لأننا late-binding عبر dynamic)
        private const int XlLandscape = 2;
        private const int XlLeft = -4131;
        private const int XlCenter = -4108;
        private const int XlRight = -4152;
        private const int XlEdgeTop = 8;
        private const int XlThick = 4;
        private const int XlContinuous = 1;

        public static void Run(Establishment establishment, Employee employee, bool monthly,
                               List<EmployeeWageReportRow> rows, ExcelAction action, string pdfPath)
        {
            Type excelType = ComHelper.GetExcelType();

            try
            {
                RunCore(excelType, establishment, employee, monthly, rows, action, pdfPath);
            }
            finally
            {
                // نفس تنظيف ExcelPrinter: منجبر الـ GC يحرّر أي كائن COM ضايل حتى ما يضل EXCEL.EXE
                // مخفي شغّال بالخلفية.
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private static void RunCore(Type excelType, Establishment est, Employee emp, bool monthly,
                                    List<EmployeeWageReportRow> rows, ExcelAction action, string pdfPath)
        {
            dynamic app = null;
            dynamic wb = null;
            dynamic ws = null;
            bool leaveOpen = false;

            try
            {
                app = Activator.CreateInstance(excelType);
                // نفس فخّ "Call was rejected by callee" يلي بيصير مع ExcelPrinter (Excel لسا عم يبلّش):
                // منعيد المحاولة تلقائياً. كل استدعاء COM هون لازم يضل بأقواس {} إلزامية — راجع
                // ComHelper.Retry للسبب (استدعاء ميثود عادي بمعامل dynamic بيصير كامل الاستدعاء dynamic).
                ComHelper.Retry(() => { app.DisplayAlerts = false; });
                // Excel ظاهر لفتح التقرير وللطباعة المباشرة كمان (حتى يشوف المستخدم الجدول قبل ما
                // يطبع، مش يصير الطبع بالخفاء بلا ما يبين شي)؛ مخفي بس لتصدير PDF.
                ComHelper.Retry(() => { app.Visible = (action == ExcelAction.OpenInExcel || action == ExcelAction.Print); });

                wb = ComHelper.Retry<object>(() => app.Workbooks.Add());
                ws = ComHelper.Retry<object>(() => wb.Worksheets[1]);

                // اتجاه الشيت من اليمين لليسار: التوثيق الرسمي لـ Worksheet.DisplayRightToLeft بيقول
                // "Read-only"، بس فحصنا ماكرو VBA حقيقي مستخرج من ملف إكسل فعلي بيعمل بالظبط
                // "ActiveWindow.DisplayRightToLeft = False" — يعني هي فعلياً قابلة للتغيير بس عن طريق
                // النافذة (Window)، مش الشيت مباشرة. فمنعتمد النافذة كطريقة أساسية موثوقة، ومنجرّب
                // الشيت كمحاولة إضافية بس منتجاهل فشلها (لأنو مو أساسي، وممكن يرمي خطأ حقيقي "read-only"
                // مش خطأ عابر، فمنعزلها بمحاولة/مسك عادية بلا Retry حتى ما توقف كل التقرير).
                ComHelper.Retry(() => { app.ActiveWindow.DisplayRightToLeft = true; });
                try { ws.DisplayRightToLeft = true; } catch (Exception) { }

                ComHelper.Retry(() => { RemoveExtraSheets(wb); });     // بعض إصدارات Excel بتعمل 3 شيتات افتراضياً
                ComHelper.Retry(() => { Fill(ws, est, emp, monthly, rows); });
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

        // شيتات إضافية افتراضية (بعض إصدارات Excel بتعمل أكتر من شيت وحدة بـ Workbooks.Add) بتنمسح
        private static void RemoveExtraSheets(dynamic wb)
        {
            int count = wb.Worksheets.Count;
            for (int i = count; i >= 2; i--)
            {
                dynamic s = wb.Worksheets[i];
                s.Delete();
                ComHelper.Release((object)s);
            }
        }

        // ------------------------------------------------------------------
        //  محتوى التقرير
        // ------------------------------------------------------------------
        private static void Fill(dynamic ws, Establishment est, Employee emp, bool monthly, List<EmployeeWageReportRow> rows)
        {
            int lastCol = monthly ? 7 : 6;
            string empName = string.IsNullOrWhiteSpace(emp.FullName) ? "(بدون اسم)" : emp.FullName;

            // ---- الاتجاه من اليمين لليسار مضبوط أصلاً بـ RunCore (النافذة + محاولة الشيت) ----
            ws.Name = "أجور الموظف";

            ws.Cells.Font.Name = "Arial";
            ws.Cells.Font.Size = 11;
            ws.Cells.NumberFormat = "@";     // نص افتراضياً بكل الشيت، حتى ما Excel يفسّر أرقام الفترة/السنة كتاريخ

            int r = 1;

            // ---- الترويسة: أجور الموظف <الاسم> / رقم الضمان / نوع التقرير ----
            SetMergedTitle(ws, r, lastCol, "أجور الموظف: " + empName, 16); r++;
            SetMergedTitle(ws, r, lastCol,
                "رقم الضمان: " + OrDash(emp.NssfNumber) + "        نوع التقرير: " + (monthly ? "شهري" : "فصلي"), 12);
            ws.Range[ws.Cells[r, 1], ws.Cells[r, lastCol]].Borders[XlEdgeTop].LineStyle = XlContinuous;
            r += 2;

            // ---- بيانات المؤسسة ----
            SetSectionHeader(ws, r++, lastCol, "بيانات المؤسسة");
            SetInfoRow(ws, r++, lastCol, "اسم المؤسسة", OrDash(est.Name));
            SetInfoRow(ws, r++, lastCol, "رقم المؤسسة بالضمان", OrDash(est.NssfNumber));
            SetInfoRow(ws, r++, lastCol, "صاحب المؤسسة", OrDash(est.OwnerName));
            SetInfoRow(ws, r++, lastCol, "رقم الهاتف", OrDash(est.OwnerPhone));
            r++;

            // ---- بيانات الموظف ----
            SetSectionHeader(ws, r++, lastCol, "بيانات الموظف");
            SetInfoRow(ws, r++, lastCol, "الاسم الثلاثي", empName);
            SetInfoRow(ws, r++, lastCol, "رقم الضمان", OrDash(emp.NssfNumber));
            string marital = emp.IsMarried
                ? ("متأهل" + (emp.ChildrenCount.HasValue ? " (" + emp.ChildrenCount.Value.ToString(CultureInfo.InvariantCulture) + " أولاد)" : ""))
                : "أعزب";
            SetInfoRow(ws, r++, lastCol, "الحالة الاجتماعية", marital);
            SetInfoRow(ws, r++, lastCol, "تاريخ الاستخدام", emp.StartDate.HasValue ? emp.StartDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : "—");
            SetInfoRow(ws, r++, lastCol, "تاريخ الترك", emp.EndDate.HasValue ? emp.EndDate.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : "—");
            r++;

            // ---- رأس الجدول ----
            int headerRow = r;
            int col = 1;
            SetHeaderCell(ws, headerRow, col++, "السنة");
            SetHeaderCell(ws, headerRow, col++, "الفترة");
            if (monthly) SetHeaderCell(ws, headerRow, col++, "الشهر");
            SetHeaderCell(ws, headerRow, col++, "مرض وأمومة");
            SetHeaderCell(ws, headerRow, col++, "نهاية خدمة");
            SetHeaderCell(ws, headerRow, col++, "تعويضات عائلية");
            SetHeaderCell(ws, headerRow, col++, "التعويض المدفوع");
            r++;

            decimal totalSm = 0m, totalEos = 0m, totalFa = 0m, totalPaid = 0m;
            foreach (EmployeeWageReportRow row in rows)
            {
                col = 1;
                SetDataCell(ws, r, col++, row.Year.ToString(CultureInfo.InvariantCulture));
                SetDataCell(ws, r, col++, row.Period);
                if (monthly) SetDataCell(ws, r, col++, row.Month);
                SetNumberCell(ws, r, col++, row.SickMaternity);
                SetNumberCell(ws, r, col++, row.EndOfService);
                SetNumberCell(ws, r, col++, row.FamilyAllowance);
                SetNumberCell(ws, r, col++, row.FamilyAllowancePaid);

                totalSm += row.SickMaternity;
                totalEos += row.EndOfService;
                totalFa += row.FamilyAllowance;
                totalPaid += row.FamilyAllowancePaid;
                r++;
            }

            // ---- سطر المجموع ----
            int totalRow = r;
            col = 1;
            SetDataCell(ws, totalRow, col++, "");
            SetDataCell(ws, totalRow, col++, "المجموع");
            if (monthly) SetDataCell(ws, totalRow, col++, "");
            SetNumberCell(ws, totalRow, col++, totalSm);
            SetNumberCell(ws, totalRow, col++, totalEos);
            SetNumberCell(ws, totalRow, col++, totalFa);
            SetNumberCell(ws, totalRow, col++, totalPaid);

            dynamic headerRange = ws.Range[ws.Cells[headerRow, 1], ws.Cells[headerRow, lastCol]];
            headerRange.Font.Bold = true;
            headerRange.Interior.Color = 0xD9D9D9;   // رمادي فاتح (BGR)
            headerRange.HorizontalAlignment = XlCenter;

            dynamic totalRange = ws.Range[ws.Cells[totalRow, 1], ws.Cells[totalRow, lastCol]];
            totalRange.Font.Bold = true;
            totalRange.Interior.Color = 0xD9D9D9;
            totalRange.Borders[XlEdgeTop].LineStyle = XlContinuous;
            totalRange.Borders[XlEdgeTop].Weight = XlThick;

            if (rows.Count > 0)
            {
                dynamic tableRange = ws.Range[ws.Cells[headerRow, 1], ws.Cells[totalRow, lastCol]];
                tableRange.Borders.LineStyle = XlContinuous;
                tableRange.HorizontalAlignment = XlCenter;
            }
            r = totalRow + 2;

            // ---- تاريخ الطباعة ----
            SetDataCell(ws, r, 1, "تاريخ الطباعة: " + DateTime.Now.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture));
            dynamic dateCell = ws.Cells[r, 1];
            dateCell.Font.Bold = true;
            dynamic dateRange = ws.Range[ws.Cells[r, 1], ws.Cells[r, lastCol]];
            dateRange.Merge();
            dateRange.HorizontalAlignment = XlRight;

            // ---- عرض الأعمدة ----
            ws.Columns.ColumnWidth = 15;
            ws.Columns[1].ColumnWidth = 22;   // عمود التسميات (اسم المؤسسة:، رقم الضمان:...)
            if (monthly) ws.Columns[3].ColumnWidth = 16;   // عمود اسم الشهر (تشرين الثاني أطول اسم)

            // ---- إعداد الصفحة: أفقي، وبعرض صفحة وحدة ----
            ws.PageSetup.Orientation = XlLandscape;
            ws.PageSetup.Zoom = false;
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = false;
            ws.PageSetup.CenterHorizontally = true;
        }

        private static string OrDash(string s)
        {
            return string.IsNullOrWhiteSpace(s) ? "—" : s;
        }

        private static void SetMergedTitle(dynamic ws, int row, int lastCol, string text, int fontSize)
        {
            ws.Cells[row, 1].Value2 = text;
            dynamic range = ws.Range[ws.Cells[row, 1], ws.Cells[row, lastCol]];
            range.Merge();
            range.Font.Bold = true;
            range.Font.Size = fontSize;
            range.HorizontalAlignment = XlCenter;
        }

        private static void SetSectionHeader(dynamic ws, int row, int lastCol, string text)
        {
            ws.Cells[row, 1].Value2 = text;
            dynamic range = ws.Range[ws.Cells[row, 1], ws.Cells[row, lastCol]];
            range.Merge();
            range.Font.Bold = true;
            range.Font.Size = 12;
            range.Interior.Color = 0xF2F2F2;   // رمادي فاتح جداً
            range.HorizontalAlignment = XlRight;
            range.Borders[XlEdgeTop].LineStyle = XlContinuous;
        }

        // تسمية بعمود 1 (تظهر يمين بفضل RTL) وقيمة مدموجة من عمود 2 لآخر عمود (تظهر يسارها)
        private static void SetInfoRow(dynamic ws, int row, int lastCol, string label, string value)
        {
            dynamic labelCell = ws.Cells[row, 1];
            labelCell.Value2 = label + ":";
            labelCell.Font.Bold = true;
            labelCell.HorizontalAlignment = XlRight;

            ws.Cells[row, 2].Value2 = value ?? "—";
            dynamic valueRange = ws.Range[ws.Cells[row, 2], ws.Cells[row, lastCol]];
            valueRange.Merge();
            valueRange.HorizontalAlignment = XlRight;
        }

        private static void SetHeaderCell(dynamic ws, int row, int col, string text)
        {
            ws.Cells[row, col].Value2 = text;
        }

        private static void SetDataCell(dynamic ws, int row, int col, string text)
        {
            ws.Cells[row, col].Value2 = text ?? "";
        }

        private static void SetNumberCell(dynamic ws, int row, int col, decimal value)
        {
            dynamic cell = ws.Cells[row, col];
            cell.NumberFormat = "#,##0";
            cell.Value2 = (double)value;
        }
    }
}
