using System;
using System.Runtime.InteropServices;
using System.Threading;

namespace NssfPayroll
{
    /// <summary>
    /// أدوات مشتركة للتعامل الآمن مع Excel عبر COM (dynamic): إعادة محاولة الأخطاء العابرة، وتحرير
    /// كائنات COM. مشتركة بين كل شي بيفتح Excel (ExcelPrinter و EmployeeReportPrinter) حتى ما تنكرر
    /// نفس الغلطة يلي انصلحت قبل بمكان وتنسى تنصلح بمكان تاني.
    /// </summary>
    public static class ComHelper
    {
        private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
        private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

        private static bool IsTransientComError(Exception ex)
        {
            COMException com = ex as COMException;
            return com != null && (com.ErrorCode == RPC_E_CALL_REJECTED || com.ErrorCode == RPC_E_SERVERCALL_RETRYLATER);
        }

        // بيعيد المحاولة لغاية 40 مرة (كل ربع ثانية، يعني تقريباً 10 ثواني بالمجموع) إذا Excel رفض
        // الأمر لأنو مشغول أو لسا عم يجهّز. أي خطأ تاني (حقيقي) بيطلع فوراً بدون ما ننتظر.
        //
        // تحذير مهم لأي استدعاء جديد: لازم تلفّ أي عملية COM بـ Retry(() => { ... }) بأقواس {}
        // إلزامية — حتى لو عم تستدعي ميثود عادي (مو dynamic) بس بوخد معامل dynamic. لأنو C# بهالحالة
        // بيعتبر الاستدعاء كامل dynamic وبيختار الطرف الغلط (Func<T> بدل Action) بدون الأقواس، وبينهار
        // وقت التشغيل بـ "Cannot implicitly convert type 'void' to 'object'". صار هيدا الخطأ فعلياً
        // مرتين بهالمشروع (بـ ExcelPrinter.Fill وKeepOnlySheet)، فأي كود جديد بيستعمل Retry لازم ينتبه.
        public static T Retry<T>(Func<T> action)
        {
            int attempt = 0;
            while (true)
            {
                try
                {
                    return action();
                }
                catch (Exception ex) when (IsTransientComError(ex) && attempt < 40)
                {
                    attempt++;
                    Thread.Sleep(250);
                }
            }
        }

        public static void Retry(Action action)
        {
            Retry<object>(delegate { action(); return null; });
        }

        public static void Release(object o)
        {
            try
            {
                if (o != null && Marshal.IsComObject(o))
                    Marshal.FinalReleaseComObject(o);
            }
            catch (Exception) { }   // ما منخلّي خطأ التنظيف يغطّي على الخطأ الأصلي
        }

        /// <summary>
        /// نوع Excel.Application، أو استثناء بالرسالة الواضحة إذا Excel مو منصّب. بيلفّ الفحص بمحاولة
        /// دفاعية لأنو Type.GetTypeFromProgID موثّق إنو بيرجّع null إذا الـ ProgID مو مسجّل — بس بعض
        /// البيئات (اختبرنا هيك فعلياً) بترمي استثناء بدل ما ترجّع null، فمنوحّد السلوك بمحاولة/مسك.
        /// </summary>
        public static Type GetExcelType()
        {
            Type excelType;
            try
            {
                excelType = Type.GetTypeFromProgID("Excel.Application");
            }
            catch (Exception)
            {
                excelType = null;
            }

            if (excelType == null)
                throw new InvalidOperationException("Microsoft Excel مو منصّب على هالكمبيوتر.");
            return excelType;
        }
    }
}
