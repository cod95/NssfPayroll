using System;
using System.Drawing;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

namespace NssfPayroll
{
    /// <summary>قراءة الأرقام المكتوبة باليد: بتقبل 28000000 أو 28,000,000 أو الأرقام العربية ٢٨٠٠٠٠٠٠.</summary>
    public static class NumberInput
    {
        public static string Normalize(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char ch in (s ?? "").Trim())
            {
                if (ch >= '\u0660' && ch <= '\u0669') sb.Append((char)('0' + (ch - '\u0660')));        // ٠-٩
                else if (ch >= '\u06F0' && ch <= '\u06F9') sb.Append((char)('0' + (ch - '\u06F0')));   // ۰-۹
                else if (ch == '\u066B') sb.Append('.');                                               // ٫ فاصلة عشرية عربية
                else if (ch == ',' || ch == '\u066C' || ch == '\u060C' || char.IsWhiteSpace(ch)) continue;
                else sb.Append(ch);
            }
            return sb.ToString();
        }

        /// <summary>النص الفاضي = 0 وبيرجّع true. بيرجّع false إذا النص مو رقم.</summary>
        public static bool TryParseDecimal(string text, out decimal value)
        {
            string s = Normalize(text);
            if (s.Length == 0)
            {
                value = 0m;
                return true;
            }
            return decimal.TryParse(s, NumberStyles.Number, CultureInfo.InvariantCulture, out value);
        }

        /// <summary>النص الفاضي = null. بيرجّع false إذا النص مو رقم صحيح.</summary>
        public static bool TryParseInt(string text, out int? value)
        {
            string s = Normalize(text);
            if (s.Length == 0)
            {
                value = null;
                return true;
            }
            int n;
            if (int.TryParse(s, NumberStyles.None, CultureInfo.InvariantCulture, out n))
            {
                value = n;
                return true;
            }
            value = null;
            return false;
        }

        /// <summary>للسماح بالأرقام والفواصل ومفاتيح التحكم بس (KeyPress).</summary>
        public static bool IsNumericKey(char c)
        {
            return char.IsControl(c) || char.IsDigit(c) || c == ',' || c == '.'
                || c == '\u066B' || c == '\u066C' || c == '\u060C';
        }

        public static string Format(decimal v)
        {
            return v.ToString("N0", CultureInfo.InvariantCulture);
        }

        public static string FormatOrBlank(decimal v)
        {
            return v == 0m ? "" : Format(v);
        }
    }

    /// <summary>أدوات واجهة مشتركة (أزرار، عناوين، رسائل بالعربي من اليمين لليسار).</summary>
    public static class UiHelpers
    {
        private static readonly MessageBoxOptions RtlOptions =
            MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign;

        /// <summary>للاختبارات الآلية بس: بدون رسائل على الشاشة (والأسئلة بتنجاوب "نعم").</summary>
        public static bool Silent;
        public static string LastMessage;

        public static Label MakeLabel(string text)
        {
            Label l = new Label();
            l.Text = text;
            l.AutoSize = true;
            l.Margin = new Padding(14, 9, 4, 0);
            return l;
        }

        public static Button MakeButton(string text, EventHandler onClick)
        {
            Button b = new Button();
            b.Text = text;
            b.AutoSize = true;
            b.Padding = new Padding(10, 4, 10, 4);
            b.Margin = new Padding(6);
            b.Click += onClick;
            return b;
        }

        public static void Info(IWin32Window owner, string text)
        {
            if (Silent) { LastMessage = text; return; }
            MessageBox.Show(owner, text, "معلومة", MessageBoxButtons.OK, MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1, RtlOptions);
        }

        public static void Warn(IWin32Window owner, string text)
        {
            if (Silent) { LastMessage = text; return; }
            MessageBox.Show(owner, text, "انتبه", MessageBoxButtons.OK, MessageBoxIcon.Warning,
                MessageBoxDefaultButton.Button1, RtlOptions);
        }

        public static void Error(IWin32Window owner, string text)
        {
            if (Silent) { LastMessage = text; return; }
            MessageBox.Show(owner, text, "خطأ", MessageBoxButtons.OK, MessageBoxIcon.Error,
                MessageBoxDefaultButton.Button1, RtlOptions);
        }

        /// <summary>سؤال نعم/لا. الافتراضي "لا" لحتى ما يصير حذف بالغلط.</summary>
        public static DialogResult Ask(IWin32Window owner, string text)
        {
            if (Silent) { LastMessage = text; return DialogResult.Yes; }
            return MessageBox.Show(owner, text, "تأكيد", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2, RtlOptions);
        }
    }
}
