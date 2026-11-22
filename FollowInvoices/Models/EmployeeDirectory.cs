using System;
using System.Collections.Generic;

namespace FollowInvoices.Models
{
    public static class EmployeeDirectory
    {
        // القائمة الحالية في التطبيق (منطقة القاهرة)
        public static readonly IReadOnlyList<string> CairoEmployees = new List<string>
        {
            "خالد يوسف",
            "محمد الهاكع",
            "الشيماء ابوبكر",
            "مني شاهين",
            "ريهام الشيح",
            "ايات اسامة",
            "تامر عبد الفتاح",
            "مني الشهاوي",
            "مدحت سمير",
            "سامح مصطفي",
            "ندا سمير",
            "شيماء ماجد",
            "اسلام علي",
            "رانيا ممدوح",
            "محمد اسماعيل ",
            "نجلاء سعد",
            "ايمن آمين"
        };

        // الموظفين لمنطقة الأجد (من صور المرفق)
        public static readonly IReadOnlyList<string> AlexEmployees = new List<string>
        {
            "نيهال محمد",
            "محمود نايل",
            "اسلام ابوشال",
            "حسن عباس",
            "محمد كشك",
            "ايه زغلول",
            "سارة خميس",
            "عبير",
            "ايناس ماهر",
            "مصطفى الشقشق",
            "هشام الدماصى",
            "عمرو زغلول",
            "سامح الزرية",
            "إيناس عبد الهادى",
            "مروة نبيل"
        };

        public static IReadOnlyList<string> EmployeesForArea(string area)
        {
            return string.Equals(area?.Trim(), "alex", StringComparison.OrdinalIgnoreCase)
                ? AlexEmployees
                : CairoEmployees;
        }
    }
}
