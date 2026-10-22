"""Arabic user-interface texts (no personal data is ever formatted here)."""

from __future__ import annotations

from binaa_reconciler.core.models import Direction, IssueCode

PRIVACY_NOTICE = "🔒 المعالجة محلية — لا يتم رفع الملفات إلى الإنترنت"
APP_TITLE = "تسوية الحسابات"

DIRECTION_AR = {Direction.IN: "دخول", Direction.OUT: "خروج"}

REASONS_AR = {
    "AMOUNT_EXACT": "نفس المبلغ",
    "DIRECTION_OK": "اتجاه متوافق",
    "DATE_SAME": "نفس اليوم",
    "DATE_1": "فرق يوم واحد",
    "DATE_2": "فرق يومين",
    "DATE_3": "فرق 3 أيام",
    "DATE_4": "فرق 4 أيام",
    "DATE_5": "فرق 5 أيام",
    "DATE_6": "فرق 6 أيام",
    "DATE_7": "فرق 7 أيام",
    "DATE_LAG_REVIEW": "فرق تاريخ من 4 إلى 7 أيام: للمراجعة فقط (دليل قوي آخر)",
    "TIMING_GROUP": "مجموعة متساوية (نفس المبلغ والعدد والمجموع خلال 7 أيام): فرق توقيت",
    "DUPLICATE_KEY": "عمليات متطابقة المبلغ والتاريخ والاتجاه: الاقتران مجمَّع، راجع الاقتران الفردي",
    "REFERENCE_MATCH": "رقم العملية متطابق",
    "PHONE_MATCH": "رقم الهاتف متطابق",
    "PHONE_CONFLICT": "رقما الهاتف مختلفان",
    "ALIAS_MATCH": "اسم بديل مؤكد",
    "NAME_SIMILAR": "تشابه في الاسم",
    "DESCRIPTION_SIMILAR": "تشابه في الوصف",
    "UNIQUE_AMOUNT": "المبلغ فريد في الفترة",
    "ROUND_AMOUNT": "مبلغ دائري شائع",
    "GENERIC_DESCRIPTION": "وصف عام",
}

ISSUES_AR = {
    IssueCode.FILE_UNREADABLE: "تعذر قراءة الملف",
    IssueCode.NO_TEXT_LAYER: "ملف PDF لا يحتوي نصًا قابلًا للقراءة (قد يكون صورة ممسوحة)",
    IssueCode.PAGE_WITHOUT_HEADER: "صفحة تحتوي بيانات بدون عناوين أعمدة",
    IssueCode.ORPHAN_LINE: "سطر لا يمكن ربطه بعملية",
    IssueCode.UNALIGNED_TEXT: "نص خارج الأعمدة المعروفة",
    IssueCode.OPENING_BALANCE_MISMATCH: "مبلغ الرصيد المنقول (B/F) لا يساوي رصيده",
    IssueCode.HEADER_NOT_FOUND: "لم يتم العثور على عناوين الأعمدة المطلوبة",
    IssueCode.AMBIGUOUS_HEADER: "عناوين أعمدة مكررة",
    IssueCode.MULTIPLE_CANDIDATE_SHEETS: "الملف يحتوي أكثر من ورقة قيود",
    IssueCode.SHEET_NOT_FOUND: "الورقة المطلوبة غير موجودة",
    IssueCode.AMBIGUOUS_DATE_ORDER: "صيغة التاريخ غامضة (يوم/شهر)",
    IssueCode.BAD_DATE: "تاريخ غير صالح",
    IssueCode.BAD_AMOUNT: "مبلغ غير مقروء",
    IssueCode.AMBIGUOUS_AMOUNT: "مبلغ غامض (فاصلة أو نقطة)",
    IssueCode.NO_AMOUNT: "سطر بلا مبلغ",
    IssueCode.SPLIT_ROW_MERGED: "عملية مقسومة على سطرين — دُمجت تلقائيًا",
    IssueCode.BOTH_SIDES: "مدين ودائن في نفس السطر",
    IssueCode.NEGATIVE_AMOUNT: "مبلغ سالب",
    IssueCode.BALANCE_CHAIN_BREAK: "انقطاع في سلسلة الأرصدة",
    IssueCode.FORMULA_WITHOUT_CACHED_VALUE: "صيغة Excel بدون قيمة محفوظة",
    IssueCode.FORMULA_ERROR_VALUE: "خطأ Excel في خلية",
    IssueCode.INVALID_COLUMN_MAPPING: "ربط أعمدة غير صالح",
    IssueCode.MISSING_ACCOUNT_PROFILE: "نوع الحساب غير محدد",
    IssueCode.INVALID_SCOPE: "نهاية الفترة قبل بدايتها",
}

STEPS = (
    "قراءة الطرف الأول",
    "قراءة الطرف الثاني",
    "التحقق من البيانات",
    "تحديد نطاق التسوية",
    "البحث عن المرشحين",
    "المطابقة",
    "تحليل العمليات غير المطابقة",
    "إعداد النتائج",
)


def issue_text(code: IssueCode) -> str:
    return ISSUES_AR.get(code, code.value)
