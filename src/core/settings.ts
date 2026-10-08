/** Everything on the sheet that is not derived from the masters (brief §7, §11). */
export interface Settings {
  /** Line under the title. */
  semester: string;
  /** Prefix of the `القسم` line: `الهندسة الكهربية (هندسة …)`. */
  department: string;
  /** The four signature blocks, right to left. */
  signatures: { title: string; name: string }[];
}

export const DEFAULT_SETTINGS: Settings = {
  semester: 'الفصل الدراسي الأول 2027-2026',
  department: 'الهندسة الكهربية',
  signatures: [
    { title: 'لجنة الجدول', name: 'د. أسماء راضي & م.م تامر الشرقاوي & م.م غادة المنوفي' },
    { title: 'رئيس مجلس القسم', name: 'أ. م. د. جمعة عثمان' },
    { title: 'وكيل الكلية لشئون التعليم والطلاب', name: 'أ. د. علي الطوانسي' },
    { title: 'عميد الكلية', name: 'أ. د. محمـــــد خيـــرت' },
  ],
};
