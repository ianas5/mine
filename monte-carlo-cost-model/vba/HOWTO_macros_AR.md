# دليل تركيب الماكروز — MonteCarloCostModel_VBA

## الماكروز اللي تحتاجها (3 ملفات فقط)
استورد هذي الثلاثة بس (الباقي لملف ثاني، لا تستوردها هنا):
- `modEngine.bas` — المحرّك الرئيسي
- `modDistributions.bas` — دوال السحب العشوائي (داخلية)
- `modExport.bas` — تصدير PDF

---

## الخطوة 1: افتح الملف وفعّل الماكرو
1. افتح `MonteCarloCostModel_VBA.xlsx` في Excel على الكمبيوتر (مو الويب/الجوال).
2. لو ظهر شريط أصفر "Enable Content / تمكين المحتوى" → اضغطه.
3. لو الملف نزل من النت ومنبلوك: اضغط يمين على الملف ▸ Properties ▸ علّم Unblock ▸ OK.

## الخطوة 2: استورد الوحدات (Alt+F11)
1. اضغط Alt + F11 (يفتح محرر VBA).
2. File ▸ Import File…
3. اختر `modEngine.bas` ▸ Open. كرّرها لـ `modDistributions.bas` و `modExport.bas`.
4. لازم تشوفها على اليسار تحت Modules: modEngine, modDistributions, modExport.
5. سكّر المحرر (رجوع لـ Excel).

## الخطوة 3: احفظه كـ ‎.xlsm
- File ▸ Save As ▸ من نوع الملف اختر Excel Macro-Enabled Workbook (*.xlsm) ▸ Save.
- مهم: لو حفظته xlsx عادي بيحذف الماكروز.

## الخطوة 4: فعّل تبويب Developer (مرة وحدة)
- File ▸ Options ▸ Customize Ribbon ▸ علّم Developer على اليمين ▸ OK.

## الخطوة 5: ضيف الأزرار واربط كل ماكرو
في ورقة Dashboard فيه 3 خلايا ملوّنة جاهزة في الصف 7. لكل واحدة:
1. Developer ▸ Insert ▸ Button (Form Control) (أول أيقونة).
2. ارسم الزر فوق الخلية.
3. تطلع نافذة Assign Macro → اختر الماكرو من الجدول ▸ OK.

| مكان الزر | اسم الزر | الماكرو اللي تربطه |
|---|---|---|
| A7:C7 | ▶ Run Simulation | `RunSimulation` |
| D7:E7 | 📅 Apply Years | `ApplyYears` |
| F7:G7 | ⬇ Export Report | `ExportReport` |

تبي زر رابع لـ `SyncTables`؟ اختياري — ارسم زر في أي مكان فاضي واربطه فيه. وإلا يشتغل تلقائيًا مع Run.

---

## شرح كل ماكرو ووظيفته
| الماكرو | وين | وش يسوي | تشغّله متى |
|---|---|---|---|
| `RunSimulation` | زر Run / Alt+F8 | يشغّل المحاكاة كاملة ويكتب كل النتائج والرسوم | كل ما تبي نتائج |
| `SyncTables` | تلقائي + Alt+F8 | يضيف صف Profiling لكل بند/خطر جديد ويربط الاسم | بعد ما تضيف بنود |
| `ApplyYears` | زر Apply Years | يخفي/يظهر أعمدة السنوات حسب Number of years | لما تغيّر عدد السنوات |
| `ExportReport` | زر Export | يصدّر Dashboard + Results كـ PDF | لما تبي تقرير |

ملاحظة: الدوال داخل `modDistributions` (مثل SampleTriangular) ما راح تظهر في قائمة الماكرو ولا تربطها بأزرار — هي داخلية يستخدمها المحرّك تلقائيًا.

---

## جرّبها بسرعة
1. اضغط زر Run Simulation (أو Alt+F8 ▸ RunSimulation).
2. لازم تطلع رسالة "Simulation complete" وتتعبّى الأرقام والرسوم.
3. الأرقام المتوقعة تقريبًا: المتوسط ~464,000 و P80 ~504,000.

لو طلع أي خطأ (مثل "RunSimulation error: ...") انسخ نص الرسالة وابعثه لي — غالبًا تصليح سطر بسيط.
