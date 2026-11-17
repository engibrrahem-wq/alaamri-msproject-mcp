# دليل تنصيب Alaamri MS Project MCP

تطوير **د. إبراهيم م. أ. العامري** (Dr Ibrahim M.A. Alaamri)، ومرخّص برخصة MIT المفتوحة: [LinkedIn](https://www.linkedin.com/in/dribrahim-alameri/).

هذه الأداة تربط **Claude** ببرنامج **Microsoft Project** على حاسبتك. بعد تنصيبها تكتب لـ Claude بالعربي أو الإنجليزي، وهو يبني الجدول ويعدّله داخل البرنامج مباشرة، كما في الفيديو.

## المتطلبات
- Windows 10 أو 11.
- **Microsoft Project** (نسخة سطح المكتب) منصّب ومفعّل.
- **Python 3.10** أو أحدث، من [python.org](https://www.python.org/downloads/). أثناء التنصيب أشّر على خيار **Add python.exe to PATH**.
- **Claude Desktop** من [claude.ai/download](https://claude.ai/download).

## التنصيب (5 دقائق)
1. حمّل الأداة من GitHub: **https://github.com/engibrrahem-wq/alaamri-msproject-mcp**
   - اضغط الزر الأخضر **Code** ثم **Download ZIP**، أو حمّلها مباشرة من: https://github.com/engibrrahem-wq/alaamri-msproject-mcp/archive/refs/heads/main.zip
   - فكّ الضغط في مكان ثابت، مثل `C:\alaamri-msproject-mcp`.
2. داخل المجلد الذي فككت ضغطه شغّل ملف `install.bat` بنقرة مزدوجة. سيُنشئ بيئة Python وينصّب الأداة، ثم يطبع لك نص الإعداد جاهزاً بمسار مجلدك.
   - بديل للمحترفين: `git clone` ثم `pip install -e .`، كما في README على GitHub.
3. في Claude Desktop افتح **Settings ← Developer ← Edit Config**. ينفتح ملف `claude_desktop_config.json`.
4. الصق النص الذي طبعه `install.bat` واحفظ الملف.
   - إن كان في الملف إعدادات سابقة، فأضف قسم `"alaamri-msproject"` داخل `"mcpServers"` الموجود ولا تكرّره.
5. **أغلق Claude Desktop تماماً**: من أيقونته بجانب الساعة اختر Quit. ثم افتحه من جديد.
6. تحقق: في نافذة المحادثة، ضمن أيقونة الأدوات، يجب أن تظهر أدوات **alaamri-msproject**.

## الاستخدام
1. افتح MS Project، وافتح **نسخة** من الجدول أو ملفاً فارغاً. لا تشتغل على ملف العميل الأصلي.
2. اكتب طلبك لـ Claude، مثل البرومبت الموجود في [`prompts/villa-schedule-prompt-ar.txt`](../../prompts/villa-schedule-prompt-ar.txt).
3. عند أول استخدام سيطلب Claude إذنك لتشغيل الأداة. اختر **Allow**.

## الأدوات المتاحة
| الأداة | الوظيفة |
|---|---|
| `status` | يتأكد أن Project شغّال، ويعرض المشاريع المفتوحة |
| `open_project` / `new_project` | يفتح ملف .mpp أو .xml، أو ينشئ مشروعاً جديداً |
| `project_summary` | التواريخ، وعدد المهام، والمهام الحرجة، ونسبة الإنجاز |
| `list_tasks` | المهام مع المدد والعلاقات والسماحية |
| `add_tasks` / `update_task` / `delete_task` | إضافة المهام وتعديلها وحذفها |
| `set_status_date` / `save_baseline` | تاريخ الحالة وحفظ خط الأساس |
| `save_project` | الحفظ، أو الحفظ باسم جديد |

## ملاحظات مهمة
- **راجع كل جدول** قبل اعتماده. الذكاء الاصطناعي يسرّع العمل، والقرار الهندسي يبقى عليك.
- **اشتغل دائماً على نسخة**، واحفظ باسم جديد.
- **الأداة محلية:** تعمل على حاسبتك فقط. أما نص المحادثة فيُرسل إلى Claude كأي محادثة عادية، فلا تضع فيه بيانات سرّية لعميل دون إذنه.
- **مشكلة شائعة:** إن لم تظهر الأدوات، فتأكد أن المسار في ملف الإعداد صحيح وأنك أغلقت Claude Desktop بالكامل. Project يجب أن يكون نسخة سطح المكتب، لا نسخة الويب.
