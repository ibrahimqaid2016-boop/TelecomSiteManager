# Telecom Site Manager

تطبيق Kivy لإدارة بيانات مواقع الاتصالات، ويشمل:
- Site Database
- 2G System
- LTE System
- Microwave System
- Tributary / E1 Management
- Power System
- Alarm System
- Excel Export / Import لقاعدة البيانات كاملة

## إنشاء APK

المشروع مجهز لـ Buildozer. يمكن بناء APK باستخدام Linux/WSL أو GitHub Actions.

الأمر الأساسي:

```bash
buildozer -v android debug
```

بعد نجاح البناء ستجد ملف APK داخل مجلد `bin/`.

ملاحظة: قاعدة البيانات يتم حفظها داخل مجلد بيانات التطبيق القابل للكتابة، وليس بجانب `main.py`، حتى تعمل النسخة المجمعة بشكل صحيح على Android.
