; سكريبت Inno Setup لبرنامج جداول الاشتراكات (الضمان الاجتماعي).
; هيدا الملف لازم ينفتح بـ Inno Setup Compiler (على ويندوز) بعد ما تعمل dotnet publish.
; شوف تعليمات "التوزيع كملف تثبيت (Inno Setup)" بملف README.md.
;
; ملاحظة تشفير: احفظ هالملف بصيغة "UTF-8 with BOM" (فيه أصلاً BOM) لحتى العربي ما ينكتب مبعثر
; بواجهة التثبيت. إذا انفتح بمحرر وحفظتو من جديد، تأكد التشفير ضل UTF-8 with BOM.

#define MyAppName "برنامج جداول الاشتراكات - الضمان الاجتماعي"
#define MyAppNameEn "NssfPayroll"
#define MyAppVersion "1.0.0"
#define MyAppExeName "NssfPayroll.exe"

; مجلد الـ publish: النتيجة يلي طلعت من:
;   dotnet publish -c Release -r win-x64 --self-contained true -o bin\publish\win-x64
#define PublishDir "..\bin\publish\win-x64"

[Setup]
AppId={{8AA76D86-6B7C-4D86-936E-A6FBA0DFDAA4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppName}
DefaultDirName={localappdata}\Programs\{#MyAppNameEn}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
; PrivilegesRequired=lowest: التثبيت بمجلد المستخدم بلا صلاحيات مدير (Admin) أبداً.
; PrivilegesRequiredOverridesAllowed: بيسمح للمستخدم يختار "تثبيت لكل المستخدمين" إذا بدو وعندو صلاحية.
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
OutputDir=Output
OutputBaseFilename=NssfPayrollSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
; ما في ملف LICENSE.txt، فمنشيل سطر الترخيص. ضيفو إذا بدك:
; LicenseFile=LICENSE.txt

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "إنشاء أيقونة على سطح المكتب"; GroupDescription: "أيقونات إضافية:"

[Files]
; بينسخ كل ملفات الـ publish (البرنامج + مكتبات .NET + قالب الإكسل) كاملين
Source: "{#PublishDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\{cm:UninstallProgram,{#MyAppName}}"; Filename: "{uninstallexe}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "تشغيل البرنامج الآن"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
; إذا بدك تحذف بيانات المستخدم (%AppData%\NssfPayroll) عند الحذف، فك التعليق عن السطر:
; Type: filesandordirs; Name: "{userappdata}\NssfPayroll"
