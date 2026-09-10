[app]
# (str) Title of your application
title = Telecom Site Manager

# (str) Package name
package.name = telecomsitemanager

# (str) Package domain (unique identifier)
package.domain = com.sabafon

# (str) Source code directory
source.dir = .

# (list) Source files to include
source.include_exts = py,png,jpg,jpeg,kv,atlas,xlsx,db

# (str) Application version
version = 1.0.0

# (list) Application requirements
requirements = python3,kivy==2.3.1,openpyxl

# (str) Supported orientation
orientation = portrait

# (bool) Fullscreen mode
fullscreen = 0

# (list) Permissions required for Excel export/import in shared storage
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# Android build settings
android.api = 28
android.minapi = 21
android.ndk = 25b
android.accept_sdk_license = True

# (str) Presplash
# presplash.filename = %(source.dir)s/presplash.png

# (str) Icon
# icon.filename = %(source.dir)s/icon.png

# (str) Supported architectures
android.archs = arm64-v8a,armeabi-v7a
# Python-for-Android custom source
p4a.source_dir = %(source.dir)s/python-for-android

# (bool) Enable Android logcat on build errors
log_level = 2

[buildozer]
# (int) Log level (0 = error only, 1 = warning, 2 = info, 3 = debug)
log_level = 2
warn_on_root = 1
