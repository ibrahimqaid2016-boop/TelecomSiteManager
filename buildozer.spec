[app]

# (str) Title of your application
title = Telecom Site Management System

# (str) Package name
package.name = telecomsitemanager

# (str) Package domain
package.domain = org.telecomsitemanager

# (str) Source code directory
source.dir = .

# (str) Main Python file
source.main = main.py

# (str) Application version
version = 1.0

# (str) Python requirements
requirements = python3,kivy==2.3.1,openpyxl,plyer

# (str) Supported orientation
orientation = portrait

# (bool) Fullscreen mode
fullscreen = 0

# Presplash and icon can be added later if desired.
# presplash.filename = %(source.dir)s/data/presplash.png
# icon.filename = %(source.dir)s/data/icon.png


[buildozer]

# (str) Log level
log_level = 2

# (bool) Warn if running as root
warn_on_root = 1


[android]

# (str) Android API
android.api = 35

# (str) Minimum Android API
android.minapi = 23

# (str) Android NDK version
android.ndk = 25b

# Build only the modern 64-bit ABI.
# This avoids the armeabi-v7a SDL2_mixer build failure encountered previously.
android.archs = arm64-v8a

# (bool) Accept Android SDK licenses automatically
android.accept_sdk_license = True

# Permissions used for the current Excel import/export implementation.
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE

# Keep app data in the app's private storage.
android.private_storage = True

# Use the SDL2 bootstrap required by Kivy.
p4a.bootstrap = sdl2


[app:python]
# sqlite3, os, tempfile and datetime are part of Python.
# pyjnius is provided by python-for-android/Kivy and is used by main.py.
