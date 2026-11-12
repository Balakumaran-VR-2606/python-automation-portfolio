# Daymark native reminder apps

These are separate, device-local apps. They do not use Google sign-in, Firebase, or cloud sync.

## What is included

- `android/`: native Android app using local storage and exact OS alarms. A foreground service plays the beep for the duration chosen for each reminder. Alarms are restored after a phone reboot.
- `windows/`: windowed Windows app that stays in the notification area when its window is closed. Windows Task Scheduler launches a small alarm window at the chosen time, so the main app does not need to be open. Each task repeats until marked complete.
- `.github/workflows/build-daymark.yml`: builds an Android debug APK and a Windows EXE on GitHub-hosted runners. Android Studio is not needed on the laptop.

## Build without Android Studio

1. Commit and push the source and workflow to the repository. This repository is public, so source files will be visible after pushing; no reminder data is included in the source.
2. In GitHub, open **Actions → Build Daymark apps → Run workflow**. A push touching `native/daymark/**` also starts the build automatically.
3. When both jobs finish, open the run and download `Daymark-Android-APK` and/or `Daymark-Windows-EXE` from **Artifacts**. Artifacts expire according to GitHub's retention setting.

The APK from this workflow is a debug build for personal installation. A Play Store release needs a separately signed release bundle, store listing, and Play Console review/testing.

## Android first run

Install `app-debug.apk`, open Daymark, and allow notifications. Then tap **Allow background reminders** and allow **Alarms & reminders** in Android settings. If the phone's battery settings restrict background activity, allow Daymark to run without battery restrictions. The exact alarm is stored locally on this phone; the laptop has its own independent list.

Android may defer sound while the device is powered off. Force-stopping the app or revoking exact-alarm permission can prevent alarms until Daymark is opened and permission is restored.

## Windows first run

Extract/run `Daymark.exe`. Enable **Start Daymark when I sign in to Windows** if you want the tray app ready after sign-in. Closing the window hides it to the notification area; choose **Exit** from the tray icon to stop the manager. The alarm itself is a Windows scheduled task and opens a small alert window at the set time. Data is saved in `%LOCALAPPDATA%\Daymark\reminders.json`.

If Windows is asleep or powered off at the scheduled time, it cannot play the beep then. It may deliver a missed scheduled task when the device wakes, depending on Windows power/task settings.

## Current reminder fields

Task title, first reminder date/time, repeat interval in minutes/hours/days, and beep duration (1–600 seconds). Marking a task complete cancels future reminders on that device.
