# Daymark task reminders

Daymark saves reminders locally in the browser on each device. There is no account or cloud sync: the phone and laptop keep separate task lists.

## Reminder behaviour and limits

- A task repeats at its chosen interval until you mark it complete.
- The sound length is configurable from 1 to 600 seconds. The in-app sound is generated locally and runs while the app page is active and allowed to play audio.
- Browser notifications and background timers are controlled by Android/Windows and the browser. A browser may suspend timers when the app is fully closed or the device is in battery-saving mode. The configurable long sound is not guaranteed while the app is closed; enable notifications and keep the app installed/open for the most reliable alerts.
- Installing the app separately on the phone and laptop gives each device its own local reminder list.

## Run the app

On a laptop, double-click `index.html` to open the app in a browser. To install it on Android, this folder must be hosted at an HTTPS website; open that address in Chrome and choose **Install app** / **Add to Home screen**. The hosting only provides the app files; task data stays in each device's browser storage. The app can use the same website address on both devices while keeping their task lists separate.
