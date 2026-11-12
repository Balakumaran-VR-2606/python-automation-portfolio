package com.daymark.reminder;

import android.app.Notification;
import android.app.NotificationChannel;
import android.app.NotificationManager;
import android.app.PendingIntent;
import android.app.Service;
import android.content.Intent;
import android.media.AudioManager;
import android.media.ToneGenerator;
import android.os.Build;
import android.os.Handler;
import android.os.IBinder;
import android.os.Looper;

import java.util.HashMap;
import java.util.Map;

public class AlarmSoundService extends Service {
    public static final String ACTION_RING = "com.daymark.reminder.RING";
    public static final String ACTION_STOP = "com.daymark.reminder.STOP";
    private static final String CHANNEL = "daymark_alarm_channel";
    private final Handler handler = new Handler(Looper.getMainLooper());
    private final Map<String, ToneGenerator> tones = new HashMap<>();
    private final Map<String, Runnable> loops = new HashMap<>();
    private int notificationId = 4401;

    @Override public void onCreate() { super.onCreate(); createChannel(); }

    @Override public int onStartCommand(Intent intent, int flags, int startId) {
        if (intent == null) { stopIfIdle(); return START_NOT_STICKY; }
        String action = intent.getAction();
        if (ACTION_STOP.equals(action)) {
            stopTask(intent.getStringExtra("taskId"));
        } else if (ACTION_RING.equals(action)) {
            String id = intent.getStringExtra("taskId");
            String title = intent.getStringExtra("title");
            int seconds = Math.max(1, Math.min(600, intent.getIntExtra("seconds", 20)));
            if (id != null && title != null) {
                startForeground(notificationId, notification(title));
                ring(id, seconds);
            }
        }
        stopIfIdle();
        return START_NOT_STICKY;
    }

    private void ring(String id, int seconds) {
        stopTask(id, false);
        try {
            ToneGenerator tone = new ToneGenerator(AudioManager.STREAM_ALARM, 100);
            tones.put(id, tone);
            long end = System.currentTimeMillis() + seconds * 1000L;
            Runnable loop = new Runnable() {
                @Override public void run() {
                    if (!tones.containsKey(id)) return;
                    if (System.currentTimeMillis() >= end) { stopTask(id); return; }
                    tone.startTone(ToneGenerator.TONE_PROP_BEEP, 350);
                    handler.postDelayed(this, 900);
                }
            };
            loops.put(id, loop);
            loop.run();
        } catch (RuntimeException ignored) { stopTask(id); }
    }

    private void stopTask(String id) {
        stopTask(id, true);
    }

    private void stopTask(String id, boolean stopServiceWhenIdle) {
        if (id == null) return;
        Runnable loop = loops.remove(id);
        if (loop != null) handler.removeCallbacks(loop);
        ToneGenerator tone = tones.remove(id);
        if (tone != null) { tone.stopTone(); tone.release(); }
        if (stopServiceWhenIdle && tones.isEmpty()) {
            stopForeground(STOP_FOREGROUND_REMOVE);
            stopSelf();
        }
    }

    private void stopIfIdle() {
        if (tones.isEmpty()) {
            stopForeground(STOP_FOREGROUND_REMOVE);
            stopSelf();
        }
    }

    private void createChannel() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
            NotificationChannel channel = new NotificationChannel(CHANNEL, "Task reminders", NotificationManager.IMPORTANCE_HIGH);
            channel.setDescription("Shows a notification while a reminder sound is playing");
            channel.setSound(null, null);
            NotificationManager manager = getSystemService(NotificationManager.class);
            if (manager != null) manager.createNotificationChannel(channel);
        }
    }

    private Notification notification(String title) {
        Intent open = new Intent(this, MainActivity.class).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_CLEAR_TOP);
        PendingIntent content = PendingIntent.getActivity(this, 4402, open,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) {
            return new Notification.Builder(this, CHANNEL).setSmallIcon(android.R.drawable.ic_lock_idle_alarm)
                    .setContentTitle("Daymark reminder").setContentText(title).setContentIntent(content)
                    .setCategory(Notification.CATEGORY_ALARM).setOngoing(true).build();
        }
        return new Notification.Builder(this).setSmallIcon(android.R.drawable.ic_lock_idle_alarm)
                .setContentTitle("Daymark reminder").setContentText(title).setContentIntent(content)
                .setCategory(Notification.CATEGORY_ALARM).setOngoing(true).build();
    }

    @Override public void onDestroy() {
        for (String id : loops.keySet().toArray(new String[0])) stopTask(id);
        super.onDestroy();
    }

    @Override public IBinder onBind(Intent intent) { return null; }
}
