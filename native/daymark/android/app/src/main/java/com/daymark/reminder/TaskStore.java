package com.daymark.reminder;

import android.app.AlarmManager;
import android.app.PendingIntent;
import android.content.Context;
import android.content.Intent;
import android.os.Build;
import android.provider.Settings;

import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

import java.util.ArrayList;
import java.util.List;

public final class TaskStore {
    private static final String PREFS = "daymark_local_tasks";
    private static final String TASKS = "tasks";
    private TaskStore() {}

    public static synchronized List<ReminderTask> all(Context context) {
        ArrayList<ReminderTask> result = new ArrayList<>();
        String raw = context.getSharedPreferences(PREFS, Context.MODE_PRIVATE).getString(TASKS, "[]");
        try {
            JSONArray array = new JSONArray(raw);
            for (int i = 0; i < array.length(); i++) result.add(ReminderTask.fromJson(array.getJSONObject(i)));
        } catch (JSONException ignored) { }
        return result;
    }

    private static synchronized void save(Context context, List<ReminderTask> tasks) {
        JSONArray array = new JSONArray();
        for (ReminderTask task : tasks) {
            try { array.put(task.toJson()); } catch (JSONException ignored) { }
        }
        context.getSharedPreferences(PREFS, Context.MODE_PRIVATE).edit().putString(TASKS, array.toString()).apply();
    }

    public static synchronized void add(Context context, ReminderTask task) {
        List<ReminderTask> tasks = all(context);
        tasks.add(0, task);
        save(context, tasks);
        schedule(context, task);
    }

    public static synchronized ReminderTask find(Context context, String id) {
        for (ReminderTask task : all(context)) if (task.id.equals(id)) return task;
        return null;
    }

    public static synchronized void setCompleted(Context context, String id) {
        List<ReminderTask> tasks = all(context);
        for (ReminderTask task : tasks) if (task.id.equals(id)) task.completed = true;
        save(context, tasks);
        cancel(context, id);
        Intent stop = new Intent(context, AlarmSoundService.class).setAction(AlarmSoundService.ACTION_STOP).putExtra("taskId", id);
        context.startService(stop);
    }

    public static synchronized void delete(Context context, String id) {
        List<ReminderTask> tasks = all(context);
        tasks.removeIf(task -> task.id.equals(id));
        save(context, tasks);
        cancel(context, id);
        Intent stop = new Intent(context, AlarmSoundService.class).setAction(AlarmSoundService.ACTION_STOP).putExtra("taskId", id);
        context.startService(stop);
    }

    public static synchronized void scheduleAll(Context context) {
        for (ReminderTask task : all(context)) if (!task.completed) schedule(context, task);
    }

    public static synchronized void advanceAndSchedule(Context context, String id) {
        List<ReminderTask> tasks = all(context);
        long now = System.currentTimeMillis();
        for (ReminderTask task : tasks) {
            if (task.id.equals(id) && !task.completed) {
                long steps = Math.max(1, ((now - task.nextAt) / task.repeatMillis) + 1);
                task.nextAt += steps * task.repeatMillis;
                save(context, tasks);
                schedule(context, task);
                return;
            }
        }
    }

    public static boolean canScheduleExact(Context context) {
        if (Build.VERSION.SDK_INT < Build.VERSION_CODES.S) return true;
        AlarmManager manager = (AlarmManager) context.getSystemService(Context.ALARM_SERVICE);
        return manager != null && manager.canScheduleExactAlarms();
    }

    public static void schedule(Context context, ReminderTask task) {
        if (task.completed || !canScheduleExact(context)) return;
        AlarmManager manager = (AlarmManager) context.getSystemService(Context.ALARM_SERVICE);
        if (manager == null) return;
        Intent intent = new Intent(context, AlarmReceiver.class).setAction("com.daymark.reminder.ALARM").putExtra("taskId", task.id);
        PendingIntent pending = PendingIntent.getBroadcast(context, task.id.hashCode(), intent,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
        Intent openApp = new Intent(context, MainActivity.class).addFlags(Intent.FLAG_ACTIVITY_NEW_TASK | Intent.FLAG_ACTIVITY_CLEAR_TOP);
        PendingIntent show = PendingIntent.getActivity(context, task.id.hashCode() ^ 0x44D,
                openApp, PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
        long at = Math.max(task.nextAt, System.currentTimeMillis() + 2_000);
        AlarmManager.AlarmClockInfo info = new AlarmManager.AlarmClockInfo(at, show);
        manager.setAlarmClock(info, pending);
    }

    public static void cancel(Context context, String id) {
        AlarmManager manager = (AlarmManager) context.getSystemService(Context.ALARM_SERVICE);
        if (manager == null) return;
        Intent intent = new Intent(context, AlarmReceiver.class).setAction("com.daymark.reminder.ALARM").putExtra("taskId", id);
        PendingIntent pending = PendingIntent.getBroadcast(context, id.hashCode(), intent,
                PendingIntent.FLAG_UPDATE_CURRENT | PendingIntent.FLAG_IMMUTABLE);
        manager.cancel(pending);
        pending.cancel();
    }
}
