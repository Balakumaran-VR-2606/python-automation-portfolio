package com.daymark.reminder;

import android.content.BroadcastReceiver;
import android.content.Context;
import android.content.Intent;
import android.os.Build;

public class AlarmReceiver extends BroadcastReceiver {
    @Override public void onReceive(Context context, Intent intent) {
        String id = intent.getStringExtra("taskId");
        if (id == null) return;
        ReminderTask task = TaskStore.find(context, id);
        if (task == null || task.completed) return;
        TaskStore.advanceAndSchedule(context, id);
        Intent sound = new Intent(context, AlarmSoundService.class)
                .setAction(AlarmSoundService.ACTION_RING)
                .putExtra("taskId", task.id)
                .putExtra("title", task.title)
                .putExtra("seconds", task.soundSeconds);
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.O) context.startForegroundService(sound);
        else context.startService(sound);
    }
}
