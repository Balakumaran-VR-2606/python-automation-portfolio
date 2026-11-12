package com.daymark.reminder;

import android.Manifest;
import android.app.Activity;
import android.app.AlertDialog;
import android.app.DatePickerDialog;
import android.app.TimePickerDialog;
import android.content.Intent;
import android.graphics.Color;
import android.graphics.Typeface;
import android.graphics.drawable.GradientDrawable;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.provider.Settings;
import android.view.Gravity;
import android.view.View;
import android.view.ViewGroup;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.EditText;
import android.widget.LinearLayout;
import android.widget.ScrollView;
import android.widget.Spinner;
import android.widget.TextView;
import android.widget.Toast;

import java.text.DateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Locale;
import java.util.UUID;

public class MainActivity extends Activity {
    private static final int BG = Color.rgb(13, 21, 34);
    private static final int CARD = Color.rgb(20, 31, 48);
    private static final int FIELD = Color.rgb(15, 25, 40);
    private static final int LINE = Color.rgb(42, 57, 80);
    private static final int TEXT = Color.rgb(242, 242, 237);
    private static final int MUTED = Color.rgb(161, 173, 189);
    private static final int MINT = Color.rgb(156, 229, 195);
    private EditText titleInput, repeatInput, soundInput;
    private Spinner unitSpinner;
    private TextView startLabel, alarmPermissionLabel, taskCount;
    private LinearLayout taskRows;
    private long startAt;
    private boolean showCompleted = false;

    @Override protected void onCreate(Bundle state) {
        super.onCreate(state);
        startAt = System.currentTimeMillis() + 2 * 60_000L;
        requestNotificationPermission();
        buildUi();
        refresh();
    }

    @Override protected void onResume() {
        super.onResume();
        TaskStore.scheduleAll(this);
        updatePermissionLabel();
        if (taskRows != null) refresh();
    }

    private void buildUi() {
        ScrollView scroll = new ScrollView(this);
        scroll.setFillViewport(true);
        scroll.setBackgroundColor(BG);
        LinearLayout page = new LinearLayout(this);
        page.setOrientation(LinearLayout.VERTICAL);
        page.setPadding(dp(20), dp(16), dp(20), dp(28));
        scroll.addView(page);

        LinearLayout brand = new LinearLayout(this);
        brand.setGravity(Gravity.CENTER_VERTICAL);
        TextView mark = text("d", 18, Color.rgb(16, 37, 30), true);
        mark.setGravity(Gravity.CENTER);
        mark.setBackground(round(MINT, 10));
        brand.addView(mark, new LinearLayout.LayoutParams(dp(32), dp(32)));
        TextView brandName = text("  daymark", 19, TEXT, true);
        brand.addView(brandName);
        TextView local = text("Saved on this phone", 12, MUTED, false);
        local.setGravity(Gravity.RIGHT | Gravity.CENTER_VERTICAL);
        brand.addView(local, new LinearLayout.LayoutParams(0, dp(36), 1));
        page.addView(brand);

        addSpace(page, 28);
        page.addView(text("MAKE ROOM FOR WHAT MATTERS", 10, MINT, true));
        TextView headline = text("Keep the things\nyou mean to do.", 33, TEXT, true);
        headline.setLineSpacing(dp(2), 1f);
        page.addView(headline);
        TextView sub = text("Set a rhythm. Reminders repeat until you mark them done.", 13, MUTED, false);
        sub.setPadding(0, dp(8), 0, dp(20));
        page.addView(sub);

        LinearLayout form = card();
        form.setPadding(dp(16), dp(16), dp(16), dp(16));
        page.addView(form, matchWrap());
        form.addView(text("A NEW REMINDER", 10, MINT, true));
        TextView prompt = text("What needs doing?", 20, TEXT, true);
        prompt.setPadding(0, dp(5), 0, dp(15));
        form.addView(prompt);

        titleInput = input("Task name, e.g. Take medicine", false);
        form.addView(titleInput, matchWrap());
        addSpace(form, 11);
        startLabel = text("Starts: " + formatDate(startAt), 13, TEXT, false);
        Button timeButton = button("Choose start date and time", false);
        timeButton.setOnClickListener(v -> pickDate());
        form.addView(timeButton, matchWrap());
        form.addView(startLabel);
        addSpace(form, 12);

        TextView repeatHeading = text("Repeat every", 12, TEXT, true);
        form.addView(repeatHeading);
        LinearLayout repeatRow = new LinearLayout(this);
        repeatInput = input("2", true);
        repeatInput.setText("2");
        repeatRow.addView(repeatInput, new LinearLayout.LayoutParams(0, dp(48), 1));
        unitSpinner = new Spinner(this);
        String[] units = {"hours", "minutes", "days"};
        ArrayAdapter<String> adapter = new ArrayAdapter<>(this, android.R.layout.simple_spinner_item, units);
        adapter.setDropDownViewResource(android.R.layout.simple_spinner_dropdown_item);
        unitSpinner.setAdapter(adapter);
        unitSpinner.setBackground(round(FIELD, 9));
        LinearLayout.LayoutParams sp = new LinearLayout.LayoutParams(0, dp(48), 1);
        sp.setMargins(dp(8), 0, 0, 0);
        repeatRow.addView(unitSpinner, sp);
        form.addView(repeatRow, matchWrap());
        addSpace(form, 12);

        form.addView(text("Sound duration (seconds)", 12, TEXT, true));
        soundInput = input("20", true);
        soundInput.setText("20");
        form.addView(soundInput, matchWrap());
        addSpace(form, 15);
        TextView keepGoing = text("↻  Keeps repeating until you mark it done", 11, MUTED, false);
        form.addView(keepGoing);
        Button add = button("＋   Add reminder", true);
        add.setOnClickListener(v -> addTask());
        LinearLayout.LayoutParams addLp = matchWrap();
        addLp.topMargin = dp(12);
        form.addView(add, addLp);

        addSpace(page, 18);
        LinearLayout permissions = card();
        permissions.setPadding(dp(14), dp(12), dp(14), dp(12));
        alarmPermissionLabel = text("Checking alert permissions…", 12, MUTED, false);
        permissions.addView(alarmPermissionLabel, matchWrap());
        Button enable = button("Allow background reminders", false);
        enable.setOnClickListener(v -> openAlarmPermission());
        LinearLayout.LayoutParams enableLp = matchWrap();
        enableLp.topMargin = dp(8);
        permissions.addView(enable, enableLp);
        page.addView(permissions, matchWrap());

        addSpace(page, 24);
        LinearLayout listHeader = new LinearLayout(this);
        listHeader.setGravity(Gravity.CENTER_VERTICAL);
        listHeader.addView(text("REMINDERS", 10, MINT, true));
        taskCount = text("", 12, MUTED, false);
        taskCount.setGravity(Gravity.RIGHT | Gravity.CENTER_VERTICAL);
        listHeader.addView(taskCount, new LinearLayout.LayoutParams(0, dp(28), 1));
        page.addView(listHeader, matchWrap());

        LinearLayout tabs = new LinearLayout(this);
        Button activeTab = button("In progress", false);
        Button completedTab = button("Completed", false);
        activeTab.setOnClickListener(v -> { showCompleted = false; refresh(); });
        completedTab.setOnClickListener(v -> { showCompleted = true; refresh(); });
        tabs.addView(activeTab, new LinearLayout.LayoutParams(0, dp(42), 1));
        tabs.addView(completedTab, new LinearLayout.LayoutParams(0, dp(42), 1));
        page.addView(tabs);
        taskRows = new LinearLayout(this);
        taskRows.setOrientation(LinearLayout.VERTICAL);
        page.addView(taskRows, matchWrap());
        setContentView(scroll);
        updatePermissionLabel();
    }

    private void addTask() {
        String title = titleInput.getText().toString().trim();
        int repeat;
        int sound;
        try { repeat = Integer.parseInt(repeatInput.getText().toString().trim()); }
        catch (NumberFormatException ex) { repeat = 0; }
        try { sound = Integer.parseInt(soundInput.getText().toString().trim()); }
        catch (NumberFormatException ex) { sound = 0; }
        if (title.isEmpty()) { toast("Enter a task name."); return; }
        if (repeat < 1 || repeat > 999) { toast("Repeat interval must be 1–999."); return; }
        if (sound < 1 || sound > 600) { toast("Sound duration must be 1–600 seconds."); return; }
        if (startAt <= System.currentTimeMillis()) { toast("Choose a future start time."); return; }
        long multiplier;
        switch (unitSpinner.getSelectedItem().toString()) {
            case "minutes": multiplier = 60_000L; break;
            case "days": multiplier = 86_400_000L; break;
            default: multiplier = 3_600_000L;
        }
        ReminderTask task = new ReminderTask(UUID.randomUUID().toString(), title, startAt, repeat * multiplier, sound, false);
        TaskStore.add(this, task);
        titleInput.setText("");
        Toast.makeText(this, TaskStore.canScheduleExact(this) ? "Reminder saved on this phone." : "Saved. Allow exact alarms to receive alerts.", Toast.LENGTH_LONG).show();
        if (!TaskStore.canScheduleExact(this)) openAlarmPermission();
        refresh();
    }

    private void pickDate() {
        Calendar c = Calendar.getInstance(); c.setTimeInMillis(startAt);
        new DatePickerDialog(this, (view, y, m, d) -> {
            Calendar chosen = Calendar.getInstance();
            chosen.set(y, m, d, 0, 0, 0); chosen.set(Calendar.MILLISECOND, 0);
            pickTime(chosen);
        }, c.get(Calendar.YEAR), c.get(Calendar.MONTH), c.get(Calendar.DAY_OF_MONTH)).show();
    }

    private void pickTime(Calendar chosen) {
        Calendar current = Calendar.getInstance();
        new TimePickerDialog(this, (view, h, m) -> {
            chosen.set(Calendar.HOUR_OF_DAY, h); chosen.set(Calendar.MINUTE, m);
            startAt = chosen.getTimeInMillis();
            startLabel.setText("Starts: " + formatDate(startAt));
        }, current.get(Calendar.HOUR_OF_DAY), current.get(Calendar.MINUTE), false).show();
    }

    private void refresh() {
        if (taskRows == null) return;
        taskRows.removeAllViews();
        List<ReminderTask> tasks = TaskStore.all(this);
        int active = 0;
        for (ReminderTask task : tasks) if (!task.completed) active++;
        taskCount.setText(active + " active");
        boolean any = false;
        for (ReminderTask task : tasks) {
            if (task.completed != showCompleted) continue;
            any = true;
            taskRows.addView(taskRow(task), matchWrap());
        }
        if (!any) {
            TextView empty = text(showCompleted ? "No completed tasks yet." : "No reminders yet. Add one above.", 13, MUTED, false);
            empty.setPadding(0, dp(18), 0, dp(18));
            taskRows.addView(empty);
        }
        updatePermissionLabel();
    }

    private View taskRow(ReminderTask task) {
        LinearLayout row = card();
        row.setPadding(dp(13), dp(13), dp(13), dp(10));
        TextView name = text(task.title, 15, TEXT, true);
        row.addView(name);
        TextView detail = text((task.completed ? "Completed" : "Next · " + formatDate(task.nextAt))
                + "\nRepeats every " + repeatLabel(task.repeatMillis) + " · " + task.soundSeconds + "s sound", 11, MUTED, false);
        detail.setPadding(0, dp(5), 0, dp(9));
        row.addView(detail);
        LinearLayout buttons = new LinearLayout(this);
        buttons.setGravity(Gravity.RIGHT);
        if (!task.completed) {
            Button done = button("Mark done", true);
            done.setOnClickListener(v -> new AlertDialog.Builder(this).setTitle("Mark completed?")
                    .setMessage("This task's repeating reminders will stop on this phone.")
                    .setNegativeButton("Cancel", null).setPositiveButton("Mark done", (d, w) -> {
                        TaskStore.setCompleted(this, task.id); refresh();
                    }).show());
            buttons.addView(done);
        }
        Button delete = button("Delete", false);
        delete.setOnClickListener(v -> new AlertDialog.Builder(this).setTitle("Delete reminder?")
                .setMessage("This only removes it from this phone.")
                .setNegativeButton("Cancel", null).setPositiveButton("Delete", (d, w) -> {
                    TaskStore.delete(this, task.id); refresh();
                }).show());
        LinearLayout.LayoutParams deleteLp = new LinearLayout.LayoutParams(ViewGroup.LayoutParams.WRAP_CONTENT, dp(42));
        deleteLp.leftMargin = dp(7);
        buttons.addView(delete, deleteLp);
        row.addView(buttons);
        return row;
    }

    private void updatePermissionLabel() {
        if (alarmPermissionLabel == null) return;
        boolean exact = TaskStore.canScheduleExact(this);
        boolean notif = Build.VERSION.SDK_INT < 33 || checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS) == android.content.pm.PackageManager.PERMISSION_GRANTED;
        if (exact && notif) alarmPermissionLabel.setText("Background alarms and notifications are enabled.");
        else if (!exact) alarmPermissionLabel.setText("Allow Alarms & reminders so exact-time beeps can run in the background.");
        else alarmPermissionLabel.setText("Allow notifications to see reminders on your lock screen.");
    }

    private void openAlarmPermission() {
        if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.S && !TaskStore.canScheduleExact(this)) {
            new AlertDialog.Builder(this).setTitle("Allow exact reminders")
                    .setMessage("Android needs Alarms & reminders permission to ring at the chosen time while Daymark is closed.")
                    .setNegativeButton("Later", null).setPositiveButton("Open settings", (d, w) -> {
                        Intent settings = new Intent(Settings.ACTION_REQUEST_SCHEDULE_EXACT_ALARM,
                                Uri.parse("package:" + getPackageName()));
                        startActivity(settings);
                    }).show();
        } else if (Build.VERSION.SDK_INT >= 33 && checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS) != android.content.pm.PackageManager.PERMISSION_GRANTED) {
            requestPermissions(new String[]{Manifest.permission.POST_NOTIFICATIONS}, 71);
        }
    }

    private void requestNotificationPermission() {
        if (Build.VERSION.SDK_INT >= 33 && checkSelfPermission(Manifest.permission.POST_NOTIFICATIONS) != android.content.pm.PackageManager.PERMISSION_GRANTED)
            requestPermissions(new String[]{Manifest.permission.POST_NOTIFICATIONS}, 71);
    }

    private String formatDate(long millis) {
        if (millis <= 0) return "just now";
        return DateFormat.getDateTimeInstance(DateFormat.MEDIUM, DateFormat.SHORT, Locale.getDefault()).format(new Date(millis));
    }

    private String repeatLabel(long millis) {
        if (millis % 86_400_000L == 0) return (millis / 86_400_000L) + " day(s)";
        if (millis % 3_600_000L == 0) return (millis / 3_600_000L) + " hour(s)";
        return (millis / 60_000L) + " minute(s)";
    }

    private EditText input(String hint, boolean number) {
        EditText field = new EditText(this);
        field.setSingleLine(true); field.setTextColor(TEXT); field.setHintTextColor(MUTED); field.setHint(hint);
        field.setTextSize(14); field.setPadding(dp(12), 0, dp(12), 0); field.setBackground(round(FIELD, 9));
        if (number) field.setInputType(android.text.InputType.TYPE_CLASS_NUMBER);
        else field.setInputType(android.text.InputType.TYPE_CLASS_TEXT | android.text.InputType.TYPE_TEXT_FLAG_CAP_SENTENCES);
        return field;
    }

    private Button button(String label, boolean primary) {
        Button b = new Button(this); b.setText(label); b.setAllCaps(false); b.setTextSize(12); b.setTypeface(null, Typeface.BOLD);
        b.setTextColor(primary ? Color.rgb(16, 37, 30) : TEXT); b.setBackground(round(primary ? MINT : Color.rgb(26, 41, 61), 9));
        return b;
    }

    private TextView text(String value, int size, int color, boolean bold) {
        TextView t = new TextView(this); t.setText(value); t.setTextSize(size); t.setTextColor(color);
        if (bold) t.setTypeface(null, Typeface.BOLD); return t;
    }

    private LinearLayout card() {
        LinearLayout l = new LinearLayout(this); l.setOrientation(LinearLayout.VERTICAL); l.setBackground(round(CARD, 15)); return l;
    }

    private GradientDrawable round(int color, int radius) {
        GradientDrawable d = new GradientDrawable(); d.setColor(color); d.setCornerRadius(dp(radius));
        if (color == FIELD) d.setStroke(dp(1), LINE); return d;
    }

    private LinearLayout.LayoutParams matchWrap() { return new LinearLayout.LayoutParams(ViewGroup.LayoutParams.MATCH_PARENT, ViewGroup.LayoutParams.WRAP_CONTENT); }
    private int dp(int value) { return (int) (value * getResources().getDisplayMetrics().density + 0.5f); }
    private void addSpace(LinearLayout layout, int height) { View space = new View(this); layout.addView(space, new LinearLayout.LayoutParams(1, dp(height))); }
    private void toast(String message) { Toast.makeText(this, message, Toast.LENGTH_SHORT).show(); }
}
