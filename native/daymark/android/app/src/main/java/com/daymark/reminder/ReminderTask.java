package com.daymark.reminder;

import org.json.JSONException;
import org.json.JSONObject;

public class ReminderTask {
    public String id;
    public String title;
    public long nextAt;
    public long repeatMillis;
    public int soundSeconds;
    public boolean completed;

    public ReminderTask(String id, String title, long nextAt, long repeatMillis, int soundSeconds, boolean completed) {
        this.id = id;
        this.title = title;
        this.nextAt = nextAt;
        this.repeatMillis = repeatMillis;
        this.soundSeconds = soundSeconds;
        this.completed = completed;
    }

    public JSONObject toJson() throws JSONException {
        JSONObject json = new JSONObject();
        json.put("id", id);
        json.put("title", title);
        json.put("nextAt", nextAt);
        json.put("repeatMillis", repeatMillis);
        json.put("soundSeconds", soundSeconds);
        json.put("completed", completed);
        return json;
    }

    public static ReminderTask fromJson(JSONObject json) throws JSONException {
        return new ReminderTask(json.getString("id"), json.getString("title"), json.getLong("nextAt"),
                json.getLong("repeatMillis"), json.getInt("soundSeconds"), json.optBoolean("completed", false));
    }
}
