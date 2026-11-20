/**
 * Family Monthly Dashboard & Planner - Google Apps Script Backend
 * Connects Google Sheets (Meals, Lists, Settings) + Shared Google Calendar
 * Deploy as Web App: Execute as "Me", Who has access: "Anyone" (secured via API_SECRET)
 */

const API_SECRET = "Family@Mazurka30"; // <-- Set a private secret token here and in the app Settings!
const DEFAULT_CALENDAR_ID = "simarwalia09@gmail.com"; // Shared calendar "Yash and Sim <3"

function doGet(e) {
  return handleRequest(e, "GET");
}

function doPost(e) {
  return handleRequest(e, "POST");
}

function handleRequest(e, method) {
  try {
    const params = e.parameter || {};
    const body = (method === "POST" && e.postData && e.postData.contents)
      ? JSON.parse(e.postData.contents)
      : {};
    const token = params.token || body.token || "";

    if (API_SECRET !== "CHANGE_ME_TO_YOUR_FAMILY_PIN_2026" && token !== API_SECRET) {
      return jsonResponse({ ok: false, error: "Unauthorized: Invalid API token" });
    }

    const action = params.action || body.action || "sync";
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    if (action === "sync") {
      runAutoCleanup(ss);
      const settings = getSettingsMap(ss);
      const calIdsStr = settings.calendar_ids || DEFAULT_CALENDAR_ID;
      const events = fetchCalendarEvents(calIdsStr);
      const meals = sheetToObjects(ss.getSheetByName("Meal_Pool"));
      const schedule = sheetToObjects(ss.getSheetByName("Meal_Schedule"));
      const lists = sheetToObjects(ss.getSheetByName("Lists"));

      return jsonResponse({
        ok: true,
        serverTime: new Date().toISOString(),
        settings: settings,
        events: events,
        meals: meals,
        schedule: schedule,
        lists: lists
      });
    }

    if (action === "addEvent") {
      const calId = body.calendarId || DEFAULT_CALENDAR_ID;
      const cal = CalendarApp.getCalendarById(calId) || CalendarApp.getDefaultCalendar();
      const start = new Date(body.start);
      const end = body.end ? new Date(body.end) : new Date(start.getTime() + 60 * 60 * 1000);
      const ev = body.allDay
        ? cal.createAllDayEvent(body.title, start, { location: body.location || "", description: body.description || "" })
        : cal.createEvent(body.title, start, end, { location: body.location || "", description: body.description || "" });
      return jsonResponse({ ok: true, id: ev.getId() });
    }

    if (action === "addListItem") {
      const sheet = ss.getSheetByName("Lists");
      const id = (body.listType === "grocery" ? "G" : "T") + Date.now();
      sheet.appendRow([
        id,
        body.listType || "todo",
        body.text,
        body.category || "General",
        "FALSE",
        "",
        body.addedBy || "Family"
      ]);
      return jsonResponse({ ok: true, id: id });
    }

    if (action === "toggleListItem") {
      const sheet = ss.getSheetByName("Lists");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(body.id)) {
          const isDone = Boolean(body.done);
          sheet.getRange(i + 1, 5).setValue(isDone ? "TRUE" : "FALSE");
          sheet.getRange(i + 1, 6).setValue(isDone ? new Date().toISOString() : "");
          break;
        }
      }
      return jsonResponse({ ok: true });
    }

    if (action === "deleteListItem") {
      const sheet = ss.getSheetByName("Lists");
      const data = sheet.getDataRange().getValues();
      for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(body.id)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
      return jsonResponse({ ok: true });
    }

    if (action === "cleanupNow") {
      const removed = runAutoCleanup(ss, true);
      return jsonResponse({ ok: true, removed: removed });
    }

    if (action === "saveSchedule") {
      const sheet = ss.getSheetByName("Meal_Schedule");
      const rows = body.schedule || [];
      sheet.clearContents();
      sheet.appendRow(["Date", "Breakfast_ID", "Breakfast_Name", "Lunch_ID", "Lunch_Name", "Dinner_ID", "Dinner_Name"]);
      rows.forEach(r => {
        sheet.appendRow([
          r.Date,
          r.Breakfast_ID || "",
          r.Breakfast_Name || "",
          r.Lunch_ID || "",
          r.Lunch_Name || "",
          r.Dinner_ID || "",
          r.Dinner_Name || ""
        ]);
      });
      return jsonResponse({ ok: true });
    }

    if (action === "addMeal") {
      const sheet = ss.getSheetByName("Meal_Pool");
      const isBreakfast = String(body.Cuisine || "").toLowerCase() === "breakfast" || String(body.Type || "").toLowerCase() === "breakfast";
      const prefix = isBreakfast ? "B" : "M";
      const id = body.ID || (prefix + Math.floor(100 + Math.random() * 900));
      sheet.appendRow([
        id,
        body.Name,
        body.Cuisine || "Indian",
        body.Type || "Quick (<25m)",
        Number(body.Prep_Mins) || 25,
        body.Ingredients || "",
        "TRUE"
      ]);
      return jsonResponse({ ok: true, id: id });
    }

    if (action === "toggleMealActive") {
      const sheet = ss.getSheetByName("Meal_Pool");
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(body.id)) {
          sheet.getRange(i + 1, 7).setValue(Boolean(body.active) ? "TRUE" : "FALSE");
          break;
        }
      }
      return jsonResponse({ ok: true });
    }

    if (action === "saveSettings") {
      const sheet = ss.getSheetByName("Settings");
      const updates = body.settings || {};
      const data = sheet.getDataRange().getValues();
      const existing = {};
      for (let i = 1; i < data.length; i++) {
        existing[data[i][0]] = i + 1;
      }
      Object.keys(updates).forEach(k => {
        if (existing[k]) {
          sheet.getRange(existing[k], 2).setValue(String(updates[k]));
        } else {
          sheet.appendRow([k, String(updates[k])]);
        }
      });
      return jsonResponse({ ok: true });
    }

    return jsonResponse({ ok: false, error: "Unknown action: " + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function fetchCalendarEvents(calIdsStr) {
  const ids = calIdsStr.split(",").map(s => s.trim()).filter(Boolean);
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const startWindow = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const endWindow = new Date(now.getFullYear(), now.getMonth() + 2, 7);
  const allEvents = [];

  ids.forEach((calId, idx) => {
    try {
      const cal = CalendarApp.getCalendarById(calId);
      if (!cal) return;
      const calName = cal.getName() || calId;
      const events = cal.getEvents(startWindow, endWindow);

      events.forEach(ev => {
        const isAllDay = ev.isAllDayEvent();
        const startTime = ev.getStartTime();
        const endTime = ev.getEndTime();

        const startDateStr = Utilities.formatDate(startTime, tz, "yyyy-MM-dd");
        const effectiveEnd = new Date(endTime.getTime() - (isAllDay ? 1 : 0));
        const endDateStr = Utilities.formatDate(effectiveEnd, tz, "yyyy-MM-dd");

        let cursor = new Date(startTime.getFullYear(), startTime.getMonth(), startTime.getDate(), 12, 0, 0);
        const lastDay = new Date(effectiveEnd.getFullYear(), effectiveEnd.getMonth(), effectiveEnd.getDate(), 12, 0, 0);

        let dayNum = 1;
        const totalDays = Math.round((lastDay - cursor) / (24 * 60 * 60 * 1000)) + 1;

        while (cursor <= lastDay) {
          const dayStr = Utilities.formatDate(cursor, tz, "yyyy-MM-dd");
          const isFirstDay = (dayStr === startDateStr);
          const isLastDay = (dayStr === endDateStr);

          const startStr = isAllDay
            ? dayStr + "T00:00:00"
            : (isFirstDay ? Utilities.formatDate(startTime, tz, "yyyy-MM-dd'T'HH:mm:ss") : dayStr + "T00:00:00");

          const endStr = isAllDay
            ? dayStr + "T23:59:59"
            : (isLastDay ? Utilities.formatDate(endTime, tz, "yyyy-MM-dd'T'HH:mm:ss") : dayStr + "T23:59:59");

          const dayLabel = totalDays > 1 ? ` (Day ${dayNum}/${totalDays})` : "";

          allEvents.push({
            id: ev.getId() + "_" + dayStr,
            calendarId: calId,
            calendarName: calName,
            colorIndex: idx,
            title: ev.getTitle() + dayLabel,
            start: startStr,
            end: endStr,
            allDay: isAllDay || !isFirstDay,
            location: ev.getLocation() || ""
          });

          cursor.setDate(cursor.getDate() + 1);
          dayNum++;
        }
      });
    } catch (e) {
      // Skip inaccessible calendar
    }
  });

  allEvents.sort((a, b) => a.start.localeCompare(b.start));
  return allEvents;
}

function runAutoCleanup(ss, forceAllDone) {
  const settings = getSettingsMap(ss);
  const mode = settings.cleanup_mode || "midnight";
  if (!forceAllDone && mode === "manual") return 0;

  const sheet = ss.getSheetByName("Lists");
  const data = sheet.getDataRange().getValues();
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  let removed = 0;

  for (let i = data.length - 1; i >= 1; i--) {
    const isDone = String(data[i][4]).toUpperCase() === "TRUE";
    if (!isDone) continue;

    if (forceAllDone) {
      sheet.deleteRow(i + 1);
      removed++;
      continue;
    }

    const completedAtStr = data[i][5];
    if (!completedAtStr) continue;
    const completedTime = new Date(completedAtStr).getTime();

    if (mode === "midnight" && completedTime < startOfToday) {
      sheet.deleteRow(i + 1);
      removed++;
    } else if (mode === "24h" && (now.getTime() - completedTime) >= 24 * 60 * 60 * 1000) {
      sheet.deleteRow(i + 1);
      removed++;
    }
  }
  return removed;
}

function sheetToObjects(sheet) {
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      obj[h] = val;
    });
    return obj;
  });
}

function getSettingsMap(ss) {
  const rows = sheetToObjects(ss.getSheetByName("Settings"));
  const map = {};
  rows.forEach(r => { if (r.Key) map[r.Key] = String(r.Value); });
  return map;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}