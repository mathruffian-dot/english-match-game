var WORDS = [
  ['apple', '蘋果'], ['baby', '嬰兒'], ['bird', '鳥'], ['book', '書'], ['box', '盒子'],
  ['cake', '蛋糕'], ['cat', '貓'], ['chair', '椅子'], ['dog', '狗'], ['door', '門'],
  ['egg', '蛋'], ['eye', '眼睛'], ['fish', '魚'], ['foot', '腳'], ['garden', '花園'],
  ['girl', '女孩'], ['hat', '帽子'], ['house', '房子'], ['ice', '冰'], ['island', '島'],
  ['key', '鑰匙'], ['king', '國王'], ['lamp', '燈'], ['love', '愛'], ['milk', '牛奶'],
  ['moon', '月亮'], ['name', '名字'], ['nose', '鼻子'], ['orange', '柳橙'], ['park', '公園'],
  ['pen', '筆'], ['queen', '女王'], ['rain', '雨'], ['red', '紅色'], ['star', '星星'],
  ['sun', '太陽'], ['table', '桌子'], ['tree', '樹'], ['uncle', '叔叔'], ['village', '村莊'],
  ['water', '水'], ['window', '窗戶'], ['year', '年'], ['yellow', '黃色'], ['zoo', '動物園'],
  ['blue', '藍色'], ['green', '綠色'], ['bread', '麵包'], ['river', '河流'], ['school', '學校']
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('英文單字配對遊戲')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getWords() {
  var shuffled = WORDS.slice();
  for (var i = shuffled.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = shuffled[i];
    shuffled[i] = shuffled[j];
    shuffled[j] = tmp;
  }
  var picked = shuffled.slice(0, 10);
  return picked.map(function(pair, idx) {
    return { id: idx, en: pair[0], zh: pair[1] };
  });
}

function setupSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.clear();
  sheet.appendRow(['時間戳記', '班級', '座號', '分數', '正確數', '錯誤次數', '花費時間(秒)']);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#4a90d9').setFontColor('#ffffff');
  sheet.setColumnWidth(1, 180);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 80);
  sheet.setColumnWidth(7, 120);
}

function getSheetUrl() {
  return SpreadsheetApp.getActiveSpreadsheet().getUrl();
}

function recordResult(className, seat, score, correct, wrong, timeSec) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var timestamp = new Date();
  sheet.appendRow([timestamp, className, parseInt(seat), parseInt(score), parseInt(correct), parseInt(wrong), parseInt(timeSec)]);
  return { success: true };
}

// ==============================
// Email 函式（Agent 用）
// ==============================

function getRecentEmails(count) {
  count = count || 10;
  var threads = GmailApp.getInboxThreads(0, count);
  return threads.map(function(thread) {
    var msg = thread.getMessages()[0];
    return {
      threadId: thread.getId(),
      subject: thread.getFirstMessageSubject(),
      from: msg.getFrom(),
      date: msg.getDate().toISOString(),
      snippet: msg.getPlainBody().substring(0, 200),
      isUnread: thread.isUnread(),
      messageCount: thread.getMessageCount()
    };
  });
}

function getUnreadEmails(count) {
  count = count || 10;
  var threads = GmailApp.search('is:unread', 0, count);
  return threads.map(function(thread) {
    var msg = thread.getMessages()[0];
    return {
      threadId: thread.getId(),
      subject: thread.getFirstMessageSubject(),
      from: msg.getFrom(),
      date: msg.getDate().toISOString(),
      snippet: msg.getPlainBody().substring(0, 200),
      messageCount: thread.getMessageCount()
    };
  });
}

function searchEmails(query, count) {
  count = count || 10;
  var threads = GmailApp.search(query, 0, count);
  return threads.map(function(thread) {
    var msg = thread.getMessages()[0];
    return {
      threadId: thread.getId(),
      subject: thread.getFirstMessageSubject(),
      from: msg.getFrom(),
      date: msg.getDate().toISOString(),
      snippet: msg.getPlainBody().substring(0, 200),
      messageCount: thread.getMessageCount()
    };
  });
}

function getEmailDetail(threadId) {
  var thread = GmailApp.getThreadById(threadId);
  if (!thread) return { error: '找不到該郵件' };
  var messages = thread.getMessages().map(function(msg) {
    return {
      from: msg.getFrom(),
      to: msg.getTo(),
      date: msg.getDate().toISOString(),
      subject: msg.getSubject(),
      body: msg.getPlainBody().substring(0, 3000),
      hasAttachment: msg.getAttachments().length > 0
    };
  });
  return {
    threadId: threadId,
    subject: thread.getFirstMessageSubject(),
    messageCount: thread.getMessageCount(),
    messages: messages
  };
}

// ==============================
// 行事曆函式（Agent 用）
// ==============================

function getUpcomingEvents(days) {
  days = days || 7;
  var now = new Date();
  var end = new Date(now.getTime() + days * 24 * 60 * 60 * 1000);
  var events = CalendarApp.getDefaultCalendar().getEvents(now, end);
  return events.map(function(e) {
    return {
      id: e.getId(),
      title: e.getTitle(),
      startTime: e.getStartTime().toISOString(),
      endTime: e.getEndTime().toISOString(),
      isAllDay: e.isAllDayEvent(),
      location: e.getLocation(),
      description: e.getDescription().substring(0, 500),
      guests: e.getGuestList().map(function(g) { return g.getEmail(); })
    };
  });
}

function getTodayEvents() {
  var now = new Date();
  var start = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
  var end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  var events = CalendarApp.getDefaultCalendar().getEvents(start, end);
  return events.map(function(e) {
    return {
      id: e.getId(),
      title: e.getTitle(),
      startTime: e.getStartTime().toISOString(),
      endTime: e.getEndTime().toISOString(),
      isAllDay: e.isAllDayEvent(),
      location: e.getLocation(),
      description: e.getDescription().substring(0, 300)
    };
  });
}

function getCalendarEvents(startDate, endDate) {
  var start = new Date(startDate);
  var end = new Date(endDate);
  var events = CalendarApp.getDefaultCalendar().getEvents(start, end);
  return events.map(function(e) {
    return {
      id: e.getId(),
      title: e.getTitle(),
      startTime: e.getStartTime().toISOString(),
      endTime: e.getEndTime().toISOString(),
      isAllDay: e.isAllDayEvent(),
      location: e.getLocation(),
      description: e.getDescription().substring(0, 500)
    };
  });
}

function createCalendarEvent(title, startTime, endTime, description, location) {
  var event = CalendarApp.getDefaultCalendar().createEvent(
    title,
    new Date(startTime),
    new Date(endTime),
    { description: description || '', location: location || '' }
  );
  return {
    id: event.getId(),
    title: event.getTitle(),
    startTime: event.getStartTime().toISOString(),
    endTime: event.getEndTime().toISOString(),
    htmlLink: 'https://calendar.google.com/calendar/event?eid=' + event.getId().replace('@google.com', '')
  };
}

// ==============================
//  Agent 儀表板：一次拿摘要
// ==============================

function getAgentDashboard() {
  return {
    todayEvents: getTodayEvents(),
    upcomingEvents: getUpcomingEvents(7),
    unreadEmails: getUnreadEmails(10),
    recentEmails: getRecentEmails(10)
  };
}
