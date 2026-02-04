/**
 * REFINED BACKEND: Code.gs
 * Ensures Rooms and Reservations are treated as separate data entities.
 */

const SPREADSHEET_ID = '1N5KV1UADGGNF7jhnwQKn3A5DAd9m2E5BeKjSQjmOptA';

function doGet(e) {
  const userEmail = Session.getActiveUser().getEmail();
  const userData = getUserData(userEmail);
  const template = HtmlService.createTemplateFromFile('index');
  template.userEmail = userEmail;
  template.userRole = userData ? userData.role : 'Participant';
  template.userName = userData ? userData.name : userEmail.split('@')[0];
  return template.evaluate()
    .setTitle('Enterprise Hub')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initializeSystem() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = {
    'Users': ['Email', 'Name', 'Role', 'Department'],
    'Rooms': ['RoomID', 'RoomName', 'Capacity', 'Location', 'Notes', 'Status'],
    'RoomReservations': ['ReservationID', 'EmployeeID', 'UserEmail', 'RoomID', 'Date', 'StartTime', 'EndTime', 'Status'],
    'Schedules': ['ID', 'HostEmail', 'ResourceName', 'Details', 'Date', 'Category', 'Status'],
    'TimeSlots': ['SlotID', 'ScheduleID', 'StartTime', 'EndTime', 'Status'],
    'Appointments': ['AppointmentID', 'HostID', 'UserEmail', 'SlotID', 'Status']
  };
  for (let n in sheets) {
    if (!ss.getSheetByName(n)) ss.insertSheet(n);
    let s = ss.getSheetByName(n);
    s.clear().getRange(1, 1, 1, sheets[n].length).setValues([sheets[n]]).setFontWeight('bold').setBackground('#f8fafc');
  }
  return "System Initialized.";
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const userEmail = Session.getActiveUser().getEmail();
    
    const fetch = (name) => {
      const s = ss.getSheetByName(name);
      return (s && s.getLastRow() > 0) ? s.getDataRange().getValues() : [];
    };

    const rRaw = fetch('Rooms');
    const resRaw = fetch('RoomReservations');
    const slotRaw = fetch('TimeSlots');
    const schRaw = fetch('Schedules');
    const aptRaw = fetch('Appointments');

    // 1. MASTER ROOM LIST (Independent of bookings)
    const rooms = rRaw.length > 1 ? rRaw.slice(1).map(r => ({
      id: String(r[0]), name: r[1], capacity: r[2], location: r[3], notes: r[4], status: r[5] || 'Active'
    })) : [];

    // 2. RESERVATIONS (Used only for blocking time)
    const roomReservations = resRaw.length > 1 ? resRaw.slice(1).map(r => ({
      id: r[0], email: r[2], roomId: String(r[3]), 
      date: r[4] ? Utilities.formatDate(new Date(r[4]), "GMT", "yyyy-MM-dd") : '', 
      start: r[5], end: r[6], status: r[7]
    })).filter(r => r.status !== 'Cancelled') : [];

    // 3. APPOINTMENT SLOTS
    const appointmentSlots = slotRaw.length > 1 ? slotRaw.slice(1).map(slot => {
      const sch = schRaw.find(s => s[0] === slot[1]);
      return {
        slotId: slot[0], resourceName: sch ? sch[2] : 'Host', 
        date: sch ? Utilities.formatDate(new Date(sch[4]), "GMT", "yyyy-MM-dd") : '',
        start: slot[2], end: slot[3], status: slot[4]
      };
    }) : [];

    // 4. MY REQUESTS
    const myApts = [
      ...(resRaw.length > 1 ? resRaw.slice(1).filter(r => r[2] === userEmail).map(r => ({
        id: r[0], resource: `Room: ${r[3]}`, 
        time: r[4] ? `${Utilities.formatDate(new Date(r[4]), "GMT", "MMM dd")} (${r[5]} - ${r[6]})` : 'TBD', 
        category: 'Room Reservation'
      })) : []),
      ...(aptRaw.length > 1 ? aptRaw.slice(1).filter(a => a[2] === userEmail).map(a => {
        const slot = slotRaw.find(s => s[0] === a[3]);
        return {
          id: a[0], resource: a[1], time: slot ? new Date(slot[2]).toLocaleString() : 'Confirmed', category: 'Host Appointment'
        };
      }) : [])
    ];

    return { rooms, roomReservations, appointmentSlots, myApts };
  } catch (e) {
    throw new Error("Critical Data Fetch Error: " + e.message);
  }
}

function toggleRoomStatus(roomId, newStatus) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName('Rooms');
  const data = sheet.getDataRange().getValues();
  const idx = data.findIndex(r => String(r[0]).trim() === String(roomId).trim());
  if (idx > -1) {
    sheet.getRange(idx + 1, 6).setValue(newStatus);
    return { success: true };
  }
  throw new Error("Room ID " + roomId + " not found.");
}

function bookRoom(roomId, date, startTime, endTime) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('RoomReservations');
    const existing = sheet.getLastRow() > 1 ? sheet.getDataRange().getValues().slice(1) : [];
    
    const newStart = new Date(date + 'T' + startTime);
    const newEnd = new Date(date + 'T' + endTime);

    const collision = existing.some(r => {
      if (String(r[3]) === String(roomId) && r[7] !== 'Cancelled') {
        const exDate = Utilities.formatDate(new Date(r[4]), "GMT", "yyyy-MM-dd");
        if (exDate === date) {
          const exStart = new Date(date + 'T' + r[5]);
          const exEnd = new Date(date + 'T' + r[6]);
          return (newStart < exEnd && newEnd > exStart);
        }
      }
      return false;
    });

    if (collision) throw new Error("This room is already reserved for the selected time.");

    sheet.appendRow([Utilities.getUuid(), "EmpID", Session.getActiveUser().getEmail(), roomId, date, startTime, endTime, 'Booked']);
    return { success: true };
  } finally { lock.releaseLock(); }
}

function bookAppointment(slotId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const slotSheet = ss.getSheetByName('TimeSlots');
  const slots = slotSheet.getDataRange().getValues();
  const idx = slots.findIndex(s => s[0] === slotId);
  const sched = ss.getSheetByName('Schedules').getDataRange().getValues().find(s => s[0] === slots[idx][1]);
  slotSheet.getRange(idx + 1, 5).setValue('Booked');
  ss.getSheetByName('Appointments').appendRow([Utilities.getUuid(), sched[2], Session.getActiveUser().getEmail(), slotId, "Confirmed"]);
  return { success: true };
}

function addRoom(name, capacity, location, notes) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const trimmedName = String(name).trim();
  // Using Room Name as ID if UUID is too complex for matching
  ss.getSheetByName('Rooms').appendRow([trimmedName, trimmedName, capacity, location, notes, 'Active']);
  return { success: true };
}

function createAppointmentSlots(hostName, date, start, end, duration) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const schedId = "SCH-" + Utilities.getUuid();
  ss.getSheetByName('Schedules').appendRow([schedId, Session.getActiveUser().getEmail(), hostName, "", date, "Appointment", "Active"]);
  let curr = new Date(date + 'T' + start);
  const stop = new Date(date + 'T' + end);
  while (curr < stop) {
    let next = new Date(curr.getTime() + duration * 60000);
    if (next > stop) break;
    ss.getSheetByName('TimeSlots').appendRow([Utilities.getUuid(), schedId, curr.toISOString(), next.toISOString(), 'Available']);
    curr = next;
  }
  return { success: true };
}

function getUserData(email) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data = ss.getSheetByName('Users').getDataRange().getValues();
    const user = data.find(r => r[0] === email);
    return user ? { email: user[0], name: user[1], role: user[2] } : null;
  } catch (e) { return null; }
}
