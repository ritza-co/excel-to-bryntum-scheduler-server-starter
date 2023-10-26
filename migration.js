import xlsx from "xlsx";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { dirname } from "path";

function excelToJSDate(excelDate) {
  const dateObj = xlsx.SSF.parse_date_code(excelDate);
  return new Date(
    dateObj.y,
    dateObj.m - 1,
    dateObj.d,
    dateObj.H,
    dateObj.M,
    dateObj.S
  );
}

function addDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result.toISOString().slice(0, 10);
}

function createBryntumRows(data) {
  const monthDate = excelToJSDate(data[1]["__EMPTY_2"]);
  let resourcesStore = [];
  let eventsStore = [];
  let resourceTimeRangesStore = [];
  const timeRangesStore = [
    {
      id: 1,
      name: "",
      recurrenceRule: "FREQ=WEEKLY;BYDAY=SA,SU;",
      startDate: addDays(monthDate, 1),
      endDate: addDays(monthDate, 2),
    },
  ];
  const monthDays = data[3];
  const firstResource = "Jane";

  let state = {
    isFirstResourceFound: false,
    resourceId: 0,
    eventId: 0,
    resourceTimeRangesId: 0,
  };

  state = resourcesLoop(
    data,
    monthDate,
    monthDays,
    firstResource,
    state,
    resourcesStore,
    eventsStore,
    resourceTimeRangesStore
  );

  return {
    resourcesStore,
    eventsStore,
    resourceTimeRangesStore,
    timeRangesStore,
  };
}

function resourcesLoop(
  data,
  monthDate,
  monthDays,
  firstResource,
  state,
  resourcesStore,
  eventsStore,
  resourceTimeRangesStore
) {
  for (let i = 0; i < data.length; i++) {
    if (
      data[i].hasOwnProperty("__EMPTY_1") &&
      data[i]["__EMPTY_1"] === firstResource
    ) {
      state.isFirstResourceFound = true;
    }
    if (!state.isFirstResourceFound) {
      continue;
    }
    if (!data[i].hasOwnProperty("__EMPTY_1")) {
      break;
    }

    const resource = createResource(data, i, state.resourceId);
    state.resourceId = resource.id;
    resourcesStore.push(resource);

    // Inner loop
    state = eventsLoop(
      data,
      i,
      monthDate,
      monthDays,
      state,
      eventsStore,
      resourceTimeRangesStore
    );
  }
  return state;
}

function createResource(data, index, resourceId) {
  return {
    id: ++resourceId,
    name: data[index]["__EMPTY_1"],
    availableDays: data[index]["__EMPTY_2"],
  };
}

function eventsLoop(
  data,
  rowIndex,
  monthDate,
  monthDays,
  state,
  eventsStore,
  resourceTimeRangesStore
) {
  let event = {};
  let lastDay = -1;
  let lastEventName = "";
  let index = 0;

  for (const [key, value] of Object.entries(data[rowIndex])) {
    if (key === "__EMPTY_1" || key === "__EMPTY_2") {
      continue;
    }
    const { isStartOfEvent, isEndOfEvent, dayOfMonth, eventName } =
      processEventDetails(
        data,
        rowIndex,
        key,
        value,
        monthDays,
        lastDay,
        lastEventName,
        index
      );

    if (isStartOfEvent) {
      event = startEvent(
        eventName,
        monthDate,
        dayOfMonth,
        state.eventId,
        state.resourceTimeRangesId,
        state.resourceId
      );

      if (eventName === "X") {
        state.resourceTimeRangesId = event.id;
      } else {
        state.eventId = event.id;
      }
    }

    if (isEndOfEvent) {
      event.endDate = addDays(monthDate, dayOfMonth + 1);
      if (event.name === "X") {
        event.name = "";
        resourceTimeRangesStore.push(event);
      } else {
        eventsStore.push(event);
      }
      event = {};
    }

    lastEventName = eventName;
    lastDay = dayOfMonth;
    index++;
  }

  return state;
}

function processEventDetails(
  data,
  rowIndex,
  key,
  value,
  monthDays,
  lastDay,
  lastEventName,
  index
) {
  const eventsObjectValues = Object.entries(data[rowIndex]);
  const eventsObjectLength = eventsObjectValues.length;
  const dayOfMonth = monthDays[key];
  const eventName = value;
  const nextEventName = eventsObjectValues[index + 3]
    ? eventsObjectValues[index + 3][1]
    : "";
  const nextEventDay = eventsObjectValues[index + 3]
    ? monthDays[eventsObjectValues[index + 3][0]]
    : -1;

  const isStartOfEvent =
    eventName !== lastEventName ||
    (eventName === lastEventName && dayOfMonth - 1 !== lastDay);
  const isEndOfEvent =
    index === eventsObjectLength - 1 ||
    eventName !== nextEventName ||
    dayOfMonth + 1 !== nextEventDay;

  return {
    isStartOfEvent,
    isEndOfEvent,
    dayOfMonth,
    eventName,
    nextEventName,
    nextEventDay,
  };
}

function startEvent(
  eventName,
  monthDate,
  dayOfMonth,
  eventId,
  resourceTimeRangesId,
  resourceId
) {
  let event = {};
  if (eventName === "X") {
    event.id = ++resourceTimeRangesId;
    event.timeRangeColor = "red";
  } else {
    event.id = ++eventId;
  }
  event.name = eventName;
  event.startDate = addDays(monthDate, dayOfMonth);
  event.resourceId = resourceId;

  return event;
}

// read the Excel file
const workbook = xlsx.readFile("./scheduler.xlsx");
const sheetName = workbook.SheetNames[1]; // select the sheet you want
const worksheet = workbook.Sheets[sheetName];
const jsonData = xlsx.utils.sheet_to_json(worksheet);
const rows = createBryntumRows(jsonData);
// convert JSON data to the expected Bryntum Scheduler load response structure
const schedulerLoadResponse = {
  success: true,
  events: {
    rows: rows.eventsStore,
  },
  resources: {
    rows: rows.resourcesStore,
  },
  resourceTimeRanges: {
    rows: rows.resourceTimeRangesStore,
  },
  timeRanges: {
    rows: rows.timeRangesStore,
  },
};

// save the data to a JSON file
const dataJson = JSON.stringify(schedulerLoadResponse, null, 2); // convert the data to JSON, indented with 2 spaces
// define the path to the data folder
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);
const dataPath = path.join(__dirname, "data");
// ensure the data folder exists
if (!fs.existsSync(dataPath)) {
  fs.mkdirSync(dataPath);
}

// define the path to the JSON file in the data folder
const filePath = path.join(dataPath, "scheduler-data.json");
// write the JSON string to a file in the data directory
fs.writeFile(filePath, dataJson, (err) => {
  if (err) throw err;
  console.log("JSON data written to file");
});
