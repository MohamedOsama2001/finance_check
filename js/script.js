let egyptBtn = document.querySelector(".egypt");
let dubaiBtn = document.querySelector(".dubai");
let ksaBtn = document.querySelector(".ksa");
let table = document.querySelector(".resultTable .table tbody");
let tableToday = document.querySelector(".todayTable .table tbody");
let result = [];
function drawTable(result) {
  let ele = result.map((item) => {
    return `
        <tr>
        <td>${item["Airline PNR"]}</td>
        <td>${item["GDS PNR"]}</td>
        <td>${item["Client Name"]}</td>
        <td>${item["Nationality"]}</td>
        <td>${item["Branch"]}</td>
        <td>${item["Booking Date"]}</td>
        <td>${item["Travel Date"]}</td>
        <td>${item["Segments"]}</td>
        </tr>
        `;
  });
  table.innerHTML = ele.join("");
}
//////////////////////////////////
function drawTodayTable(result) {
    let ele = result.map((item) => {
      return `
          <tr>
          <td>${item["Airline PNR"]}</td>
          <td>${item["GDS PNR"]}</td>
          <td>${item["Client Name"]}</td>
          <td>${item["Nationality"]}</td>
          <td>${item["Branch"]}</td>
          <td>${item["Booking Date"]}</td>
          <td>${item["Travel Date"]}</td>
          <td>${item["Segments"]}</td>
          </tr>
          `;
    });
    tableToday.innerHTML = ele.join("");
}
function readExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];

    const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    const keys = jsonData[0];
    const values = jsonData.slice(1);

    result = values.map((row) => {
      let obj = {};
      keys.forEach((key, index) => {
        if (key === "Booking Date") {
          const dateValue = row[index];
          obj[key] = formatDateWithAMPM(dateValue,true);
        } else if (key === "Travel Date") {
          const dateValue = row[index];
          obj[key] = formatDateWithAMPM(dateValue,false);
        } else {
          obj[key] = row[index];
        }
      });
      return obj;
    });
  };
  reader.readAsArrayBuffer(file);
}
// date
function formatDateWithAMPM(dateValue, showTime) {
  const date = new Date((dateValue - 25569) * 86400 * 1000);
  const options = {
    year: "numeric",
    month: "numeric",
    day: "numeric",
    timeZone: "UTC",
    ...(showTime
      ? {
          hour: "2-digit",
          minute: "2-digit",
          hour12: true,
        }
      : {}),
  };
  return date.toLocaleString("en-US", options);
}
// sort
function sortByBranch(data) {
  return data.sort((a, b) => {
    const branchA = a["Branch"].toLowerCase();
    const branchB = b["Branch"].toLowerCase();
    if (branchA < branchB) return -1;
    if (branchA > branchB) return 1;
    return 0;
  });
}
// filter egypt refrences
function sotoEgypt(stop) {
  stop.preventDefault();
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  const todayDateString = today.toISOString().split("T")[0];
  const tomorrowDateString = tomorrow.toISOString().split("T")[0];
  const egyptAirline = [
    "ABS",
    "ATZ",
    "ASW",
    "HBE",
    "CAI",
    "HRG",
    "LXR",
    "RMF",
    "SSH",
    "HMB",
    "SPX",
  ];
  const branches = ["Dubai", "wonder saudia"];
  const dateFilteredData = result.filter((item) => {
    const travelDateValue = item["Travel Date"];
  
    // تحقق إذا كانت القيمة صالحة لتكون تاريخًا
    if (!travelDateValue) return false; // إذا كانت القيمة غير موجودة
    
    const travelDate = new Date(travelDateValue);
    
    // تحقق من صحة التاريخ
    if (isNaN(travelDate.getTime())) return false; // إذا كان التاريخ غير صحيح
  
    const travelDateString = travelDate.toISOString().split("T")[0];
    
    return item["Status"] === "S" 
      && item["Booking Type"] === "NDC Reservation" 
      && !branches.includes(item["Branch"]) 
      && (travelDateString === todayDateString || travelDateString === tomorrowDateString);
  });
  const filteredData = result.filter((item) => {
    return (
      item["Status"] === "S" &&
      item["Booking Type"] === "NDC Reservation" &&
      !branches.includes(item["Branch"]) &&
      !egyptAirline.includes(item["Origin"]) &&
      !egyptAirline.includes(item["Final Destination"]) &&
      !egyptAirline.includes(item["Segment 1"]) &&
      !egyptAirline.includes(item["Segment 2"]) &&
      !egyptAirline.includes(item["Segment 3"]) &&
      !egyptAirline.includes(item["Segment 4"]) &&
      !egyptAirline.includes(item["Segment 5"])
    );
  });
  if (filteredData.length === 0) {
    alert("No Data!");
  }
  drawTable(sortByBranch(filteredData));
  drawTodayTable(sortByBranch(dateFilteredData));
}
// filter dubai refrences
function sotoDubai(stop) {
  stop.preventDefault();
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  const todayDateString = today.toISOString().split("T")[0];
  const tomorrowDateString = tomorrow.toISOString().split("T")[0];
  const dubaiAirline = [
    "AUH",
    "AAN",
    "DHF",
    "DWC",
    "NHD",
    "AZI",
    "DXB",
    "FJR",
    "RKT",
    "SHJ",
    "XSB",
  ];
  const dateFilteredData = result.filter((item) => {
    const travelDateValue = item["Travel Date"];
  
    // تحقق إذا كانت القيمة صالحة لتكون تاريخًا
    if (!travelDateValue) return false; // إذا كانت القيمة غير موجودة
    
    const travelDate = new Date(travelDateValue);
    
    // تحقق من صحة التاريخ
    if (isNaN(travelDate.getTime())) return false; // إذا كان التاريخ غير صحيح
  
    const travelDateString = travelDate.toISOString().split("T")[0];
    return item["Status"] === "S" && item["Booking Type"] === "NDC Reservation" && item["Branch"] === "Dubai" && (travelDateString === todayDateString || travelDateString === tomorrowDateString);
  });
  const filteredData = result.filter((item) => {
    return (
      item["Status"] === "S" &&
      item["Booking Type"] === "NDC Reservation" &&
      item["Branch"] === "Dubai" &&
      !dubaiAirline.includes(item["Origin"]) &&
      !dubaiAirline.includes(item["Final Destination"]) &&
      !dubaiAirline.includes(item["Segment 1"]) &&
      !dubaiAirline.includes(item["Segment 2"]) &&
      !dubaiAirline.includes(item["Segment 3"]) &&
      !dubaiAirline.includes(item["Segment 4"]) &&
      !dubaiAirline.includes(item["Segment 5"])
    );
  });
  if (filteredData.length === 0) {
    alert("No Data!");
  }
  drawTable(filteredData);
  drawTodayTable(dateFilteredData)
}
// filter ksa refrences
function sotoKsa(stop) {
  stop.preventDefault();
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);
  const todayDateString = today.toISOString().split("T")[0];
  const tomorrowDateString = tomorrow.toISOString().split("T")[0];
  const ksaAirline = [
    "AHB",
    "HOF",
    "ABT",
    "AQI",
    "EJH",
    "AJF",
    "RAE",
    "BHH",
    "ELQ",
    "URY",
    "HAS",
    "GIZ",
    "QJB",
    "DHA",
    "JED",
    "DMM",
    "KMX",
    "RUH",
    "KMC",
    "DWD",
    "ULH",
    "EAM",
    "YNB",
    "MED",
    "AKH",
    "RAH",
    "XXN",
    "SHW",
    "TIF",
    "TUU",
    "TUI",
    "WAE",
  ];
  const dateFilteredData = result.filter((item) => {
    const travelDateValue = item["Travel Date"];
  
    // تحقق إذا كانت القيمة صالحة لتكون تاريخًا
    if (!travelDateValue) return false; // إذا كانت القيمة غير موجودة
    
    const travelDate = new Date(travelDateValue);
    // تحقق من صحة التاريخ
    if (isNaN(travelDate.getTime())) return false; // إذا كان التاريخ غير صحيح
  
    const travelDateString = travelDate.toISOString().split("T")[0];
    return item["Status"] === "S" && item["Booking Type"] === "NDC Reservation" && item["Branch"] === "wonder saudia" && (travelDateString === todayDateString || travelDateString === tomorrowDateString);
  });
  const filteredData = result.filter((item) => {
    return (
      item["Status"] === "S" &&
      item["Booking Type"] === "NDC Reservation" &&
      item["Branch"] === "wonder saudia" &&
      !ksaAirline.includes(item["Origin"]) &&
      !ksaAirline.includes(item["Final Destination"]) &&
      !ksaAirline.includes(item["Segment 1"]) &&
      !ksaAirline.includes(item["Segment 2"]) &&
      !ksaAirline.includes(item["Segment 3"]) &&
      !ksaAirline.includes(item["Segment 4"]) &&
      !ksaAirline.includes(item["Segment 5"])
    );
  });
  if (filteredData.length === 0) {
    alert("No Data!");
  }
  drawTable(filteredData);
  drawTodayTable(dateFilteredData)
}

egyptBtn.addEventListener("click", sotoEgypt);
dubaiBtn.addEventListener("click", sotoDubai);
ksaBtn.addEventListener("click", sotoKsa);
document.getElementById("file-input").addEventListener("change", (event) => {
  const file = event.target.files[0];
  if (file) {
    readExcelFile(file);
  }
});
