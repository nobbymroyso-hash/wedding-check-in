const SHEET_NAME = "Guests";

function doGet(e) {
  const action = e.parameter.action;

  if (action === "guest") {
    return jsonResponse(findGuest(e.parameter.trackingCode));
  }

  if (action === "dashboard") {
    return jsonResponse(getDashboard());
  }

  return jsonResponse({success:false,message:"Invalid action"});
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);

  if (data.action === "checkin") {
    return jsonResponse(checkInGuest(data.trackingCode, data.staff || ""));
  }

  if (data.action === "status") {
    return jsonResponse(updateStatus(data.trackingCode, data.status));
  }

  return jsonResponse({success:false});
}

function sheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}

function findGuest(code){
  const values = sheet().getDataRange().getValues();
  for(let i=1;i<values.length;i++){
    if(values[i][1] == code){
      return {
        success:true,
        row:i+1,
        guest:values[i][0],
        accommodation:values[i][2],
        status:values[i][4]
      };
    }
  }
  return {success:false};
}

function checkInGuest(code, staff){
  const s = sheet();
  const values = s.getDataRange().getValues();

  for(let i=1;i<values.length;i++){
    if(values[i][1] == code){

      if(values[i][4] === "Checked In"){
        return {
          success:false,
          message:"Already checked in"
        };
      }

      s.getRange(i+1,5).setValue("Checked In");
      s.getRange(i+1,7).setValue(new Date());

      return {
        success:true,
        guest:values[i][0]
      };
    }
  }

  return {success:false,message:"Guest not found"};
}

function updateStatus(code,status){
  const s = sheet();
  const values = s.getDataRange().getValues();

  for(let i=1;i<values.length;i++){
    if(values[i][1] == code){
      s.getRange(i+1,5).setValue(status);
      return {success:true};
    }
  }

  return {success:false};
}

function getDashboard(){
  const values = sheet().getDataRange().getValues();

  let stats = {
    registered:0,
    arrived:0,
    checkedIn:0,
    seated:0,
    noShow:0
  };

  for(let i=1;i<values.length;i++){
    const st = values[i][4];

    if(st==="Registered") stats.registered++;
    if(st==="Arrived") stats.arrived++;
    if(st==="Checked In") stats.checkedIn++;
    if(st==="Seated") stats.seated++;
    if(st==="No Show") stats.noShow++;
  }

  return stats;
}

function jsonResponse(obj){
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
