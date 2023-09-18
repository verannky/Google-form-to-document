function main(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (e.values[8] === "Photogrammetry Survey" && e.values[1] === "Indonesia (id)") {
    photogrammetrySurveyGenerateDocsIndonesia(e, ss);
  } else if (e.values[8] === "LiDAR Survey" && e.values[1] === "Indonesia (id)") {
    lidarSurveyGenerateDocsIndonesia(e, ss);
  } else if (e.values[8] === "Photogrammetry Survey" && e.values[1] === "English (En)") {
    photogrammetrySurveyGenerateDocsEnglish(e, ss);
  } else if (e.values[8] === "LiDAR Survey" && e.values[1] === "English (En)") {
    lidarSurveyGenerateDocsEnglish(e, ss);
  }
}

function photogrammetrySurveyGenerateDocsIndonesia(e, ss) {
  var timeStamp = e.values[0];
  var projecttitle = e.values[2];
  var projectInChargePIC = e.values[6];
  var userClientName = e.values[3];
  var months = e.values[4];
  var year = e.values[5];
  var projectManagerInChargePMIC = e.values[7];
  var Province = e.values[11];
  var city = e.values[10];
  var area = e.values[12];
  var gsdPlan = e.values[9];
  var gpsName = e.values[18];
  var zonaUTM = e.values[26];
  var photoDroneName = e.values[16];
  var numberOfCP = e.values[19];
  var numberOfBM = e.values[20];
  var zonaUTM = e.values[26];
  var numberOfFlights = e.values[23];
  var numberofDays = e.values[24];
  var orthoGSD = e.values[31];
  var dsmGSD = e.values[41];
  var horizontalRMSE = e.values[33];
  var numberOfGCP = e.values[21];
  var numberOfICP = e.values[22];
  var VerticalRMSE = e.values[34];
  var documentation1Description = e.values[46];
  var documentation2Description = e.values[48];
  var documentation3Description = e.values[50];

  //image link
  var Detailed_AOI = e.values[13]; // Assuming this is the URL of the image in the spreadsheet
  var OverviewAOIImage = e.values[14];
  var tableOfDailyProgressReport = e.values[25];
  var cpGeodeticCoordinateTable = e.values[27];
  var cpUtmCoordinateTable = e.values[28];
  var sampleRawPhoto = e.values[29];
  var orthophotoImage = e.values[30];
  var horizontalAccuracyTest = e.values[32];
  var dsmImage = e.values[40];
  var mapLayout = e.values[42];
  var agisoftProcessingReport = e.values[43];
  var gcpForm = e.values[44];
  var documentation1 = e.values[45];
  var documentation2 = e.values[47];
  var documentation3 = e.values[49];
  

  //timestamp split
  var timeStampSplit = timeStamp.split(" ")[0].split("/");
  var dayTimeStamp = timeStampSplit[0];
  var monthTimeStamp = timeStampSplit[1];
  var yearTimeStamp = timeStampSplit[2]; //ambil 2 angka dibelakang

  var templateFile = DriveApp.getFileById("TemplateID");
  var templateresponsefolder = DriveApp.getFolderById("OutputFolderID");

  var copy = templateFile.makeCopy(yearTimeStamp + '-' + monthTimeStamp + '_' + userClientName + '_'+ 'Photogrammetry_Final Report', templateresponsefolder);
  var documentId = copy.getId();
  var documentApp = DocumentApp.openById(documentId);
  var documentBody = documentApp.getBody();

  documentBody.replaceText("{{Project Title}}", projecttitle);
  documentBody.replaceText("{{Data PIC}}", projectInChargePIC);
  documentBody.replaceText("{{Month}}", months);
  documentBody.replaceText("{{Year}}", year);
  documentBody.replaceText("{{dd/mm/yyyy}}", `${dayTimeStamp}/${monthTimeStamp}/${yearTimeStamp}`);
  documentBody.replaceText("{{PMIC}}", projectManagerInChargePMIC);
  documentBody.replaceText("{{User}}", userClientName);
  documentBody.replaceText("{{City}}", city);
  documentBody.replaceText("{{Area}}", area);
  documentBody.replaceText("{{Province}}", Province);
  documentBody.replaceText("{{GSD Plan}}", gsdPlan);
  documentBody.replaceText("{{GPS Name}}", gpsName);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{Photo Drone Name}}", photoDroneName);
  documentBody.replaceText("{{Number of CP}}", numberOfCP);
  documentBody.replaceText("{{Number of BM}}", numberOfBM);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{Number of Flights}}", numberOfFlights);
  documentBody.replaceText("{{Number of Days}}", numberofDays);
  documentBody.replaceText("{{Ortho GSD}}", orthoGSD);
  documentBody.replaceText("{{DSM GSD}}", dsmGSD);
  documentBody.replaceText("{{Horizontal RMSE}}", horizontalRMSE);
  documentBody.replaceText("{{Number of GCP}}", numberOfGCP);
  documentBody.replaceText("{{Number of ICP}}", numberOfICP);
  documentBody.replaceText("{{Vertical RMSE}}", VerticalRMSE);
  documentBody.replaceText("{{Documentation 1 Description}}", documentation1Description);
  documentBody.replaceText("{{Documentation 2 Description}}", documentation2Description);
  documentBody.replaceText("{{Documentation 3 Description}}", documentation3Description);

  //input from spreadsheet
  var picData = searchDataFromGoogleSheet(ss, "Team Database", "PIC", projectInChargePIC); // Use 'ss' parameter here
  if (picData) {
    documentBody.replaceText("{{PIC Position}}", picData[1]);
    replaceTextToImageBody(documentBody, "{{PIC TTD}}", picData[2], 5, 3);
  }

  var pmicData = searchDataFromGoogleSheet(ss, "PMIC Database", "PMIC", projectManagerInChargePMIC); // Use 'ss' parameter here
  if (pmicData) {
    replaceTextToImageBody(documentBody, "{{PMIC TTD}}", pmicData[1], 5, 3);
  }

  var gpsData = searchDataFromGoogleSheet(ss, "GPS Database", "GPS Name", gpsName); // Use 'ss' parameter here
  if (gpsData) {
    documentBody.replaceText("{{GPS Description}}", gpsData[2]);
    documentBody.replaceText("{{Specification Tabel of GPS}}", gpsData[3]);
    replaceTextToImageBody(documentBody, "{{GPS Image}}", gpsData[1], 10, 8);
  }

  var droneData = searchDataFromGoogleSheet(ss, "Photo Drone Database", "Photo Drone Name", photoDroneName); // Use 'ss' parameter here
  if (droneData) {
    documentBody.replaceText("{{Photo Drone Description}}", droneData[2]);
    documentBody.replaceText("{{Specification Table of Photo Drone}}", droneData[3]);
    replaceTextToImageBody(documentBody, "{{Photo Drone Image}}", droneData[1], 10, 8);
  }

  //input image

  if (Detailed_AOI) {
    replaceTextToImageBody(documentBody, "{{Detailed AOI Image}}", Detailed_AOI, 10, 8);
  }

  if (OverviewAOIImage) {
    replaceTextToImageBody(documentBody, "{{Overview AOI Image}}", OverviewAOIImage, 10, 8);
  }

  if (tableOfDailyProgressReport) {
    replaceTextToImageBody(documentBody, "{{Table of Daily Progress Report}}", tableOfDailyProgressReport, 10, 8);
  }

  if (cpGeodeticCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP Geodetic Coordinate Table}}", cpGeodeticCoordinateTable, 10, 8);
  }

  if (cpUtmCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP UTM Coordinate Table}}", cpUtmCoordinateTable, 10, 8);
  }

  if (sampleRawPhoto) {
    replaceTextToImageBody(documentBody, "{{Sample Raw Photo}}", sampleRawPhoto, 10, 8);
  }

  if (orthophotoImage) {
    replaceTextToImageBody(documentBody, "{{Orthophoto Image}}", orthophotoImage, 10, 8);
  }

  if (horizontalAccuracyTest) {
    replaceTextToImageBody(documentBody, "{{Horizontal Accuracy Test}}", horizontalAccuracyTest, 10, 8);
  }

  if (dsmImage) {
    replaceTextToImageBody(documentBody, "{{DSM Image}}", dsmImage, 10, 8);
  }

  if (mapLayout) {
    replaceTextToImageBody(documentBody, "{{Map Layout}}", mapLayout, 10, 8);
  }

  if (agisoftProcessingReport) {
    replaceTextToImageBody(documentBody, "{{Agisoft Processing Report}}", agisoftProcessingReport, 10, 8);
  }

  if (gcpForm) {
    replaceTextToImageBody(documentBody, "{{GCP Form}}", gcpForm, 10, 8);
  }

  if (documentation1) {
    replaceTextToImageBody(documentBody, "{{Documentation 1}}", documentation1, 10, 8);
  }

  if (documentation2) {
    replaceTextToImageBody(documentBody, "{{Documentation 2}}", documentation2, 10, 8);
  }

  if (documentation3) {
    replaceTextToImageBody(documentBody, "{{Documentation 3}}", documentation3, 10, 8);
  }

  //input header
  replaceTextToTextHeader(documentApp, "{{Project Title}}", projecttitle);
  replaceTextToTextHeader(documentApp, "{{Province}}", Province);

  documentApp.saveAndClose();
}

function photogrammetrySurveyGenerateDocsEnglish(e, ss) {
  var timeStamp = e.values[0];
  var projecttitle = e.values[2];
  var projectInChargePIC = e.values[6];
  var userClientName = e.values[3];
  var months = e.values[4];
  var year = e.values[5];
  var projectManagerInChargePMIC = e.values[7];
  var Province = e.values[11];
  var city = e.values[10];
  var area = e.values[12];
  var gsdPlan = e.values[9];
  var gpsName = e.values[18];
  var zonaUTM = e.values[26];
  var photoDroneName = e.values[16];
  var numberOfCP = e.values[19];
  var zonaUTM = e.values[26];
  var numberOfFlights = e.values[23];
  var numberofDays = e.values[24];
  var orthoGSD = e.values[31];
  var dsmGSD = e.values[41];
  var numberOfGCP = e.values[21];
  var numberOfICP = e.values[22];
  var horizontalRMSE = e.values[33];
  var VerticalRMSE = e.values[34];
  var documentation1Description = e.values[46];
  var documentation2Description = e.values[48];
  var documentation3Description = e.values[50];

  //image link
  var Detailed_AOI = e.values[13]; // Assuming this is the URL of the image in the spreadsheet
  var OverviewAOIImage = e.values[14];
  var tableOfDailyProgressReport = e.values[25];
  var cpGeodeticCoordinateTable = e.values[27];
  var cpUtmCoordinateTable = e.values[28];
  var sampleRawPhoto = e.values[29];
  var orthophotoImage = e.values[30];
  var horizontalAccuracyTest = e.values[32];
  var dsmImage = e.values[40];
  var mapLayout = e.values[42];
  var agisoftProcessingReport = e.values[43];
  var gcpForm = e.values[44];
  var documentation1 = e.values[45];
  var documentation2 = e.values[47];
  var documentation3 = e.values[49];
  
  //timestamp split
  var timeStampSplit = timeStamp.split(" ")[0].split("/");
  var dayTimeStamp = timeStampSplit[1];
  var monthTimeStamp = timeStampSplit[0];
  var yearTimeStamp = timeStampSplit[2];

  var templateFile = DriveApp.getFileById("templateID");
  var templateresponsefolder = DriveApp.getFolderById("OutputFolderID");

  var copy = templateFile.makeCopy(yearTimeStamp + '-' + monthTimeStamp + '_' + userClientName + '-'+ 'PRJ_Photogrammetry_Final Report', templateresponsefolder);
  var documentId = copy.getId();
  var documentApp = DocumentApp.openById(documentId);
  var documentBody = documentApp.getBody();

  documentBody.replaceText("{{Project Title}}", projecttitle);
  documentBody.replaceText("{{Data PIC}}", projectInChargePIC);
  documentBody.replaceText("{{Month}}", months);
  documentBody.replaceText("{{Year}}", year);
  documentBody.replaceText("{{dd/mm/yyyy}}", `${dayTimeStamp}/${monthTimeStamp}/${yearTimeStamp}`);
  documentBody.replaceText("{{PMIC}}", projectManagerInChargePMIC);
  documentBody.replaceText("{{User}}", userClientName);
  documentBody.replaceText("{{City}}", city);
  documentBody.replaceText("{{Area}}", area);
  documentBody.replaceText("{{Province}}", Province);
  documentBody.replaceText("{{GSD Plan}}", gsdPlan);
  documentBody.replaceText("{{GPS Name}}", gpsName);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{Photo Drone Name}}", photoDroneName);
  documentBody.replaceText("{{Number of CP}}", numberOfCP);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{Number of Flights}}", numberOfFlights);
  documentBody.replaceText("{{Number of Days}}", numberofDays);
  documentBody.replaceText("{{Ortho GSD}}", orthoGSD);
  documentBody.replaceText("{{DSM GSD}}", dsmGSD);
  documentBody.replaceText("{{Horizontal RMSE}}", horizontalRMSE);
  documentBody.replaceText("{{Vertical RMSE}}", VerticalRMSE);
  documentBody.replaceText("{{Number of GCP}}", numberOfGCP);
  documentBody.replaceText("{{Number of ICP}}", numberOfICP);
  documentBody.replaceText("{{Documentation 1 Description}}", documentation1Description);
  documentBody.replaceText("{{Documentation 2 Description}}", documentation2Description);
  documentBody.replaceText("{{Documentation 3 Description}}", documentation3Description);

  //input from spreadsheet
  var picData = searchDataFromGoogleSheet(ss, "Team Database", "PIC", projectInChargePIC); // Use 'ss' parameter here
  if (picData) {
    documentBody.replaceText("{{PIC Position}}", picData[1]);
    replaceTextToImageBody(documentBody, "{{PIC TTD}}", picData[2], 3, 2);
  }

  var pmicData = searchDataFromGoogleSheet(ss, "PMIC Database", "PMIC", projectManagerInChargePMIC); // Use 'ss' parameter here
  if (pmicData) {
    replaceTextToImageBody(documentBody, "{{PMIC TTD}}", pmicData[1], 3, 2);
  }

  var gpsData = searchDataFromGoogleSheet(ss, "GPS Database", "GPS Name", gpsName); // Use 'ss' parameter here
  if (gpsData) {
    documentBody.replaceText("{{GPS Description}}", gpsData[2]);
    documentBody.replaceText("{{Specification Tabel of GPS}}", gpsData[3]);
    replaceTextToImageBody(documentBody, "{{GPS Image}}", gpsData[1], 10, 8);
  }

  var droneData = searchDataFromGoogleSheet(ss, "Photo Drone Database", "Photo Drone Name", photoDroneName); // Use 'ss' parameter here
  if (droneData) {
    documentBody.replaceText("{{Photo Drone Description}}", droneData[2]);
    documentBody.replaceText("{{Specification Table of Photo Drone}}", droneData[3]);
    replaceTextToImageBody(documentBody, "{{Photo Drone Image}}", droneData[1], 10, 8);
  }

  //input image

  if (Detailed_AOI) {
    replaceTextToImageBody(documentBody, "{{Detailed AOI Image}}", Detailed_AOI, 10, 8);
  }

  if (OverviewAOIImage) {
    replaceTextToImageBody(documentBody, "{{Overview AOI Image}}", OverviewAOIImage, 10, 8);
  }

  if (tableOfDailyProgressReport) {
    replaceTextToImageBody(documentBody, "{{Table of Daily Progress Report}}", tableOfDailyProgressReport, 10, 8);
  }

  if (cpGeodeticCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP Geodetic Coordinate Table}}", cpGeodeticCoordinateTable, 10, 8);
  }

  if (cpUtmCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP UTM Coordinate Table}}", cpUtmCoordinateTable, 10, 8);
  }

  if (sampleRawPhoto) {
    replaceTextToImageBody(documentBody, "{{Sample Raw Photo}}", sampleRawPhoto, 10, 8);
  }

  if (orthophotoImage) {
    replaceTextToImageBody(documentBody, "{{Orthophoto Image}}", orthophotoImage, 10, 8);
  }

  if (horizontalAccuracyTest) {
    replaceTextToImageBody(documentBody, "{{Horizontal Accuracy Test}}", horizontalAccuracyTest, 10, 8);
  }

  if (dsmImage) {
    replaceTextToImageBody(documentBody, "{{DSM Image}}", dsmImage, 10, 8);
  }

  if (mapLayout) {
    replaceTextToImageBody(documentBody, "{{Map Layout}}", mapLayout, 10, 8);
  }

  if (agisoftProcessingReport) {
    replaceTextToImageBody(documentBody, "{{Agisoft Processing Report}}", agisoftProcessingReport, 10, 8);
  }

  if (gcpForm) {
    replaceTextToImageBody(documentBody, "{{GCP Form}}", gcpForm, 10, 8);
  }

  if (documentation1) {
    replaceTextToImageBody(documentBody, "{{Documentation 1}}", documentation1, 10, 8);
  }

  if (documentation2) {
    replaceTextToImageBody(documentBody, "{{Documentation 2}}", documentation2, 10, 8);
  }

  if (documentation3) {
    replaceTextToImageBody(documentBody, "{{Documentation 3}}", documentation3, 10, 8);
  }

  //input header
  replaceTextToTextHeader(documentApp, "{{Project Title}}", projecttitle);
  replaceTextToTextHeader(documentApp, "{{Province}}", Province);

  documentApp.saveAndClose();
}

function lidarSurveyGenerateDocsIndonesia(e, ss) {
  var timeStamp = e.values[0];
  var projecttitle = e.values[2];
  var projectInChargePIC = e.values[6];
  var userClientName = e.values[3];
  var months = e.values[4];
  var year = e.values[5];
  var projectManagerInChargePMIC = e.values[7];
  var Province = e.values[11];
  var city = e.values[10];
  var area = e.values[12];
  var gsdPlan = e.values[9];
  var gpsName = e.values[18];
  var zonaUTM = e.values[26];
  var LiDARDroneName = e.values[17];
  var numberOfCP = e.values[19];
  var zonaUTM = e.values[26];
  var numberOfFlights = e.values[23];
  var numberofDays = e.values[24];
  var orthoGSD = e.values[31];
  var dsmGSD = e.values[41];
  var numberOfGCP = e.values[21];
  var numberOfICP = e.values[22];
  var horizontalRMSE = e.values[33];
  var VerticalRMSE = e.values[34];
  var documentation1Description = e.values[46];
  var documentation2Description = e.values[48];
  var documentation3Description = e.values[50];

  //image link
  var Detailed_AOI = e.values[13]; // Assuming this is the URL of the image in the spreadsheet
  var OverviewAOIImage = e.values[14];
  var tableOfDailyProgressReport = e.values[25];
  var cpGeodeticCoordinateTable = e.values[27];
  var cpUtmCoordinateTable = e.values[28];
  var sampleRawPhoto = e.values[29];
  var orthophotoImage = e.values[30];
  var horizontalAccuracyTest = e.values[32];
  var dsmImage = e.values[40];
  var mapLayout = e.values[42];
  var agisoftProcessingReport = e.values[43];
  var gcpForm = e.values[44];
  var documentation1 = e.values[45];
  var documentation2 = e.values[47];
  var documentation3 = e.values[49];
  
  //timestamp split
  var timeStampSplit = timeStamp.split(" ")[0].split("/");
  var dayTimeStamp = timeStampSplit[1];
  var monthTimeStamp = timeStampSplit[0];
  var yearTimeStamp = timeStampSplit[2];

  var templateFile = DriveApp.getFileById("templateID");
  var templateresponsefolder = DriveApp.getFolderById("OutputFolderID");

  var copy = templateFile.makeCopy(yearTimeStamp + '-' + monthTimeStamp + '_' + userClientName + '-'+ 'PRJ_LiDAR_Final Report', templateresponsefolder);
  var documentId = copy.getId();
  var documentApp = DocumentApp.openById(documentId);
  var documentBody = documentApp.getBody();

  documentBody.replaceText("{{Project Title}}", projecttitle);
  documentBody.replaceText("{{Data PIC}}", projectInChargePIC);
  documentBody.replaceText("{{Month}}", months);
  documentBody.replaceText("{{Year}}", year);
  documentBody.replaceText("{{dd/mm/yyyy}}", `${dayTimeStamp}/${monthTimeStamp}/${yearTimeStamp}`);
  documentBody.replaceText("{{PMIC}}", projectManagerInChargePMIC);
  documentBody.replaceText("{{User}}", userClientName);
  documentBody.replaceText("{{City}}", city);
  documentBody.replaceText("{{Area}}", area);
  documentBody.replaceText("{{Province}}", Province);
  documentBody.replaceText("{{GSD Plan}}", gsdPlan);
  documentBody.replaceText("{{GPS Name}}", gpsName);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{LiDAR Drone Name}}", LiDARDroneName);
  documentBody.replaceText("{{Number of CP}}", numberOfCP);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{Number of Flights}}", numberOfFlights);
  documentBody.replaceText("{{Number of Days}}", numberofDays);
  documentBody.replaceText("{{Ortho GSD}}", orthoGSD);
  documentBody.replaceText("{{DSM GSD}}", dsmGSD);
  documentBody.replaceText("{{Horizontal RMSE}}", horizontalRMSE);
  documentBody.replaceText("{{Vertical RMSE}}", VerticalRMSE);
  documentBody.replaceText("{{Number of GCP}}", numberOfGCP);
  documentBody.replaceText("{{Number of ICP}}", numberOfICP);
  documentBody.replaceText("{{Documentation 1 Description}}", documentation1Description);
  documentBody.replaceText("{{Documentation 2 Description}}", documentation2Description);
  documentBody.replaceText("{{Documentation 3 Description}}", documentation3Description);

  //input from spreadsheet
  var picData = searchDataFromGoogleSheet(ss, "Team Database", "PIC", projectInChargePIC); // Use 'ss' parameter here
  if (picData) {
    documentBody.replaceText("{{PIC Position}}", picData[1]);
    replaceTextToImageBody(documentBody, "{{PIC TTD}}", picData[2], 3, 2);
  }

  var pmicData = searchDataFromGoogleSheet(ss, "PMIC Database", "PMIC", projectManagerInChargePMIC); // Use 'ss' parameter here
  if (pmicData) {
    replaceTextToImageBody(documentBody, "{{PMIC TTD}}", pmicData[1], 3, 2);
  }

  var gpsData = searchDataFromGoogleSheet(ss, "GPS Database", "GPS Name", gpsName); // Use 'ss' parameter here
  if (gpsData) {
    documentBody.replaceText("{{GPS Description}}", gpsData[2]);
    documentBody.replaceText("{{Specification Table of GPS}}", gpsData[3]);
    replaceTextToImageBody(documentBody, "{{GPS Image}}", gpsData[1], 10, 8);
  }

  var droneData = searchDataFromGoogleSheet(ss, "LiDAR Drone Database", "LiDAR Drone Name", LiDARDroneName); // Use 'ss' parameter here
  if (droneData) {
    documentBody.replaceText("{{LiDAR Drone Description}}", droneData[2]);
    documentBody.replaceText("{{Specification Table of LiDAR Drone}}", droneData[3]);
    replaceTextToImageBody(documentBody, "{{LiDAR Drone Image}}", droneData[1], 10, 8);
  }

  //input image

  if (Detailed_AOI) {
    replaceTextToImageBody(documentBody, "{{Detailed AOI Image}}", Detailed_AOI, 10, 8);
  }

  if (OverviewAOIImage) {
    replaceTextToImageBody(documentBody, "{{Overview AOI Image}}", OverviewAOIImage, 10, 8);
  }

  if (tableOfDailyProgressReport) {
    replaceTextToImageBody(documentBody, "{{Table of Daily Progress Report}}", tableOfDailyProgressReport, 10, 8);
  }

  if (cpGeodeticCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP Geodetic Coordinate Table}}", cpGeodeticCoordinateTable, 10, 8);
  }

  if (cpUtmCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP UTM Coordinate Table}}", cpUtmCoordinateTable, 10, 8);
  }

  if (sampleRawPhoto) {
    replaceTextToImageBody(documentBody, "{{Sample Raw Photo}}", sampleRawPhoto, 10, 8);
  }

  if (orthophotoImage) {
    replaceTextToImageBody(documentBody, "{{Orthophoto Image}}", orthophotoImage, 10, 8);
  }

  if (horizontalAccuracyTest) {
    replaceTextToImageBody(documentBody, "{{Horizontal Accuracy Test}}", horizontalAccuracyTest, 10, 8);
  }

  if (dsmImage) {
    replaceTextToImageBody(documentBody, "{{DSM Image}}", dsmImage, 10, 8);
  }

  if (mapLayout) {
    replaceTextToImageBody(documentBody, "{{Map Layout}}", mapLayout, 30, 20);
  }

  if (agisoftProcessingReport) {
    replaceTextToImageBody(documentBody, "{{Agisoft Processing Report}}", agisoftProcessingReport, 10, 8);
  }

  if (gcpForm) {
    replaceTextToImageBody(documentBody, "{{GCP Form}}", gcpForm, 10, 8);
  }

  if (documentation1) {
    replaceTextToImageBody(documentBody, "{{Documentation 1}}", documentation1, 10, 8);
  }

  if (documentation2) {
    replaceTextToImageBody(documentBody, "{{Documentation 2}}", documentation2, 10, 8);
  }

  if (documentation3) {
    replaceTextToImageBody(documentBody, "{{Documentation 3}}", documentation3, 10, 8);
  }

  //input header
  replaceTextToTextHeader(documentApp, "{{Project Title}}", projecttitle);
  replaceTextToTextHeader(documentApp, "{{Province}}", Province);

  documentApp.saveAndClose();
}

function lidarSurveyGenerateDocsEnglish(e, ss) {
  var timeStamp = e.values[0];
  var projecttitle = e.values[2];
  var projectInChargePIC = e.values[6];
  var userClientName = e.values[3];
  var months = e.values[4];
  var year = e.values[5];
  var projectManagerInChargePMIC = e.values[7];
  var Province = e.values[11];
  var city = e.values[10];
  var area = e.values[12];
  var gsdPlan = e.values[9];
  var gpsName = e.values[18];
  var zonaUTM = e.values[26];
  var LiDARDroneName = e.values[17];
  var numberOfCP = e.values[19];
  var zonaUTM = e.values[26];
  var numberOfFlights = e.values[23];
  var numberofDays = e.values[24];
  var orthoGSD = e.values[31];
  var dsmGSD = e.values[41];
  var horizontalRMSE = e.values[33];
  var VerticalRMSE = e.values[34];
  var numberOfGCP = e.values[21];
  var numberOfICP = e.values[22];
  var documentation1Description = e.values[46];
  var documentation2Description = e.values[48];
  var documentation3Description = e.values[50];

  //image link
  var Detailed_AOI = e.values[13]; // Assuming this is the URL of the image in the spreadsheet
  var OverviewAOIImage = e.values[14];
  var tableOfDailyProgressReport = e.values[25];
  var cpGeodeticCoordinateTable = e.values[27];
  var cpUtmCoordinateTable = e.values[28];
  var sampleRawPhoto = e.values[29];
  var orthophotoImage = e.values[30];
  var horizontalAccuracyTest = e.values[32];
  var dsmImage = e.values[40];
  var mapLayout = e.values[42];
  var agisoftProcessingReport = e.values[43];
  var gcpForm = e.values[44];
  var documentation1 = e.values[45];
  var documentation2 = e.values[47];
  var documentation3 = e.values[49];
  
  //timestamp split
  var timeStampSplit = timeStamp.split(" ")[0].split("/");
  var dayTimeStamp = timeStampSplit[1];
  var monthTimeStamp = timeStampSplit[0];
  var yearTimeStamp = timeStampSplit[2];

  var templateFile = DriveApp.getFileById("templateID");
  var templateresponsefolder = DriveApp.getFolderById("outputFolderID");

  var copy = templateFile.makeCopy(yearTimeStamp + '-' + monthTimeStamp + '_' + userClientName + '-'+ 'PRJ_LiDAR_Final Report', templateresponsefolder);
  var documentId = copy.getId();
  var documentApp = DocumentApp.openById(documentId);
  var documentBody = documentApp.getBody();

  documentBody.replaceText("{{Project Title}}", projecttitle);
  documentBody.replaceText("{{Data PIC}}", projectInChargePIC);
  documentBody.replaceText("{{Month}}", months);
  documentBody.replaceText("{{Year}}", year);
  documentBody.replaceText("{{dd/mm/yyyy}}", `${dayTimeStamp}/${monthTimeStamp}/${yearTimeStamp}`);
  documentBody.replaceText("{{PMIC}}", projectManagerInChargePMIC);
  documentBody.replaceText("{{User}}", userClientName);
  documentBody.replaceText("{{City}}", city);
  documentBody.replaceText("{{Area}}", area);
  documentBody.replaceText("{{Province}}", Province);
  documentBody.replaceText("{{GSD Plan}}", gsdPlan);
  documentBody.replaceText("{{GPS Name}}", gpsName);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{LiDAR Drone Name}}", LiDARDroneName);
  documentBody.replaceText("{{Number of CP}}", numberOfCP);
  documentBody.replaceText("{{Zona UTM}}", zonaUTM);
  documentBody.replaceText("{{Number of Flights}}", numberOfFlights);
  documentBody.replaceText("{{Number of Days}}", numberofDays);
  documentBody.replaceText("{{Ortho GSD}}", orthoGSD);
  documentBody.replaceText("{{DSM GSD}}", dsmGSD);
  documentBody.replaceText("{{Horizontal RMSE}}", horizontalRMSE);
  documentBody.replaceText("{{Vertical RMSE}}", VerticalRMSE);
  documentBody.replaceText("{{Number of GCP}}", numberOfGCP);
  documentBody.replaceText("{{Number of ICP}}", numberOfICP);
  documentBody.replaceText("{{Documentation 1 Description}}", documentation1Description);
  documentBody.replaceText("{{Documentation 2 Description}}", documentation2Description);
  documentBody.replaceText("{{Documentation 3 Description}}", documentation3Description);

  //input from spreadsheet
  var picData = searchDataFromGoogleSheet(ss, "Team Database", "PIC", projectInChargePIC); // Use 'ss' parameter here
  if (picData) {
    documentBody.replaceText("{{PIC Position}}", picData[1]);
    replaceTextToImageBody(documentBody, "{{PIC TTD}}", picData[2], 3, 2);
  }

  var pmicData = searchDataFromGoogleSheet(ss, "PMIC Database", "PMIC", projectManagerInChargePMIC); // Use 'ss' parameter here
  if (pmicData) {
    replaceTextToImageBody(documentBody, "{{PMIC TTD}}", pmicData[1], 3, 2);
  }

  var gpsData = searchDataFromGoogleSheet(ss, "GPS Database", "GPS Name", gpsName); // Use 'ss' parameter here
  if (gpsData) {
    documentBody.replaceText("{{GPS Description}}", gpsData[2]);
    documentBody.replaceText("{{Specification Tabel of GPS}}", gpsData[3]);
    replaceTextToImageBody(documentBody, "{{GPS Image}}", gpsData[1], 10, 8);
  }

  var droneData = searchDataFromGoogleSheet(ss, "LiDAR Drone Database", "LiDAR Drone Name", LiDARDroneName); // Use 'ss' parameter here
  if (droneData) {
    documentBody.replaceText("{{LiDAR Drone Description}}", droneData[2]);
    documentBody.replaceText("{{Specification Table of LiDAR Drone}}", droneData[3]);
    replaceTextToImageBody(documentBody, "{{LiDAR Drone Image}}", droneData[1], 10, 8);
  }

  //input image

  if (Detailed_AOI) {
    replaceTextToImageBody(documentBody, "{{Detailed AOI Image}}", Detailed_AOI, 10, 8);
  }

  if (OverviewAOIImage) {
    replaceTextToImageBody(documentBody, "{{Overview AOI Image}}", OverviewAOIImage, 10, 8);
  }

  if (tableOfDailyProgressReport) {
    replaceTextToImageBody(documentBody, "{{Table of Daily Progress Report}}", tableOfDailyProgressReport, 10, 8);
  }

  if (cpGeodeticCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP Geodetic Coordinate Table}}", cpGeodeticCoordinateTable, 10, 8);
  }

  if (cpUtmCoordinateTable) {
    replaceTextToImageBody(documentBody, "{{CP UTM Coordinate Table}}", cpUtmCoordinateTable, 10, 8);
  }

  if (sampleRawPhoto) {
    replaceTextToImageBody(documentBody, "{{Sample Raw Photo}}", sampleRawPhoto, 10, 8);
  }

  if (orthophotoImage) {
    replaceTextToImageBody(documentBody, "{{Orthophoto Image}}", orthophotoImage, 10, 8);
  }

  if (horizontalAccuracyTest) {
    replaceTextToImageBody(documentBody, "{{Horizontal Accuracy Test}}", horizontalAccuracyTest, 10, 8);
  }

  if (dsmImage) {
    replaceTextToImageBody(documentBody, "{{DSM Image}}", dsmImage, 10, 8);
  }

  if (mapLayout) {
    replaceTextToImageBody(documentBody, "{{Map Layout}}", mapLayout, 30, 20);
  }

  if (agisoftProcessingReport) {
    replaceTextToImageBody(documentBody, "{{Agisoft Processing Report}}", agisoftProcessingReport, 10, 8);
  }

  if (gcpForm) {
    replaceTextToImageBody(documentBody, "{{GCP Form}}", gcpForm, 10, 8);
  }

  if (documentation1) {
    replaceTextToImageBody(documentBody, "{{Documentation 1}}", documentation1, 10, 8);
  }

  if (documentation2) {
    replaceTextToImageBody(documentBody, "{{Documentation 2}}", documentation2, 10, 8);
  }

  if (documentation3) {
    replaceTextToImageBody(documentBody, "{{Documentation 3}}", documentation3, 10, 8);
  }

  //input header
  replaceTextToTextHeader(documentApp, "{{Project Title}}", projecttitle);
  replaceTextToTextHeader(documentApp, "{{Province}}", Province);

  documentApp.saveAndClose();
}

function searchDataFromGoogleSheet(SpreadsheetApp, sheetName, columnName, searchValue) {
  try {
    const sheetLidarDroneDatabase = SpreadsheetApp.getSheetByName(sheetName);
    const data =  sheetLidarDroneDatabase.getDataRange().getValues();
    var headers = data[0];
    var searchCol = headers.indexOf(columnName);
    var resultRow = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][searchCol] == searchValue) {
        resultRow = i;
        break;
      }
    }
    if (resultRow > -1) {
      var resultData = data[resultRow];
      return resultData;
    } else {
      return null;
    }
  } catch (e) {
    Logger.log(e);
    Logger.log(e.stack);
    Logger.log(`Error when to search spreadsheet sheet name ${sheetName}, column name ${columnName} and search value ${searchValue}`);
  }
}

function replaceTextToImageBody(documentBody, searchText, imageFileLink, widthInCm = 9, heightInCm = 6) {
  if (typeof imageFileLink === 'undefined' || imageFileLink === null) return;

  try {
    const imageFileId = imageFileLink.split("?id=")[1];
    const image = DriveApp.getFileById(imageFileId).getBlob();
    const next = documentBody.findText(searchText);
    if (!next) {
      Logger.log(`Text not found: ${searchText}`);
      return;
    }
    const r = next.getElement();
    r.asText().setText("");
    const img = r.getParent().asParagraph().insertInlineImage(0, image);

    const widthInPoints = widthInCm * 30;
    const heightInPoints = heightInCm * 30;

    img.setWidth(widthInPoints);
    img.setHeight(heightInPoints);

  } catch (e) {
    Logger.log(e);
    Logger.log(e.stack);
    Logger.log(`Error when replacing image ${imageFileLink} in this text body ${searchText}`);
  }
}

function replaceTextToTextHeader(documentApp, searchText, textField) {
  if (typeof textField === 'undefined' || textField === null) return;

  try {
    const documentParent = documentApp.getHeader().getParent();
    for (let i = 0; i < documentParent.getNumChildren(); i += 1) {
      const child = documentParent.getChild(i);
      const childType = child.getType();
      if (childType === DocumentApp.ElementType.HEADER_SECTION) {
        child.asHeaderSection().replaceText(searchText, textField);
      }
    }
  } catch (e) {
    Logger.log(e);
    Logger.log(e.stack);
    Logger.log(`Error when to replace text ${textField} to this text header ${searchText}`);
  }
}