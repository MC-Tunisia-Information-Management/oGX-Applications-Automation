function dataExtraction(query) {
  var requestOptions = {
    method: "post",
    payload: query,
    contentType: "application/json",
    headers: {
      access_token: "",
    },
  };
  var response = UrlFetchApp.fetch(
    `https://gis-api.aiesec.org/graphql?access_token=${requestOptions["headers"]["access_token"]}`,
    requestOptions
  );
  var recievedDate = JSON.parse(response.getContentText());
  return recievedDate.data.allOpportunityApplication.data;
}

function dataUpdating() {
  var sheetInterface =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name
  var sheetApplications =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("1Applications"); // write sheet name
  var startDate = Utilities.formatDate(
    sheetInterface.getRange(7, 2).getValue(),
    "GMT+1",
    "dd/MM/yyyy"
  );
  //  var startDate = "01/07/2023"
  var page_number = 1;
  var allData = [];

  var queryApplications = `query{\n  allOpportunityApplication(\n    filters:{\n     created_at:{from:\"${startDate}\"}\n  person_home_mc:1559   \n programmes:[7,8,9]\n   }\n    page:${page_number} \n    per_page:100\n  ){\n    data{\n      id\n      person{\n        full_name\n        id\n        contact_detail{        phone\n        }        home_lc{\n          name\n        }\n  home_mc{\n          name\n        }\n cvs{\n          url\n        }\n   }\n opportunity{\n    id\n   title\n    programme{\n          short_name_display\n        }\n      }\n      created_at\n      status\n      host_lc_name\n  home_mc{\n        name\n      }\n   }\n  }\n}`;
  var query = JSON.stringify({ query: queryApplications });
  var data = dataExtraction(query);
  allData.push(data);
  var newRows = [];

  for (let data of allData) {
    console.log(data);
    for (let i = 0; i < data.length; i++) {
      var duplicatedRowIndex = sheetApplications
        .createTextFinder(data[i].id)
        .matchEntireCell(true)
        .findAll()
        .map((x) => x.getRow());
      let duplicated = duplicatedRowIndex.length == 0 ? false : true;
      if (!duplicated) {
        console.log(i);
        newRows.push([
          data[i].created_at,
          data[i].id ? data[i].id : "",
          data[i].person.full_name ? data[i].person.full_name : "",
          data[i].person.contact_detail.phone
            ? data[i].person.contact_detail.phone
            : "",
          data[i].person.id ? data[i].person.id : "",
          data[i].opportunity.id ? data[i].opportunity.id : "",
          data[i].opportunity.title ? data[i].opportunity.title : "",
          data[i].opportunity.programme.short_name_display
            ? data[i].opportunity.programme.short_name_display
            : "",
          data[i].status ? data[i].status : "",
          data[i].person.home_lc.name,
          data[i].person.home_mc.name,
          data[i].host_lc_name,
          data[i].host_mc_name,
          data[i].person.cvs[0] ? data[i].person.cvs[0].url : "-",
        ]);
      } else {
        sheetApplications
          .getRange(duplicatedRowIndex[0], 9)
          .setValue(data[i].status);
      }
    }
  }
  if (newRows.length > 0) {
    sheetApplications
      .getRange(
        sheetApplications.getLastRow() + 1,
        1,
        newRows.length,
        newRows[0].length
      )
      .setValues(newRows);
  }
  var now = new Date();
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Interface"); // write sheet name
  updateDate = sheet.getRange(7, 4).setValue(now);
  updateDate = sheet.getRange(7, 3).setValue("Succeeded");
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu("Run").addItem("Run the code", "dataUpdating").addToUi();
}
