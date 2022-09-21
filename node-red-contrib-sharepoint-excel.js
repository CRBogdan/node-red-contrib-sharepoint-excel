const sprequest = require("sp-request");
const excel = require("exceljs");

module.exports = function (RED) {
  function SharepointExcel(config) {
    RED.nodes.createNode(this, config);
    var node = this;
    node.name = config.sharepoint;
    node.username = config.username;
    node.password = config.password;

    this.on("input", function (msg) {
      var SharepointFlowNode = RED.nodes.getNode(node.name);
      var serviceUri = config.serviceUri || msg.sharepointUri;

      if (
        SharepointFlowNode == null ||
        SharepointFlowNode.credentials == null
      ) {
        node.warn("Sharepoint credentials are missing.");
        return;
      }
      if (
        serviceUri == "" ||
        serviceUri == null ||
        serviceUri.indexOf("_api") < 0
      ) {
        node.error(
          'Service URL must be specified in Sharepoint node or passed by stream mode in msg.spURL. And URL must contain "_api" path'
        );
        return;
      }

      let credentialOptions = {
        username: SharepointFlowNode.credentials.username,
        password: SharepointFlowNode.credentials.password,
      };

      const wb = new excel.Workbook();
      let spr = sprequest.create(credentialOptions);
      let objectToSend = [];
      let sheets = [];
      let tableHeader = {
        rowIndex: 0,
        colIndex: [],
      };

      spr
        .get(encodeURI(serviceUri), { responseType: "buffer" })
        .then((response) => {
          wb.xlsx.load(response.body).then(() => {
            wb.eachSheet((sheet, id) => {
              let isFirst = true;
              sheet.eachRow((row, rowNum) => {
                let tempObj = {};

                if (isFirst == true) {
                  tableHeader.rowIndex = rowNum;
                  row.eachCell((cell, cellNum) => {
                    tableHeader.colIndex.push(cellNum);
                  });
                } else {
                  tableHeader.colIndex.forEach((index) => {
                    tempObj = {
                      ...tempObj,
                      [sheet.getRow(tableHeader.rowIndex).getCell(index).value]:
                        row.getCell(index).value,
                    };
                  });
                  objectToSend.push(tempObj);
                }
                isFirst = false;
              });

              sheets.push({
                sheetName: sheet.name,
                table: objectToSend,
              });
            });
            msg.payload = sheets;
            node.send(msg);
          });
        });
    });
  }

  RED.nodes.registerType("sharepoint-excel-read", SharepointExcel);

  function nodeRedSharepointSettings(n) {
    RED.nodes.createNode(this, n);
  }

  RED.nodes.registerType("sharepoint-excel-config", nodeRedSharepointSettings, {
    credentials: {
      username: {
        type: "text",
      },
      password: {
        type: "password",
      },
    },
  });
};
