const vscode = require("vscode");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");
// var a = /\.xlsx|\.xls$/;
function activate(context) {
  vscode.commands.registerCommand("mergeXlsx.input", function() {
    if (vscode.workspace.workspaceFolders.length === 1) {
      const basePath = vscode.workspace.workspaceFolders[0].uri.fsPath;
      vscode.window
        .showInputBox({
          prompt: "merge xlsx: Please enter the first xlsx or xls file name."
        })
        .then(function(text) {
          if (!/\.xlsx|\.xls$/.test(text)) {
            vscode.window.showErrorMessage("Please choose a xlsx or xls file!");
            return;
          }
          fs.stat(path.join(basePath, text), function(err, file) {
            if (err) {
              vscode.window.showErrorMessage(err.message);
              return;
            }
            if (file) {
              vscode.window
                .showInputBox({
                  prompt: "merge xlsx: Please enter the second xlsx or xls file name."
                })
                .then(function(text2) {
                  fs.stat(path.join(basePath, text), function(err, file) {
                    if (err) {
                      vscode.window.showErrorMessage(err.message);
                      return;
                    }
                    if (file) {
                      let workbook = XLSX.readFile(path.join(basePath, text));
                      let workbook2 = XLSX.readFile(path.join(basePath, text2));
                      let sheetNames = workbook.SheetNames;
                      let sheetNames2 = workbook2.SheetNames;
                      let worksheet = workbook.Sheets[sheetNames[0]];
                      let worksheet2 = workbook2.Sheets[sheetNames2[0]];
                      let sheet1 = XLSX.utils.sheet_to_json(worksheet);
                      let sheet2 = XLSX.utils.sheet_to_json(worksheet2);
                      let key1 = worksheet.A1 ? worksheet.A1.h : "__EMPTY";
                      const result = sheet1.concat(sheet2).reduce((prev, next) => {
                        let index = prev.findIndex((elem, i) => elem[key1] === next[key1]);

                        if (index === -1) {
                          return prev.concat(next);
                        } else {
                          prev[index] = Object.assign({}, prev[index], next);
                          return prev;
                        }
                      }, []);
                      var wb = {
                        SheetNames: ["mySheet"],
                        Sheets: {
                          mySheet: XLSX.utils.json_to_sheet(result)
                        }
                      };

                      XLSX.writeFile(wb, path.join(basePath, `output-${new Date().getTime()}.xlsx`));
                    } else {
                      vscode.window.showErrorMessage(`There is no ${text2} in the workspace root folder!`);
                    }
                  });
                })
                .catch(function(err) {
                  vscode.window.showErrorMessage(err);
                });
            } else {
              vscode.window.showErrorMessage(`There is no ${text} in the workspace root folder!`);
            }
          });
        });
    } else {
      vscode.window.showErrorMessage("目前只支持工作区包含一个根文件夹!");
    }
  });
}
exports.activate = activate;

function deactivate() {}

module.exports = {
  activate,
  deactivate
};
