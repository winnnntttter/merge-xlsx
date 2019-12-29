const vscode = require("vscode");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

function activate(context) {
  vscode.commands.registerCommand("mergeXlsx.input", function() {
    if (vscode.workspace.workspaceFolders.length === 1) {
      const basePath = vscode.workspace.workspaceFolders[0].uri.fsPath;
      vscode.window
        .showInputBox({
          prompt: "merge xlsx: 请输入当前工作区根目录第一个xlsx文件名。"
        })
        .then(function(text) {
          fs.exists(path.join(basePath, text), function(exists) {
            if (exists) {
              vscode.window
                .showInputBox({
                  prompt: "merge xlsx: 请输入当前工作区根目录第二个xlsx文件名。"
                })
                .then(function(text2) {
                  fs.exists(path.join(basePath, text), function(exists2) {
                    if (exists2) {
                      let workbook = XLSX.readFile(path.join(basePath, text));
                      let workbook2 = XLSX.readFile(
                        path.join(basePath, text2)
                      );
                      let sheetNames = workbook.SheetNames;
                      let sheetNames2 = workbook2.SheetNames;
                      let worksheet = workbook.Sheets[sheetNames[0]];
                      let worksheet2 = workbook2.Sheets[sheetNames2[0]];
                      let sheet1 = XLSX.utils.sheet_to_json(worksheet);
                      let sheet2 = XLSX.utils.sheet_to_json(worksheet2);
                      let key1 = worksheet.A1.h;
                      const result = sheet1
                        .concat(sheet2)
                        .reduce((prev, next) => {
                          let index = prev.findIndex(
                            (elem, i) => elem[key1] === next[key1]
                          );

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
                      vscode.window.showErrorMessage(
                        `当前工作区根路径不存在${text2}!`
                      );
                    }
                  });
                });
            } else {
              vscode.window.showErrorMessage(`当前工作区根路径不存在${text}!`);
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
