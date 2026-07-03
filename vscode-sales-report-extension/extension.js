const vscode = require('vscode');
const cp = require('child_process');
const path = require('path');

function activate(context) {
  let disposable = vscode.commands.registerCommand('extension.generateSalesReport', function () {
    const workspaceFolders = vscode.workspace.workspaceFolders;
    if (!workspaceFolders) {
      vscode.window.showErrorMessage('Please open a workspace first.');
      return;
    }
    const scriptPath = path.join(workspaceFolders[0].uri.fsPath, 'vscode-sales-report-extension', 'report.py');
    const terminal = vscode.window.createTerminal('Sales Report');
    terminal.show();
    terminal.sendText(`python "${scriptPath}"`);
  });
  context.subscriptions.push(disposable);
}

function deactivate() {}

module.exports = { activate, deactivate };
