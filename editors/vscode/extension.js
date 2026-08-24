"use strict";

const vscode = require("vscode");
const service = require("./language-service");

const selector = [
  {language: "yaml", scheme: "file"},
  {language: "json", scheme: "file"},
  {language: "jsonc", scheme: "file"},
  {language: "toml", scheme: "file"},
  {pattern: "**/*.rusdox.yaml", scheme: "file"},
  {pattern: "**/*.rusdox.yml", scheme: "file"},
  {pattern: "**/*.rusdox.json", scheme: "file"},
  {pattern: "**/*.rusdox.toml", scheme: "file"}
];

function isRusdox(document) {
  return /\.rusdox\.(ya?ml|json|toml)$/i.test(document.fileName) ||
    /(^|\n)\s*version\s*[:=]\s*1\b/.test(document.getText());
}

function activate(context) {
  const collection = vscode.languages.createDiagnosticCollection("rusdox");
  context.subscriptions.push(collection);

  const refresh = (document) => {
    if (!isRusdox(document)) {
      collection.delete(document.uri);
      return;
    }
    const severity = {
      error: vscode.DiagnosticSeverity.Error,
      warning: vscode.DiagnosticSeverity.Warning
    };
    const diagnostics = service.analyze(document.getText()).map((item) => {
      const range = new vscode.Range(item.line, item.start, item.line, item.end);
      const diagnostic = new vscode.Diagnostic(range, item.message, severity[item.severity]);
      diagnostic.source = "RusDox";
      diagnostic.code = "rusdox-spec-v1";
      return diagnostic;
    });
    collection.set(document.uri, diagnostics);
  };

  vscode.workspace.textDocuments.forEach(refresh);
  context.subscriptions.push(vscode.workspace.onDidOpenTextDocument(refresh));
  context.subscriptions.push(vscode.workspace.onDidChangeTextDocument((event) => refresh(event.document)));
  context.subscriptions.push(vscode.workspace.onDidCloseTextDocument((document) => collection.delete(document.uri)));

  context.subscriptions.push(vscode.languages.registerCompletionItemProvider(selector, {
    provideCompletionItems(document, position) {
      if (!isRusdox(document)) return [];
      const prefix = document.lineAt(position.line).text.slice(0, position.character);
      return service.completionItems(prefix).map((item) => {
        const completion = new vscode.CompletionItem(item.label, vscode.CompletionItemKind.EnumMember);
        completion.detail = item.detail;
        completion.insertText = item.label;
        return completion;
      });
    }
  }, ":", "\""));

  context.subscriptions.push(vscode.languages.registerHoverProvider(selector, {
    provideHover(document, position) {
      if (!isRusdox(document)) return null;
      const range = document.getWordRangeAtPosition(position);
      const help = range && service.hoverFor(document.getText(range));
      return help ? new vscode.Hover(new vscode.MarkdownString(help), range) : null;
    }
  }));

  context.subscriptions.push(vscode.commands.registerCommand("rusdox.showSchemaVersion", () => {
    vscode.window.showInformationMessage("RusDox document spec schema version: 1");
  }));
}

function deactivate() {}

module.exports = {activate, deactivate};
