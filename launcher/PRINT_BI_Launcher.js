ObjC.import("Foundation");

function shQuote(value) {
  return "'" + String(value).replace(/'/g, "'\\''") + "'";
}

function zshLoginCommand(command) {
  return "/bin/zsh -lic " + shQuote(command);
}

function getProjectDir() {
  var bundlePath = ObjC.unwrap($.NSBundle.mainBundle.bundlePath);
  if (bundlePath && bundlePath.length > 0) {
    return ObjC.unwrap($(bundlePath).stringByDeletingLastPathComponent);
  }

  var cwd = ObjC.unwrap($.NSFileManager.defaultManager.currentDirectoryPath);
  return cwd;
}

function run() {
  var app = Application.currentApplication();
  app.includeStandardAdditions = true;

  try {
    var projectDir = getProjectDir();
    var selected = app.chooseFromList(
      ["Login", "Captura completa", "Abrir ultima saida"],
      {
        withTitle: "Automacao de Prints",
        withPrompt: "PRINT BI - Qlik Sense\nSelecione a acao:",
        defaultItems: ["Captura completa"],
        okButtonName: "Executar",
        cancelButtonName: "Cancelar"
      }
    );
    if (!selected) return;
    var choice = selected[0];

    if (choice === "Abrir ultima saida") {
      var latest = app.doShellScript(
        "cd " + shQuote(projectDir) + " && ls -1dt output/* 2>/dev/null | head -n 1 || true"
      );
      if (latest && latest.trim() !== "") {
        app.doShellScript("open " + shQuote(latest.trim()));
      } else {
        app.displayDialog("Nenhuma saida encontrada em output/.", {
          withTitle: "PRINT BI",
          buttons: ["OK"],
          defaultButton: "OK"
        });
      }
      return;
    }

    var actionCmd =
      choice === "Login"
        ? "npm run capture -- --config config.json --login"
        : "npm run capture -- --config config.json";

    var runCmd = zshLoginCommand(
      "cd " +
        shQuote(projectDir) +
        " && command -v npm >/dev/null 2>&1 || { echo 'ERRO: npm nao encontrado no PATH.'; echo 'Instale Node.js/NPM ou ajuste seu PATH no zsh.'; exit 1; };" +
        " if [ ! -d node_modules ]; then npm install; fi && " +
        actionCmd
    );

    if (choice === "Login") {
      app.displayDialog(
        "Sera aberto o navegador para autenticacao Intraer.\n\n1) Faca login com usuario/senha.\n2) Volte ao Terminal.\n3) Pressione ENTER para salvar a sessao.",
        {
          withTitle: "PRINT BI - Login",
          buttons: ["OK"],
          defaultButton: "OK"
        }
      );
    }

    var terminal = Application("Terminal");
    terminal.activate();
    terminal.doScript(runCmd);
  } catch (error) {
    // -128 = user canceled
    if (error && String(error).indexOf("-128") >= 0) return;
    app.displayDialog("Erro no launcher:\n" + error, {
      withTitle: "PRINT BI",
      buttons: ["OK"],
      defaultButton: "OK"
    });
  }
}
