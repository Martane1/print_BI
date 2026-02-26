#!/usr/bin/env node
/* eslint-disable no-console */
const fs = require("fs");
const path = require("path");
const readline = require("readline");
const { chromium } = require("playwright");
const { PDFDocument } = require("pdf-lib");

function parseArgs(argv) {
  const options = {
    configPath: "config.json",
    loginOnly: false
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === "--config" && argv[i + 1]) {
      options.configPath = argv[i + 1];
      i += 1;
      continue;
    }

    if (arg === "--login") {
      options.loginOnly = true;
      continue;
    }
  }

  return options;
}

function normalizeMaybeConcatenatedUrl(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return trimmed;

  const firstHttps = trimmed.indexOf("https://");
  const secondHttps = firstHttps >= 0 ? trimmed.indexOf("https://", firstHttps + 1) : -1;
  const firstHttp = trimmed.indexOf("http://");
  const secondHttp = firstHttp >= 0 ? trimmed.indexOf("http://", firstHttp + 1) : -1;

  const candidates = [secondHttps, secondHttp].filter((x) => x > 0);
  if (candidates.length === 0) return trimmed;

  const cutAt = Math.min(...candidates);
  return trimmed.slice(0, cutAt);
}

function loadConfig(configPath) {
  const absolutePath = path.resolve(configPath);
  if (!fs.existsSync(absolutePath)) {
    throw new Error(`Arquivo de config nao encontrado: ${absolutePath}`);
  }

  const raw = JSON.parse(fs.readFileSync(absolutePath, "utf8"));
  const defaults = {
    outputDir: "output",
    login: {
      storageStatePath: "auth/storage-state.json",
      headless: false
    },
    browser: {
      channel: "",
      headlessCapture: true,
      proxy: {
        server: "",
        bypass: "",
        username: "",
        password: ""
      }
    },
    auth: {
      httpUsername: "",
      httpPassword: ""
    },
    capture: {
      viewport: { width: 1920, height: 1080 },
      fullPage: true,
      waitAfterNavigationMs: 4500,
      waitAfterSelectionMs: 4500,
      navigationTimeoutMs: 120000
    },
    qlik: {
      omField: "OM",
      omValues: [],
      sheetIds: [],
      sheetTitleRegex: "",
      maxFieldRows: 20000,
      fieldPageSize: 5000
    }
  };

  const merged = {
    ...defaults,
    ...raw,
    login: { ...defaults.login, ...(raw.login || {}) },
    browser: {
      ...defaults.browser,
      ...(raw.browser || {}),
      proxy: { ...defaults.browser.proxy, ...((raw.browser && raw.browser.proxy) || {}) }
    },
    auth: { ...defaults.auth, ...(raw.auth || {}) },
    capture: { ...defaults.capture, ...(raw.capture || {}) },
    qlik: { ...defaults.qlik, ...(raw.qlik || {}) }
  };

  merged.baseUrl = normalizeMaybeConcatenatedUrl(merged.baseUrl);

  if (!merged.baseUrl) {
    throw new Error("Config invalida: 'baseUrl' e obrigatorio.");
  }

  return merged;
}

function ensureDir(targetPath) {
  fs.mkdirSync(targetPath, { recursive: true });
}

function buildLaunchOptions(config, loginMode) {
  const options = {
    headless: loginMode ? Boolean(config.login.headless) : Boolean(config.browser.headlessCapture)
  };

  if (config.browser.channel) {
    options.channel = String(config.browser.channel);
  }

  const proxyCfg = (config.browser && config.browser.proxy) || {};
  if (proxyCfg.server) {
    options.proxy = {
      server: String(proxyCfg.server)
    };

    if (proxyCfg.bypass) options.proxy.bypass = String(proxyCfg.bypass);
    if (proxyCfg.username) options.proxy.username = String(proxyCfg.username);
    if (proxyCfg.password) options.proxy.password = String(proxyCfg.password);
  }

  return options;
}

function buildContextAuthOptions(config) {
  const username = String((config.auth && config.auth.httpUsername) || "").trim();
  const password = String((config.auth && config.auth.httpPassword) || "").trim();
  if (!username || !password) {
    return {};
  }

  return {
    httpCredentials: {
      username,
      password
    }
  };
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function sanitize(input, fallback = "item") {
  const text = String(input || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9._-]+/g, "_")
    .replace(/^_+|_+$/g, "");
  return text ? text.slice(0, 90) : fallback;
}

function extractAppId(url) {
  const match = String(url).match(/\/app\/([^/]+)/);
  return match ? match[1] : "";
}

function resolveSheetUrlTemplate(baseUrl, explicitTemplate) {
  if (explicitTemplate && explicitTemplate.includes("{sheetId}")) {
    return explicitTemplate;
  }

  if (/\/sheet\/[^/]+/.test(baseUrl)) {
    return baseUrl.replace(/\/sheet\/[^/]+/, "/sheet/{sheetId}");
  }

  throw new Error(
    "Nao foi possivel gerar sheetUrlTemplate automaticamente. Defina 'sheetUrlTemplate' com {sheetId} na config."
  );
}

async function waitForEnter(promptText) {
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
  });

  await new Promise((resolve) => {
    rl.question(promptText, () => resolve());
  });

  rl.close();
}

async function qlikEval(page, appId, payload) {
  return page.evaluate(
    async ({ appIdIn, payloadIn }) => {
      const req = window.require || window.requirejs;
      if (!req) {
        throw new Error("requirejs nao encontrado na pagina.");
      }

      const qlik = await new Promise((resolve, reject) => {
        req(
          ["js/qlik"],
          (qlikModule) => resolve(qlikModule),
          (err) => reject(err || new Error("Falha ao carregar js/qlik."))
        );
      });

      let app = null;
      try {
        app = qlik.currApp && qlik.currApp();
      } catch (_err) {
        app = null;
      }

      if (!app && appIdIn && qlik.openApp) {
        app = qlik.openApp(appIdIn);
      }

      if (!app) {
        throw new Error("Nao foi possivel obter o app do Qlik.");
      }

      const action = payloadIn.action;

      if (action === "ping") {
        return new Promise((resolve, reject) => {
          try {
            app.getAppLayout((layout) => {
              resolve({
                appId: (layout && layout.qInfo && layout.qInfo.qId) || app.id || appIdIn || null
              });
            });
          } catch (error) {
            reject(error);
          }
        });
      }

      if (action === "sheetList") {
        return new Promise((resolve, reject) => {
          try {
            app.getAppObjectList("sheet", (reply) => {
              const items = (reply && reply.qAppObjectList && reply.qAppObjectList.qItems) || [];
              const parsed = items.map((item) => ({
                id: item && item.qInfo ? item.qInfo.qId : "",
                title:
                  (item && item.qMeta && item.qMeta.title) ||
                  (item && item.qData && item.qData.title) ||
                  (item && item.qInfo && item.qInfo.qId) ||
                  ""
              }));
              resolve(parsed.filter((x) => x.id));
            });
          } catch (error) {
            reject(error);
          }
        });
      }

      if (action === "fieldValues") {
        const fieldName = payloadIn.fieldName;
        const maxRows = Number(payloadIn.maxRows || 20000);
        const pageSize = Number(payloadIn.pageSize || 5000);

        if (!fieldName) throw new Error("fieldName vazio.");

        const values = [];
        let top = 0;
        let totalRows = Number.MAX_SAFE_INTEGER;

        while (top < totalRows && top < maxRows) {
          const chunk = await new Promise((resolve, reject) => {
            try {
              app.createList(
                {
                  qDef: { qFieldDefs: [fieldName] },
                  qInitialDataFetch: [
                    {
                      qTop: top,
                      qLeft: 0,
                      qWidth: 1,
                      qHeight: pageSize
                    }
                  ]
                },
                (reply) => {
                  const listObject = (reply && reply.qListObject) || {};
                  const page = (listObject.qDataPages && listObject.qDataPages[0]) || {};
                  const matrix = page.qMatrix || [];
                  const size = (listObject.qSize && listObject.qSize.qcy) || matrix.length;
                  const objectId = reply && reply.qInfo ? reply.qInfo.qId : null;

                  if (objectId && app.destroySessionObject) {
                    app.destroySessionObject(objectId);
                  }

                  resolve({ matrix, size });
                }
              );
            } catch (error) {
              reject(error);
            }
          });

          totalRows = chunk.size;
          for (const row of chunk.matrix) {
            const cell = row && row[0];
            if (!cell || cell.qIsNull) continue;
            const value = String(cell.qText || "").trim();
            if (!value || value === "-") continue;
            values.push(value);
          }

          if (chunk.matrix.length === 0) break;
          top += chunk.matrix.length;
        }

        return Array.from(new Set(values));
      }

      if (action === "selectFieldValue") {
        const fieldName = payloadIn.fieldName;
        const value = payloadIn.value;
        const field = app.field(fieldName);
        await Promise.resolve(field.clear());
        await Promise.resolve(field.selectMatch(value, false));
        return { ok: true };
      }

      if (action === "clearField") {
        const fieldName = payloadIn.fieldName;
        const field = app.field(fieldName);
        await Promise.resolve(field.clear());
        return { ok: true };
      }

      throw new Error(`Acao Qlik nao suportada: ${action}`);
    },
    { appIdIn: appId, payloadIn: payload }
  );
}

async function ensureQlikReady(page, appId, timeoutMs) {
  await page.waitForFunction(() => Boolean(window.require || window.requirejs), {
    timeout: timeoutMs
  });
  await qlikEval(page, appId, { action: "ping" });
}

async function createPdfFromImages(imagePaths, outputPdfPath) {
  if (!imagePaths.length) return;

  const pdfDoc = await PDFDocument.create();
  for (const imagePath of imagePaths) {
    const bytes = fs.readFileSync(imagePath);
    const ext = path.extname(imagePath).toLowerCase();
    const image =
      ext === ".jpg" || ext === ".jpeg" ? await pdfDoc.embedJpg(bytes) : await pdfDoc.embedPng(bytes);

    const page = pdfDoc.addPage([image.width, image.height]);
    page.drawImage(image, {
      x: 0,
      y: 0,
      width: image.width,
      height: image.height
    });
  }

  fs.writeFileSync(outputPdfPath, await pdfDoc.save());
}

async function performLogin(config) {
  if (config.browser && config.browser.proxy && config.browser.proxy.server) {
    console.log(`Proxy habilitado: ${config.browser.proxy.server}`);
  }

  const browser = await chromium.launch(buildLaunchOptions(config, true));

  const context = await browser.newContext({
    ignoreHTTPSErrors: true,
    viewport: config.capture.viewport,
    ...buildContextAuthOptions(config)
  });

  const page = await context.newPage();
  page.setDefaultTimeout(config.capture.navigationTimeoutMs);

  console.log(`Abrindo pagina para login: ${config.baseUrl}`);
  await page.goto(config.baseUrl, { waitUntil: "domcontentloaded" });

  await waitForEnter(
    "\nFinalize o login no navegador e pressione ENTER aqui para salvar a sessao autenticada.\n"
  );

  const statePath = path.resolve(config.login.storageStatePath);
  ensureDir(path.dirname(statePath));
  await context.storageState({ path: statePath });

  await browser.close();
  console.log(`Sessao salva em: ${statePath}`);
}

function timestampId() {
  return new Date().toISOString().replace(/[:.]/g, "-");
}

function buildSheetData(config, sheetIdsOrObjects) {
  const template = resolveSheetUrlTemplate(config.baseUrl, config.sheetUrlTemplate);

  return sheetIdsOrObjects
    .map((entry) => {
      if (typeof entry === "string") {
        return { id: entry, title: entry, url: template.replace("{sheetId}", entry) };
      }

      const id = entry.id;
      if (!id) return null;
      return {
        id,
        title: entry.title || entry.id,
        url: template.replace("{sheetId}", id)
      };
    })
    .filter(Boolean);
}

async function runCapture(config) {
  const appId = config.appId || extractAppId(config.baseUrl);
  if (!appId) {
    throw new Error(
      "Nao foi possivel detectar appId pela URL. Defina 'appId' explicitamente na config."
    );
  }

  const storageStatePath = path.resolve(config.login.storageStatePath);
  if (!fs.existsSync(storageStatePath)) {
    throw new Error(
      `Sessao autenticada nao encontrada em ${storageStatePath}. Rode primeiro com --login.`
    );
  }

  if (config.browser && config.browser.proxy && config.browser.proxy.server) {
    console.log(`Proxy habilitado: ${config.browser.proxy.server}`);
  }

  const browser = await chromium.launch(buildLaunchOptions(config, false));

  const context = await browser.newContext({
    ignoreHTTPSErrors: true,
    viewport: config.capture.viewport,
    storageState: storageStatePath,
    ...buildContextAuthOptions(config)
  });

  const page = await context.newPage();
  page.setDefaultTimeout(config.capture.navigationTimeoutMs);

  console.log(`Abrindo app: ${config.baseUrl}`);
  await page.goto(config.baseUrl, { waitUntil: "domcontentloaded" });
  await ensureQlikReady(page, appId, config.capture.navigationTimeoutMs);

  let sheets = [];
  if (Array.isArray(config.qlik.sheetIds) && config.qlik.sheetIds.length > 0) {
    sheets = buildSheetData(config, config.qlik.sheetIds);
    console.log(`Usando ${sheets.length} sheets da config.`);
  } else {
    const discoveredSheets = await qlikEval(page, appId, { action: "sheetList" });
    sheets = buildSheetData(config, discoveredSheets);
    console.log(`Sheets descobertas automaticamente: ${sheets.length}`);
  }

  if (config.qlik.sheetTitleRegex) {
    const regex = new RegExp(config.qlik.sheetTitleRegex, "i");
    sheets = sheets.filter((sheet) => regex.test(sheet.title));
    console.log(`Sheets apos filtro por titulo: ${sheets.length}`);
  }

  if (!sheets.length) {
    throw new Error("Nenhuma sheet encontrada para captura.");
  }

  let oms = [];
  if (Array.isArray(config.qlik.omValues) && config.qlik.omValues.length > 0) {
    oms = config.qlik.omValues;
    console.log(`Usando ${oms.length} OMs da config.`);
  } else {
    oms = await qlikEval(page, appId, {
      action: "fieldValues",
      fieldName: config.qlik.omField,
      maxRows: config.qlik.maxFieldRows,
      pageSize: config.qlik.fieldPageSize
    });
    console.log(`OMs descobertas automaticamente no campo '${config.qlik.omField}': ${oms.length}`);
  }

  if (!oms.length) {
    throw new Error(
      `Nenhuma OM encontrada no campo '${config.qlik.omField}'. Verifique o nome do campo na config.`
    );
  }

  const runOutputDir = path.resolve(config.outputDir, timestampId());
  const imagesRootDir = path.join(runOutputDir, "images");
  const pdfRootDir = path.join(runOutputDir, "pdf");
  ensureDir(imagesRootDir);
  ensureDir(pdfRootDir);

  const allImages = [];
  for (const om of oms) {
    const omName = String(om);
    const omSafe = sanitize(omName, "OM");
    const omDir = path.join(imagesRootDir, omSafe);
    ensureDir(omDir);

    console.log(`\nOM: ${omName}`);
    await qlikEval(page, appId, {
      action: "selectFieldValue",
      fieldName: config.qlik.omField,
      value: omName
    });
    await sleep(config.capture.waitAfterSelectionMs);

    const omImages = [];
    for (let i = 0; i < sheets.length; i += 1) {
      const sheet = sheets[i];
      const fileName = `${String(i + 1).padStart(2, "0")}_${sanitize(sheet.title, sheet.id)}.png`;
      const filePath = path.join(omDir, fileName);

      console.log(`  [${i + 1}/${sheets.length}] ${sheet.title}`);
      await page.goto(sheet.url, { waitUntil: "domcontentloaded" });
      await ensureQlikReady(page, appId, config.capture.navigationTimeoutMs);
      await sleep(config.capture.waitAfterNavigationMs);
      await page.screenshot({
        path: filePath,
        fullPage: Boolean(config.capture.fullPage)
      });

      omImages.push(filePath);
      allImages.push(filePath);
    }

    const omPdfPath = path.join(pdfRootDir, `${omSafe}.pdf`);
    await createPdfFromImages(omImages, omPdfPath);
    console.log(`PDF OM gerado: ${omPdfPath}`);
  }

  await qlikEval(page, appId, {
    action: "clearField",
    fieldName: config.qlik.omField
  }).catch(() => {});

  const consolidatedPdfPath = path.join(pdfRootDir, "TODAS_OMS.pdf");
  await createPdfFromImages(allImages, consolidatedPdfPath);
  console.log(`\nPDF consolidado gerado: ${consolidatedPdfPath}`);
  console.log(`Imagens e PDFs salvos em: ${runOutputDir}`);

  await browser.close();
}

async function main() {
  try {
    const options = parseArgs(process.argv.slice(2));
    const config = loadConfig(options.configPath);

    if (options.loginOnly) {
      await performLogin(config);
      return;
    }

    await runCapture(config);
  } catch (error) {
    console.error("\nERRO:", error.message || error);
    process.exitCode = 1;
  }
}

main();
