# PRINT BI (Qlik Sense)

Automacao para:
- login no Qlik Sense (Intraer)
- captura de todas as telas por OM
- geracao de PDF por OM
- geracao de PDF consolidado

## Estrutura

- `scripts/capture-clicksense.js`: motor principal da captura
- `config.example.json`: modelo de configuracao
- `ABRIR_PRINT_BI.command`: launcher macOS
- `INSTALAR_WINDOWS.bat`: instalador por clique no Windows
- `windows/install-windows.ps1`: instala dependencias e cria atalhos
- `windows/run-print-bi.ps1`: launcher/execucao no Windows

## Configuracao

1. Copie o arquivo de exemplo:

```bash
cp config.example.json config.json
```

`config.json` nao sobe para o GitHub (esta no `.gitignore`), para evitar enviar credenciais/config local.

2. Ajuste os campos principais:
- `baseUrl`
- `appId`
- `sheetUrlTemplate`
- `qlik.omField`

3. Para rede Intraer, ajuste proxy se necessario:
- `browser.proxy.server` (ex.: `http://proxybrasilia.intraer:8080`)
- `browser.proxy.bypass` (ja vem com `*.intraer`)
- `browser.proxy.username` e `browser.proxy.password` apenas se seu proxy exigir autenticacao

4. `auth.httpUsername` e `auth.httpPassword`:
- deixe vazio na maioria dos casos
- use somente se houver autenticao HTTP basica na aplicacao

## Uso no macOS

Duplo clique em:
- `PRINT_BI.app`
- ou `ABRIR_PRINT_BI.command` (recria o app e abre)

Acoes do launcher:
- `Login`
- `Captura completa`
- `Abrir ultima saida`

## Uso no Windows (instalador)

1. Depois de clonar o projeto, execute por duplo clique:
- `INSTALAR_WINDOWS.bat`

2. O instalador faz:
- valida `npm`
- roda `npm install`
- instala `chromium` do Playwright
- cria atalhos na Area de Trabalho:
  - `PRINT BI - Launcher`
  - `PRINT BI - Login`
  - `PRINT BI - Captura`
  - `PRINT BI - Ultima Saida`

3. Use o atalho `PRINT BI - Launcher` para operar por botao.

## Execucao manual (qualquer SO)

1. Instalar dependencias:

```bash
npm install
```

2. Salvar sessao autenticada (primeira vez):

```bash
npm run capture -- --config config.json --login
```

3. Rodar captura:

```bash
npm run capture -- --config config.json
```

Saida:
- `output/<timestamp>/images/<OM>/...png`
- `output/<timestamp>/pdf/<OM>.pdf`
- `output/<timestamp>/pdf/TODAS_OMS.pdf`

## Subir para GitHub

Se o repositorio ainda nao estiver criado localmente:

```bash
git init
git add .
git commit -m "feat: automacao print bi com launcher mac e instalador windows"
```

Depois, conecte ao repositorio remoto e envie:

```bash
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/SEU_REPOSITORIO.git
git push -u origin main
```

## Troubleshooting rapido

- `npm nao encontrado`: instale Node.js LTS e reabra o terminal/sessao.
- erro de login no BI: rode `Login` novamente e confirme ENTER no terminal.
- sem OMs encontradas: confira `qlik.omField`.
- renderizacao incompleta: aumente `capture.waitAfterNavigationMs`.
