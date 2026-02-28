# xllify project

This project contains custom Excel functions. Functions live as `.luau` files in the `functions/` directory. Tests live in `tests/`.

## Working with functions

**To create a new function**, run `xllify suggest "<description>"` redirecting stdout to `work/` and stderr to a separate file:
```
xllify suggest "<description>" > work/<suggested-name>.luau 2>work/<suggested-name>.stderr
```
The CLI prints the Luau code to stdout and a summary to stderr. **Always redirect them separately** — combining them (e.g. `2>&1`) breaks the output. After running, read both files and present the output nicely: show the summary from the `.stderr` file, examples, and the code in a fenced luau code block. Then offer to test it by calling `xllify-lua --load work/<name>.luau call <FUNCNAME> <args...>` using one of the example inputs. If the user likes it and asks to save, move it to `functions/<name>.luau`.

Never write or edit Luau function code directly. Always go through `xllify suggest` so generation is handled correctly.

## Testing functions locally

Use `xllify-lua` to run and test functions without building an XLL:

**Call a function:**
```
xllify-lua --load functions/<name>.luau call <FUNCNAME> <args...>
```

**List / describe functions:**
```
xllify-lua --load functions/<name>.luau list
xllify-lua --load functions/<name>.luau desc <FUNCNAME>
```

**Run tests:**
```
xllify-lua --load functions/<name>.luau test tests/<name>_test.luau
```

Tests use `xllify.TestSuite` / `xllify.Test` / `xllify.Assert.Equal`. See `tests/hello_test.luau` for an example.

## Building

All build output goes to the `builds/` directory.

When the user asks to build without specifying a target, ask which type they want:
- **XLL** — Windows desktop Excel
- **Office Add-in** — Office.js zip for Excel Online + Desktop
- **Office Add-in (dev)** — same but pointed at a local dev server

If the user's request implies running or testing the add-in locally (e.g. "run locally", "run the add-in locally", "export for dev", "export this", "try it out", "test it in Excel"), ask whether they want:
- **Office Add-in (dev)** — runs the add-in locally as an Office.js add-in (Excel Online + Desktop via a local dev server)
- **XLL** — builds a standalone `.xll` file for Windows desktop Excel

If they choose Office Add-in, proceed with **Office Add-in (dev)** using the existing flow. If they choose XLL, proceed with the **XLL** build — the user can load the resulting `.xll` directly from `builds/` in Excel on Windows.

Before running any build command, read `xllify.json` to pre-fill known values (`name`, `namespace`, `base_url`). If required values are missing from `xllify.json`, ask for them all in a single message, then write them into `xllify.json` before proceeding.

Required values per build type:
- **XLL**: `name`, `namespace`
- **Office Add-in**: `name`, `namespace`, `base_url`
- **Office Add-in (dev)**: `name`, `namespace`, `base_url`

**XLL (Windows desktop Excel)**
```
xllify build xll --title "<name>" --namespace <ns> -o builds/<name>.xll functions/*.luau
```

**Office Add-in**
```
xllify build officejs --title "<name>" --namespace <ns> -o builds/<name>.zip functions/*.luau
```

**Office Add-in (dev)**
```
xllify build officejs --title "<name>" --namespace <ns> --dev-url <url> -o builds/<name>.zip functions/*.luau
```

## File layout

```
functions/           # .luau source files
tests/               # _test.luau test files
work/                # staging area for new functions before saving (create if missing)
builds/              # build output (create if missing)
xllify.json          # add-in config (name, namespace, base_url, app_id, ...)
.env                 # XLLIFY_DEV_KEY=your_key_here  (never commit this)
```

## After building

- **XLL**: Tell the user to copy the `.xll` file to a Windows machine, then visit https://xllify.com/xll-need-to-know for next steps.
- **Office Add-in**: Build and leave the `.zip` in `builds/`. Remind the user that the contents need to be hosted at the `base_url` specified in `xllify.json`.
- **Office Add-in (dev)**: Before building, check if `package.json` exists in the repo root. If it does, tell the user the add-in is already set up and skip the build entirely. If it does not exist, run the build, then unzip the `.zip` into a temp directory and copy the following into the repo root (skip any item that already exists):
  - `office-addin/` directory → repo root
  - `package.json` → repo root
  - `scripts/` directory → repo root
  - `README.md` from the archive → repo root as `ADDIN.md`

  Do not overwrite existing files.

## Rules

- Never commit `.env`
- If `.env` does not exist, create it with `XLLIFY_DEV_KEY=` and tell the user:
  1. Visit https://xllify.com/web#userinfo to generate an API key
  2. They can either paste it here (pasting it here is safe — Claude Code runs locally and your input never leaves your machine) or add it to `.env` themselves as `XLLIFY_DEV_KEY=<their-key>`
- Before running any `xllify` or `xllify-lua` command, run it directly — the CLI reads `.env` automatically. Only if it fails due to a missing key, check `.env` exists and `XLLIFY_DEV_KEY` is set. If missing or empty, follow the `.env` setup rule above before proceeding.
- Before running any `xllify` or `xllify-lua` command, check it's available with `which xllify-lua`. If not found, ask the user if they'd like to install it, then run the appropriate installer:
  - **macOS/Linux**: `curl -fsSL https://xllify.com/install.sh | bash`
  - **Windows**: `irm https://xllify.com/install.ps1 | iex`
- When the user asks to run `npm run start` (or `npm start`), first run `which xllify-lua`. If not found, run `npm run install-xllify` before proceeding.
- After setting up the dev add-in for the first time (copying files from the zip), tell the user to run:
  1. `npm install`
  2. `npm run certs` (first run only — installs localhost HTTPS certificates)
  3. `npm run install-xllify` only if `which xllify-lua` returns nothing
  4. `npm start`
- If asked how to write or modify a function, always use `xllify suggest` rather than answering directly — the CLI handles generation correctly
- If a build fails, show the error verbatim — don't attempt to fix Luau syntax manually
