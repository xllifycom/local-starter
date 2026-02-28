# xllify project

This project contains custom Excel functions. Functions live as `.luau` files in the `functions/` directory. Tests live in `tests/`.

## Working with functions

**To create a new function**, run `xllify suggest "<description>"` redirecting stdout to `work/`:
```
xllify suggest "<description>" > work/<suggested-name>.luau
```
The CLI prints the Luau code to stdout and a summary to stderr. After running, read `work/<name>.luau` and present the output nicely: show the summary from stderr, examples, and the code in a fenced luau code block. Then offer to test it by calling `xllify-lua --load work/<name>.luau call <FUNCNAME> <args...>` using one of the example inputs. If the user likes it and asks to save, move it to `functions/<name>.luau`.

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
office-addin/
  manifest.xml       # generated — do not edit by hand, use npm run generate-manifest
  public/            # served to Excel
xll/                 # XLL build output placeholder
xllify.json          # add-in config (name, namespace, base_url, app_id, ...)
.env                 # XLLIFY_API_KEY=your_key_here  (never commit this)
```

## After building

- **XLL**: Tell the user to copy the `.xll` file to a Windows machine, then visit https://xllify.com/xll-need-to-know for next steps.
- **Office Add-in** / **Office Add-in (dev)**: No special instructions needed.

## Rules

- Never commit `.env`
- Before running any `xllify` or `xllify-lua` command, check it's available with `which xllify-lua`. If not found, ask the user if they'd like to install it, then run the appropriate installer:
  - **macOS/Linux**: `curl -fsSL https://xllify.com/install.sh | bash`
  - **Windows**: `irm https://xllify.com/install.ps1 | iex`
- When the user asks to run `npm run start` (or `npm start`), first run `which xllify-lua`. If not found, run `npm run install-xllify` before proceeding.
- If asked how to write or modify a function, always use `xllify suggest` rather than answering directly — the CLI handles generation correctly
- If a build fails, show the error verbatim — don't attempt to fix Luau syntax manually
