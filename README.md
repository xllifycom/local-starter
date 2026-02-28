# xllify Local Starter

[xllify](https://xllify.com) lets you build custom Excel functions in plain language — no C++ or VBA required. You describe what you want, and it generates the function code and packages it as a native Excel add-in (.xll for Windows desktop, or an Office.js add-in for Excel Online).

This repo is for developers who prefer to work in Claude Code: you get the same AI-powered function generation and build pipeline, but driven from your editor with your functions stored in version control. You can of course write your function code by hand.

## Setup

1. Install the xllify CLI:
   ```bash
   # macOS / Linux as non-root user
   # you can inspect install.sh with curl https://xllify.com/install.sh i
   curl -fsSL https://xllify.com/install.sh | bash

   # Windows (PowerShell)
   irm https://xllify.com/install.ps1 | iex
   ```
2. Clone this repo (or use **Use this template** on GitHub)
3. Copy `.env.example` to `.env` and set `XLLIFY_DEV_KEY` to your key from [xllify.com](https://xllify.com)
4. Open in Claude Code

## Building

```bash
make xll        # builds/xllify.xll  (Windows desktop Excel)
make officejs   # builds/xllify.zip  (Office Add-in)
make dev        # builds/xllify-dev.zip (Office Add-in, pointed at localhost:3000)
make clean      # remove all build outputs
```

Values are read from `xllify.json`. Override on the command line if needed:

```bash
make officejs BASE_URL=https://myserver.com
```

Targets are file-based — Make skips rebuilds when nothing in `functions/` has changed.

---

## Usage

Just describe what you want in Claude Code:

> "Create a function that calculates compound interest"

> "Add a function that strips HTML tags from a cell"

> "Build an XLL from all my functions"

Claude Code will use the xllify CLI to generate code and produce deployable add-ins.

## Structure

```
functions/      # Your .luau function files
tests/          # Test files (_test.luau)
.claude/        # Claude Code config and instructions
```

---

## CLI Reference

xllify ships two CLI tools: **`xllify`** (the build/suggest CLI) and **`xllify-lua`** (the local function runner).

### `xllify` — build & suggest

Requires `XLLIFY_DEV_KEY` set in your `.env`.

#### `xllify suggest <query>`

AI-generates a Luau function from a plain-language description. Prints Luau code to **stdout** and a summary/examples to **stderr**.

```bash
xllify suggest "compound interest calculator" > work/compound_interest.luau
```

If the query can be answered by a native Excel formula, the CLI notes that on stderr and exits.

#### `xllify build xll [flags] <files...>`

Builds a Windows XLL add-in. **Builds run remotely on the xllify API** — your `.luau` source files are sent to the server and the compiled artifact is returned. Fully on-premises builds are available by arrangement; contact [xllify.com](https://xllify.com) for details.

```bash
xllify build xll --title "My Addin" --namespace myaddin -o builds/my-addin.xll functions/*.luau
```

| Flag | Description |
|------|-------------|
| `--title <title>` | Add-in display title |
| `--namespace <ns>` | Function namespace prefix (e.g. `xfy`) |
| `-o <path>` | Output file path (defaults to `<title>.xll`) |

#### `xllify build officejs [flags] <files...>`

Builds an Office.js add-in zip (Excel Online + Desktop). Like `build xll`, this runs remotely on the xllify API.

```bash
xllify build officejs --title "My Addin" --namespace myaddin -o builds/my-addin.zip functions/*.luau

# Dev build pointed at a local server
xllify build officejs --title "My Addin" --namespace myaddin --dev-url http://localhost:3000 -o builds/my-addin.zip functions/*.luau
```

| Flag | Description |
|------|-------------|
| `--title <title>` | Add-in display title (required) |
| `--namespace <ns>` | Function namespace prefix (required) |
| `--dev-url <url>` | Base URL for dev builds |
| `-o <path>` | Output file path |

#### Environment variables

| Variable | Description |
|----------|-------------|
| `XLLIFY_DEV_KEY` | Your xllify dev key (required) |
| `XLLIFY_API_URL` | Override API base URL (default: `https://xllify.com`) |

---

### `xllify-lua` — local function runner

Tests and runs Luau functions locally without building an XLL.

#### Load a file

```bash
xllify-lua --load functions/myfunc.luau
```

`--load` can be specified multiple times to load several files.

#### Commands

| Command | Description |
|---------|-------------|
| `list` | List all registered functions |
| `desc <FUNCNAME>` | Show function signature, parameters, and examples |
| `call <FUNCNAME> [args...]` | Call a function with arguments |
| `test <testfile.luau>` | Run a test file |

#### Calling functions

```bash
# Scalar args
xllify-lua --load functions/bmi.luau call BMI 70 1.75

# String args (quote them)
xllify-lua --load functions/text.luau call StripHTML "hello world"

# Array args (JSON syntax)
xllify-lua --load functions/math.luau call Sum [1,2,3,4,5]

# 2D array / matrix
xllify-lua --load functions/matrix.luau call Transpose [[1,2,3],[4,5,6]]
```

Argument types parsed automatically: numbers, strings (quoted), booleans (`true`/`false`), 1D arrays `[...]`, 2D arrays `[[...],[...]]`.

#### JSON output

```bash
xllify-lua --json --load functions/bmi.luau call BMI 70 1.75
# Returns raw JSON value, useful for scripting
```

#### Output formats

| Return type | Example output |
|-------------|---------------|
| Number | `22.857143` |
| String | `"Hello"` |
| Boolean | `true` |
| Matrix | Formatted table with column labels and row numbers |
| Error | `{"error":"message"}` (JSON mode) |

---

### Testing with `xllify-lua`

Test files use the `xllify.TestSuite` / `xllify.Test` / `xllify.Assert` API.

**Example test file** (`tests/bmi_test.luau`):

```lua
xllify.TestSuite("BMI Tests", function()

    xllify.Test("normal BMI", function()
        local result = xllify.Call("BMI", 70, 1.75)
        xllify.Assert.Equal(22.857143, result)
    end)

    xllify.Test("returns a number", function()
        local result = xllify.Call("BMI", 80, 1.80)
        xllify.Assert.Equal(false, result == nil)
    end)

end)
```

**Run tests:**

```bash
# Human-readable
xllify-lua --load functions/bmi.luau test tests/bmi_test.luau

# JSON output
xllify-lua --load functions/bmi.luau --json test tests/bmi_test.luau

# JUnit XML (for CI)
xllify-lua --load functions/bmi.luau --junit test tests/bmi_test.luau > test-results.xml
```

**Assertion helpers:**

| Assertion | Description |
|-----------|-------------|
| `xllify.Assert.Equal(expected, actual)` | Deep equality; numbers use epsilon (`1e-9`) |
| `xllify.Assert.IsMatrix(value)` | Assert value is a 2D array |
| `xllify.Assert.Throws(fn)` | Assert function raises an error |
| `xllify.Assert.IsNumber(value)` | Assert value is a number |
| `xllify.Assert.True(condition, message?)` | Assert condition is truthy |

**Utility:**

```lua
local dims = xllify.GetDimensions(matrix)  -- returns {rows, cols}
```

---

## Luau `xllify` namespace

Every `.luau` function file runs inside the xllify runtime with a global `xllify` table pre-populated with these functions, forming a small standard library. Here is the full API.

### Registering functions

#### `xllify.ExcelFunction(config, fn)` / `xllify.fn(config, fn)`

Register a Luau function as an Excel-callable function.

```lua
xllify.ExcelFunction({
    name = "BMI",
    description = "Calculate Body Mass Index",
    category = "Health",
    parameters = {
        { name = "weight_kg", type = "number", description = "Weight in kilograms" },
        { name = "height_m",  type = "number", description = "Height in metres" },
    },
    result = { type = "number", dimensionality = "scalar" },
}, function(weight_kg, height_m)
    return weight_kg / (height_m * height_m)
end)
```

**`config` fields:**

| Field | Type | Description |
|-------|------|-------------|
| `name` | `string` | Function name as it appears in Excel (required) |
| `description` | `string?` | Shown in the Excel function wizard |
| `category` | `string?` | Category grouping in the wizard |
| `execution_type` | `"sync"\|"async"` | Defaults to `"sync"` |
| `parameters` | `param[]?` | Parameter descriptors (see below) |
| `result` | `result?` | Return type descriptor |

**Parameter descriptor:**

| Field | Type | Description |
|-------|------|-------------|
| `name` | `string` | Parameter name |
| `type` | `"number"\|"string"\|"boolean"\|"table"` | Value type |
| `description` | `string?` | Shown in Excel wizard |
| `dimensionality` | `"scalar"\|"matrix"` | Scalar (default) or 2D range |

**Result descriptor:**

| Field | Type | Description |
|-------|------|-------------|
| `type` | `"number"\|"string"\|"boolean"\|"any"` | Return type |
| `dimensionality` | `"scalar"\|"matrix"` | Scalar (default) or 2D array |

### Built-in utilities

These helpers are always available inside your Luau files:

#### JSON

```lua
local tbl, err = xllify.json_parse('{"key": 42}')   -- table or nil + error string
local str       = xllify.json_stringify(tbl)          -- JSON string
```

#### Text

```lua
xllify.strip_html("<b>hello</b>")            -- "hello"
xllify.spell_number(1234.5)                  -- "one thousand two hundred thirty-four point five"
xllify.regex_match(text, pattern, n?)        -- nth match (1-indexed, default 1), ECMAScript regex
xllify.regex_replace(text, pattern, repl)    -- replace matches
```

#### String distance

```lua
xllify.levenshtein_distance("kitten", "sitting")  -- 3
```

#### Introspection

```lua
xllify.version          -- runtime version string, e.g. "0.9.10"
xllify._namespace       -- namespace prefix applied to registered names
xllify.GetRegisteredFunctions()  -- metadata table for all registered functions
```
