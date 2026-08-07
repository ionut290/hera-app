from pathlib import Path

path = Path("app.js")
source = path.read_text(encoding="utf-8")
function_marker = "async function loadStartupCoreCollections() {"
function_start = source.find(function_marker)
if function_start < 0:
    raise SystemExit("Funzione loadStartupCoreCollections non trovata")

next_async = source.find("\nasync function ", function_start + len(function_marker))
next_plain = source.find("\nfunction ", function_start + len(function_marker))
candidates = [position for position in (next_async, next_plain) if position >= 0]
next_function = min(candidates) if candidates else len(source)
block = source[function_start:next_function]
marker = "await Promise.all(["
marker_pos = block.find(marker)
if marker_pos < 0:
    raise SystemExit("Promise.all iniziale non trovata nella funzione")

array_open = marker_pos + len("await Promise.all(")
depth = 0
array_close = None
for index in range(array_open, len(block)):
    if block[index] == "[":
        depth += 1
    elif block[index] == "]":
        depth -= 1
        if depth == 0:
            array_close = index
            break
if array_close is None:
    raise SystemExit("Chiusura array Promise.all non trovata")

statement_end = block.find(");", array_close)
if statement_end < 0:
    raise SystemExit("Fine Promise.all non trovata")
statement_end += 2

tasks_body = block[array_open + 1:array_close]
replacement = """const startupTasks = [""" + tasks_body + """];
    const startupTaskNames = [
      "personale",
      "qualifiche/corsi",
      "sicurezza",
      "squadre",
      "commesse",
      "mezzi"
    ];
    const startupTimeoutMs = 10000;
    const startupResults = await Promise.allSettled(
      startupTasks.map((task, index) => Promise.race([
        Promise.resolve(task),
        new Promise((resolve) => {
          setTimeout(() => {
            const sectionName = startupTaskNames[index] || `sezione-${index + 1}`;
            console.warn(`[Startup] caricamento parziale, sezione lenta: ${sectionName}`);
            resolve(false);
          }, startupTimeoutMs);
        })
      ]))
    );
    startupResults.forEach((result, index) => {
      if (result.status === "rejected") {
        const sectionName = startupTaskNames[index] || `sezione-${index + 1}`;
        console.warn(`[Startup] caricamento parziale, sezione non disponibile: ${sectionName}`, result.reason);
      }
    });
    const availableCommesse = Array.isArray(appState?.commesse)
      ? appState.commesse.length
      : (commesseById instanceof Map ? commesseById.size : 0);
    console.log(`[Startup] commesse iniziali pronte: ${availableCommesse}`);
    if (availableCommesse > 0) {
      console.log("[Startup] Home sbloccata con commesse disponibili");
      renderCommesseHomeList();
      updateHomeStatus();
    }"""

updated_block = block[:marker_pos] + replacement + block[statement_end:]
updated = source[:function_start] + updated_block + source[next_function:]
path.write_text(updated, encoding="utf-8")
print("Patch applicata a loadStartupCoreCollections")
