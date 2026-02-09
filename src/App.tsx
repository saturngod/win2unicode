import { useMemo, useState } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import "./App.css";

const SUPPORTED_EXTENSIONS = ["txt", "docx", "xlsx", "pptx"];

function App() {
  const [sourceFont, setSourceFont] = useState("Win Innwa");
  const [selectedFile, setSelectedFile] = useState<string | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const fileLabel = useMemo(() => {
    if (!selectedFile) return "No file selected.";
    const parts = selectedFile.split(/[\\/]/);
    return parts[parts.length - 1] ?? selectedFile;
  }, [selectedFile]);

  async function pickFile() {
    setStatus(null);
    const result = await open({
      multiple: false,
      filters: [
        {
          name: "Documents",
          extensions: SUPPORTED_EXTENSIONS,
        },
      ],
    });

    if (typeof result === "string") {
      setSelectedFile(result);
    }
  }

  async function convertNow() {
    if (!selectedFile) return;
    setStatus(null);
    setBusy(true);

    try {
      const defaultName = fileLabel.replace(/\.[^.]+$/, "");
      const ext = selectedFile.split(".").pop()?.toLowerCase() ?? "";
      const suggestedName = ext ? `${defaultName}.${ext}` : defaultName;

      const target = await save({
        defaultPath: suggestedName,
        filters: [
          {
            name: "Documents",
            extensions: ext && SUPPORTED_EXTENSIONS.includes(ext) ? [ext] : SUPPORTED_EXTENSIONS,
          },
        ],
      });

      if (!target) {
        setBusy(false);
        return;
      }

      await invoke("convert_file", {
        sourcePath: selectedFile,
        targetPath: target,
        sourceFont,
      });

      setStatus("Conversion completed.");
    } catch (err) {
      setStatus(err instanceof Error ? err.message : "Conversion failed.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="app">
      <header className="header">
        <div>
          <p className="eyebrow">Win to Unicode Converter</p>
          <h1>Win Font to Myanmar Unicode</h1>
          <p className="sub">
            Select a file and replace any Win font family with Myanmar Unicode.
          </p>
        </div>
      </header>

      <section className="panel">
        <label className="field">
          <span>Font Family To Replace</span>
          <input
            value={sourceFont}
            onChange={(e) => setSourceFont(e.currentTarget.value)}
            placeholder="Win Innwa"
          />
        </label>

        <div className="field">
          <span>Selected File</span>
          <div className="file-row">
            <div className="file-name">{fileLabel}</div>
            <button type="button" onClick={pickFile} className="secondary">
              Select File
            </button>
          </div>
          <p className="hint">Supported: txt, docx, xlsx, pptx</p>
        </div>

        <div className="actions">
          <button
            type="button"
            onClick={convertNow}
            disabled={!selectedFile || busy}
          >
            {busy ? "Converting..." : "Convert Now"}
          </button>
          {status && <p className="status">{status}</p>}
        </div>
      </section>
    </main>
  );
}

export default App;
