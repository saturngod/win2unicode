import { useMemo, useState, useEffect, useRef } from "react";
import { invoke } from "@tauri-apps/api/core";
import { open, save } from "@tauri-apps/plugin-dialog";
import { listen, type UnlistenFn } from "@tauri-apps/api/event";
import "./App.css";

const SUPPORTED_EXTENSIONS = ["txt", "docx", "xlsx", "pptx"];

interface ConversionProgress {
  current: number;
  total: number;
  percentage: number;
  message: string;
}

function App() {
  const [sourceFont, setSourceFont] = useState("Win Innwa");
  const [selectedFile, setSelectedFile] = useState<string | null>(null);
  const [status, setStatus] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<ConversionProgress | null>(null);
  const [page, setPage] = useState<"file" | "text">("file");
  const [winText, setWinText] = useState("");
  const [unicodeText, setUnicodeText] = useState("");
  const [textBusy, setTextBusy] = useState(false);
  const unlistenRef = useRef<UnlistenFn | null>(null);

  const fileLabel = useMemo(() => {
    if (!selectedFile) return "No file selected.";
    const parts = selectedFile.split(/[\\/]/);
    return parts[parts.length - 1] ?? selectedFile;
  }, [selectedFile]);

  // Setup progress listener
  useEffect(() => {
    const setupListener = async () => {
      const unlisten = await listen<ConversionProgress>("conversion-progress", (event) => {
        setProgress(event.payload);
        console.log(`Progress: ${event.payload.percentage.toFixed(1)}% - ${event.payload.message}`);
      });
      unlistenRef.current = unlisten;
    };

    setupListener();

    return () => {
      if (unlistenRef.current) {
        unlistenRef.current();
      }
    };
  }, []);

  // Reset progress when not busy
  useEffect(() => {
    if (!busy) {
      setTimeout(() => setProgress(null), 2000);
    }
  }, [busy]);

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
    setProgress(null);
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

      setProgress({ current: 0, total: 1, percentage: 0, message: "Starting conversion..." });

      await invoke("convert_file", {
        sourcePath: selectedFile,
        targetPath: target,
        sourceFont,
      });

      setStatus("Conversion completed successfully!");
      setProgress({ current: 1, total: 1, percentage: 100, message: "Done!" });
    } catch (err) {
      setStatus(err instanceof Error ? err.message : "Conversion failed.");
    } finally {
      setBusy(false);
    }
  }

  useEffect(() => {
    if (page !== "text") return;
    let cancelled = false;
    const timer = setTimeout(async () => {
      setTextBusy(true);
      try {
        const result = await invoke<string>("convert_text", { input: winText });
        if (!cancelled) {
          setUnicodeText(result);
        }
      } catch (err) {
        if (!cancelled) {
          setUnicodeText(err instanceof Error ? err.message : "Conversion failed.");
        }
      } finally {
        if (!cancelled) {
          setTextBusy(false);
        }
      }
    }, 150);

    return () => {
      cancelled = true;
      clearTimeout(timer);
    };
  }, [winText, page]);

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
        <div className="header-actions">
          <button
            type="button"
            className={page === "file" ? "secondary" : "ghost"}
            onClick={() => setPage("file")}
          >
            File Converter
          </button>
          <button
            type="button"
            className={page === "text" ? "secondary" : "ghost"}
            onClick={() => setPage("text")}
          >
            Text Converter
          </button>
        </div>
      </header>

      {page === "file" ? (
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

            {progress && (
              <div className="progress-container">
                <div className="progress-bar-wrapper">
                  <div className="progress-bar">
                    <div
                      className="progress-fill"
                      style={{ width: `${progress.percentage}%` }}
                    />
                  </div>
                  <span className="progress-percentage">
                    {progress.percentage.toFixed(1)}%
                  </span>
                </div>
                <p className="progress-message">
                  {progress.message}
                </p>
              </div>
            )}

            {status && !progress && <p className="status">{status}</p>}
          </div>
        </section>
      ) : (
        <section className="panel text-panel">
          <div className="text-grid">
            <label className="field">
              <span>Win Innwa</span>
              <textarea
                value={winText}
                onChange={(e) => setWinText(e.currentTarget.value)}
                placeholder="Type Win Innwa text here..."
              />
            </label>
            <label className="field">
              <span>Unicode</span>
              <textarea
                value={unicodeText}
                readOnly
                placeholder="Unicode output..."
              />
            </label>
          </div>
          <div className="text-status">
            <p className="hint">
              {textBusy ? "Converting..." : "Instant conversion as you type."}
            </p>
          </div>
        </section>
      )}
    </main>
  );
}

export default App;
