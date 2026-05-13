import * as vscode from "vscode";
import * as fs from "fs";
import * as path from "path";
import { spawn, exec } from "child_process";
import * as http from "http";
import * as os from "os";

function existsExe(p: string): boolean {
    try { return fs.statSync(p).isFile(); } catch { return false; }
}

function findChrome(): string | undefined {
    const candidates = [
        process.env["ProgramFiles"] && path.join(process.env["ProgramFiles"]!, "Google", "Chrome", "Application", "chrome.exe"),
        process.env["ProgramFiles(x86)"] && path.join(process.env["ProgramFiles(x86)"]!, "Google", "Chrome", "Application", "chrome.exe"),
        process.env["LOCALAPPDATA"] && path.join(process.env["LOCALAPPDATA"]!, "Google", "Chrome", "Application", "chrome.exe"),
    ].filter(Boolean) as string[];
    return candidates.find(existsExe);
}

function checkCdp(): Promise<boolean> {
    return new Promise((resolve) => {
        const req = http.get("http://localhost:9222/json/version", { timeout: 2000 }, (res) => {
            res.resume();
            resolve(res.statusCode === 200);
        });
        req.on("error", () => resolve(false));
        req.on("timeout", () => { req.destroy(); resolve(false); });
    });
}

export async function startChromeCdp(out: vscode.OutputChannel): Promise<void> {
    out.show(true);
    out.appendLine("\n=== Start Chrome with Remote Debugging (port 9222) ===");

    if (await checkCdp()) {
        out.appendLine("Chrome CDP is already running on :9222.");
        vscode.window.showInformationMessage("Chrome CDP is already running.");
        return;
    }

    const cfg = vscode.workspace.getConfiguration("d365FormExtractor");
    let chrome = cfg.get<string>("chromePath") || "";
    if (!chrome) chrome = findChrome() || "";
    if (!chrome) {
        const input = await vscode.window.showInputBox({
            prompt: "Could not auto-detect chrome.exe. Enter full path.",
            placeHolder: "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
            ignoreFocusOut: true,
        });
        if (!input) return;
        chrome = input;
        await cfg.update("chromePath", chrome, vscode.ConfigurationTarget.Global);
    }

    const warn = await vscode.window.showWarningMessage(
        "Chrome must be restarted with remote debugging enabled. All existing Chrome windows will be closed. Continue?",
        { modal: true },
        "Close Chrome and Launch",
        "Cancel"
    );
    if (warn !== "Close Chrome and Launch") return;

    out.appendLine("Closing existing Chrome processes...");
    await new Promise<void>((resolve) => {
        exec("taskkill /F /IM chrome.exe", () => resolve());
    });
    await new Promise((r) => setTimeout(r, 1500));

    // Chrome 136+ blocks --remote-debugging-port on the default user-data-dir.
    // Use a dedicated profile dir (workspace-local if available, else under temp).
    const ws = vscode.workspace.workspaceFolders?.[0]?.uri.fsPath;
    const userDataDir = ws
        ? path.join(ws, "_chrome_profile")
        : path.join(os.tmpdir(), "d365-chrome-profile");
    try { fs.mkdirSync(userDataDir, { recursive: true }); } catch { /* ignore */ }
    const args = [
        "--remote-debugging-port=9222",
        `--user-data-dir=${userDataDir}`,
        "--no-first-run",
        "--no-default-browser-check",
    ];
    out.appendLine(`Launching: "${chrome}" ${args.join(" ")}`);
    const p = spawn(chrome, args, { detached: true, stdio: "ignore", windowsHide: false });
    p.on("error", (e) => out.appendLine(`spawn error: ${e.message}`));
    p.unref();

    // Poll CDP
    for (let i = 0; i < 15; i++) {
        await new Promise((r) => setTimeout(r, 1000));
        if (await checkCdp()) {
            out.appendLine("Chrome CDP is up on :9222.");
            vscode.window.showInformationMessage(
                "Chrome launched. Sign in to D365 in the new window, then run 'D365: Extract Form'."
            );
            return;
        }
    }
    vscode.window.showErrorMessage("Chrome started but CDP did not respond on :9222. Check Chrome path.");
}
