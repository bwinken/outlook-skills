# 安裝疑難排解

## marketplace add 報「network source differs from the one declared in settings」

```
/plugin marketplace remove outlook-skills
/plugin marketplace add bwinken/outlook-skills
```

還是失敗就到 `~/.claude/settings.json`（或專案的 `.claude/settings.json`、`settings.local.json`）刪掉 `extraKnownMarketplaces` 底下的 `outlook-skills`，再 add。

## 設了 HTTP_PROXY / HTTPS_PROXY 連不到 GitHub

先只讓 GitHub 走直連，在同一個視窗啟動 Claude Code：

```powershell
$env:NO_PROXY = "github.com,api.github.com,objects.githubusercontent.com"
claude
```

不行再只對這個視窗清掉 proxy：

```powershell
Remove-Item Env:HTTP_PROXY, Env:HTTPS_PROXY -ErrorAction SilentlyContinue
claude
```

cmd：`set HTTP_PROXY=`、`set HTTPS_PROXY=`。git 自己的 proxy：`git config --global --unset http.proxy`、`--unset https.proxy`。

## 完全連不到 GitHub：手動安裝

用任何方式把 repo 放到本機（`git clone`，或 GitHub 頁面 Code → Download ZIP），假設在 `C:\tools\outlook-skills`。二選一：

```
/plugin marketplace add C:\tools\outlook-skills
/plugin install outlook@outlook-skills
```

或不經 marketplace，複製成個人 skills（指令變成 `/outlook-status`，沒有 `outlook:` 前綴）：

```
python C:\tools\outlook-skills\install.py --claude
```

更新：重新下載或 `git pull`，再重跑同一個指令。

## PowerShell 相關

- skill 一律透過 `scripts/run.py` 執行 .ps1，不需要也不應該改 execution policy；被 GPO 擋時它會自動改用 script block 方式。
- Claude Code 的 PowerShell 工具被公司政策封鎖時，用 Bash 工具跑 `python .../run.py ...` 即可，效果相同。
- 0.1.2 起 .ps1 都帶 UTF-8 BOM。舊版在 Windows PowerShell 5.1 會把中文當 ANSI 讀，出現「Array index expression is missing」這類假的 parse error；更新 plugin 即可。
- AppLocker / WDAC 的 Constrained Language Mode 會擋 COM，只有 `outlook-open-msg` 和 `outlook-status -SkipCom` 能用。

## 郵件在 .pst 而不在 Exchange 信箱

`outlook-status` 會看到 Exchange 收件匣幾乎是空的、某個 PST store 有大量郵件。設成預設 store，之後搜尋、摘要、掃描都會用它：

```
python <plugin>/scripts/settings.py set store "20230731"
```

或單次用 `-Store "20230731"`、`-AllStores`。中文 Outlook 的資料夾名（收件匣、寄件備份）與英文名都可以用。
