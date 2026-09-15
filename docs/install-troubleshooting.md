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

## 執行環境

- 腳本全是 Python，只需要 `pip install pywin32`（COM 用）。沒有 PowerShell、execution policy、編碼的問題。
- 出現「pywin32 is not installed」就裝它；出現「Cannot start Outlook COM automation」表示不是 Classic Outlook，或 AppLocker / WDAC 擋了 COM，這時只有 `outlook-open-msg` 和 `outlook_status.py -SkipCom` 能用。
- Claude Code 的 PowerShell 工具被封鎖時無所謂，用 Bash 工具跑 `python ...` 即可。

## pywin32 與 64 位元

`pip install pywin32` 會依 Python 的位元數自動挑 wheel（64 位元 Python 抓 `win_amd64`，32 位元抓 `win32`，ARM 機器抓 `win_arm64`），Windows 本身是幾位元、Outlook 是幾位元都不影響，COM 是跨 process 的。pip 被擋時手動下載 wheel，先確認 Python 版本與位元：

```
python -c "import sys, struct; print(sys.version_info[:2], struct.calcsize('P')*8, 'bit')"
```

例如 `(3, 11) 64 bit` 就到 PyPI 抓檔名含 `cp311` 與 `win_amd64` 的 wheel，`pip install <檔案>.whl` 不需要網路。

## 郵件在 .pst 而不在 Exchange 信箱

`outlook-status` 會看到 Exchange 收件匣幾乎是空的、某個 PST store 有大量郵件。設成預設 store，之後搜尋、摘要、掃描都會用它：

```
python <plugin>/scripts/settings.py set store "20230731"
```

或單次用 `-Store "20230731"`、`-AllStores`。中文 Outlook 的資料夾名（收件匣、寄件備份）與英文名都可以用。
