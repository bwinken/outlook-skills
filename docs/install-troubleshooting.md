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

## PowerShell 說 running scripts is disabled

skill 會自動改用不受 execution policy 限制的呼叫方式，不需要改系統設定。若機器啟用 AppLocker / WDAC 的 Constrained Language Mode，COM 會被擋，只有 `outlook-open-msg` 和 `outlook-status -SkipCom` 能用。
