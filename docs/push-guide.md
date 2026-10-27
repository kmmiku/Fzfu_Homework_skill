# 一键推送 Fzfu_Homework_skill 到 GitHub

仓库目录：`D:\DSH_Wrodkstation\Git_work\Fzfu_Homework_skill`

## 使用方法

把下面的 `YOUR_PAT` 换成你的 Personal Access Token（需勾选 `repo` 权限），然后执行：

```powershell
cd D:\DSH_Wrodkstation\Git_work\Fzfu_Homework_skill
$PAT = "YOUR_PAT"
git push "https://kmmiku:$PAT@github.com/kmmiku/Fzfu_Homework_skill.git" main
```

## 说明

- **PAT 生成**：https://github.com/settings/tokens → Generate new token (classic) → 勾选 `repo` → 生成后立即复制（只显示一次）。
- **代理**：部分网络环境 `github.com:443` 直连被阻断，需要走本地代理时加 `-c http.proxy=http://127.0.0.1:7897 -c https.proxy=http://127.0.0.1:7897`。
- **仓库不存在时**：GitHub 在首次 push 时不会自动建仓，需先在网页 https://github.com/new 创建空仓库
  `Fzfu_Homework_skill`（不要勾选 Add README / .gitignore / license）。
- **避免泄露**：命令行带 PAT 会进入 PowerShell 历史。用完后可执行
  `Remove-Item (Get-PSReadlineOption).HistorySavePath` 清理历史，或改用 `git credential-manager` 存储。

## 验证推送结果

```powershell
git ls-remote "https://kmmiku:$PAT@github.com/kmmiku/Fzfu_Homework_skill.git" main
```

返回一个 sha 即表示上传成功。