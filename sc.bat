# 1. 确保你在 main 分支
git checkout main

# 2. 添加所有本地文件
git add .

# 3. 提交本地更改（必须先有本地提交）
git commit -m "强制同步本地网站文件"

# 4. 强制推送本地分支到远程（关键！）
git push --force origin main