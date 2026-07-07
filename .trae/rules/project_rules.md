# 项目规则

## 每次修复/新功能完成后的流程

每次完成一个 bug 修复或新功能实现后，必须执行以下步骤：

1. **更新 README.md**：在 `## 版本变更` 最上方新增版本记录，包含版本号、日期、变更摘要
2. **更新版本号**：修改 `app/main.py` 中的 `__VERSION__`
3. **提交代码**：`git add` + `git commit`（中文描述，标注 feat: / fix: / refactor: 前缀 + 版本号）
4. **推送到远程**：`git push`

### 提交信息格式

```
feat: 新功能描述 vX.Y.Z
fix: 修复描述
refactor: 重构描述
```

其中 version bump 提交附带版本号，小修复可不带。
