# Interview Presentation Template

使用 PptxGenJS 由 JSON 資料產生 16:9 面試簡報。

## 使用

```bash
npm install
cp profile.example.json profile.json
# 編輯 profile.json
npm run build
```

可透過環境變數指定檔案：

```bash
PRESENTATION_PROFILE=/path/to/profile.json \
PRESENTATION_OUTPUT=/path/to/interview.pptx \
npm run build
```

`profile.json` 與產出的 `.pptx` 可能含個資或公司資訊，均不會提交至 Git。
