# DeepSeek 本地配置

将本文件复制为同目录的 `DeepSeek配置.md`，再填写自己的 API Key。实际配置文件已由 Git 忽略，不会提交到 GitHub。

```json
{
  "apiKey": "",
  "model": "deepseek-flash"
}
```

保存后下一次解析立即生效，无需修改网页或重启服务。API 地址固定为官方 `https://api.deepseek.com/chat/completions`。

- 密钥优先级：页面临时填写 > `DEEPSEEK_API_KEY` 环境变量 > 本地配置文件。
- 模型优先级：`DEEPSEEK_MODEL` 环境变量 > 本地配置文件 > 默认值。
- 配置不作为提示词发送给模型，也不通过本机网页提供下载。
- HTML 数据导入和持久保存不需要 API Key；DeepSeek 仅用于粘贴文本解析。
