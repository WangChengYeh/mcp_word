# MCP Word Proxy Server

This project provides a Node.js proxy server for an Office Word Add-in, enabling real-time interaction between the Task Pane client via WebSocket and integrating with Claude MCP SDK (or CLI) to process `EditTask` requests.

---

## 1. Installation

1. Install Node.js (v16+ recommended)
2. In the project root, install dependencies:
  ```bash
  npm install
  ```

## 2. Configuration

Obtain your Claude API key and set the environment variable:

```bash
export CLAUDE_API_KEY=your_api_key_here
```

## 3. Start the Server

```bash
npm start
```
By default, the server listens on http://localhost:3000

## 4. Trigger EditTask via Claude CLI

You can send edit requests through the Anthropic CLI (`claude`) and forward them to the local proxy:

```bash
claude --model=claude-2 --json -p '{"task":"EditTask","content":"Please rewrite the following text as formal business correspondence: Thank you for your letter."}' \
  | curl -s -X POST http://localhost:3000/mcp \
    -H 'Content-Type: application/json' -d @-
```

Alternatively, create a simple script `edit.sh`:
```bash
#!/usr/bin/env bash
prompt="$1"
claude --model=claude-2 --json -p "{\"task\":\"EditTask\",\"content\":\"$prompt\"}" \
  | curl -s -X POST http://localhost:3000/mcp -H 'Content-Type: application/json' -d @-
```

Make the script executable:
```bash
chmod +x edit.sh
```
Run:
```bash
./edit.sh "Provide a concise summary of this week's meeting conclusions."
```

## 5. Example Response

```bash
$ ./edit.sh "Please convert the following text into a professional business letter: We have received your payment."
{
  "status": "ok",
  "result": {
    "content": "Dear Customer,\n\nThank you for your payment. We have received the funds and will proceed with the next steps shortly.\n\nSincerely,\nYour Company"
  }
}
```
The Task Pane Add-in will receive the same JSON over WebSocket and insert the content into the Word document.

接著，Office Add-in Task Pane 便會透過 WebSocket 接收同樣的 JSON，並在 Word 文件中插入結果。
