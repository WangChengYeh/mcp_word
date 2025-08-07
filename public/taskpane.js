Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // 建立 WebSocket 連線
    const socket = io();

    // 監聽 ai-cmd 事件
    socket.on('ai-cmd', async (payload) => {
      try {
        await Word.run(async (context) => {
          const body = context.document.body;
          // 插入接收到的文字內容
          const text = payload.content || JSON.stringify(payload);
          body.insertText(text, Word.InsertLocation.end);
          await context.sync();
        });
      } catch (err) {
        console.error('Error inserting text:', err);
      }
    });
  }
});
