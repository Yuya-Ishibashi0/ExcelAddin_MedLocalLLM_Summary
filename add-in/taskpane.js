/* global Office, Excel */
(() => {
  const API_URL = "http://127.0.0.1:8787/chat";
  const STREAM_URL = "http://127.0.0.1:8787/chat/stream";
  const TIMEOUT_MS = 120000;

  let lastAnswer = "";
  let isBusy = false;
  let selectionContext = "";

  const sendBtn = document.getElementById("send-btn");
  const writeBtn = document.getElementById("write-btn");
  const statusEl = document.getElementById("status");
  const chatLogEl = document.getElementById("chat-log");
  const inputEl = document.getElementById("user-input");
  const contextBtn = document.getElementById("context-btn");
  const contextToggle = document.getElementById("context-toggle");
  const contextStatusEl = document.getElementById("context-status");
  const inferenceToggle = document.getElementById("inference-toggle");

  class ApiError extends Error {
    constructor(status, message) {
      super(message);
      this.status = status;
    }
  }

  Office.onReady(() => {
    sendBtn.addEventListener("click", onSend);
    writeBtn.addEventListener("click", onWrite);
    contextBtn.addEventListener("click", onUpdateContext);
    contextToggle.addEventListener("change", updateContextStatus);
  });

  function setBusy(busy) {
    isBusy = busy;
    sendBtn.disabled = busy;
    contextBtn.disabled = busy;
    writeBtn.disabled = busy || !lastAnswer;
  }

  function setStatus(message, type = "") {
    statusEl.textContent = message;
    statusEl.parentElement.classList.toggle("error", type === "error");
  }

  function isInferenceModeEnabled() {
    return Boolean(inferenceToggle && inferenceToggle.checked);
  }

  function valuesToTSV(values) {
    if (!Array.isArray(values)) {
      return "";
    }
    return values
      .map((row) =>
        row
          .map((cell) => {
            if (cell === null || cell === undefined) {
              return "";
            }
            return String(cell);
          })
          .join("\t")
      )
      .join("\n");
  }

  async function getSelectionTSV() {
    return Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load("values");
      await context.sync();
      return valuesToTSV(range.values);
    });
  }

  async function callLocalApi(payload) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), TIMEOUT_MS);

    try {
      const response = await fetch(API_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });

      if (!response.ok) {
        let detail = "";
        try {
          const data = await response.json();
          detail = data.detail || data.error || "";
        } catch {
          detail = "";
        }
        throw new ApiError(response.status, detail || response.statusText);
      }

      const data = await response.json();
      return typeof data.answer === "string" ? data.answer : "";
    } catch (err) {
      if (err.name === "AbortError") {
        throw new ApiError(504, "応答が遅延しています。モデルを軽くする/入力範囲を減らす…");
      }
      if (err instanceof TypeError) {
        throw new ApiError(
          0,
          "ローカルAPIが起動していません（port 8787）。起動コマンド: uvicorn server:app --host 127.0.0.1 --port 8787"
        );
      }
      throw err;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  async function callLocalApiStream(payload, onChunk) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), TIMEOUT_MS);

    try {
      const response = await fetch(STREAM_URL, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
        signal: controller.signal,
      });

      if (!response.ok) {
        let detail = "";
        try {
          const data = await response.json();
          detail = data.detail || data.error || "";
        } catch {
          detail = "";
        }
        throw new ApiError(response.status, detail || response.statusText);
      }

      if (!response.body) {
        throw new ApiError(0, "ストリーミングに失敗しました。");
      }

      const reader = response.body.getReader();
      const decoder = new TextDecoder("utf-8");

      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        const chunk = decoder.decode(value, { stream: true });
        if (chunk) {
          onChunk(chunk);
        }
      }
    } catch (err) {
      if (err.name === "AbortError") {
        throw new ApiError(504, "応答が遅延しています。モデルを軽くする/入力範囲を減らす…");
      }
      if (err instanceof TypeError) {
        throw new ApiError(
          0,
          "ローカルAPIが起動していません（port 8787）。起動コマンド: uvicorn server:app --host 127.0.0.1 --port 8787"
        );
      }
      throw err;
    } finally {
      clearTimeout(timeoutId);
    }
  }

  function appendMessage(role, content) {
    const messageEl = document.createElement("div");
    messageEl.className = `message ${role}`;

    const roleEl = document.createElement("div");
    roleEl.className = "role";
    roleEl.textContent = role === "user" ? "あなた" : "アシスタント";

    const contentEl = document.createElement("div");
    contentEl.className = "content";
    contentEl.textContent = content;

    messageEl.appendChild(roleEl);
    messageEl.appendChild(contentEl);
    chatLogEl.appendChild(messageEl);
    chatLogEl.scrollTop = chatLogEl.scrollHeight;
    return { messageEl, contentEl };
  }

  function updateContextStatus() {
    const hasContext = selectionContext.trim().length > 0;
    const contextText = hasContext ? `設定済み (${selectionContext.length}文字)` : "未設定";
    const toggleText = contextToggle.checked ? "送信ON" : "送信OFF";
    contextStatusEl.textContent = `${contextText} / ${toggleText}`;
  }

  async function onUpdateContext() {
    if (isBusy) return;

    setBusy(true);
    setStatus("コンテキスト更新中…");

    try {
      const tsv = await getSelectionTSV();
      if (!tsv.trim()) {
        setStatus("選択範囲が空です。セルを選択してください。", "error");
        return;
      }
      selectionContext = tsv;
      updateContextStatus();
      setStatus("コンテキスト更新完了");
    } catch (err) {
      setStatus("コンテキスト取得に失敗しました。", "error");
    } finally {
      setBusy(false);
    }
  }

  async function onSend() {
    if (isBusy) return;

    const text = inputEl.value.trim();
    if (!text) {
      setStatus("メッセージを入力してください。", "error");
      return;
    }

    appendMessage("user", text);
    inputEl.value = "";

    setBusy(true);
    setStatus("準備中…");

    try {
      setStatus("生成中…");
      const payload = {
        messages: [{ role: "user", content: text }],
        selection_context: contextToggle.checked ? selectionContext : "",
        model: null,
        max_tokens: null,
        inference_mode: isInferenceModeEnabled(),
      };

      const { contentEl } = appendMessage("assistant", "");
      let answer = "";

      await callLocalApiStream(payload, (chunk) => {
        answer += chunk;
        contentEl.textContent = answer;
        chatLogEl.scrollTop = chatLogEl.scrollHeight;
      });

      if (!answer) {
        contentEl.textContent = "(結果が空です)";
      }

      lastAnswer = answer;
      writeBtn.disabled = !lastAnswer;
      setStatus("完了");
    } catch (err) {
      const lastMessage = chatLogEl.lastElementChild;
      if (lastMessage && lastMessage.classList.contains("assistant")) {
        lastMessage.remove();
      }
      const message = resolveErrorMessage(err);
      setStatus(message, "error");
    } finally {
      setBusy(false);
    }
  }

  async function onWrite() {
    if (isBusy) return;
    if (!lastAnswer) {
      setStatus("書き込み内容がありません。", "error");
      return;
    }

    setBusy(true);
    setStatus("選択セルへ書き込み中…");

    try {
      await Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange().getCell(0, 0);
        range.values = [[lastAnswer]];
        range.select();
        await context.sync();
      });
      setStatus("書き込み完了");
    } catch (err) {
      setStatus("書き込みに失敗しました。", "error");
    } finally {
      setBusy(false);
    }
  }

  function resolveErrorMessage(err) {
    if (err instanceof ApiError) {
      if (err.status === 503) {
        return "Ollamaが起動していません。ollama serve を実行してください";
      }
      if (err.status === 404) {
        return "モデルが見つかりません。ollama pull <model> を実行してください";
      }
      return err.message || "エラーが発生しました。";
    }

    return "予期しないエラーが発生しました。";
  }

  updateContextStatus();
})();
