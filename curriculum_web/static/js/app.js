(() => {
  const fileInput = document.getElementById("fileInput");
  const fileName = document.getElementById("fileName");
  const runBtn = document.getElementById("runBtn");
  const statusText = document.getElementById("statusText");
  const progress = document.getElementById("progress");
  const tabsEl = document.getElementById("tabs");
  const panelsEl = document.getElementById("panels");
  const emptyHint = document.getElementById("emptyHint");

  let selectedFile = null;

  function setBusy(busy) {
    runBtn.disabled = busy || !selectedFile;
    fileInput.disabled = busy;
    progress.hidden = !busy;
  }

  function clearResults() {
    tabsEl.innerHTML = "";
    panelsEl.innerHTML = "";
    emptyHint.classList.remove("hidden");
  }

  function escapeHtml(text) {
    return String(text)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;");
  }

  function renderResults(data) {
    clearResults();
    emptyHint.classList.add("hidden");

    const tabNames = data.tab_names || [];
    const tabs = data.tabs || {};

    tabNames.forEach((name, index) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "tab" + (index === 0 ? " active" : "");
      btn.setAttribute("role", "tab");
      btn.textContent = name;
      btn.addEventListener("click", () => activateTab(name));
      tabsEl.appendChild(btn);

      const panel = document.createElement("div");
      panel.className = "panel" + (index === 0 ? " active" : "");
      panel.dataset.tab = name;
      panel.setAttribute("role", "tabpanel");

      const lines = tabs[name] || [];
      if (!lines.length) {
        panel.innerHTML = '<span class="line INFO">표시할 내용이 없습니다.</span>';
      } else {
        panel.innerHTML = lines
          .map((ln) => {
            const tag = ln.tag || "INFO";
            const text = ln.text === "" ? " " : ln.text;
            return `<span class="line ${tag}">${escapeHtml(text)}</span>`;
          })
          .join("");
      }
      panelsEl.appendChild(panel);
    });
  }

  function activateTab(name) {
    tabsEl.querySelectorAll(".tab").forEach((btn) => {
      btn.classList.toggle("active", btn.textContent === name);
    });
    panelsEl.querySelectorAll(".panel").forEach((panel) => {
      panel.classList.toggle("active", panel.dataset.tab === name);
    });
  }

  fileInput.addEventListener("change", () => {
    const file = fileInput.files && fileInput.files[0];
    if (!file) {
      selectedFile = null;
      fileName.textContent = "선택된 파일 없음";
      runBtn.disabled = true;
      return;
    }
    selectedFile = file;
    fileName.textContent = file.name;
    runBtn.disabled = false;
    statusText.textContent = "파일 선택됨 — 검사 실행을 눌러주세요";
  });

  runBtn.addEventListener("click", async () => {
    if (!selectedFile) {
      alert("먼저 엑셀 파일을 선택하세요.");
      return;
    }

    setBusy(true);
    statusText.textContent = "검사 중...";

    const form = new FormData();
    form.append("file", selectedFile);

    try {
      const res = await fetch("/api/check", { method: "POST", body: form });
      const data = await res.json();

      if (!res.ok || !data.ok) {
        clearResults();
        statusText.textContent = "오류 발생";
        alert(data.error || "검사 중 오류가 발생했습니다.");
        return;
      }

      renderResults(data);
      statusText.textContent = data.status || "검사 완료";
    } catch (err) {
      clearResults();
      statusText.textContent = "오류 발생";
      alert("서버와 통신할 수 없습니다.\n" + (err && err.message ? err.message : err));
    } finally {
      setBusy(false);
    }
  });
})();
