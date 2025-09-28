async function postForm(url, formData) {
  const resp = await fetch(url, { method: "POST", body: formData });
  return resp.json();
}

document.getElementById("upload-form").addEventListener("submit", async (e) => {
  e.preventDefault();
  const input = document.getElementById("file-input");
  if (!input.files || input.files.length === 0) {
    alert("請選擇一個 Excel 檔案");
    return;
  }
  const file = input.files[0];
  const fd = new FormData();
  fd.append("file", file);

  const resultDiv = document.getElementById("upload-result");
  resultDiv.textContent = "Uploading...";

  try {
    const res = await postForm("/api/upload", fd);
    if (res.ok) {
      resultDiv.textContent = `成功插入 ${res.inserted} 筆`;
    } else {
      resultDiv.textContent = `錯誤: ${res.error || JSON.stringify(res)}`;
    }
  } catch (err) {
    resultDiv.textContent = "Upload failed: " + err.message;
  }
});

document.getElementById("list-btn").addEventListener("click", async () => {
  const resp = await fetch("/api/items");
  const data = await resp.json();
  const container = document.getElementById("items-table");
  const countDiv = document.getElementById("items-count");
  countDiv.textContent = `共 ${data.count} 筆（最多顯示 1000）`;

  if (!data.items || data.items.length === 0) {
    container.innerHTML = "<p>目前沒有資料</p>";
    return;
  }

  // Build a table dynamically with keys from first item
  const keys = Object.keys(data.items[0]);
  let html = "<table><thead><tr>";
  keys.forEach(k => html += `<th>${k}</th>`);
  html += "</tr></thead><tbody>";
  data.items.forEach(item => {
    html += "<tr>";
    keys.forEach(k => html += `<td>${item[k] === null || item[k] === undefined ? "" : item[k]}</td>`);
    html += "</tr>";
  });
  html += "</tbody></table>";
  container.innerHTML = html;
});

document.getElementById("clear-btn").addEventListener("click", async () => {
  if (!confirm("這會刪除整個 collection 的所有資料，確定要執行嗎？")) return;
  const resultDiv = document.getElementById("clear-result");
  resultDiv.textContent = "Clearing...";
  try {
    const resp = await fetch("/api/clear", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ confirm: true })
    });
    const data = await resp.json();
    if (data.ok) {
      resultDiv.textContent = data.message || "清除成功";
      // refresh list
      document.getElementById("items-table").innerHTML = "";
      document.getElementById("items-count").textContent = "共 0 筆";
    } else {
      resultDiv.textContent = "Error: " + (data.error || JSON.stringify(data));
    }
  } catch (err) {
    resultDiv.textContent = "Clear failed: " + err.message;
  }
});