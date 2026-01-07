import run from "./movie-metadata.js";

const titleInput = document.getElementById("movie-title");
const fileInput = document.getElementById("movie-file");
const statusEl = document.getElementById("status");
const resultEl = document.getElementById("result");
const downloadLink = document.getElementById("download");
const searchTitleBtn = document.getElementById("search-title");
const searchFileBtn = document.getElementById("search-file");
const resetBtn = document.getElementById("reset");

function setStatus(message, isError = false) {
  statusEl.textContent = message;
  statusEl.classList.toggle("error", isError);
}

function resetOutput() {
  resultEl.innerHTML = "";
  downloadLink.hidden = true;
  downloadLink.removeAttribute("href");
  downloadLink.removeAttribute("download");
}

function setLoading(isLoading) {
  searchTitleBtn.disabled = isLoading;
  searchFileBtn.disabled = isLoading;
}

function normalizeMetadata(metadata) {
  if (!metadata) return [];
  if (Array.isArray(metadata)) return metadata;
  return [metadata];
}

function renderTable(metadataList) {
  if (!metadataList.length) {
    resultEl.innerHTML = "<p>결과가 없습니다.</p>";
    return;
  }

  const rows = metadataList.map(item => {
    const error = item.error ? `<div class="error">${item.error}</div>` : "";
    return `
      <tr>
        <td><span class="badge">${item.source ?? "N/A"}</span></td>
        <td>${item.title ?? ""} ${error}</td>
        <td>${item.englishTitle ?? ""}</td>
        <td>${item.year ?? ""}</td>
        <td>${item.director ?? ""}</td>
        <td>${item.cast ?? ""}</td>
        <td>${item.genre ?? ""}</td>
        <td>${item.rating ?? ""}</td>
        <td>${item.plot ?? ""}</td>
        <td>${item.country ?? ""}</td>
        <td>${item.releaseDate ?? ""}</td>
      </tr>
    `;
  }).join("");

  resultEl.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Source</th>
          <th>Title</th>
          <th>English Title</th>
          <th>Year</th>
          <th>Director</th>
          <th>Cast</th>
          <th>Genre</th>
          <th>Rating</th>
          <th>Plot</th>
          <th>Country</th>
          <th>Release Date</th>
        </tr>
      </thead>
      <tbody>
        ${rows}
      </tbody>
    </table>
  `;
}

function setDownload(fileBuffer) {
  const blob = new Blob([fileBuffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const url = URL.createObjectURL(blob);
  downloadLink.href = url;
  downloadLink.download = `movie-metadata-${Date.now()}.xlsx`;
  downloadLink.hidden = false;
}

async function handleSearch(payload) {
  resetOutput();
  setLoading(true);
  setStatus("조회 중...");

  try {
    const output = await run(payload);
    if (output.error) {
      setStatus(output.error, true);
      return;
    }

    const metadataList = normalizeMetadata(output.metadata);
    renderTable(metadataList);
    if (output.excelFile) {
      setDownload(output.excelFile);
    }
    setStatus(`총 ${metadataList.length}건 조회 완료.`);
  } catch (error) {
    setStatus(`오류가 발생했습니다: ${error.message}`, true);
  } finally {
    setLoading(false);
  }
}

searchTitleBtn.addEventListener("click", () => {
  const title = titleInput.value.trim();
  if (!title) {
    setStatus("영화 제목을 입력하세요.", true);
    return;
  }
  handleSearch({ movieTitle: title });
});

searchFileBtn.addEventListener("click", () => {
  const file = fileInput.files[0];
  if (!file) {
    setStatus("엑셀 파일을 선택하세요.", true);
    return;
  }
  handleSearch({ files: [file] });
});

resetBtn.addEventListener("click", () => {
  titleInput.value = "";
  fileInput.value = "";
  resetOutput();
  setStatus("대기 중...");
});
