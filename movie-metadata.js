// ==============================
// SheetJS(XLSX) 사용
// ==============================
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";

// API Keys
const OMDB_KEY = "ce0ca871";
const KMDB_KEY = "66C7AXU2KBQEJ6Y5LX1U";

const CORS = "https://api.allorigins.win/raw?url=";
const isBrowser = typeof window !== "undefined";

function wrap(url) {
  return isBrowser ? CORS + encodeURIComponent(url) : url;
}

// ==============================
// 실행 엔트리
// ==============================
export default async function run(input) {
  if (input.movieTitle) {
    const meta = await getMovieMetadata(input.movieTitle);
    return {
      metadata: meta,
      excelFile: createExcel([meta])
    };
  }

  if (input.files && input.files.length > 0) {
    const file = input.files[0];
    const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // 첫 번째 열에서 제목 추출
    const titles = rows.map(r => r[0]).filter(Boolean);

    const results = [];
    for (const t of titles) {
      const meta = await getMovieMetadata(t);
      results.push(meta);
    }

    return {
      metadata: results,
      excelFile: createExcel(results)
    };
  }

  return { error: "제목 또는 파일이 필요합니다." };
}

// ==============================
// 영화 메타데이터 조회 (KMDB → OMDb 순)
// ==============================
async function getMovieMetadata(title) {
  const kmdb = await fetchFromKMDB(title);
  if (kmdb && !kmdb.error) return kmdb;

  const omdb = await fetchFromOMDb(title);
  if (omdb && !omdb.error) return omdb;

  return { source: "N/A", title, error: "메타데이터를 찾을 수 없습니다." };
}

// ==============================
// KMDB 조회
// ==============================
async function fetchFromKMDB(title) {
  const url = wrap(
    `https://api.koreafilm.or.kr/openapi-data2/wisenut/search_api/search_json2.jsp?collection=kmdb_new2&detail=Y&query=${encodeURIComponent(
      title
    )}&ServiceKey=${KMDB_KEY}`
  );

  try {
    const res = await fetch(url);
    const json = await res.json();

    if (!json.Data || !json.Data[0]?.Result?.length) return { error: "KMDB 없음" };

    const m = json.Data[0].Result[0];

    let actors = m.actors?.actor?.map(a => a.actorNm) || [];
    actors = actors.slice(0, 4).join(", ");

    return {
      source: "KMDB",
      title: cleanTitle(m.title),
      englishTitle: (m.titleEng || "").toUpperCase(),
      year: m.prodYear || "",
      director: m.directors?.director?.[0]?.directorNm || "",
      cast: actors,
      genre: m.genre || "",
      rating: normalizeRating(m.rating),
      plot: m.plots?.plot?.[0]?.plotText || "",
      country: m.nation || "",
      releaseDate: m.repRlsDate || ""
    };
  } catch {
    return { error: "KMDB 오류" };
  }
}

// ==============================
// OMDb 조회
// ==============================
async function fetchFromOMDb(title) {
  const url = wrap(
    `https://www.omdbapi.com/?t=${encodeURIComponent(title)}&apikey=${OMDB_KEY}&plot=full`
  );

  try {
    const res = await fetch(url);
    const data = await res.json();

    if (data.Response === "False") return { error: "OMDb 검색 실패" };

    let actors = (data.Actors || "").split(",").map(a => a.trim());
    actors = actors.slice(0, 4).join(", ");

    return {
      source: "OMDb",
      title: data.Title || "",
      englishTitle: (data.Title || "").toUpperCase(),
      year: data.Year || "",
      director: data.Director || "",
      cast: actors,
      genre: data.Genre || "",
      rating: normalizeRating(data.Rated),
      plot: data.Plot || "",
      country: data.Country || "",
      releaseDate: data.Released || ""
    };
  } catch {
    return { error: "OMDb 오류" };
  }
}

// ==============================
// 제목 clean
// ==============================
function cleanTitle(t) {
  return (t || "").replace(/!HS|!HE/g, "").trim();
}

// ==============================
// 등급 변환 로직
// ==============================
function normalizeRating(r) {
  if (!r) return "";
  r = r.replace("관람가", "").trim();

  if (r.includes("청불") || r.includes("청소년")) return "청불";
  if (r.includes("19") || r.includes("R")) return "18세";
  if (r.includes("15")) return "15세";
  if (r.includes("12") || r.includes("PG-13")) return "12세";

  return r;
}

// ==============================
// Excel 생성 (SheetJS)
// ==============================
function createExcel(list) {
  const header = [
    "Source",
    "Title",
    "English Title",
    "Year",
    "Director",
    "Cast",
    "Genre",
    "Rating",
    "Plot",
    "Country",
    "Release Date"
  ];

  const rows = list.map(m => [
    m.source,
    m.title,
    m.englishTitle,
    m.year,
    m.director,
    m.cast,
    m.genre,
    m.rating,
    m.plot,
    m.country,
    m.releaseDate
  ]);

  const worksheet = XLSX.utils.aoa_to_sheet([header, ...rows]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Movies");

  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}
