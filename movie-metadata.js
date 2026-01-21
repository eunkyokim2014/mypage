const OMDB_KEY = "ce0ca871";
const KMDB_KEY = "66C7AXU2KBQEJ6Y5LX1U";
const CORS_PROXY = "https://api.allorigins.win/raw?url=";

const isBrowser = typeof window !== "undefined";

function cors(url) {
  return isBrowser ? `${CORS_PROXY}${encodeURIComponent(url)}` : url;
}

export default async function run(input) {

  // 단일 제목 검색
  if (input.movieTitle) {
    const meta = await getMetadata(input.movieTitle);
    return {
      metadata: meta,
      excelFile: await createExcel([meta])
    };
  }

  // 엑셀 파일 업로드
  if (input.files?.length) {
    const file = input.files[0];
    const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });

    const titles = rows.map(r => r[0]).filter(Boolean);

    const results = [];
    for (const title of titles) {
      results.push(await getMetadata(title));
    }

    return {
      metadata: results,
      excelFile: await createExcel(results)
    };
  }

  return { error: "제목 또는 파일이 필요합니다." };
}

// ======================================================
// 통합 검색 (KMDB → OMDb)
// ======================================================
async function getMetadata(title) {
  const cleanTitle = title.trim();

  const kmdb = await fetchKMDB(cleanTitle);
  if (kmdb && !kmdb.error) return kmdb;

  const omdb = await fetchOMDb(cleanTitle);
  if (omdb && !omdb.error) return omdb;

  return { title, error: "메타데이터를 찾을 수 없습니다." };
}

// ======================================================
// KMDB
// ======================================================
async function fetchKMDB(title) {
  const url =
    `https://api.koreafilm.or.kr/openapi-data2/wisenut/search_api/search_json2.jsp?collection=kmdb_new2&detail=Y&query=${encodeURIComponent(
      title
    )}&ServiceKey=${KMDB_KEY}`;

  const res = await fetch(cors(url));
  const data = await res.json();

  if (!data.Data?.[0]?.Result?.length) return { error: "KMDB 검색 실패" };

  const m = data.Data[0].Result[0];

  return {
    source: "KMDB",
    title: (m.title || "").replace(/!HS|!HE/g, "").trim(),
    englishTitle: ((m.titleEng || "").toUpperCase()),
    year: m.prodYear || "",
    director: m.directors?.director?.[0]?.directorNm || "",
    cast: (m.actors?.actor?.map(a => a.actorNm).slice(0, 4).join(", ")) || "",
    genre: m.genre || "",
    rating: cleanRating(m.rating || ""),
    plot: m.plots?.plot?.[0]?.plotText || "",
    country: m.nation || "",
    releaseDate: m.repRlsDate || "",
    poster: m.posters?.split("|")[0] || ""
  };
}

// ======================================================
// OMDb
// ======================================================
async function fetchOMDb(title) {
  const url = `https://www.omdbapi.com/?t=${encodeURIComponent(
    title
  )}&apikey=${OMDB_KEY}&plot=full&r=json`;

  const res = await fetch(cors(url));
  const data = await res.json();

  if (data.Response === "False") return { error: "OMDb 검색 실패" };

  return {
    source: "OMDb",
    title: data.Title || "",
    englishTitle: (data.Title || "").toUpperCase(),
    year: data.Year || "",
    director: data.Director || "",
    cast: data.Actors?.split(",").slice(0, 4).join(", "),
    genre: data.Genre || "",
    rating: cleanRating(data.Rated || ""),
    plot: data.Plot || "",
    country: data.Country || "",
    releaseDate: data.Released || "",
    runtime: data.Runtime || "",
    imdbRating: data.imdbRating || ""
  };
}

// ======================================================
// 등급 정리
// ======================================================
function cleanRating(r) {
  if (!r) return "";

  if (r.includes("12")) return "12세";
  if (r.includes("15")) return "15세";
  if (r.includes("18") || r.includes("19") || r.includes("청불")) return "18세";

  return r;
}

// ======================================================
// 엑셀 생성
// ======================================================
async function createExcel(list) {
  const header = [
    "Source", "Title", "English Title", "Year", "Director", "Cast",
    "Genre", "Rating", "Plot", "Country", "Release Date", "Poster/Runtime", "IMDB Rating"
  ];

  const rows = list.map(m => [
    m.source || "",
    m.title || "",
    m.englishTitle || "",
    m.year || "",
    m.director || "",
    m.cast || "",
    m.genre || "",
    m.rating || "",
    m.plot || "",
    m.country || "",
    m.releaseDate || "",
    m.poster || m.runtime || "",
    m.imdbRating || ""
  ]);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);
  XLSX.utils.book_append_sheet(wb, ws, "Result");

  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}
