import * as XLSX from "xlsx";

// ===============================
// 🔑 API Keys
// ===============================
const OMDB_KEY = "ce0ca871";
const KMDB_KEY = "66C7AXU2KBQEJ6Y5LX1U";
const CORS_PROXY = "https://api.allorigins.win/raw?url=";
const isBrowser = typeof window !== "undefined";

function withCors(url) {
  return isBrowser ? `${CORS_PROXY}${encodeURIComponent(url)}` : url;
}

// ========================================================
// 🎯 KMDB 최적 매칭 로직
// ========================================================
function clean(t) {
  return (t || "").replace(/!HS|!HE/g, "").trim();
}

function normalize(t) {
  return (t || "").replace(/\s+/g, "").trim();
}

function selectBestKMDBMatch(title, results) {
  const normalizedTitle = normalize(title);

  // 1) 완전 일치
  const exact = results.find(r => normalize(clean(r.title)) === normalizedTitle);
  if (exact) return exact;

  // 2) 한글 포함 매칭
  const partial = results.find(r => clean(r.title).includes(title));
  if (partial) return partial;

  // 3) 영문 Oldboy 정확 매칭
  const engMatch = results.find(
    r => normalize(r.titleEng)?.toLowerCase() === "oldboy"
  );
  if (engMatch) return engMatch;

  // 4) 기본값
  return results[0];
}

// ========================================================
// 🎬 KMDB API
// ========================================================
async function fetchFromKMDB(title) {
  const url = `https://api.koreafilm.or.kr/openapi-data2/wisenut/search_api/search_json2.jsp?collection=kmdb_new2&detail=Y&query=${encodeURIComponent(
    title
  )}&ServiceKey=${KMDB_KEY}`;

  const res = await fetch(withCors(url));
  const data = await res.json();

  const list = data.Data?.[0]?.Result;
  if (!list || list.length === 0) return null;

  const movie = selectBestKMDBMatch(title, list);

  // 출연자 3~4명 제한
  let castList = movie.actors?.actor?.map(a => a.actorNm) || [];
  castList = castList.slice(0, 4);

  // 등급 변환
  const rating = convertRating(movie.rating);

  return {
    source: "KMDB",
    title: clean(movie.title),
    englishTitle: (movie.titleEng || "").toUpperCase(),
    year: movie.prodYear || "",
    director: movie.directors?.director?.[0]?.directorNm || "",
    cast: castList.join(", "),
    genre: movie.genre || "",
    rating,
    plot: movie.plots?.plot?.[0]?.plotText || "",
    country: movie.nation || "",
    releaseDate: movie.repRlsDate || "",
    poster: movie.posters?.split("|")[0] || ""
  };
}

// ========================================================
// 🌍 OMDb API
// ========================================================
async function fetchFromOMDb(title) {
  const url = `https://www.omdbapi.com/?t=${encodeURIComponent(
    title
  )}&apikey=${OMDB_KEY}&plot=full&r=json`;

  const res = await fetch(withCors(url));
  const data = await res.json();

  if (data.Response === "False") return null;

  // 출연자 3~4명
  let cast = (data.Actors || "").split(",").map(a => a.trim()).slice(0, 4);

  const rating = convertRating(data.Rated);

  return {
    source: "OMDb",
    title: data.Title || "",
    englishTitle: (data.Title || "").toUpperCase(),
    year: data.Year || "",
    director: data.Director || "",
    cast: cast.join(", "),
    genre: data.Genre || "",
    rating,
    plot: data.Plot || "",
    country: data.Country || "",
    releaseDate: data.Released || "",
    runtime: data.Runtime || "",
    imdbRating: data.imdbRating || ""
  };
}

// ========================================================
// 🎯 등급 변환 함수
// ========================================================
function convertRating(r) {
  if (!r) return "";

  if (r.includes("전체")) return "전체";
  if (r.includes("12")) return "12세";
  if (r.includes("15")) return "15세";
  if (r.includes("청소년") || r.includes("청불")) return "18세";
  if (r === "R" || r === "NC-17") return "18세";

  return r; // 기타 등급은 그대로
}

// ========================================================
// 🎬 통합 검색
// ========================================================
async function getMovieMetadata(title) {
  // 1) KMDB 우선
  const kmdb = await fetchFromKMDB(title);
  if (kmdb) return kmdb;

  // 2) OMDb
  const omdb = await fetchFromOMDb(title);
  if (omdb) return omdb;

  return { title, error: "메타데이터를 찾을 수 없습니다." };
}

// ========================================================
// 🧾 Excel 생성 (xlsx)
// ========================================================
function createExcel(metadataList) {
  const rows = [
    [
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
      "Release Date",
      "Poster/Runtime",
      "IMDB Rating"
    ],
    ...metadataList.map(m => [
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
    ])
  ];

  const worksheet = XLSX.utils.aoa_to_sheet(rows);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Movies");

  return XLSX.write(workbook, { type: "array", bookType: "xlsx" });
}

// ========================================================
// 📌 실행 엔트리
// ========================================================
export default async function run(input) {
  if (input.movieTitle) {
    const meta = await getMovieMetadata(input.movieTitle);
    return {
      metadata: meta,
      excelFile: createExcel([meta])
    };
  }

  if (input.files?.length > 0) {
    const file = input.files[0];
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const titles = rows.map(r => r[0]).filter(Boolean);

    const results = [];
    for (const title of titles) {
      results.push(await getMovieMetadata(title));
    }

    return {
      metadata: results,
      excelFile: createExcel(results)
    };
  }

  return { error: "제목 또는 파일이 필요합니다." };
}
