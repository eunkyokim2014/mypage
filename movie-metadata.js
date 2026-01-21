import { readXlsx, writeXlsx } from "xlsx-populate";

// ===============================
// 🔑 API Keys
// ===============================
const OMDB_KEY = "ce0ca871";
const KMDB_KEY = "66C7AXU2KBQEJ6Y5LX1U";
const CORS_PROXY = "https://api.allorigins.win/raw?url=";
const isBrowser = typeof window !== "undefined";

function withCorsProxy(url) {
  return isBrowser ? `${CORS_PROXY}${encodeURIComponent(url)}` : url;
}

// ===============================
// 🎯 제목 정확도 검사 함수
// ===============================
function isTitleMatch(input, resultTitle) {
  if (!resultTitle) return false;

  const cleanInput = input.replace(/\s+/g, "").toLowerCase();
  const cleanResult = resultTitle.replace(/\s+/g, "").toLowerCase();

  // 입력어가 결과에 포함되거나, 결과가 입력어에 포함되면 OK
  return cleanResult.includes(cleanInput) || cleanInput.includes(cleanResult);
}

// ===============================
// 🎯 등급 통일 함수
// ===============================
function normalizeRating(rating) {
  if (!rating) return "";

  if (rating.includes("전체")) return "전체";
  if (rating.includes("12")) return "12세";
  if (rating.includes("15")) return "15세";
  if (rating.includes("청소년") || rating.includes("불가")) return "청불";
  if (rating.includes("19")) return "18세";

  return rating;
}

// ===============================
// ⭐ MAIN RUN FUNCTION
// ===============================
export default async function run(input) {
  // 1) 단일 제목 검색
  if (input.movieTitle) {
    const meta = await getMovieMetadata(input.movieTitle);
    return {
      metadata: meta,
      excelFile: await createExcel([meta])
    };
  }

  // 2) 파일 업로드 검색
  if (input.files && input.files.length > 0) {
    const file = input.files[0];
    const workbook = await readXlsx(await file.arrayBuffer());
    const sheet = workbook.sheet(0);

    const titles = [];
    sheet.usedRange().value().forEach(row => {
      if (row[0]) titles.push(row[0]);
    });

    const results = [];
    for (const t of titles) {
      const meta = await getMovieMetadata(t);
      results.push(meta);
    }

    return {
      metadata: results,
      excelFile: await createExcel(results)
    };
  }

  return { error: "제목 또는 파일이 필요합니다." };
}

// ========================================================
// 🎬 통합 메타데이터 조회
// ========================================================
async function getMovieMetadata(title) {
  // KMDB 먼저 조회
  const kmdb = await fetchFromKMDB(title);
  if (kmdb && !kmdb.error && isTitleMatch(title, kmdb.title)) {
    return kmdb;
  }

  // OMDb 조회
  const omdb = await fetchFromOMDb(title);
  if (omdb && !omdb.error && isTitleMatch(title, omdb.title)) {
    return omdb;
  }

  return { title, error: "메타데이터를 찾을 수 없습니다." };
}

// ========================================================
// 🎥 KMDB API 호출
// ========================================================
async function fetchFromKMDB(title) {
  const url = `https://api.koreafilm.or.kr/openapi-data2/wisenut/search_api/search_json2.jsp?collection=kmdb_new2&detail=Y&query=${encodeURIComponent(title)}&ServiceKey=${KMDB_KEY}`;

  try {
    const res = await fetch(withCorsProxy(url));
    const data = await res.json();

    if (!data.Data || !data.Data[0]?.Result?.length) {
      return { error: "KMDB 검색 실패" };
    }

    const movie = data.Data[0].Result[0];

    return {
      source: "KMDB",
      title: movie.title?.replace(/!HS|!HE/g, "").trim(),
      englishTitle: (movie.titleEng || "").toUpperCase(),
      year: movie.prodYear || "",
      director: movie.directors?.director?.[0]?.directorNm || "",
      cast: movie.actors?.actor?.slice(0, 4).map(a => a.actorNm).join(", ") || "",
      genre: movie.genre || "",
      rating: normalizeRating(movie.rating || ""),
      plot: movie.plots?.plot?.[0]?.plotText || "",
      country: movie.nation || "",
      releaseDate: movie.repRlsDate || "",
      poster: movie.posters?.split("|")[0] || ""
    };
  } catch {
    return { error: "KMDB 호출 오류" };
  }
}

// ========================================================
// 🌍 OMDb API 호출
// ========================================================
async function fetchFromOMDb(title) {
  const url = `https://www.omdbapi.com/?t=${encodeURIComponent(title)}&apikey=${OMDB_KEY}&plot=full&r=json`;

  try {
    const res = await fetch(withCorsProxy(url));
    const data = await res.json();

    if (data.Response === "False") {
      return { error: "OMDb 검색 실패" };
    }

    return {
      source: "OMDb",
      title: data.Title || "",
      englishTitle: (data.Title || "").toUpperCase(),
      year: data.Year || "",
      director: data.Director || "",
      cast: data.Actors?.split(",").slice(0, 4).join(", ") || "",
      genre: data.Genre || "",
      rating: normalizeRating(data.Rated || ""),
      plot: data.Plot || "",
      country: data.Country || "",
      releaseDate: data.Released || "",
      runtime: data.Runtime || "",
      imdbRating: data.imdbRating || ""
    };
  } catch {
    return { error: "OMDb 호출 오류" };
  }
}

// ========================================================
// 🧾 엑셀 생성
// ========================================================
async function createExcel(metadataList) {
  const workbook = await writeXlsx();
  const sheet = workbook.sheet(0);

  sheet.cell("A1").value([
    [
      "Source", "Title", "English Title", "Year", "Director", "Cast",
      "Genre", "Rating", "Plot", "Country", "Release Date", "Poster/Runtime", "IMDB Rating"
    ],
    ...metadataList.map(m => [
      m.source ?? "",
      m.title ?? "",
      m.englishTitle ?? "",
      m.year ?? "",
      m.director ?? "",
      m.cast ?? "",
      m.genre ?? "",
      m.rating ?? "",
      m.plot ?? "",
      m.country ?? "",
      m.releaseDate ?? "",
      m.poster ?? m.runtime ?? "",
      m.imdbRating ?? ""
    ])
  ]);

  return workbook.outputAsync();
}
