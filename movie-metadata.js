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

export default async function run(input) {
  // 1) 단일 제목 조회
  if (input.movieTitle) {
    const meta = await getMovieMetadata(input.movieTitle);
    return {
      metadata: meta,
      excelFile: await createExcel([meta])
    };
  }

  // 2) 엑셀 파일 조회
  if (input.files && input.files.length > 0) {
    const titles = await readExcelTitles(input.files[0]);

    const results = [];
    for (const t of titles) {
      results.push(await getMovieMetadata(t));
    }

    return {
      metadata: results,
      excelFile: await createExcel(results)
    };
  }

  return { error: "제목 또는 파일이 필요합니다." };
}


// ========================================================
// 📘 엑셀 읽기 (SheetJS 기반)
// ========================================================
async function readExcelTitles(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = function (e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: "array" });

      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      // 2D 배열로 가져오기
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // 첫 번째 컬럼만 가져오기
      const titles = rows
        .map(row => row[0])
        .filter(v => v && typeof v === "string");

      resolve(titles);
    };

    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}


// ========================================================
// 🎬 통합 검색: KMDB → OMDb
// ========================================================
async function getMovieMetadata(title) {
  const kmdb = await fetchFromKMDB(title);
  if (kmdb && !kmdb.error) return kmdb;

  const omdb = await fetchFromOMDb(title);
  if (omdb && !omdb.error) return omdb;

  return { title, error: "메타데이터를 찾을 수 없습니다." };
}


// ========================================================
// 🎥 KMDB API
// ========================================================
async function fetchFromKMDB(title) {
  const url = `https://api.koreafilm.or.kr/openapi-data2/wisenut/search_api/search_json2.jsp?collection=kmdb_new2&detail=Y&query=${encodeURIComponent(title)}&ServiceKey=${KMDB_KEY}`;

  const res = await fetch(withCorsProxy(url));
  const data = await res.json();

  if (!data.Data || !data.Data[0]?.Result?.length) {
    return { error: "KMDB 검색 실패" };
  }

  const movie = data.Data[0].Result[0];

  return {
    source: "KMDB",
    title: movie.title?.replace(/!HS|!HE/g, "").trim(),
    englishTitle: movie.titleEng || "",
    year: movie.prodYear || "",
    director: movie.directors?.director?.[0]?.directorNm || "",
    cast: movie.actors?.actor?.map(a => a.actorNm).join(", "),
    genre: movie.genre || "",
    rating: movie.rating || "",
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
  const url = `https://www.omdbapi.com/?t=${encodeURIComponent(title)}&apikey=${OMDB_KEY}&plot=full&r=json`;

  const res = await fetch(withCorsProxy(url));
  const data = await res.json();

  if (data.Response === "False") return { error: "OMDb 검색 실패" };

  return {
    source: "OMDb",
    title: data.Title,
    englishTitle: data.Title,
    year: data.Year,
    director: data.Director,
    cast: data.Actors,
    genre: data.Genre,
    rating: data.Rated,
    plot: data.Plot,
    country: data.Country,
    releaseDate: data.Released,
    runtime: data.Runtime,
    imdbRating: data.imdbRating
  };
}


// ========================================================
// 🧾 엑셀 생성 (SheetJS 기반)
// ========================================================
async function createExcel(metadataList) {
  const wb = XLSX.utils.book_new();

  const rows = [
    [
      "Source", "Title", "English Title", "Year", "Director", "Cast",
      "Genre", "Rating", "Plot", "Country", "Release Date", "Poster / Runtime", "IMDB Rating"
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
  ];

  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, sheet, "Results");

  return XLSX.write(wb, { type: "array", bookType: "xlsx" });
}
