// ==================== 東橫INN官網爬蟲函式 ====================

const TOYOKO_PRIMARY_API =
    "https://www.toyoko-inn.com/api/hotels/list?lang=zh-tw";
const TOYOKO_FALLBACK_APIS = [
    "https://www.toyoko-inn.com/api/v1/hotels",
    "https://www.toyoko-inn.com/api/hotels",
    "https://www.toyoko-inn.com/china/api/hotels/list",
];
const TOYOKO_WEBSITE_URL = "https://www.toyoko-inn.com/china/hotel_list/";
const TOYOKO_PROXY_BASE = "https://r.jina.ai/";

function scrapeToyokoInnBranches() {
    Logger.log("========== 開始爬取東橫INN分店資訊 ==========");
    const branches = acquireToyokoInnBranches();
    if (!branches || branches.length === 0) {
        throw new Error("無法從官網取得任何分店資料");
    }
    Logger.log(`✓ 成功取得 ${branches.length} 家分店`);
    const result = updateBranchesSheet(branches);
    Logger.log("========== 爬取完成 ==========");
    return result;
}

function acquireToyokoInnBranches() {
    const steps = [
        () => fetchToyokoInnApi(TOYOKO_PRIMARY_API, buildPrimaryHeaders()),
        () => tryFallbackApis(),
        () => fetchToyokoInnProxyApi(TOYOKO_PRIMARY_API),
        () => fetchToyokoInnWebsite(),
        () => fetchToyokoInnWebsiteMirror(),
    ];
    const errors = [];
    for (const step of steps) {
        try {
            const data = step();
            if (Array.isArray(data) && data.length > 0) {
                return dedupeBranches(data);
            }
        } catch (error) {
            Logger.log(`⚠️ 資料來源失敗: ${error.message}`);
            errors.push(error.message);
        }
    }
    throw new Error(`所有資料來源皆失敗:\n${errors.join("\n")}`);
}

function tryFallbackApis() {
    for (const url of TOYOKO_FALLBACK_APIS) {
        Logger.log(`🔄 嘗試替代 API: ${url}`);
        try {
            const branches = fetchToyokoInnApi(url, buildFallbackHeaders());
            if (branches.length > 0) {
                Logger.log(`✓ 成功使用替代 API`);
                return branches;
            }
        } catch (error) {
            Logger.log(`  ✗ 替代 API 失敗: ${error.message}`);
        }
    }
    throw new Error("所有替代 API 皆不可用");
}

function fetchToyokoInnProxyApi(url) {
    const proxyUrl = buildProxyUrl(url);
    Logger.log(`🔁 使用代理抓取 API: ${proxyUrl}`);
    return fetchToyokoInnApi(proxyUrl, buildProxyHeaders());
}

function fetchToyokoInnApi(url, headers) {
    const options = {
        method: "get",
        muteHttpExceptions: true,
        headers: headers,
        timeout: 30,
    };
    const response = UrlFetchApp.fetch(url, options);
    const status = response.getResponseCode();
    Logger.log(`📊 API 回應碼 (${url}): ${status}`);
    if (status === 403) {
        throw new Error("403 Forbidden (API)");
    }
    if (status !== 200) {
        throw new Error(`API 狀態碼 ${status}`);
    }
    const payload = response.getContentText("utf-8");
    let parsed = safeJsonParse(payload);
    if (!isLikelyApiPayload(parsed)) {
        const extracted = extractJsonPayload(payload);
        if (extracted) {
            parsed = safeJsonParse(extracted);
        }
    }
    if (!isLikelyApiPayload(parsed)) {
        throw new Error("API 回應不是有效 JSON");
    }
    return parseApiResponse(parsed);
}

function fetchToyokoInnWebsite() {
    Logger.log("🔄 API 失敗，嘗試改爬 HTML 網頁");
    const options = {
        method: "get",
        muteHttpExceptions: true,
        headers: buildWebsiteHeaders(),
        timeout: 30,
    };
    const response = UrlFetchApp.fetch(TOYOKO_WEBSITE_URL, options);
    const status = response.getResponseCode();
    Logger.log(`📊 HTML 回應碼: ${status}`);
    if (status !== 200) {
        throw new Error(`HTML 狀態碼 ${status}`);
    }
    const html = response.getContentText("utf-8");
    if (!html || html.length === 0) {
        throw new Error("HTML 內容為空");
    }
    Logger.log(`✓ 取得 HTML (${html.length} 字元)`);
    const branches = parseWebsitePayload(html);
    if (!branches.length) {
        throw new Error("HTML 無法解析分店資料");
    }
    return branches;
}

function fetchToyokoInnWebsiteMirror() {
    const proxyUrl = buildProxyUrl(TOYOKO_WEBSITE_URL);
    Logger.log(`🔁 使用代理抓取 HTML: ${proxyUrl}`);
    const options = {
        method: "get",
        muteHttpExceptions: true,
        headers: buildProxyHeaders(),
        timeout: 30,
    };
    const response = UrlFetchApp.fetch(proxyUrl, options);
    const status = response.getResponseCode();
    Logger.log(`📊 代理 HTML 回應碼: ${status}`);
    if (status !== 200) {
        throw new Error(`代理 HTML 狀態碼 ${status}`);
    }
    const html = response.getContentText("utf-8");
    if (!html || html.length === 0) {
        throw new Error("代理 HTML 內容為空");
    }
    Logger.log(`✓ 取得代理 HTML (${html.length} 字元)`);
    const branches = parseWebsitePayload(html);
    if (!branches.length) {
        throw new Error("代理 HTML 無法解析分店資料");
    }
    return branches;
}

function parseApiResponse(data) {
    const branches = [];
    Object.keys(data || {}).forEach((region) => {
        const hotels = Array.isArray(data[region]) ? data[region] : [];
        hotels.forEach((hotel) => {
            const record = normalizeHotelRecord(hotel, region);
            if (record) {
                branches.push(record);
            }
        });
    });
    Logger.log(`✓ API 解析出 ${branches.length} 家分店`);
    return branches;
}

function parseWebsitePayload(html) {
    const nuxtBranches = extractFromNuxtPayload(html);
    if (nuxtBranches.length) {
        Logger.log(`✓ 從 __NUXT__ 取出 ${nuxtBranches.length} 家分店`);
        return nuxtBranches;
    }
    const datasetBranches = extractFromDataAttributes(html);
    if (datasetBranches.length) {
        Logger.log(`✓ 從 data-* 屬性取出 ${datasetBranches.length} 家分店`);
        return datasetBranches;
    }
    const anchorBranches = extractFromAnchors(html);
    if (anchorBranches.length) {
        Logger.log(`✓ 從連結取出 ${anchorBranches.length} 家分店`);
        return anchorBranches;
    }
    const markdownBranches = extractFromMarkdown(html);
    if (markdownBranches.length) {
        Logger.log(`✓ 從 Markdown 取出 ${markdownBranches.length} 家分店`);
        return markdownBranches;
    }
    return anchorBranches;
}

function extractFromNuxtPayload(html) {
    const match = html.match(/window\.__NUXT__=([\s\S]*?);<\/script>/);
    if (!match) {
        return [];
    }
    const sanitized = sanitizeNuxtPayload(match[1]);
    const nuxt = safeJsonParse(sanitized);
    if (!nuxt) {
        return [];
    }
    const buckets = [];
    const dataBlocks = Array.isArray(nuxt.data) ? nuxt.data : [];
    dataBlocks.forEach((block) => buckets.push(block));
    if (nuxt.state) {
        buckets.push(nuxt.state);
    }
    return flattenHotels(buckets);
}

function sanitizeNuxtPayload(raw) {
    return raw
        .replace(/undefined/g, "null")
        .replace(/NaN/g, "null")
        .replace(/new Date\([^\)]*\)/g, "null");
}

function flattenHotels(buckets) {
    const results = [];
    const visited = new Set();
    const walk = (node) => {
        if (!node || visited.has(node)) {
            return;
        }
        if (typeof node !== "object") {
            return;
        }
        visited.add(node);
        if (Array.isArray(node)) {
            node.forEach((item) => walk(item));
            return;
        }
        const record = normalizeHotelRecord(node);
        if (record) {
            results.push(record);
        }
        Object.keys(node).forEach((key) => walk(node[key]));
    };
    buckets.forEach((bucket) => walk(bucket));
    return dedupeBranches(results);
}

function extractFromDataAttributes(html) {
    const pattern =
        /data-name="([^"]+)"[^>]*data-code="([^"]+)"[^>]*data-region="([^"]+)"[^>]*data-pref="([^"]+)"/g;
    const branches = [];
    let match;
    while ((match = pattern.exec(html)) !== null) {
        branches.push(
            buildBranchRecord({
                name: match[1],
                code: match[2],
                region: match[3],
                prefecture: match[4],
            })
        );
    }
    return branches;
}

function extractFromAnchors(html) {
    const pattern = /href="\/china\/hotel\/([^"]+)"[^>]*>([^<]+)<\/a>/g;
    const branches = [];
    let match;
    while ((match = pattern.exec(html)) !== null) {
        const code = match[1];
        const name = match[2]
            .replace(/\s*\(\d+\)\s*$/, "")
            .replace(/\s*（\d+）\s*$/, "")
            .trim();
        if (!code || !name) {
            continue;
        }
        const [region, prefecture] = deriveRegionPrefecture(code);
        branches.push(
            buildBranchRecord({
                name: name,
                code: code,
                region: region,
                prefecture: prefecture,
            })
        );
    }
    return branches;
}

function extractFromMarkdown(markdown) {
    const lines = markdown.split(/\r?\n/);
    const branches = [];
    let currentRegion = "未分類";
    let currentPrefecture = "未分類";

    for (let i = 0; i < lines.length; i += 1) {
        const line = lines[i];
        if (!line) {
            continue;
        }
        const trimmed = line.trim();
        if (!trimmed) {
            continue;
        }

        const next = lines[i + 1] ? lines[i + 1].trim() : "";
        if (/^[-]{2,}$/.test(next) && !trimmed.startsWith("#")) {
            currentRegion = trimmed;
            currentPrefecture = trimmed;
            continue;
        }

        if (/^###\s+/.test(trimmed)) {
            currentPrefecture =
                trimmed.replace(/^###\s+/, "").trim() || currentRegion;
            continue;
        }

        const linkMatch = trimmed.match(/\[([^\]]+)\]\((https?:\/\/[^\)]+)\)/);
        if (linkMatch) {
            const name = sanitizeHotelName(linkMatch[1]);
            const url = linkMatch[2];
            const codeMatch = url.match(/\/(\d{3,})\/?$/);
            const code = codeMatch
                ? codeMatch[1]
                : url.split("/").filter(Boolean).pop();
            if (name && code) {
                branches.push(
                    buildBranchRecord({
                        name: name,
                        code: code,
                        region: currentRegion,
                        prefecture: currentPrefecture,
                    })
                );
            }
            continue;
        }

        if (trimmed.startsWith("*")) {
            const bracketLine = (lines[i + 1] || "").trim();
            const bracketMatch = bracketLine.match(
                /\[([^\]]+)\]\((https?:\/\/[^\)]+)\)/
            );
            if (bracketMatch) {
                const name = sanitizeHotelName(bracketMatch[1]);
                const url = bracketMatch[2];
                const codeMatch = url.match(/\/(\d{3,})\/?$/);
                const code = codeMatch
                    ? codeMatch[1]
                    : url.split("/").filter(Boolean).pop();
                if (name && code) {
                    branches.push(
                        buildBranchRecord({
                            name: name,
                            code: code,
                            region: currentRegion,
                            prefecture: currentPrefecture,
                        })
                    );
                }
            }
        }
    }

    return branches;
}

function deriveRegionPrefecture(code) {
    const parts = code.split("-");
    if (parts.length >= 3) {
        return [`${parts[0]}-${parts[1]}`, parts.slice(2).join("-")];
    }
    if (parts.length === 2) {
        return [parts[0], parts[1]];
    }
    return ["未分類", "未分類"];
}

function normalizeHotelRecord(raw, fallbackRegion) {
    if (!raw) {
        return null;
    }
    const code =
        raw.code ||
        raw.hotelCode ||
        raw.hotel_code ||
        raw.hotelNo ||
        raw.id ||
        raw.hotel_id;
    const name = raw.name || raw.hotelName || raw.hotel_name || raw.hotelTitle;
    if (!code || !name) {
        return null;
    }
    const region =
        raw.region ||
        raw.areaName ||
        raw.area ||
        raw.countryRegion ||
        raw.prefectureBlock ||
        fallbackRegion ||
        "未分類";
    const prefecture =
        raw.prefecture ||
        raw.pref ||
        raw.prefectureName ||
        raw.cityName ||
        raw.city ||
        "未分類";
    return buildBranchRecord({
        name: sanitizeHotelName(name),
        code: normalizeHotelCode(code),
        region: String(region).trim() || "未分類",
        prefecture: String(prefecture).trim() || "未分類",
    });
}

function normalizeHotelCode(code) {
    const cleaned = String(code).replace(/\D/g, "");
    if (cleaned.length >= 5) {
        return cleaned.substring(0, 5);
    }
    return cleaned.padStart(5, "0");
}

function sanitizeHotelName(name) {
    return String(name)
        .replace(/^東橫INN\s*/, "")
        .replace(/\s*\(\d+\)\s*$/, "")
        .replace(/\s*（\d+）\s*$/, "")
        .replace(/＜.+?＞/g, "")
        .trim();
}

function buildBranchRecord({ name, code, region, prefecture }) {
    return {
        name: name,
        code: code,
        region: region || "未分類",
        prefecture: prefecture || "未分類",
    };
}

function dedupeBranches(branches) {
    const map = new Map();
    branches.forEach((branch) => {
        if (!branch || !branch.code) {
            return;
        }
        map.set(branch.code, branch);
    });
    return Array.from(map.values()).sort((a, b) =>
        a.code.localeCompare(b.code)
    );
}

function extractJsonPayload(text) {
    if (!text) {
        return null;
    }
    const objectMatch = text.match(/({[\s\S]*})\s*$/);
    if (objectMatch) {
        return objectMatch[1];
    }
    const arrayMatch = text.match(/(\[[\s\S]*\])\s*$/);
    if (arrayMatch) {
        return arrayMatch[1];
    }
    const braceIndex = text.indexOf("{");
    const bracketIndex = text.indexOf("[");
    let start = -1;
    if (
        braceIndex !== -1 &&
        (bracketIndex === -1 || braceIndex < bracketIndex)
    ) {
        start = braceIndex;
    } else {
        start = bracketIndex;
    }
    if (start === -1) {
        return null;
    }
    return text.substring(start).trim();
}

function isLikelyApiPayload(data) {
    if (!data || typeof data !== "object" || Array.isArray(data)) {
        return false;
    }
    return Object.values(data).some((value) => Array.isArray(value));
}

function safeJsonParse(payload) {
    try {
        return JSON.parse(payload);
    } catch (error) {
        Logger.log(`⚠️ JSON 解析失敗: ${error.message}`);
        return null;
    }
}

function buildPrimaryHeaders() {
    return {
        "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        Accept: "application/json, text/plain, */*",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        Referer: TOYOKO_WEBSITE_URL,
        Origin: "https://www.toyoko-inn.com",
        DNT: "1",
    };
}

function buildFallbackHeaders() {
    return {
        "User-Agent":
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
        Accept: "application/json, text/plain, */*",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        DNT: "1",
    };
}

function buildWebsiteHeaders() {
    return {
        "User-Agent":
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        Accept: "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
        DNT: "1",
    };
}

function buildProxyHeaders() {
    return {
        "User-Agent":
            "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        Accept: "*/*",
        "Accept-Language": "zh-TW,zh;q=0.9,en;q=0.8",
        "Cache-Control": "no-cache",
        Pragma: "no-cache",
    };
}

function buildProxyUrl(url) {
    return `${TOYOKO_PROXY_BASE}${url}`;
}

function updateBranchesSheet(branches) {
    const spreadsheet = getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAMES.branches);
    if (!sheet) {
        initializeBranchesSheet();
        sheet = spreadsheet.getSheetByName(SHEET_NAMES.branches);
    }
    const current = sheet.getDataRange().getValues();
    if (current.length > 1) {
        sheet.deleteRows(2, current.length - 1);
    }

    // 排除「日本以外」地區的分店
    const filteredBranches = branches.filter(
        (branch) => branch.region !== "日本以外"
    );

    if (filteredBranches.length > 0) {
        // 依照「地區」、「都道府縣」、「分店名稱」排序
        const sorted = filteredBranches.slice().sort((a, b) => {
            if (a.region !== b.region) {
                return a.region.localeCompare(b.region, "zh-TW");
            }
            if (a.prefecture !== b.prefecture) {
                return a.prefecture.localeCompare(b.prefecture, "zh-TW");
            }
            return a.name.localeCompare(b.name, "zh-TW");
        });
        sheet.insertRows(2, sorted.length);
        sheet
            .getRange(2, 1, sorted.length, 4)
            .setValues(
                sorted.map((branch) => [
                    branch.name,
                    branch.code,
                    branch.region,
                    branch.prefecture,
                ])
            );
    }
    const stats = summarizeBranches(filteredBranches);
    const props = PropertiesService.getUserProperties();
    props.setProperty("BRANCHES_LAST_SCRAPED", new Date().toISOString());
    props.setProperty("BRANCHES_SCRAPE_STATS", JSON.stringify(stats));
    const summary = Object.keys(stats.byRegion)
        .map((key) => `${key}(${stats.byRegion[key]}家)`)
        .join("、");
    return {
        count: filteredBranches.length,
        summary: summary ? `已新增: ${summary}` : "未取得地區統計",
    };
}

function summarizeBranches(branches) {
    const stats = { byRegion: {}, byPrefecture: {} };
    branches.forEach((branch) => {
        stats.byRegion[branch.region] =
            (stats.byRegion[branch.region] || 0) + 1;
        stats.byPrefecture[branch.prefecture] =
            (stats.byPrefecture[branch.prefecture] || 0) + 1;
    });
    return stats;
}

function getScrapeStatus() {
    try {
        const props = PropertiesService.getUserProperties();
        const lastScraped = props.getProperty("BRANCHES_LAST_SCRAPED");
        const statsJson = props.getProperty("BRANCHES_SCRAPE_STATS");
        const stats = statsJson ? JSON.parse(statsJson) : { byRegion: {} };
        const sheet = getSheet(SHEET_NAMES.branches);
        const total = sheet.getLastRow() - 1;
        const regions = Object.keys(stats.byRegion || {}).sort();
        return {
            lastUpdated: lastScraped
                ? new Date(lastScraped).toLocaleString("zh-TW")
                : "尚未爬取",
            branchCount: Math.max(total, 0),
            regions: regions.length > 0 ? regions : ["未分類"],
        };
    } catch (error) {
        Logger.log(`⚠️ 狀態查詢失敗: ${error.message}`);
        return {
            lastUpdated: "未知",
            branchCount: 0,
            regions: [],
        };
    }
}
