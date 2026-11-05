/**
 * Search Web Worker — Chunked Architecture
 * 
 * Only fetches the sentence chunks needed for each search,
 * caches them so repeat searches are instant.
 */

let wordIndex = null;
const chunkCache = new Map();   // chunk number → { docId: [sentences] }
const searchCache = new Map();  // query key → results

// ---- Init ----
async function loadWordIndex() {
    try {
        const resp = await fetch("word_index.json");
        if (!resp.ok) throw new Error("word_index.json not found");
        wordIndex = await resp.json();
        self.postMessage({ type: "ready" });
    } catch (e) {
        self.postMessage({ type: "error", message: "Failed to load word index: " + e.message });
    }
}

// ---- Fetch a sentence chunk (cached) ----
async function getChunk(chunkNum) {
    if (chunkCache.has(chunkNum)) return chunkCache.get(chunkNum);
    try {
        const resp = await fetch("chunks/" + chunkNum + ".json");
        if (!resp.ok) return {};
        const data = await resp.json();
        chunkCache.set(chunkNum, data);
        return data;
    } catch (e) {
        return {};
    }
}

// ---- Tokenize (matches build_index.py) ----
const STOP = new Set(["a","an","and","are","as","at","be","by","for","from","has","have",
    "in","is","it","its","of","on","or","that","the","this","to","was","were","will",
    "with","not","but","they","them","their","than","also","been","can","could","did",
    "does","each","had","how","into","may","more","most","must","nor","only","other",
    "our","out","own","said","same","she","should","so","some","such","then","there",
    "these","those","through","too","under","until","very","we","what","when","where",
    "which","while","who","whom","why","would","you","your","about","after","all",
    "before","both","between","come","get","here","him","his","just","like","made",
    "make","many","much","never","now","off","often","upon","shall","thereof","herein",
    "hereof","pursuant","provided","section","subsection","paragraph"]);

function tokenize(text) {
    return (text.toLowerCase().match(/[a-zA-Z]+/g) || [])
        .filter(w => w.length >= 3 && !STOP.has(w));
}

// ---- Inverted index lookup ----
function getDocIdsForQuery(queryWords) {
    if (!wordIndex || !wordIndex.index) return null;
    let result = null;
    for (const word of queryWords) {
        const entry = wordIndex.index[word];
        let docSet;
        if (entry) {
            docSet = new Set(entry.docs);
        } else {
            // Prefix match fallback
            docSet = new Set();
            for (const key of Object.keys(wordIndex.index)) {
                if (key.startsWith(word)) {
                    for (const id of wordIndex.index[key].docs) docSet.add(id);
                }
            }
            if (docSet.size === 0) return new Set();
        }
        if (result === null) result = docSet;
        else result = new Set([...result].filter(id => docSet.has(id)));
    }
    return result || new Set();
}

// ---- TF-IDF ----
function scoreTfIdf(docId, queryWords, matchedSentences) {
    if (!wordIndex) return matchedSentences.length;
    const totalDocs = wordIndex.total_docs || 1;
    const docWordCount = (wordIndex.doc_word_counts && wordIndex.doc_word_counts[docId]) || 1;
    const allText = matchedSentences.join(" ").toLowerCase();
    let score = 0;
    for (const word of queryWords) {
        const matches = allText.match(new RegExp(word, "gi"));
        const tf = matches ? matches.length / docWordCount : 0;
        const entry = wordIndex.index[word];
        const df = entry ? entry.df : 1;
        const idf = Math.log(totalDocs / df);
        score += tf * idf;
    }
    score += matchedSentences.length * 0.1;
    return score;
}

// ---- Search ----
async function doSearch(query, candidateIds, chunkMap) {
    const cacheKey = query + "|" + candidateIds.sort().join(",");
    if (searchCache.has(cacheKey)) {
        self.postMessage({ type: "results", results: searchCache.get(cacheKey), cached: true });
        return;
    }

    const queryWords = tokenize(query);
    const queryWordsRaw = query.toLowerCase().split(/\s+/).filter(Boolean);
    if (queryWords.length === 0 && queryWordsRaw.length === 0) {
        self.postMessage({ type: "results", results: [], cached: false });
        return;
    }

    // Use inverted index
    let indexHits = getDocIdsForQuery(queryWords);
    let docsToSearch;
    if (candidateIds.length > 0 && indexHits !== null) {
        const cs = new Set(candidateIds);
        docsToSearch = [...indexHits].filter(id => cs.has(id));
    } else if (candidateIds.length > 0) {
        docsToSearch = candidateIds;
    } else if (indexHits !== null) {
        docsToSearch = [...indexHits];
    } else {
        docsToSearch = Object.keys(chunkMap);
    }

    self.postMessage({ type: "progress", pct: 5,
        note: "Found " + docsToSearch.length + " candidate docs via index..." });

    // Figure out which chunks we need
    const neededChunks = new Set();
    for (const docId of docsToSearch) {
        const c = chunkMap[docId];
        if (c !== undefined) neededChunks.add(c);
    }

    // Fetch needed chunks (in parallel, cached)
    self.postMessage({ type: "progress", pct: 10,
        note: "Loading " + neededChunks.size + " chunk(s)..." });
    await Promise.all([...neededChunks].map(c => getChunk(c)));

    // Search
    const results = [];
    const total = docsToSearch.length;
    for (let i = 0; i < total; i++) {
        const docId = docsToSearch[i];
        const chunkNum = chunkMap[docId];
        const chunk = chunkCache.get(chunkNum);
        if (!chunk) continue;
        const docSentences = chunk[docId];
        if (!docSentences) continue;

        const matched = [];
        for (const s of docSentences) {
            const sLower = s.toLowerCase();
            if (queryWordsRaw.every(w => sLower.includes(w))) {
                matched.push(s);
                if (matched.length >= 5) break;
            }
        }

        if (matched.length > 0) {
            results.push({
                docId, matchedSentences: matched, matchCount: matched.length,
                score: Math.round(scoreTfIdf(docId, queryWords, matched) * 10000) / 10000,
            });
        }

        if ((i + 1) % 50 === 0 || i === total - 1) {
            self.postMessage({ type: "progress", pct: 10 + Math.round(((i+1)/total)*90) });
        }
    }

    results.sort((a, b) => b.score - a.score);
    searchCache.set(cacheKey, results);
    self.postMessage({ type: "results", results, cached: false });
}

// ---- Message handler ----
self.onmessage = function(e) {
    const msg = e.data;
    if (msg.type === "init") loadWordIndex();
    else if (msg.type === "search") doSearch(msg.query, msg.candidateIds || [], msg.chunkMap || {});
};
