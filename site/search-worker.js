/**
 * Search Web Worker
 * Runs search off the main thread so the UI never freezes.
 *
 * Messages IN:
 *   { type: "init" }                          — fetch word_index.json
 *   { type: "search", query, candidateIds }   — run a search
 *
 * Messages OUT:
 *   { type: "ready" }
 *   { type: "progress", pct }
 *   { type: "results", results, cached }
 *   { type: "error", message }
 */

let wordIndex = null;   // { total_docs, doc_word_counts, index: { word: {df, docs} } }
let sentences = null;    // { doc_id: [sentences...] }
const searchCache = new Map();

// ---- Init: load word index ----
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

// ---- Load sentences on first search ----
async function loadSentences() {
    if (sentences !== null) return true;
    try {
        self.postMessage({ type: "progress", pct: 0, note: "Loading document text..." });
        const resp = await fetch("sentences.json");
        if (!resp.ok) throw new Error("sentences.json not found");
        sentences = await resp.json();
        return true;
    } catch (e) {
        self.postMessage({ type: "error", message: "Failed to load sentences: " + e.message });
        return false;
    }
}

// ---- Tokenize (must match build_index.py logic) ----
const STOP_WORDS = new Set([
    "a","an","and","are","as","at","be","by","for","from","has","have","in",
    "is","it","its","of","on","or","that","the","this","to","was","were",
    "will","with","not","but","they","them","their","than","also","been",
    "can","could","did","does","each","had","how","into","may","more",
    "most","must","nor","only","other","our","out","own","said","same",
    "she","should","so","some","such","then","there","these","those",
    "through","too","under","until","very","we","what","when","where",
    "which","while","who","whom","why","would","you","your","about",
    "after","all","before","both","between","come","get","here","him",
    "his","just","like","made","make","many","much","never","now","off",
    "often","upon","shall","thereof","herein","hereof","pursuant",
    "provided","section","subsection","paragraph",
]);

function tokenize(text) {
    return text.toLowerCase().match(/[a-zA-Z]+/g)
        ?.filter(w => w.length >= 3 && !STOP_WORDS.has(w)) || [];
}

// ---- Inverted index lookup ----
function getDocIdsForQuery(queryWords) {
    if (!wordIndex || !wordIndex.index) return null;

    let result = null;

    for (const word of queryWords) {
        const entry = wordIndex.index[word];
        if (!entry) {
            // Word not in index — try prefix match
            const prefixMatches = new Set();
            for (const key of Object.keys(wordIndex.index)) {
                if (key.startsWith(word)) {
                    for (const id of wordIndex.index[key].docs) {
                        prefixMatches.add(id);
                    }
                }
            }
            if (prefixMatches.size === 0) return new Set(); // no docs match this word
            if (result === null) {
                result = prefixMatches;
            } else {
                result = new Set([...result].filter(id => prefixMatches.has(id)));
            }
        } else {
            const docSet = new Set(entry.docs);
            if (result === null) {
                result = docSet;
            } else {
                // Intersect: all query words must be present
                result = new Set([...result].filter(id => docSet.has(id)));
            }
        }
    }

    return result || new Set();
}

// ---- TF-IDF scoring ----
function scoreTfIdf(docId, queryWords, matchedSentences) {
    if (!wordIndex) return matchedSentences.length;

    const totalDocs = wordIndex.total_docs || 1;
    const docWordCount = (wordIndex.doc_word_counts && wordIndex.doc_word_counts[docId]) || 1;

    // Count query term occurrences in matched sentences
    const allText = matchedSentences.join(" ").toLowerCase();

    let score = 0;
    for (const word of queryWords) {
        // TF: occurrences in this doc's matches / doc length
        const regex = new RegExp(word, "gi");
        const matches = allText.match(regex);
        const tf = matches ? matches.length / docWordCount : 0;

        // IDF: log(total docs / docs containing this word)
        const entry = wordIndex.index[word];
        const df = entry ? entry.df : 1;
        const idf = Math.log(totalDocs / df);

        score += tf * idf;
    }

    // Boost for match density (matches per sentence)
    const density = matchedSentences.length;
    score += density * 0.1;

    return score;
}

// ---- Search ----
async function doSearch(query, candidateIds) {
    const cacheKey = query + "|" + candidateIds.sort().join(",");

    // Check cache
    if (searchCache.has(cacheKey)) {
        self.postMessage({ type: "results", results: searchCache.get(cacheKey), cached: true });
        return;
    }

    // Load sentences on first search
    if (!(await loadSentences())) return;

    const queryWords = tokenize(query);
    const queryWordsRaw = query.toLowerCase().split(/\s+/).filter(Boolean);

    if (queryWords.length === 0 && queryWordsRaw.length === 0) {
        self.postMessage({ type: "results", results: [], cached: false });
        return;
    }

    // Use inverted index to narrow candidates
    let indexHits = getDocIdsForQuery(queryWords);

    // Intersect with metadata-filtered candidates
    let docsToSearch;
    if (candidateIds.length > 0 && indexHits !== null) {
        const candidateSet = new Set(candidateIds);
        docsToSearch = [...indexHits].filter(id => candidateSet.has(id));
    } else if (candidateIds.length > 0) {
        docsToSearch = candidateIds;
    } else if (indexHits !== null) {
        docsToSearch = [...indexHits];
    } else {
        docsToSearch = Object.keys(sentences);
    }

    self.postMessage({ type: "progress", pct: 10, note: "Searching " + docsToSearch.length + " candidate docs..." });

    const results = [];
    const total = docsToSearch.length;

    for (let i = 0; i < total; i++) {
        const docId = docsToSearch[i];
        const docSentences = sentences[docId];
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
            const score = scoreTfIdf(docId, queryWords, matched);
            results.push({
                docId: docId,
                matchedSentences: matched,
                matchCount: matched.length,
                score: Math.round(score * 10000) / 10000,
            });
        }

        // Progress updates every 50 docs
        if ((i + 1) % 50 === 0 || i === total - 1) {
            const pct = 10 + Math.round(((i + 1) / total) * 90);
            self.postMessage({ type: "progress", pct: pct });
        }
    }

    // Sort by TF-IDF score descending
    results.sort((a, b) => b.score - a.score);

    // Cache
    searchCache.set(cacheKey, results);

    self.postMessage({ type: "results", results: results, cached: false });
}

// ---- Message handler ----
self.onmessage = function (e) {
    const msg = e.data;
    if (msg.type === "init") {
        loadWordIndex();
    } else if (msg.type === "search") {
        doSearch(msg.query, msg.candidateIds || []);
    }
};
