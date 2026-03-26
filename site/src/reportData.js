export const REPORT_SOURCE_PATH = "src/data/rapport.json";

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function pickFirstObject(source, keys) {
  for (const key of keys) {
    if (isPlainObject(source?.[key])) return source[key];
  }
  return {};
}

function pickFirstArray(source, keys) {
  for (const key of keys) {
    if (Array.isArray(source?.[key])) return source[key];
  }
  return [];
}

function pickFirstValue(...values) {
  return values.find((value) => value !== undefined && value !== null);
}

function toNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function pickNumber(...values) {
  for (const value of values) {
    const parsed = toNumber(value);
    if (parsed !== null) return parsed;
  }
  return null;
}

function normalizeLookup(value) {
  return String(value || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function translateThemeLabel(theme, lang) {
  if (lang === "fr") return theme;

  const key = normalizeLookup(theme);
  if (key.includes("polit")) return "Politics";
  if (key.includes("catastroph") || key.includes("faits divers")) return "Disasters / incidents";
  if (key.includes("covid") || key.includes("sant")) return "COVID / Health";
  if (key.includes("technolog")) return "Technology";
  if (key.includes("deepfake") || key.startsWith("ia ") || key === "ia") return "AI / Deepfake";
  if (key.includes("relig") || key.includes("hist")) return "Religion / History";
  if (key.includes("divert")) return "Entertainment";
  if (key.includes("conspir")) return "Conspiracies";
  if (key.includes("autre") || key.includes("other")) return "Other";
  if (key.includes("brit")) return "Celebrities";
  return theme;
}

function translatePlatformLabel(platform, lang) {
  if (lang === "fr") return platform;

  const key = normalizeLookup(platform);
  if (key === "inconnu") return "Unknown";
  if (key === "presse") return "Press";
  return platform;
}

function translatePatternLabel(pattern, lang) {
  if (lang === "fr") return pattern;

  const key = normalizeLookup(pattern);
  if (key.includes("alerte") || key.includes("urgence")) return "Alert / urgency";
  if (key.includes("partage")) return "Call to share";
  if (key.includes("mort") || key.includes("deces")) return "Death claim";
  if (key.includes("complot")) return "Revealed plot";
  if (key.includes("reference") || key.includes("image") || key === "ia") return "AI / image reference";
  if (key.includes("scandale") || key.includes("choc")) return "Scandal / shock";
  return pattern;
}

function toCountArray(record, translateLabel = (label) => label) {
  return Object.entries(record || {})
    .map(([label, value]) => ({ rawLabel: label, label: translateLabel(label), value: Number(value) }))
    .filter((item) => Number.isFinite(item.value) && item.value > 0)
    .sort((a, b) => b.value - a.value || a.label.localeCompare(b.label, "fr"));
}

function formatSummaryLabel(key, lang) {
  const labels = {
    total_documents: {
      fr: "documents au total",
      en: "total documents",
    },
    classified_documents: {
      fr: "documents classes",
      en: "classified documents",
    },
    unclassified_documents: {
      fr: "documents non classes",
      en: "unclassified documents",
    },
    platform_mentions: {
      fr: "mentions de plateformes",
      en: "platform mentions",
    },
    distinct_platforms: {
      fr: "plateformes distinctes",
      en: "distinct platforms",
    },
    tracked_topics: {
      fr: "thematiques suivies",
      en: "tracked topics",
    },
    rhetorical_patterns: {
      fr: "patterns rhetoriques",
      en: "rhetorical patterns",
    },
    real_world_usages: {
      fr: "usages reels",
      en: "real-world usages",
    },
  };

  if (labels[key]) return labels[key][lang];

  return key
    .replace(/_/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildSummaryCards(raw, datasets, documents, lang) {
  const explicitSummary = pickFirstObject(raw, ["summary"]);
  const totalDocuments = pickNumber(
    explicitSummary.total_documents,
    explicitSummary.totalDocuments,
    explicitSummary.total_fichiers,
    raw.total_fichiers,
    documents.length
  );

  const baseCards = [
    { key: "total_documents", value: totalDocuments },
    { key: "distinct_platforms", value: datasets.platforms.length },
    { key: "tracked_topics", value: datasets.topics.length },
  ];

  return baseCards
    .filter((card) => card.value !== null && card.value !== undefined)
    .map((card) => ({
      id: card.key,
      label: formatSummaryLabel(card.key, lang),
      value: card.value,
    }));
}

function normalizeDocuments(rawDocuments, lang) {
  return rawDocuments.map((document, index) => {
    const filename = pickFirstValue(document.filename, document.fichier, `Document ${index + 1}`);
    const platforms = pickFirstArray(document, ["platforms", "plateformes"]).map((item) =>
      translatePlatformLabel(item, lang)
    );
    const patterns = pickFirstArray(document, ["rhetorical_patterns", "patterns_rhetoriques"]).map((item) =>
      translatePatternLabel(item, lang)
    );
    const usages = pickFirstArray(document, ["usages", "usage", "usage_labels"]);
    const topic = pickFirstValue(document.topic, document.theme, document.thematique, "");

    return {
      id: `${filename}-${index}`,
      filename,
      platforms,
      topic: translateThemeLabel(topic, lang),
      patterns,
      usages,
    };
  });
}

export async function loadLocalReport() {
  const module = await import("./data/rapport.json");
  return module.default || module;
}

export function createDashboardDataset(rawReport, lang = "fr") {
  const raw = isPlainObject(rawReport) ? rawReport : {};
  const platformCounts = pickFirstObject(raw, ["platform_counts", "plateformes"]);
  const topicCounts = pickFirstObject(raw, ["topic_counts", "thematiques"]);
  const rhetoricalCounts = pickFirstObject(raw, ["rhetorical_counts", "patterns_rhetoriques"]);
  const usageCounts = pickFirstObject(raw, ["usage_counts"]);
  const ocrWordCounts = pickFirstObject(raw, ["top_mots_ocr"]);
  const rawDocuments = pickFirstArray(raw, ["documents"]);

  const datasets = {
    platforms: toCountArray(platformCounts, (label) => translatePlatformLabel(label, lang)),
    topics: toCountArray(topicCounts, (label) => translateThemeLabel(label, lang)),
    rhetoricalPatterns: toCountArray(rhetoricalCounts, (label) => translatePatternLabel(label, lang)),
    usages: toCountArray(usageCounts),
  };

  const platformRanking = datasets.platforms
    .filter((item) => {
      const key = normalizeLookup(item.rawLabel || item.label);
      return key !== "inconnu" && key !== "unknown" && key !== "pediavenir";
    });

  const stopAtPressIndex = platformRanking.findIndex((item) => {
    const key = normalizeLookup(item.rawLabel || item.label);
    return key === "presse" || key === "press";
  });

  const ocrWords = Object.entries(ocrWordCounts || {})
    .map(([text, count]) => ({ text, count: Number(count) }))
    .filter((item) => Number.isFinite(item.count) && item.count > 0)
    .sort((a, b) => b.count - a.count || a.text.localeCompare(b.text, "fr"))
    .slice(0, 36)
    .map((item) => ({
      ...item,
      size: 16 + item.count * 0.9,
    }));

  const documents = normalizeDocuments(rawDocuments, lang);
  const summaryCards = buildSummaryCards(raw, datasets, documents, lang);

  return {
    summaryCards,
    platforms: datasets.platforms,
    platformRanking: stopAtPressIndex >= 0 ? platformRanking.slice(0, stopAtPressIndex + 1) : platformRanking,
    topics: datasets.topics,
    rhetoricalPatterns: datasets.rhetoricalPatterns,
    usages: datasets.usages,
    ocrWords,
    documents,
    sourcePath: REPORT_SOURCE_PATH,
    isEmpty:
      summaryCards.length === 0 &&
      datasets.platforms.length === 0 &&
      datasets.topics.length === 0 &&
      datasets.rhetoricalPatterns.length === 0 &&
      datasets.usages.length === 0 &&
      ocrWords.length === 0 &&
      documents.length === 0,
  };
}
