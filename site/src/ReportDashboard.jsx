import { useEffect, useMemo, useState } from "react";
import { createDashboardDataset, loadLocalReport, REPORT_SOURCE_PATH } from "./reportData";

const PIE_COLORS = ["navy", "gold", "cyan", "success", "warning", "navyLight", "textMuted"];

function formatNumber(value, lang) {
  return new Intl.NumberFormat(lang === "en" ? "en-US" : "fr-FR").format(value);
}

function buildPieData(data, otherLabel) {
  if (data.length <= 6) return data;

  const visible = data.slice(0, 5);
  const others = data.slice(5).reduce((total, item) => total + item.value, 0);

  return others > 0 ? [...visible, { label: otherLabel, value: others }] : visible;
}

function EmptyBlock({ message, colors, fonts }) {
  return (
    <div
      style={{
        border: `1px dashed ${colors.border}`,
        borderRadius: 8,
        padding: "20px 18px",
        background: `${colors.navy}03`,
        textAlign: "center",
      }}
    >
      <p style={{ margin: 0, fontFamily: fonts.sans, fontSize: 13, color: colors.textMuted, lineHeight: 1.7 }}>
        {message}
      </p>
    </div>
  );
}

function StatePanel({ title, description, colors, fonts, Label, variant = "navy" }) {
  return (
    <div
      style={{
        background: colors.warmWhite,
        border: `1px solid ${colors.border}`,
        borderRadius: 10,
        padding: "40px 28px",
        textAlign: "center",
      }}
    >
      <Label variant={variant}>{title}</Label>
      <p
        style={{
          margin: "18px auto 0",
          maxWidth: 560,
          fontFamily: fonts.sans,
          fontSize: 14,
          color: colors.textLight,
          lineHeight: 1.8,
        }}
      >
        {description}
      </p>
    </div>
  );
}

function SectionTitle({ eyebrow, title, description, colors, fonts, Label, Ornament, align = "left" }) {
  const isCenter = align === "center";
  return (
    <div style={{ textAlign: align, marginBottom: 24 }}>
      {eyebrow ? <Label variant="navy">{eyebrow}</Label> : null}
      <h2
        style={{
          margin: eyebrow ? "14px 0 0" : "0",
          fontFamily: fonts.serif,
          fontSize: "clamp(22px, 3vw, 34px)",
          fontWeight: 700,
          color: colors.navy,
          lineHeight: 1.2,
        }}
      >
        {title}
      </h2>
      <Ornament width={46} style={{ margin: isCenter ? "18px auto 0" : "18px 0 0" }} />
      {description ? (
        <p
          style={{
            margin: "18px 0 0",
            maxWidth: isCenter ? 760 : 640,
            fontFamily: fonts.sans,
            fontSize: 14,
            color: colors.textLight,
            lineHeight: 1.8,
          }}
        >
          {description}
        </p>
      ) : null}
    </div>
  );
}

function StatsCard({ item, colors, fonts, lang }) {
  return (
    <div
      style={{
        background: colors.warmWhite,
        border: `1px solid ${colors.border}`,
        borderRadius: 10,
        padding: "24px 22px",
        boxShadow: `0 10px 30px ${colors.navy}08`,
      }}
    >
      <div style={{ fontFamily: fonts.serif, fontSize: 34, fontWeight: 700, color: colors.navy, lineHeight: 1.05 }}>
        {formatNumber(item.value, lang)}
      </div>
      <div
        style={{
          marginTop: 10,
          fontFamily: fonts.sans,
          fontSize: 12,
          textTransform: "uppercase",
          letterSpacing: 0.7,
          color: colors.textMuted,
        }}
      >
        {item.label}
      </div>
    </div>
  );
}

function BarChartBlock({
  title,
  subtitle,
  data,
  colors,
  fonts,
  lang,
  tone = "navy",
  emptyLabel,
  footer,
}) {
  const maxValue = Math.max(...data.map((item) => item.value), 1);
  const total = data.reduce((sum, item) => sum + item.value, 0);
  const toneMap = {
    navy: `linear-gradient(90deg, ${colors.navy}, ${colors.navyLight})`,
    gold: `linear-gradient(90deg, ${colors.gold}, ${colors.goldLight})`,
    cyan: `linear-gradient(90deg, ${colors.cyan}, ${colors.navyLight})`,
  };

  return (
    <div
      style={{
        background: colors.warmWhite,
        border: `1px solid ${colors.border}`,
        borderRadius: 10,
        padding: "28px 24px",
      }}
    >
      <h3 style={{ margin: 0, fontFamily: fonts.serif, fontSize: 19, fontWeight: 700, color: colors.navy }}>{title}</h3>
      {subtitle ? (
        <p style={{ margin: "10px 0 0", fontFamily: fonts.sans, fontSize: 13, color: colors.textMuted, lineHeight: 1.7 }}>
          {subtitle}
        </p>
      ) : null}

      <div style={{ marginTop: 22 }}>
        {data.length === 0 ? (
          <EmptyBlock message={emptyLabel} colors={colors} fonts={fonts} />
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
            {data.map((item) => {
              const pct = total ? Math.round((item.value / total) * 100) : 0;
              return (
                <div key={item.label}>
                  <div style={{ display: "flex", justifyContent: "space-between", gap: 16, marginBottom: 7 }}>
                    <span style={{ fontFamily: fonts.sans, fontSize: 13, color: colors.navy }}>{item.label}</span>
                    <span style={{ fontFamily: fonts.sans, fontSize: 12, color: colors.textMuted }}>
                      {formatNumber(item.value, lang)}{total ? ` • ${pct}%` : ""}
                    </span>
                  </div>
                  <div style={{ height: 8, borderRadius: 999, background: colors.borderLight, overflow: "hidden" }}>
                    <div
                      style={{
                        width: `${(item.value / maxValue) * 100}%`,
                        height: "100%",
                        background: toneMap[tone] || toneMap.navy,
                      }}
                    />
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      {footer ? (
        <p style={{ margin: "18px 0 0", fontFamily: fonts.sans, fontSize: 12, color: colors.textMuted, lineHeight: 1.7 }}>
          {footer}
        </p>
      ) : null}
    </div>
  );
}

function PieChartBlock({ title, subtitle, data, colors, fonts, lang, emptyLabel, otherLabel }) {
  const pieData = buildPieData(data, otherLabel);
  const total = pieData.reduce((sum, item) => sum + item.value, 0);

  if (pieData.length === 0) {
    return (
      <div
        style={{
          background: colors.warmWhite,
          border: `1px solid ${colors.border}`,
          borderRadius: 10,
          padding: "28px 24px",
        }}
      >
        <h3 style={{ margin: 0, fontFamily: fonts.serif, fontSize: 19, fontWeight: 700, color: colors.navy }}>{title}</h3>
        {subtitle ? (
          <p style={{ margin: "10px 0 0", fontFamily: fonts.sans, fontSize: 13, color: colors.textMuted, lineHeight: 1.7 }}>
            {subtitle}
          </p>
        ) : null}
        <div style={{ marginTop: 22 }}>
          <EmptyBlock message={emptyLabel} colors={colors} fonts={fonts} />
        </div>
      </div>
    );
  }

  let start = 0;
  const gradient = pieData
    .map((item, index) => {
      const colorKey = PIE_COLORS[index % PIE_COLORS.length];
      const end = start + (item.value / total) * 360;
      const segment = `${colors[colorKey]} ${start}deg ${end}deg`;
      start = end;
      return segment;
    })
    .join(", ");

  return (
    <div
      style={{
        background: colors.warmWhite,
        border: `1px solid ${colors.border}`,
        borderRadius: 10,
        padding: "28px 24px",
      }}
    >
      <h3 style={{ margin: 0, fontFamily: fonts.serif, fontSize: 19, fontWeight: 700, color: colors.navy }}>{title}</h3>
      {subtitle ? (
        <p style={{ margin: "10px 0 0", fontFamily: fonts.sans, fontSize: 13, color: colors.textMuted, lineHeight: 1.7 }}>
          {subtitle}
        </p>
      ) : null}

      <div
        style={{
          marginTop: 24,
          display: "grid",
          gridTemplateColumns: "minmax(180px, 220px) minmax(0, 1fr)",
          gap: 24,
          alignItems: "center",
        }}
      >
        <div style={{ display: "flex", justifyContent: "center" }}>
          <div
            style={{
              width: 190,
              height: 190,
              borderRadius: "50%",
              background: `conic-gradient(${gradient})`,
              position: "relative",
              boxShadow: `inset 0 0 0 1px ${colors.borderLight}`,
            }}
          >
            <div
              style={{
                position: "absolute",
                inset: 30,
                borderRadius: "50%",
                background: colors.warmWhite,
                border: `1px solid ${colors.borderLight}`,
                display: "flex",
                flexDirection: "column",
                alignItems: "center",
                justifyContent: "center",
                textAlign: "center",
              }}
            >
              <span style={{ fontFamily: fonts.sans, fontSize: 11, color: colors.textMuted, textTransform: "uppercase", letterSpacing: 1 }}>
                {lang === "en" ? "Total" : "Total"}
              </span>
              <span style={{ fontFamily: fonts.serif, fontSize: 28, fontWeight: 700, color: colors.navy, lineHeight: 1.1 }}>
                {formatNumber(total, lang)}
              </span>
            </div>
          </div>
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {pieData.map((item, index) => {
            const colorKey = PIE_COLORS[index % PIE_COLORS.length];
            const pct = Math.round((item.value / total) * 100);
            return (
              <div
                key={item.label}
                style={{
                  display: "grid",
                  gridTemplateColumns: "14px minmax(0, 1fr) auto",
                  gap: 12,
                  alignItems: "center",
                }}
              >
                <span style={{ width: 14, height: 14, borderRadius: 3, background: colors[colorKey], display: "inline-block" }} />
                <span style={{ fontFamily: fonts.sans, fontSize: 13, color: colors.navy }}>{item.label}</span>
                <span style={{ fontFamily: fonts.sans, fontSize: 12, color: colors.textMuted }}>
                  {formatNumber(item.value, lang)} • {pct}%
                </span>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

function WordCloudBlock({ title, subtitle, words, colors, fonts, emptyLabel }) {
  const topCount = Math.max(...words.map((word) => word.count), 1);

  return (
    <div
      style={{
        background: colors.warmWhite,
        border: `1px solid ${colors.border}`,
        borderRadius: 10,
        padding: "28px 24px",
      }}
    >
      <h3 style={{ margin: 0, fontFamily: fonts.serif, fontSize: 19, fontWeight: 700, color: colors.navy }}>{title}</h3>
      {subtitle ? (
        <p style={{ margin: "10px 0 0", fontFamily: fonts.sans, fontSize: 13, color: colors.textMuted, lineHeight: 1.7 }}>
          {subtitle}
        </p>
      ) : null}

      <div style={{ marginTop: 22 }}>
        {words.length === 0 ? (
          <EmptyBlock message={emptyLabel} colors={colors} fonts={fonts} />
        ) : (
          <div
            style={{
              minHeight: 260,
              display: "flex",
              flexWrap: "wrap",
              alignItems: "center",
              justifyContent: "center",
              gap: "8px 14px",
              padding: "18px 10px",
            }}
          >
            {words.map((word, index) => {
              const opacity = 0.35 + (word.count / topCount) * 0.65;
              const color = index % 3 === 0 ? colors.gold : index % 3 === 1 ? colors.navy : colors.cyan;

              return (
                <span
                  key={word.text}
                  style={{
                    fontFamily: fonts.serif,
                    fontSize: Math.max(14, word.size),
                    fontWeight: word.count > topCount * 0.65 ? 700 : 500,
                    color,
                    opacity,
                    lineHeight: 1.25,
                  }}
                >
                  {word.text}
                </span>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

export default function ReportDashboard({ lang, colors, fonts, Label, Ornament, FadeIn }) {
  const [rawReport, setRawReport] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const isEn = lang === "en";

  useEffect(() => {
    let cancelled = false;

    async function fetchReport() {
      try {
        setLoading(true);
        setError(null);
        const report = await loadLocalReport();
        if (!cancelled) setRawReport(report);
      } catch (loadError) {
        if (!cancelled) setError(loadError instanceof Error ? loadError : new Error("Unable to load report"));
      } finally {
        if (!cancelled) setLoading(false);
      }
    }

    fetchReport();

    return () => {
      cancelled = true;
    };
  }, []);

  const dashboard = useMemo(
    () => (rawReport ? createDashboardDataset(rawReport, lang) : null),
    [rawReport, lang]
  );

  return (
    <div style={{ paddingTop: 68, minHeight: "100vh", background: colors.cream }}>
      <div style={{ maxWidth: 1080, margin: "0 auto", padding: "48px 24px 72px" }}>
        <div style={{ textAlign: "center", marginBottom: 42 }}>
          <h1
            style={{
              margin: 0,
              fontFamily: fonts.serif,
              fontSize: "clamp(26px, 4vw, 42px)",
              fontWeight: 700,
              color: colors.navy,
              lineHeight: 1.15,
            }}
          >
            {isEn ? "Dashboard" : "Tableau de bord"}
          </h1>
          <Ornament width={50} style={{ margin: "20px auto" }} />
          <p
            style={{
              margin: "0 auto",
              maxWidth: 760,
              fontFamily: fonts.sans,
              fontSize: 15,
              color: colors.textLight,
              lineHeight: 1.8,
            }}
          >
            {isEn
              ? "This dashboard presents the main results of a textual analysis conducted on a corpus of disinformation-related content. It highlights the circulation platforms, the dominant topics and the rhetorical patterns that structure the analyzed documents."
              : "Ce tableau de bord presente les principaux resultats d'une analyse textuelle menee sur un corpus de contenus lies a la desinformation. Il met en lumiere les plateformes de circulation, les thematiques dominantes et les patterns rhetoriques qui structurent les documents analyses."}
          </p>
        </div>

        {loading ? (
          <StatePanel
            title={isEn ? "Loading" : "Chargement"}
            description={
              isEn
                ? "The analysis dataset is loading."
                : "Les donnees d'analyse sont en cours de chargement."
            }
            colors={colors}
            fonts={fonts}
            Label={Label}
            variant="navy"
          />
        ) : null}

        {!loading && error ? (
          <StatePanel
            title={isEn ? "Load error" : "Erreur de chargement"}
            description={
              isEn
                ? `The dashboard could not load the expected data file (${REPORT_SOURCE_PATH}). Check that it exists and contains valid JSON.`
                : `Le dashboard n'a pas pu charger le fichier de donnees attendu (${REPORT_SOURCE_PATH}). Verifie qu'il existe bien et qu'il contient un JSON valide.`
            }
            colors={colors}
            fonts={fonts}
            Label={Label}
            variant="error"
          />
        ) : null}

        {!loading && !error && dashboard?.isEmpty ? (
          <StatePanel
            title={isEn ? "Empty report" : "Rapport vide"}
            description={
              isEn
                ? "The file was loaded, but no exploitable statistics were found in summary, count maps or documents."
                : "Le fichier a bien ete charge, mais aucune statistique exploitable n'a ete trouvee dans summary, les maps de comptage ou les documents."
            }
            colors={colors}
            fonts={fonts}
            Label={Label}
            variant="gold"
          />
        ) : null}

        {!loading && !error && dashboard && !dashboard.isEmpty ? (
          <>
            <FadeIn>
              <section style={{ marginBottom: 40 }}>
                <SectionTitle
                  eyebrow=""
                  title={isEn ? "Overview" : "Vue d'ensemble"}
                  description=""
                  colors={colors}
                  fonts={fonts}
                  Label={Label}
                  Ornament={Ornament}
                  align="center"
                />

                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 16 }}>
                  {dashboard.summaryCards.map((item) => (
                    <StatsCard key={item.id} item={item} colors={colors} fonts={fonts} lang={lang} />
                  ))}
                </div>
              </section>
            </FadeIn>

            <FadeIn delay={80}>
              <section style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 20, marginBottom: 40 }}>
                <BarChartBlock
                  title={isEn ? "Platform ranking" : "Classement des plateformes"}
                  subtitle={
                    isEn
                      ? "Ranking of the platforms from which fake news originate. The category \"Unknown\" (intentionally hidden) exists in the dataset but is not displayed here."
                      : "Classement des plateformes d'ou sont issues les fake news. La categorie \"Inconnu\" (volontairement masquee) est bien presente dans les donnees, mais elle n'est pas affichee ici."
                  }
                  data={dashboard.platformRanking}
                  colors={colors}
                  fonts={fonts}
                  lang={lang}
                  tone="gold"
                  emptyLabel={isEn ? "No platform data available in the report." : "Aucune donnee de plateforme disponible dans le rapport."}
                />

                <PieChartBlock
                  title={isEn ? "Topic distribution" : "Repartition par thematiques"}
                  subtitle=""
                  data={dashboard.topics}
                  colors={colors}
                  fonts={fonts}
                  lang={lang}
                  emptyLabel={isEn ? "No topic data available in the report." : "Aucune donnee thematique disponible dans le rapport."}
                  otherLabel={isEn ? "Other" : "Autres"}
                />
              </section>
            </FadeIn>

            <FadeIn delay={140}>
              <section style={{ marginBottom: 40 }}>
                <BarChartBlock
                  title={isEn ? "Rhetorical patterns" : "Patterns rhetoriques"}
                  subtitle={
                    isEn
                      ? "Rhetorical patterns are recurring formulations used to trigger emotion, urgency, sharing reflexes or suspicion in misleading content."
                      : "Les patterns rhetoriques designent des formulations recurrentes utilisees pour susciter l'emotion, l'urgence, le partage ou la suspicion dans les contenus trompeurs."
                  }
                  data={dashboard.rhetoricalPatterns}
                  colors={colors}
                  fonts={fonts}
                  lang={lang}
                  tone="gold"
                  emptyLabel={
                    isEn
                      ? "No rhetorical-pattern data available in the report."
                      : "Aucune donnee de pattern rhetorique disponible dans le rapport."
                  }
                />
              </section>
            </FadeIn>

            {dashboard.usages.length ? (
              <FadeIn delay={200}>
                <section style={{ marginBottom: 40 }}>
                  <BarChartBlock
                    title={isEn ? "Real-world usages" : "Usages reels"}
                    subtitle={
                      isEn
                        ? "Optional section generated from the usage_counts object when it is present in the analysis data."
                        : "Section optionnelle generee a partir de l'objet usage_counts lorsqu'il est present dans les donnees d'analyse."
                    }
                    data={dashboard.usages}
                    colors={colors}
                    fonts={fonts}
                    lang={lang}
                    tone="cyan"
                    emptyLabel={isEn ? "No usage data available in the report." : "Aucune donnee d'usage disponible dans le rapport."}
                  />
                </section>
              </FadeIn>
            ) : null}

            <FadeIn delay={260}>
              <section>
                <WordCloudBlock
                  title={isEn ? "OCR word cloud" : "Nuage de mots OCR"}
                  subtitle={
                    isEn
                      ? "Most frequent words extracted by OCR from the analyzed documents."
                      : "Mots les plus frequents extraits par OCR dans les documents analyses."
                  }
                  words={dashboard.ocrWords}
                  colors={colors}
                  fonts={fonts}
                  emptyLabel={isEn ? "No OCR vocabulary is available in the report." : "Aucun vocabulaire OCR n'est disponible dans le rapport."}
                />
              </section>
            </FadeIn>
          </>
        ) : null}
      </div>
    </div>
  );
}
