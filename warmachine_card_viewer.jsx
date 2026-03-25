import { useState, useCallback, useRef, useEffect } from "react";

// ─── Color Palette ───
const C = {
  bg: "#1a1a1a",
  card: "#2a2a2a",
  cardDark: "#222",
  headerBg: "#1e3a2f",
  headerGrad: "linear-gradient(135deg, #1e3a2f 0%, #2d5a45 50%, #1e3a2f 100%)",
  accent: "#3d8b6e",
  accentLight: "#4da67e",
  accentDim: "#2a5e4a",
  gold: "#c9a84c",
  text: "#e8e8e8",
  textDim: "#999",
  textMuted: "#666",
  ranged: "#3a6b5a",
  melee: "#4a4a4a",
  statBg: "#1a2e25",
  statBorder: "#3d8b6e",
  healthBox: "#555",
  healthBoxEmpty: "#333",
  gridBox: "#4a6a5a",
  gridMissing: "#1a1a1a",
  spellBg: "#1e2e28",
  featBg: "#2a1a1a",
  ruleBg: "#1a2a24",
  border: "#3a3a3a",
  scrollThumb: "#3d8b6e",
};

// ─── Fonts via Google ───
const fontLink = "https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=Source+Sans+3:wght@300;400;600;700&family=Fira+Code:wght@400&display=swap";

// ─── Stat Labels ───
const STAT_ORDER = ["spd","aat","mat","rat","def","arm","arc","fury","ctrl","thr","ess"];
const STAT_LABELS = {spd:"SPD",aat:"AAT",mat:"MAT",rat:"RAT",def:"DEF",arm:"ARM",arc:"ARC",fury:"FURY",ctrl:"CTRL",thr:"THR",ess:"ESS"};

const SPELL_STAT_ORDER = ["cost","rng","aoe","pow","dur","off"];
const SPELL_STAT_LABELS = {cost:"COST",rng:"RNG",aoe:"AOE",pow:"POW",dur:"DUR",off:"OFF"};

// ─── Grid System Letters ───
const GRID_COLORS = { L:"#5588cc", M:"#cc8855", C:"#cc5555", R:"#55aa55", H:"#aa55aa", A:"#8888cc", "-":"transparent", " ":"#4a6a5a" };
const GRID_LABELS = { L:"L", M:"M", C:"C", R:"R", H:"H", A:"A" };

// ─── Utility ───
const typeColor = (type) => {
  const t = (type||"").toLowerCase();
  if (t.includes("warcaster") || t.includes("warlock")) return "#8b6b3d";
  if (t.includes("warjack") || t.includes("warbeast")) return "#3d6b8b";
  if (t.includes("unit")) return "#6b3d8b";
  if (t.includes("solo")) return "#3d8b6b";
  if (t.includes("battle engine")) return "#6b6b3d";
  if (t.includes("command")) return "#8b3d6b";
  return "#5a5a5a";
};

// ─── Global Styles ───
const globalCSS = `
  @import url('${fontLink}');
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { background: ${C.bg}; color: ${C.text}; font-family: 'Source Sans 3', sans-serif; }
  ::-webkit-scrollbar { width: 6px; }
  ::-webkit-scrollbar-track { background: ${C.bg}; }
  ::-webkit-scrollbar-thumb { background: ${C.scrollThumb}; border-radius: 3px; }
  @keyframes fadeIn { from { opacity: 0; transform: translateY(8px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes slideIn { from { opacity: 0; transform: translateX(-12px); } to { opacity: 1; transform: translateX(0); } }
  @keyframes pulse { 0%,100% { opacity: 0.4; } 50% { opacity: 1; } }
`;

// ─── Components ───

function StatBar({ stats, labels = STAT_LABELS, order = STAT_ORDER, compact = false }) {
  const visible = order.filter(k => stats[k] !== undefined && stats[k] !== "");
  if (!visible.length) return null;
  return (
    <div style={{
      display: "flex", gap: compact ? 0 : 1, borderRadius: 4, overflow: "hidden",
      border: `1px solid ${C.statBorder}40`, fontSize: compact ? 11 : 13,
    }}>
      {visible.map(k => (
        <div key={k} style={{
          display: "flex", flexDirection: "column", alignItems: "center",
          minWidth: compact ? 32 : 38, flex: 1,
        }}>
          <div style={{
            background: C.statBg, color: C.textDim, padding: "2px 4px",
            width: "100%", textAlign: "center",
            fontFamily: "'Oswald', sans-serif", fontWeight: 500, fontSize: compact ? 9 : 10,
            letterSpacing: 1, textTransform: "uppercase",
          }}>{labels[k]}</div>
          <div style={{
            background: C.card, padding: "3px 4px", width: "100%", textAlign: "center",
            fontFamily: "'Oswald', sans-serif", fontWeight: 700, fontSize: compact ? 14 : 16,
            color: C.text,
          }}>{stats[k]}</div>
        </div>
      ))}
    </div>
  );
}

function HealthBoxes({ health }) {
  if (!health || !health.type) return null;
  const t = health.type;

  if (t === "Box" || t === "Box Multi") {
    const names = health.names || [];
    const values = health.values || [];
    return (
      <div style={{ padding: "8px 12px" }}>
        {names.map((name, i) => {
          const count = parseInt(values[i]?.[0] || values[i] || "0");
          return (
            <div key={i} style={{ marginBottom: i < names.length - 1 ? 8 : 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 4 }}>
                <span style={{ color: "#e55", fontFamily: "'Oswald'", fontWeight: 700, fontSize: 14 }}>♥{count}</span>
                {name && <span style={{ color: C.textDim, fontSize: 12, fontStyle: "italic" }}>{name}</span>}
              </div>
              <div style={{ display: "flex", gap: 2, flexWrap: "wrap" }}>
                {Array.from({ length: count }, (_, j) => (
                  <div key={j} style={{
                    width: 18, height: 18, borderRadius: 2,
                    background: j >= count - 2 ? `${C.accent}55` : C.healthBox,
                    border: `1px solid ${j >= count - 2 ? C.accent : C.border}`,
                  }} />
                ))}
              </div>
            </div>
          );
        })}
      </div>
    );
  }

  if (t === "Grid" || t === "Grid Large") {
    const rows = health.values || [];
    const totalBoxes = rows.reduce((s, row) => {
      const cells = row.split(",");
      return s + cells.filter(c => c.trim() !== "-").length;
    }, 0);
    return (
      <div style={{ padding: "8px 12px" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
          <span style={{ color: "#e55", fontFamily: "'Oswald'", fontWeight: 700, fontSize: 14 }}>♥{totalBoxes}</span>
        </div>
        <div style={{ display: "flex", gap: 1, marginBottom: 4 }}>
          {["", "1", "2", "3", "4", "5", "6"].map((h, i) => (
            <div key={i} style={{
              width: 28, textAlign: "center",
              fontFamily: "'Oswald'", fontSize: 11, fontWeight: 600,
              color: i === 0 ? "transparent" : C.textDim,
            }}>{h}</div>
          ))}
        </div>
        <div style={{ display: "flex", flexDirection: "column-reverse", gap: 1 }}>
          {rows.map((row, ri) => {
            const cells = row.split(",");
            return (
              <div key={ri} style={{ display: "flex", gap: 1 }}>
                <div style={{
                  width: 28, display: "flex", alignItems: "center", justifyContent: "center",
                  fontFamily: "'Oswald'", fontSize: 10, color: C.textDim,
                }}>{rows.length - ri}</div>
                {cells.map((cell, ci) => {
                  const val = cell.trim();
                  const isMissing = val === "-";
                  const letter = val && val !== " " && !isMissing ? val : "";
                  const color = GRID_COLORS[val] || GRID_COLORS[" "];
                  return (
                    <div key={ci} style={{
                      width: 28, height: 22, borderRadius: 2,
                      background: isMissing ? "transparent" : `${color}40`,
                      border: isMissing ? "1px solid transparent" : `1px solid ${color}80`,
                      display: "flex", alignItems: "center", justifyContent: "center",
                      fontSize: 10, fontWeight: 700, color: `${color}`,
                      fontFamily: "'Oswald'",
                    }}>
                      {letter && GRID_LABELS[letter] ? GRID_LABELS[letter] : ""}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
        <div style={{ display: "flex", gap: 8, marginTop: 6, flexWrap: "wrap" }}>
          {Object.entries(GRID_LABELS).map(([k, v]) => {
            const hasLetter = rows.some(r => r.includes(k));
            if (!hasLetter) return null;
            return (
              <span key={k} style={{ fontSize: 10, color: GRID_COLORS[k], fontFamily: "'Oswald'" }}>
                ■ {k === "L" ? "Left" : k === "R" ? "Right" : k === "M" ? "Movement" : k === "C" ? "Cortex" : k === "H" ? "Head" : k === "A" ? "Arc Node"}
              </span>
            );
          })}
        </div>
      </div>
    );
  }
  return null;
}

function WeaponBlock({ weapon, properties, isLast }) {
  const isRanged = weapon.type === "Ranged";
  const statOrder = isRanged ? ["rat","rng","rof","aoe","pow"] : ["mat","rng","pow"];
  const statLabels = { rat:"RAT", mat:"MAT", rng:"RNG", rof:"ROF", aoe:"AOE", pow:"POW" };
  const wStats = { ...weapon.stats };
  // For display, we combine count and name
  const displayName = (parseInt(weapon.count) > 1 ? `${weapon.name} x${weapon.count}` : weapon.name)
    + (weapon.location ? ` [${weapon.location}]` : "");

  return (
    <div style={{
      marginBottom: isLast ? 0 : 2, borderLeft: `3px solid ${isRanged ? C.ranged : C.melee}`,
      background: `${isRanged ? C.ranged : C.melee}15`,
    }}>
      <div style={{ padding: "6px 10px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
          <div>
            <div style={{
              fontFamily: "'Oswald'", fontWeight: 600, fontSize: 15,
              color: isRanged ? "#7bc4a8" : C.text,
            }}>{displayName}</div>
            <div style={{ fontSize: 10, color: C.textDim, marginTop: 1 }}>
              {isRanged ? "⚡ Ranged" : "⚔ Melee"}
            </div>
          </div>
          <div style={{ display: "flex", gap: 1 }}>
            {statOrder.map(k => {
              const v = wStats[k];
              if (v === undefined || v === "") return null;
              return (
                <div key={k} style={{ textAlign: "center", minWidth: 30 }}>
                  <div style={{
                    fontSize: 9, color: C.textDim, fontFamily: "'Oswald'",
                    letterSpacing: 0.5, fontWeight: 500,
                  }}>{statLabels[k]}</div>
                  <div style={{
                    fontSize: 14, fontWeight: 700, fontFamily: "'Oswald'",
                    color: C.text,
                  }}>{v}</div>
                </div>
              );
            })}
          </div>
        </div>
        {weapon.properties && weapon.properties.length > 0 && (
          <div style={{ marginTop: 6 }}>
            {weapon.properties.map((prop, i) => (
              <div key={i} style={{ marginBottom: 3 }}>
                <span style={{
                  fontSize: 12, fontWeight: 600,
                  color: prop.startsWith("•") || prop.startsWith("Attack Type") ? C.accentLight : C.gold,
                }}>{prop}</span>
                {properties[prop] && (
                  <span style={{ fontSize: 11, color: C.textDim, marginLeft: 4 }}>
                    — {properties[prop]}
                  </span>
                )}
              </div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function SpellBlock({ name, spell }) {
  if (!spell) return null;
  return (
    <div style={{
      background: C.spellBg, borderRadius: 4, padding: "8px 10px", marginBottom: 4,
      borderLeft: `3px solid ${C.accent}`,
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
        <div style={{
          fontFamily: "'Oswald'", fontWeight: 600, fontSize: 14, color: C.accentLight,
        }}>{name}</div>
        <div style={{ display: "flex", gap: 1 }}>
          {SPELL_STAT_ORDER.map(k => {
            const v = spell.stats?.[k];
            if (v === undefined || v === "") return null;
            return (
              <div key={k} style={{ textAlign: "center", minWidth: 26 }}>
                <div style={{
                  fontSize: 8, color: C.textDim, fontFamily: "'Oswald'",
                  letterSpacing: 0.5, fontWeight: 500,
                }}>{SPELL_STAT_LABELS[k]}</div>
                <div style={{
                  fontSize: 12, fontWeight: 700, fontFamily: "'Oswald'", color: C.text,
                }}>{v}</div>
              </div>
            );
          })}
        </div>
      </div>
      {spell.text && (
        <div style={{ fontSize: 11, color: C.textDim, marginTop: 4, lineHeight: 1.4 }}>
          {spell.text}
        </div>
      )}
    </div>
  );
}

function SectionHeader({ children, icon }) {
  return (
    <div style={{
      fontFamily: "'Oswald'", fontWeight: 600, fontSize: 11,
      letterSpacing: 2, textTransform: "uppercase",
      color: C.textMuted, padding: "8px 12px 4px",
      borderTop: `1px solid ${C.border}`,
      display: "flex", alignItems: "center", gap: 6,
    }}>
      {icon && <span style={{ fontSize: 13 }}>{icon}</span>}
      {children}
    </div>
  );
}

function ModelCard({ card, data }) {
  if (!card) return null;
  const spells = data.spells || {};
  const abilities = data.abilities || {};
  const properties = data.properties || {};
  const tc = typeColor(card.type);

  return (
    <div style={{
      width: 380, minHeight: 200, background: C.cardDark,
      borderRadius: 8, overflow: "hidden",
      border: `1px solid ${C.border}`,
      boxShadow: "0 4px 24px rgba(0,0,0,0.5)",
      animation: "fadeIn 0.3s ease-out",
      display: "flex", flexDirection: "column",
    }}>
      {/* ─── Header ─── */}
      <div style={{
        background: C.headerGrad, padding: "14px 14px 10px",
        borderBottom: `2px solid ${C.accent}40`,
        position: "relative", overflow: "hidden",
      }}>
        <div style={{
          position: "absolute", top: 0, right: 0, width: 120, height: "100%",
          background: `linear-gradient(90deg, transparent, ${tc}30)`,
        }} />
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div>
            <div style={{
              fontFamily: "'Oswald'", fontWeight: 700, fontSize: 20,
              color: "#fff", lineHeight: 1.1, textShadow: "0 1px 3px rgba(0,0,0,0.5)",
            }}>{card.name}</div>
            <div style={{
              fontSize: 11, color: `${C.text}aa`, marginTop: 2,
              fontStyle: "italic",
            }}>{card.faction} {card.type}</div>
          </div>
          <div style={{ textAlign: "right" }}>
            {card.cost && card.cost !== "0" && (
              <div style={{
                background: tc, color: "#fff", borderRadius: 4,
                padding: "2px 8px", fontFamily: "'Oswald'", fontWeight: 700,
                fontSize: 14,
              }}>
                {card.cost} pts
              </div>
            )}
          </div>
        </div>
        <div style={{ display: "flex", gap: 12, marginTop: 6, fontSize: 11, color: `${C.text}88` }}>
          <span>FA: <strong style={{ color: C.text }}>{card.fa}</strong></span>
          {card.composition && <span>{card.composition}</span>}
          {card.keywords?.filter(k=>k).length > 0 && (
            <span style={{ color: C.gold }}>{card.keywords.filter(k=>k).join(", ")}</span>
          )}
        </div>
      </div>

      {/* ─── Health ─── */}
      {card.health && card.health.type && (
        <>
          <SectionHeader icon="♥">Health</SectionHeader>
          <HealthBoxes health={card.health} />
        </>
      )}

      {/* ─── Rules (top-level card rules) ─── */}
      {card.rules && Object.keys(card.rules).length > 0 && (
        <>
          <SectionHeader icon="📋">Card Rules</SectionHeader>
          <div style={{ padding: "4px 12px 8px" }}>
            {Object.entries(card.rules).map(([name, desc]) => (
              <div key={name} style={{
                marginBottom: 4, background: C.ruleBg, borderRadius: 4,
                padding: "6px 8px", borderLeft: `2px solid ${C.gold}`,
              }}>
                <span style={{ fontWeight: 700, fontSize: 12, color: C.gold }}>{name}</span>
                <span style={{ fontSize: 11, color: C.textDim, marginLeft: 4 }}>— {desc}</span>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ─── Profiles ─── */}
      {card.profiles?.map((profile, pi) => (
        <div key={pi}>
          <SectionHeader icon="👤">{profile.name}</SectionHeader>
          <div style={{ padding: "6px 12px" }}>
            <StatBar stats={profile.stats} />
            {/* Abilities */}
            {profile.abilities?.length > 0 && (
              <div style={{ marginTop: 6 }}>
                <div style={{
                  fontSize: 11, fontWeight: 600, color: C.accentLight,
                  marginBottom: 3, display: "flex", flexWrap: "wrap", gap: 2,
                }}>
                  {profile.abilities.map((a, i) => (
                    <span key={i}>
                      <span style={{ color: C.accentLight }}>{a}</span>
                      {i < profile.abilities.length - 1 && <span style={{ color: C.textMuted }}>, </span>}
                    </span>
                  ))}
                </div>
                {profile.abilities.map((a, i) => (
                  abilities[a] ? (
                    <div key={i} style={{
                      fontSize: 11, color: C.textDim, marginBottom: 3,
                      padding: "3px 6px", background: `${C.accent}10`, borderRadius: 3,
                      lineHeight: 1.4,
                    }}>
                      <strong style={{ color: C.text }}>{a}</strong> — {abilities[a]}
                    </div>
                  ) : null
                ))}
              </div>
            )}
          </div>

          {/* Weapons */}
          {profile.weapons?.length > 0 && (
            <>
              <SectionHeader icon="⚔">Weapons</SectionHeader>
              <div style={{ padding: "4px 12px 8px" }}>
                {profile.weapons.map((w, wi) => (
                  <WeaponBlock
                    key={wi}
                    weapon={w}
                    properties={properties}
                    isLast={wi === profile.weapons.length - 1}
                  />
                ))}
              </div>
            </>
          )}
        </div>
      ))}

      {/* ─── Spells ─── */}
      {card.spells?.length > 0 && (
        <>
          <SectionHeader icon="✦">Spells</SectionHeader>
          <div style={{ padding: "4px 12px 8px" }}>
            {card.spells.map((sName, si) => (
              <SpellBlock key={si} name={sName} spell={spells[sName]} />
            ))}
          </div>
        </>
      )}

      {/* ─── Rack ─── */}
      {card.rack && card.rack !== "" && (
        <div style={{ padding: "2px 12px 6px" }}>
          <span style={{ fontSize: 10, color: C.textMuted }}>Spell Rack: {card.rack} slots</span>
        </div>
      )}

      {/* ─── Feat ─── */}
      {card.feat && Object.keys(card.feat).length > 0 && (
        <>
          <SectionHeader icon="🔥">Feat</SectionHeader>
          <div style={{ padding: "4px 12px 10px" }}>
            {Object.entries(card.feat).map(([name, desc]) => (
              <div key={name} style={{
                background: `${C.featBg}`, borderRadius: 4, padding: "8px 10px",
                borderLeft: `3px solid #c44`,
              }}>
                <div style={{
                  fontFamily: "'Oswald'", fontWeight: 700, fontSize: 15,
                  color: "#e77", marginBottom: 4,
                }}>{name}</div>
                <div style={{ fontSize: 11, color: C.textDim, lineHeight: 1.4 }}>{desc}</div>
              </div>
            ))}
          </div>
        </>
      )}

      {/* ─── Options / Hardpoints ─── */}
      {card.options?.length > 0 && (
        <>
          <SectionHeader icon="🔧">Options</SectionHeader>
          <div style={{ padding: "4px 12px 10px" }}>
            {card.options.map((opt, oi) => (
              <div key={oi} style={{
                background: `${C.statBg}`, borderRadius: 4, padding: "8px 10px",
                marginBottom: 4, borderLeft: `3px solid ${C.gold}`,
              }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div>
                    <span style={{
                      fontFamily: "'Oswald'", fontWeight: 600, fontSize: 14, color: C.gold,
                    }}>{opt.name}</span>
                    <span style={{ fontSize: 10, color: C.textMuted, marginLeft: 8 }}>
                      [{opt.hardpoints?.join(", ")}]
                    </span>
                  </div>
                  {opt.cost && opt.cost !== "0" && (
                    <span style={{
                      fontSize: 11, fontWeight: 700, color: C.gold,
                      background: `${C.gold}20`, padding: "1px 6px", borderRadius: 3,
                    }}>+{opt.cost} pts</span>
                  )}
                </div>
                {opt.abilities?.length > 0 && (
                  <div style={{ marginTop: 4 }}>
                    {opt.abilities.map((a, i) => (
                      <div key={i} style={{ fontSize: 11, color: C.textDim }}>
                        <strong style={{ color: C.accentLight }}>{a}</strong>
                        {abilities[a] && <span> — {abilities[a]}</span>}
                      </div>
                    ))}
                  </div>
                )}
                {opt.weapons?.length > 0 && (
                  <div style={{ marginTop: 4 }}>
                    {opt.weapons.map((w, wi) => (
                      <WeaponBlock key={wi} weapon={w} properties={properties} isLast={wi === opt.weapons.length - 1} />
                    ))}
                  </div>
                )}
                {opt.profiles?.length > 0 && opt.profiles.map((p, pi) => (
                  <div key={pi} style={{ marginTop: 6 }}>
                    <div style={{ fontSize: 12, fontWeight: 600, color: C.text, marginBottom: 3 }}>{p.name}</div>
                    <StatBar stats={p.stats} compact />
                    {p.abilities?.length > 0 && (
                      <div style={{ marginTop: 3 }}>
                        {p.abilities.map((a, ai) => (
                          <div key={ai} style={{ fontSize: 10, color: C.textDim }}>
                            <strong style={{ color: C.accentLight }}>{a}</strong>
                            {abilities[a] && <span> — {abilities[a]}</span>}
                          </div>
                        ))}
                      </div>
                    )}
                    {p.weapons?.length > 0 && (
                      <div style={{ marginTop: 3 }}>
                        {p.weapons.map((w, wi) => (
                          <WeaponBlock key={wi} weapon={w} properties={properties} isLast={wi === p.weapons.length - 1} />
                        ))}
                      </div>
                    )}
                  </div>
                ))}
              </div>
            ))}
          </div>
        </>
      )}

      {/* ─── Footer ─── */}
      <div style={{
        marginTop: "auto", padding: "6px 12px",
        borderTop: `1px solid ${C.border}`,
        fontSize: 9, color: C.textMuted, textAlign: "center",
        fontStyle: "italic",
      }}>
        Warmachine Card Viewer • {card.faction} — {card.type}
      </div>
    </div>
  );
}

// ─── Main App ───
export default function WarmachineCardViewer() {
  const [data, setData] = useState(null);
  const [selectedCard, setSelectedCard] = useState(null);
  const [filter, setFilter] = useState("");
  const [typeFilter, setTypeFilter] = useState("All");
  const [error, setError] = useState(null);
  const [dragging, setDragging] = useState(false);
  const fileRef = useRef(null);

  const loadJSON = useCallback((text) => {
    try {
      const parsed = JSON.parse(text);
      if (!parsed.cards || !Array.isArray(parsed.cards)) {
        setError("Invalid format: missing 'cards' array");
        return;
      }
      setData(parsed);
      setSelectedCard(parsed.cards[0] || null);
      setError(null);
    } catch (e) {
      setError(`JSON parse error: ${e.message}`);
    }
  }, []);

  const handleFile = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => loadJSON(e.target.result);
    reader.readAsText(file);
  }, [loadJSON]);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith(".json")) handleFile(file);
    else setError("Please drop a .json file");
  }, [handleFile]);

  // Load the demo data from storage
  useEffect(() => {
    // Try to auto-load if data is available via storage
    (async () => {
      try {
        const result = await window.storage.get("card-data");
        if (result?.value) {
          loadJSON(result.value);
        }
      } catch {}
    })();
  }, [loadJSON]);

  const types = data ? ["All", ...new Set(data.cards.map(c => c.type))] : [];
  const filtered = data ? data.cards.filter(c => {
    const matchType = typeFilter === "All" || c.type === typeFilter;
    const matchName = !filter || c.name.toLowerCase().includes(filter.toLowerCase());
    return matchType && matchName;
  }) : [];

  // Group by faction
  const grouped = {};
  filtered.forEach(c => {
    const f = c.faction || "Unknown";
    if (!grouped[f]) grouped[f] = [];
    grouped[f].push(c);
  });

  if (!data) {
    return (
      <div style={{ minHeight: "100vh", display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", padding: 40 }}>
        <style>{globalCSS}</style>
        <div style={{
          fontFamily: "'Oswald'", fontSize: 36, fontWeight: 700,
          color: C.accent, marginBottom: 4, letterSpacing: 2,
        }}>WARMACHINE</div>
        <div style={{
          fontFamily: "'Oswald'", fontSize: 16, fontWeight: 400,
          color: C.textDim, marginBottom: 32, letterSpacing: 4, textTransform: "uppercase",
        }}>Card Viewer</div>

        <div
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          onClick={() => fileRef.current?.click()}
          style={{
            width: 400, maxWidth: "90vw", height: 180,
            border: `2px dashed ${dragging ? C.accent : C.border}`,
            borderRadius: 12, display: "flex", flexDirection: "column",
            alignItems: "center", justifyContent: "center", gap: 12,
            cursor: "pointer", transition: "all 0.2s",
            background: dragging ? `${C.accent}10` : "transparent",
          }}
        >
          <div style={{ fontSize: 40, opacity: 0.4 }}>📂</div>
          <div style={{ color: C.textDim, fontSize: 14, textAlign: "center" }}>
            Drop a JSON file here or click to browse
          </div>
          <div style={{ color: C.textMuted, fontSize: 11 }}>
            Supports Card Creator V4 simple format
          </div>
        </div>
        <input
          ref={fileRef}
          type="file"
          accept=".json"
          style={{ display: "none" }}
          onChange={(e) => handleFile(e.target.files[0])}
        />
        {error && (
          <div style={{
            marginTop: 16, color: "#e55", fontSize: 13,
            background: "#e5551a", padding: "8px 16px", borderRadius: 6,
          }}>{error}</div>
        )}
        <div style={{ marginTop: 24, color: C.textMuted, fontSize: 11, textAlign: "center", maxWidth: 400 }}>
          Load your card data JSON file (same format as Soul Samurai's Card Creator V4).
          Image paths in the JSON are noted but images are not loaded in this viewer.
        </div>
      </div>
    );
  }

  return (
    <div style={{ display: "flex", height: "100vh", overflow: "hidden" }}>
      <style>{globalCSS}</style>

      {/* ─── Sidebar ─── */}
      <div style={{
        width: 280, minWidth: 280, background: C.cardDark,
        borderRight: `1px solid ${C.border}`,
        display: "flex", flexDirection: "column", overflow: "hidden",
      }}>
        {/* Sidebar Header */}
        <div style={{
          padding: "14px 14px 10px",
          borderBottom: `1px solid ${C.border}`,
          background: C.headerGrad,
        }}>
          <div style={{
            fontFamily: "'Oswald'", fontWeight: 700, fontSize: 16,
            color: C.accent, letterSpacing: 1,
          }}>WARMACHINE</div>
          <div style={{ fontSize: 10, color: C.textMuted, letterSpacing: 2, textTransform: "uppercase" }}>
            Card Viewer — {data.cards.length} models
          </div>
        </div>

        {/* Search + Filter */}
        <div style={{ padding: "10px 12px 6px" }}>
          <input
            value={filter}
            onChange={e => setFilter(e.target.value)}
            placeholder="Search models..."
            style={{
              width: "100%", padding: "6px 10px", borderRadius: 4,
              border: `1px solid ${C.border}`, background: C.bg,
              color: C.text, fontSize: 12, fontFamily: "'Source Sans 3'",
              outline: "none",
            }}
          />
          <div style={{ display: "flex", flexWrap: "wrap", gap: 3, marginTop: 6 }}>
            {types.map(t => (
              <button
                key={t}
                onClick={() => setTypeFilter(t)}
                style={{
                  padding: "2px 8px", borderRadius: 3, border: "none",
                  fontSize: 10, cursor: "pointer",
                  fontFamily: "'Oswald'", letterSpacing: 0.5,
                  background: typeFilter === t ? C.accent : C.bg,
                  color: typeFilter === t ? "#fff" : C.textDim,
                  transition: "all 0.15s",
                }}
              >{t}</button>
            ))}
          </div>
        </div>

        {/* Card List */}
        <div style={{ flex: 1, overflowY: "auto", padding: "4px 8px" }}>
          {Object.entries(grouped).map(([faction, cards]) => (
            <div key={faction}>
              <div style={{
                fontFamily: "'Oswald'", fontSize: 10, color: C.textMuted,
                letterSpacing: 2, textTransform: "uppercase",
                padding: "8px 6px 3px",
              }}>{faction}</div>
              {cards.map((c, i) => {
                const isSelected = selectedCard?.name === c.name;
                return (
                  <div
                    key={i}
                    onClick={() => setSelectedCard(c)}
                    style={{
                      padding: "6px 10px", borderRadius: 4, cursor: "pointer",
                      marginBottom: 2, transition: "all 0.15s",
                      background: isSelected ? `${C.accent}30` : "transparent",
                      borderLeft: isSelected ? `3px solid ${C.accent}` : "3px solid transparent",
                      animation: `slideIn 0.15s ease-out ${i * 0.03}s both`,
                    }}
                    onMouseEnter={e => { if (!isSelected) e.currentTarget.style.background = `${C.accent}15`; }}
                    onMouseLeave={e => { if (!isSelected) e.currentTarget.style.background = "transparent"; }}
                  >
                    <div style={{
                      fontSize: 13, fontWeight: isSelected ? 600 : 400,
                      color: isSelected ? C.accentLight : C.text,
                      fontFamily: "'Source Sans 3'",
                    }}>{c.name}</div>
                    <div style={{ display: "flex", gap: 6, alignItems: "center", marginTop: 1 }}>
                      <span style={{
                        fontSize: 9, padding: "0px 5px", borderRadius: 2,
                        background: `${typeColor(c.type)}40`,
                        color: `${typeColor(c.type)}`,
                        fontFamily: "'Oswald'", fontWeight: 500,
                        border: `1px solid ${typeColor(c.type)}60`,
                      }}>{c.type}</span>
                      {c.cost && c.cost !== "0" && (
                        <span style={{ fontSize: 10, color: C.textMuted }}>{c.cost}pts</span>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          ))}
        </div>

        {/* Load New */}
        <div style={{
          padding: "8px 12px", borderTop: `1px solid ${C.border}`,
        }}>
          <button
            onClick={() => { setData(null); setSelectedCard(null); setError(null); }}
            style={{
              width: "100%", padding: "6px", borderRadius: 4,
              border: `1px solid ${C.border}`, background: C.bg,
              color: C.textDim, fontSize: 11, cursor: "pointer",
              fontFamily: "'Oswald'", letterSpacing: 1,
              transition: "all 0.15s",
            }}
            onMouseEnter={e => { e.currentTarget.style.borderColor = C.accent; e.currentTarget.style.color = C.accent; }}
            onMouseLeave={e => { e.currentTarget.style.borderColor = C.border; e.currentTarget.style.color = C.textDim; }}
          >LOAD NEW FILE</button>
        </div>
      </div>

      {/* ─── Main Content ─── */}
      <div style={{
        flex: 1, overflowY: "auto", padding: 32,
        display: "flex", flexWrap: "wrap", gap: 24,
        alignContent: "flex-start", justifyContent: "center",
        background: `radial-gradient(ellipse at center, ${C.bg} 0%, #111 100%)`,
      }}>
        {selectedCard ? (
          <ModelCard card={selectedCard} data={data} />
        ) : (
          <div style={{ color: C.textMuted, fontStyle: "italic", marginTop: 80 }}>
            Select a model from the sidebar
          </div>
        )}
      </div>
    </div>
  );
}
