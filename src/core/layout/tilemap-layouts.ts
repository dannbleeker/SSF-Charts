/**
 * Tile-grid cartogram layouts: [col, row], 0-indexed, col = west→east,
 * row = north→south.
 *
 * US: the widely syndicated NYT-lineage 51-tile layout (50 states + DC) —
 * Alaska/Hawaii detached top/bottom-left, Maine top-right, Florida/Texas
 * south, so the silhouette still reads as the US.
 * EUROPE: the geofacet europe_countries_grid1 community layout (40
 * countries). "eu" is the same grid filtered to the 27 EU member states.
 * WORLD: 10 macro-regions preserving real-world relative bearing.
 */
export type TileLayout = Record<string, [number, number]>;

export const US_TILES: TileLayout = {
  AK: [0, 0],
  ME: [11, 0],
  VT: [9, 1],
  NH: [10, 1],
  MA: [11, 1],
  WA: [1, 2],
  MT: [2, 2],
  ND: [3, 2],
  SD: [4, 2],
  MN: [5, 2],
  WI: [6, 2],
  MI: [7, 2],
  NY: [9, 2],
  CT: [10, 2],
  RI: [11, 2],
  OR: [1, 3],
  ID: [2, 3],
  WY: [3, 3],
  NE: [4, 3],
  IA: [5, 3],
  IL: [6, 3],
  IN: [7, 3],
  OH: [8, 3],
  PA: [9, 3],
  NJ: [10, 3],
  CA: [0, 4],
  NV: [1, 4],
  UT: [2, 4],
  CO: [3, 4],
  KS: [4, 4],
  MO: [5, 4],
  KY: [6, 4],
  WV: [7, 4],
  DC: [8, 4],
  MD: [9, 4],
  DE: [10, 4],
  AZ: [2, 5],
  NM: [3, 5],
  OK: [4, 5],
  AR: [5, 5],
  TN: [6, 5],
  VA: [7, 5],
  NC: [8, 5],
  TX: [3, 6],
  LA: [4, 6],
  MS: [5, 6],
  AL: [6, 6],
  GA: [7, 6],
  SC: [8, 6],
  HI: [0, 7],
  FL: [7, 7],
};

export const EUROPE_TILES: TileLayout = {
  IS: [0, 0],
  NO: [4, 0],
  SE: [5, 0],
  FI: [6, 0],
  EE: [7, 1],
  IE: [0, 2],
  GB: [1, 2],
  NL: [3, 2],
  DK: [4, 2],
  LV: [7, 2],
  BE: [3, 3],
  DE: [4, 3],
  CZ: [5, 3],
  PL: [6, 3],
  LT: [7, 3],
  BY: [8, 3],
  RU: [9, 3],
  FR: [2, 4],
  LU: [3, 4],
  AT: [4, 4],
  SK: [5, 4],
  RS: [6, 4],
  RO: [7, 4],
  UA: [8, 4],
  PT: [0, 5],
  ES: [1, 5],
  CH: [3, 5],
  SI: [4, 5],
  HU: [5, 5],
  BA: [6, 5],
  BG: [7, 5],
  MD: [8, 5],
  IT: [3, 6],
  HR: [5, 6],
  ME: [6, 6],
  MK: [7, 6],
  TR: [8, 6],
  MT: [2, 7],
  AL: [6, 7],
  GR: [7, 7],
  CY: [9, 7],
};

/** The 27 EU member states (a filter of EUROPE_TILES). */
const EU_MEMBERS = new Set([
  "AT",
  "BE",
  "BG",
  "HR",
  "CY",
  "CZ",
  "DK",
  "EE",
  "FI",
  "FR",
  "DE",
  "GR",
  "HU",
  "IE",
  "IT",
  "LV",
  "LT",
  "LU",
  "MT",
  "NL",
  "PL",
  "PT",
  "RO",
  "SK",
  "SI",
  "ES",
  "SE",
]);

export const EU_TILES: TileLayout = Object.fromEntries(
  Object.entries(EUROPE_TILES).filter(([code]) => EU_MEMBERS.has(code)),
) as TileLayout;

export const WORLD_TILES: TileLayout = {
  NA: [1, 0],
  EU: [3, 0],
  CIS: [5, 0],
  MEA: [3, 1],
  SA: [4, 1],
  EA: [5, 1],
  LATAM: [1, 2],
  SSA: [3, 2],
  SEA: [5, 2],
  OCE: [5, 3],
};

export const TILE_LAYOUTS: Record<string, TileLayout> = {
  us: US_TILES,
  europe: EUROPE_TILES,
  eu: EU_TILES,
  world: WORLD_TILES,
};

/**
 * Auto-detect the layout from region codes when `map` is omitted: the layout
 * matching ≥90% of the provided codes, with a clear margin over the runner-up.
 * "world" is never auto-detected (its macro codes collide with ISO-2).
 */
export function detectLayout(codes: string[]): "us" | "eu" | "europe" | null {
  if (!codes.length) return null;
  const upper = codes.map((c) => c.trim().toUpperCase());
  const score = (layout: TileLayout) => upper.filter((c) => c in layout).length / upper.length;
  const us = score(US_TILES);
  const eu = score(EU_TILES);
  const europe = score(EUROPE_TILES);
  // The runner-up margin, across geographies only: the EU grid is a subset of
  // Europe, so an EU-only set matches both by construction and the tighter
  // frame is the intended pick. But codes like DE/MD/ME/MT are valid ISO-2 AND
  // US postal codes, and first-match-wins painted that European dataset onto
  // the US map. When both geographies clear the bar the codes do not name one —
  // fall through to the "set map: …" hint rather than guess.
  if (us >= 0.9 && Math.max(eu, europe) >= 0.9) return null;
  // Otherwise most-specific first.
  if (us >= 0.9) return "us";
  // The tighter EU frame is for an EU-ONLY set, i.e. one the EU grid matches as
  // well as Europe does. `layoutTilemap` keeps only the codes its grid has, so
  // choosing "eu" for ten codes of which one is GB drew nine of them and
  // dropped the tenth — tile, value and its share of the colour scale and
  // legend — with nothing on the chart to say so, while Europe matched all ten.
  // A code in neither grid (a typo) counts against both, so it cannot argue for
  // the wider one.
  if (eu >= 0.9 && eu >= europe) return "eu";
  if (europe >= 0.9) return "europe";
  return null;
}

/**
 * The grid a tile map draws: the pinned `map`, unless it matches NONE of the codes
 * and the codes name a layout of their own.
 *
 * A pin is a statement about the data, and a pin that none of the data agrees with
 * is a stale one. The pane has no control for `map` and its tile-map sample pins
 * `"europe"`, so typing US postal codes over the sample exported `map: "europe"`
 * and drew the Europe grid with no value on it, while the same codes with `map`
 * omitted drew the US grid. The data decides there, through `detectLayout`, which
 * needs 90% of the codes to agree — whichever grid was pinned, a world pin included.
 *
 * A pin that matches SOME of the codes stays: that is a chart with a few typos in it,
 * and the author's `map` is the better guess about which grid they meant. Nor does a
 * code set that names no layout move a pin (detection never answers "world", so a
 * world set is only ever drawn when it is pinned); there is nothing to move it to,
 * and a pin that matches nothing keeps its grid of grey tiles, as it always did.
 *
 * Own properties only, like every table keyed by a config string: `map: "constructor"`
 * is no layout, and used to reach `Object.prototype`.
 */
export function resolveTileLayout(pinned: string | undefined, codes: string[]): TileLayout | undefined {
  const named = (key: string | null | undefined): TileLayout | undefined =>
    key && Object.prototype.hasOwnProperty.call(TILE_LAYOUTS, key) ? TILE_LAYOUTS[key] : undefined;
  const pinnedLayout = named(pinned);
  if (pinnedLayout && codes.some((c) => c.trim().toUpperCase() in pinnedLayout)) return pinnedLayout;
  return named(detectLayout(codes)) ?? pinnedLayout;
}
