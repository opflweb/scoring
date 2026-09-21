"use strict";

const app = document.querySelector("#app");
const statusRegion = document.querySelector("#status");
const seasonSelect = document.querySelector("#season-select");
const navigation = document.querySelector("#primary-nav");
const menuToggle = document.querySelector("#menu-toggle");
const mainContent = document.querySelector("#main-content");

const VIEWS = new Set([
  "home", "matchups", "schedule", "standings", "rosters", "stats",
  "transactions", "drafts", "history", "rules", "banners",
]);
const CAPABILITY_BY_VIEW = {
  matchups: "matchups",
  schedule: "schedules",
  rosters: "rosters",
  stats: "player_stats",
  drafts: "drafts",
};
const VIEW_TITLES = {
  home: "Home",
  matchups: "Matchups",
  schedule: "Schedule",
  standings: "Standings",
  rosters: "Rosters",
  stats: "Stats",
  transactions: "Transactions",
  drafts: "Drafts",
  history: "League History",
  rules: "Rules",
  banners: "Banner Room",
};

let archiveIndex = null;
let renderVersion = 0;
const dataCache = new Map();

const escapeHtml = (value) => String(value ?? "").replace(/[&<>'"]/g, (character) => ({
  "&": "&amp;",
  "<": "&lt;",
  ">": "&gt;",
  "'": "&#39;",
  '"': "&quot;",
})[character]);

const formatNumber = (value, digits = 1) => {
  if (value === null || value === undefined || Number.isNaN(Number(value))) return "—";
  const number = Number(value);
  return number.toLocaleString(undefined, {
    minimumFractionDigits: Number.isInteger(number) ? 0 : digits,
    maximumFractionDigits: digits,
  });
};

const fetchJson = async (path) => {
  if (!dataCache.has(path)) {
    dataCache.set(path, fetch(path).then((response) => {
      if (!response.ok) throw new Error(`${response.status} while loading ${path}`);
      return response.json();
    }).catch((error) => {
      dataCache.delete(path);
      throw error;
    }));
  }
  return dataCache.get(path);
};

const seasonPath = (year, file) => `data/seasons/${year}/${file}`;
const sharedPath = (file) => `data/shared/${file}`;
const entryFor = (year) => archiveIndex.seasons.find((entry) => Number(entry.year) === Number(year));

function parseRoute() {
  const parts = window.location.hash.replace(/^#\/?/, "").split("/").filter(Boolean);
  const requestedYear = Number(parts[0]);
  const year = entryFor(requestedYear) ? requestedYear : archiveIndex.current_season;
  const view = VIEWS.has(parts[1]) ? parts[1] : "home";
  const weekIndex = parts.indexOf("week");
  const week = weekIndex >= 0 && Number.isInteger(Number(parts[weekIndex + 1]))
    ? Number(parts[weekIndex + 1])
    : null;
  return { year, view, week };
}

function routeHref(year, view, week = null) {
  return `#/${year}/${view}${week ? `/week/${week}` : ""}`;
}

function pageHeader(year, title, description, extra = "") {
  return `<header class="page-header">
    <div>
      <p class="eyebrow">${escapeHtml(year)} OPFL</p>
      <h1>${escapeHtml(title)}</h1>
      <p class="lede">${escapeHtml(description)}</p>
    </div>
    ${extra}
  </header>`;
}

function emptyState(title, message, link = "") {
  return `<section class="empty-state">
    <strong>${escapeHtml(title)}</strong>
    <p>${escapeHtml(message)}</p>
    ${link}
  </section>`;
}

function updateChrome(route, metadata) {
  seasonSelect.value = String(route.year);
  document.title = `${VIEW_TITLES[route.view]} · ${route.year} OPFL`;
  document.querySelectorAll("#primary-nav a").forEach((link) => {
    const view = link.dataset.view;
    const capability = CAPABILITY_BY_VIEW[view];
    link.href = routeHref(route.year, view);
    link.toggleAttribute("aria-current", view === route.view);
    if (view === route.view) link.setAttribute("aria-current", "page");
    const unavailable = capability && metadata.capabilities[capability] === false;
    if (unavailable) link.setAttribute("aria-disabled", "true");
    else link.removeAttribute("aria-disabled");
  });
  document.querySelector(".brand").href = routeHref(route.year, "home");
}

async function loadSeasonBase(year) {
  const [metadata, standingsDocument] = await Promise.all([
    fetchJson(seasonPath(year, "metadata.json")),
    fetchJson(seasonPath(year, "standings.json")),
  ]);
  return { metadata, standingsDocument, standings: standingsDocument.standings };
}

async function renderRoute() {
  const version = ++renderVersion;
  const route = parseRoute();
  statusRegion.textContent = `Loading ${route.year} ${VIEW_TITLES[route.view].toLowerCase()}…`;
  app.innerHTML = '<div class="grid two"><div class="card skeleton"></div><div class="card skeleton"></div></div>';
  navigation.classList.remove("open");
  menuToggle.setAttribute("aria-expanded", "false");
  try {
    const base = await loadSeasonBase(route.year);
    if (version !== renderVersion) return;
    updateChrome(route, base.metadata);
    const capability = CAPABILITY_BY_VIEW[route.view];
    if (capability && base.metadata.capabilities[capability] === false) {
      app.innerHTML = pageHeader(
        route.year,
        VIEW_TITLES[route.view],
        "This season is preserved at summary detail.",
        '<span class="badge">Summary archive</span>',
      ) + emptyState(
        `${VIEW_TITLES[route.view]} unavailable`,
        `The ${route.year} archive contains final standings and league finishes, but no reliable weekly, roster, or player-stat detail.`,
        `<a href="${routeHref(route.year, "standings")}">View recorded standings</a>`,
      );
      statusRegion.textContent = "Summary-season limitation shown.";
      return;
    }
    const renderers = {
      home: renderHome,
      matchups: renderMatchupsPage,
      schedule: renderSchedule,
      standings: renderStandings,
      rosters: renderRosters,
      stats: renderStats,
      transactions: renderTransactions,
      drafts: renderDrafts,
      history: renderHistory,
      rules: renderRules,
      banners: renderBanners,
    };
    const markup = await renderers[route.view](route, base);
    if (version !== renderVersion) return;
    app.innerHTML = markup;
    bindViewEvents(route, base);
    statusRegion.textContent = `${route.year} ${VIEW_TITLES[route.view]} loaded.`;
  } catch (error) {
    if (version !== renderVersion) return;
    console.error(error);
    app.innerHTML = `<section class="error-state">
      <strong>We couldn’t load this view.</strong>
      <p>${escapeHtml(error.message)}</p>
      <button id="retry-load" type="button">Try again</button>
    </section>`;
    statusRegion.textContent = "Data loading error.";
    document.querySelector("#retry-load")?.addEventListener("click", renderRoute);
  }
}

function renderStandingsTable(rows, advanced = true, limit = null) {
  const selected = limit ? rows.slice(0, limit) : rows;
  const historical = !advanced;
  return `<div class="table-wrap"><table>
    <caption class="sr-only">OPFL season standings</caption>
    <thead><tr>
      <th>Team</th><th>W-L-T</th>
      ${historical ? "<th>PF</th><th>Finish</th>" : "<th>Rank Pts</th><th>Top 6</th><th>PF</th><th>Avg</th><th>PA</th><th>Expected</th><th>Luck</th><th>SOS</th>"}
    </tr></thead>
    <tbody>${selected.map((row) => `<tr>
      <td><span class="rank">${escapeHtml(row.rank)}.</span> <span class="team-code">${escapeHtml(row.abbrev || row.name)}</span><br><small class="muted">${escapeHtml(row.name)}</small></td>
      <td>${formatNumber(row.wins, 1)}-${formatNumber(row.losses, 1)}-${formatNumber(row.ties, 1)}</td>
      ${historical
        ? `<td>${formatNumber(row.points_for)}</td><td>${escapeHtml(row.finish || "—")}</td>`
        : `<td><strong>${formatNumber(row.rank_points, 3)}</strong></td>
          <td>${formatNumber(row.top_six, 3)}</td><td>${formatNumber(row.points_for)}</td>
          <td>${formatNumber(row.average_points_for)}</td><td>${formatNumber(row.points_against)}</td>
          <td>${formatNumber(row.expected_wins)}-${formatNumber(row.expected_losses)}</td>
          <td class="${Number(row.luck) >= 0 ? "positive" : "negative"}">${Number(row.luck) > 0 ? "+" : ""}${formatNumber(row.luck)}</td>
          <td>${formatNumber(row.remaining_sos)}</td>`}
    </tr>`).join("")}</tbody>
  </table></div>`;
}

async function renderHome(route, base) {
  if (base.metadata.detail_level === "summary") {
    const finish = base.standingsDocument.finish;
    return pageHeader(route.year, "Season Archive", "Recorded standings and championship finish.", '<span class="badge">Summary</span>') +
      `<section class="grid four">
        <article class="card metric gold-line"><span>Champion</span><strong>${escapeHtml(finish.first)}</strong></article>
        <article class="card metric"><span>Runner-up</span><strong>${escapeHtml(finish.second)}</strong></article>
        <article class="card metric"><span>Oakland Bowl</span><strong>${escapeHtml(finish.championship_score || "—")}</strong></article>
        <article class="card metric"><span>League size</span><strong>${escapeHtml(finish.team_count)}</strong></article>
      </section>
      <section class="card flush" style="margin-top:1rem"><div class="card-header"><h2>Final standings</h2></div>${renderStandingsTable(base.standings, false)}</section>
      ${finish.comments ? `<section class="card" style="margin-top:1rem"><h2>Season note</h2><p class="muted">${escapeHtml(finish.comments)}</p></section>` : ""}`;
  }
  const weeks = base.metadata.weeks_available;
  const latestWeek = Math.max(...weeks);
  const [latest, previous, transactions, playoffs] = await Promise.all([
    fetchJson(seasonPath(route.year, `weeks/week_${latestWeek}.json`)),
    latestWeek > 1 ? fetchJson(seasonPath(route.year, `weeks/week_${latestWeek - 1}.json`)) : Promise.resolve(null),
    fetchJson(sharedPath("transactions.json")),
    fetchJson(seasonPath(route.year, "playoffs.json")).catch(() => null),
  ]);
  const leader = base.standings[0];
  const recentTrades = transactions.trades.filter((trade) => Number(trade.date.slice(0, 4)) <= route.year).slice(0, 3);
  const jamboreeLeader = playoffs?.jamboree?.standings?.[0];
  return pageHeader(route.year, "League Dashboard", "Scores, playoff results, standings, and recent league activity.", '<span class="badge gold">Full season</span>') +
    `<section class="grid four">
      <article class="card metric gold-line"><span>Rank leader</span><strong>${escapeHtml(leader.abbrev)}</strong><small>${formatNumber(leader.rank_points, 3)} Rank Pts</small></article>
      <article class="card metric"><span>Points leader</span><strong>${escapeHtml([...base.standings].sort((a, b) => b.points_for - a.points_for)[0].abbrev)}</strong><small>${formatNumber(Math.max(...base.standings.map((row) => row.points_for)))} PF</small></article>
      <article class="card metric green-line"><span>Oakland Bowl</span><strong>${escapeHtml(playoffs?.championship?.winner || "TBD")}</strong><small>Weeks 16–17</small></article>
      <article class="card metric red-line"><span>Jamboree</span><strong>${escapeHtml(jamboreeLeader?.abbrev || "TBD")}</strong><small>${formatNumber(jamboreeLeader?.total)} points</small></article>
    </section>
    <section class="grid two" style="margin-top:1rem">
      <article class="card"><div class="card-header"><h2>Week ${escapeHtml(latestWeek)}</h2><a href="${routeHref(route.year, "matchups", latestWeek)}">Full detail</a></div>${renderCompactMatchups(latest)}</article>
      <article class="card"><div class="card-header"><h2>${previous ? `Week ${latestWeek - 1}` : "Standings"}</h2>${previous ? `<a href="${routeHref(route.year, "matchups", latestWeek - 1)}">Review</a>` : ""}</div>${previous ? renderCompactMatchups(previous) : renderStandingsTable(base.standings, true, 5)}</article>
    </section>
    <section class="grid two" style="margin-top:1rem">
      <article class="card flush"><div class="card-header"><h2>Top five</h2><a href="${routeHref(route.year, "standings")}">Advanced standings</a></div>${renderStandingsTable(base.standings, true, 5)}</article>
      <article class="card"><div class="card-header"><h2>Recent trades</h2><a href="${routeHref(route.year, "transactions")}">Search all</a></div>${renderActivityItems(recentTrades)}</article>
    </section>`;
}

function normalizeMatchup(matchup) {
  if (matchup.away) return { team1: matchup.away, team2: matchup.home, score1: matchup.away_score, score2: matchup.home_score };
  return { team1: matchup.team1 || matchup.higher_seed, team2: matchup.team2 || matchup.lower_seed, score1: matchup.score1 ?? matchup.higher_score, score2: matchup.score2 ?? matchup.lower_score };
}

function renderCompactMatchups(weekData) {
  if (!weekData.matchups.length) return '<p class="muted">No head-to-head matchups were recorded for this week.</p>';
  return `<div class="matchup-list">${weekData.matchups.map((raw) => {
    const game = normalizeMatchup(raw);
    return `<div class="schedule-game"><span>${escapeHtml(game.team1)} <strong>${formatNumber(game.score1)}</strong></span><span>${escapeHtml(game.team2)} <strong>${formatNumber(game.score2)}</strong></span></div>`;
  }).join("")}</div>`;
}

function renderTeamLineup(team) {
  if (!team) return '<p class="muted">Roster detail unavailable.</p>';
  const ordered = [...team.roster].sort((a, b) => Number(b.starter) - Number(a.starter) || a.position.localeCompare(b.position));
  return `<h3>${escapeHtml(team.abbrev)} lineup</h3>${ordered.map((player) => {
    const breakdown = Object.entries(player.breakdown || {}).filter(([key]) => key !== "floor_applied").map(([key, value]) => `${key.replaceAll("_", " ")}: ${value}`).join("; ");
    return `<div class="lineup-row ${player.starter ? "" : "bench"}" title="${escapeHtml(breakdown || "No scoring breakdown in archive")}">
      <span class="position">${escapeHtml(player.position)}</span><span>${escapeHtml(player.name)} <small class="muted">${escapeHtml(player.nfl_team)}</small>${player.starter ? ' <span class="starter">●</span>' : ""}</span><strong>${formatNumber(player.score)}</strong>
    </div>`;
  }).join("")}`;
}

function renderDetailedMatchups(weekData) {
  const teams = Object.fromEntries(weekData.teams.map((team) => [team.abbrev, team]));
  if (!weekData.matchups.length) return emptyState("No head-to-head bracket", "This week contains individual team scores only.");
  return `<div class="matchup-list">${weekData.matchups.map((raw, index) => {
    const game = normalizeMatchup(raw);
    const winner1 = game.score1 !== null && Number(game.score1) > Number(game.score2);
    const winner2 = game.score2 !== null && Number(game.score2) > Number(game.score1);
    return `<details class="matchup" ${index === 0 ? "open" : ""}>
      <summary>
        <span class="matchup-team ${winner1 ? "winner" : ""}"><strong>${escapeHtml(game.team1)}</strong><span class="score">${formatNumber(game.score1)}</span></span>
        <span class="versus">VS</span>
        <span class="matchup-team ${winner2 ? "winner" : ""}"><strong>${escapeHtml(game.team2)}</strong><span class="score">${formatNumber(game.score2)}</span></span>
      </summary>
      <div class="matchup-detail"><div>${renderTeamLineup(teams[game.team1])}</div><div>${renderTeamLineup(teams[game.team2])}</div></div>
    </details>`;
  }).join("")}</div>`;
}

async function renderMatchupsPage(route, base) {
  const weeks = base.metadata.weeks_available;
  const week = weeks.includes(route.week) ? route.week : Math.max(...weeks);
  const weekData = await fetchJson(seasonPath(route.year, `weeks/week_${week}.json`));
  const options = weeks.map((number) => `<option value="${number}" ${number === week ? "selected" : ""}>Week ${number}${number > 15 ? " · Playoffs" : ""}</option>`).join("");
  return pageHeader(route.year, `Week ${week} Matchups`, "Expand a matchup for starter, bench, and scoring detail.", `<span class="badge">${escapeHtml(weekData.status)}</span>`) +
    `<div class="toolbar"><label class="field">Week<select id="matchup-week">${options}</select></label></div>${renderDetailedMatchups(weekData)}`;
}

async function renderSchedule(route) {
  const schedule = await fetchJson(seasonPath(route.year, "schedule.json"));
  const weekDocuments = await Promise.all(schedule.weeks.map((item) => fetchJson(seasonPath(route.year, `weeks/week_${item.week}.json`)).catch(() => null)));
  const scoreByWeek = Object.fromEntries(weekDocuments.filter(Boolean).map((week) => [week.week, Object.fromEntries(week.teams.map((team) => [team.abbrev, team.total_score]))]));
  const teams = [...new Set(schedule.weeks.flatMap((week) => week.matchups.flatMap((game) => [game.away, game.home])))].sort();
  return pageHeader(route.year, "Season Schedule", "Fifteen regular-season weeks; filter the grid to follow one team.") +
    `<div class="toolbar"><label class="field">Team<select id="schedule-team"><option value="">All teams</option>${teams.map((team) => `<option>${escapeHtml(team)}</option>`).join("")}</select></label></div>
    <div id="schedule-grid" class="schedule-grid">${renderScheduleGrid(schedule.weeks, scoreByWeek, "")}</div>`;
}

function renderScheduleGrid(weeks, scores, selected) {
  return weeks.map((week) => `<article class="schedule-week"><h3><span>Week ${week.week}</span><a href="${routeHref(parseRoute().year, "matchups", week.week)}">Details</a></h3>
    ${week.matchups.map((game) => {
      const active = selected && [game.away, game.home].includes(selected);
      return `<div class="schedule-game ${active ? "selected" : ""}"><span>${escapeHtml(game.away)} ${scores[week.week] ? `<strong>${formatNumber(scores[week.week][game.away])}</strong>` : ""}</span><span>${escapeHtml(game.home)} ${scores[week.week] ? `<strong>${formatNumber(scores[week.week][game.home])}</strong>` : ""}</span></div>`;
    }).join("")}</article>`).join("");
}

async function renderStandings(route, base) {
  const advanced = base.metadata.capabilities.advanced_standings;
  const finish = base.standingsDocument.finish;
  const historyRows = advanced ? base.standings.map((row) => `<tr><td><span class="team-code">${escapeHtml(row.abbrev)}</span></td><td class="text-left">${(row.weekly_rank_history || []).map((item) => `W${item.week}:${item.rank}`).join(" · ")}</td></tr>`).join("") : "";
  return pageHeader(route.year, "Standings", advanced ? "Rank Points, all-play expectation, luck, schedule strength, and weekly rank history." : "Recorded regular-season results; advanced weekly inputs were not archived.", `<span class="badge">${advanced ? "Advanced" : "Summary"}</span>`) +
    `<section class="card flush">${renderStandingsTable(base.standings, advanced)}</section>
    ${advanced ? `<section class="card flush" style="margin-top:1rem"><div class="card-header"><h2>Weekly rank history</h2></div><div class="table-wrap"><table><thead><tr><th>Team</th><th class="text-left">Weekly ranks</th></tr></thead><tbody>${historyRows}</tbody></table></div></section>` :
      `<section class="card" style="margin-top:1rem"><h2>League finish</h2><p><strong>${escapeHtml(finish.first)}</strong> defeated ${escapeHtml(finish.second)} · ${escapeHtml(finish.championship_score)}</p><p class="muted">${escapeHtml(finish.comments)}</p></section>`}`;
}

async function renderRosters(route, base) {
  const [rosters, futurePicks, transactions] = await Promise.all([
    fetchJson(seasonPath(route.year, "rosters.json")),
    fetchJson(sharedPath("future-picks.json")),
    fetchJson(sharedPath("transactions.json")),
  ]);
  const teams = Object.keys(rosters.teams);
  const week = Math.max(...base.metadata.weeks_available);
  return pageHeader(route.year, "Roster Browser", "Search all rosters, inspect weekly lineups and taxi slots, compare teams, and filter related activity.") +
    `<div class="toolbar">
      <label class="field">Roster<select id="roster-team"><option value="ALL">All teams</option>${teams.map((team) => `<option value="${escapeHtml(team)}">${escapeHtml(team)}</option>`).join("")}</select></label>
      <label class="field">Snapshot<select id="roster-week"><option value="current">Current roster</option>${base.metadata.weeks_available.map((number) => `<option value="${number}" ${number === week ? "selected" : ""}>Week ${number}</option>`).join("")}</select></label>
      <label class="field grow">Player search<input id="roster-search" type="search" placeholder="Name, NFL team, or position"></label>
      <label class="field">Compare with<select id="compare-team"><option value="">None</option>${teams.map((team) => `<option value="${escapeHtml(team)}">${escapeHtml(team)}</option>`).join("")}</select></label>
    </div>
    <div id="roster-results"></div>`;
}

function rosterCards(entries, query) {
  const needle = query.trim().toLowerCase();
  const players = entries.flatMap(([team, roster]) => roster.map((player) => ({ ...player, fantasy_team: team }))).filter((player) => !player.placeholder && (!needle || `${player.name} ${player.nfl_team} ${player.position} ${player.fantasy_team}`.toLowerCase().includes(needle)));
  if (!players.length) return emptyState("No roster matches", "Try a broader player or team search.");
  return `<div class="roster-grid">${players.map((player) => `<article class="player-card"><header><strong><span class="position">${escapeHtml(player.position)}</span>${escapeHtml(player.name)}</strong>${player.starter ? '<span class="starter">Starter</span>' : ""}</header><p>${escapeHtml(player.fantasy_team)} · ${escapeHtml(player.nfl_team || "No NFL team")}${player.score !== undefined ? ` · ${formatNumber(player.score)} pts` : ""}</p></article>`).join("")}</div>`;
}

function activityMatchesTeam(trade, team, teamName = "") {
  if (!team || team === "ALL") return true;
  const text = `${trade.team_a} ${trade.team_b}`.toLowerCase();
  return text.includes(team.toLowerCase()) || (teamName && text.includes(teamName.toLowerCase()));
}

function renderActivityItems(trades) {
  if (!trades.length) return '<p class="muted">No matching trades.</p>';
  return `<div class="activity-list">${trades.map((trade) => `<article class="activity-item"><header><strong>${escapeHtml(trade.team_a)} ↔ ${escapeHtml(trade.team_b)}</strong><time>${escapeHtml(trade.date)}</time></header><div class="grid two" style="margin-top:.55rem"><div><small>${escapeHtml(trade.team_a)} receives</small><ul class="asset-list">${trade.team_a_receives.map((asset) => `<li>${escapeHtml(asset)}</li>`).join("")}</ul></div><div><small>${escapeHtml(trade.team_b)} receives</small><ul class="asset-list">${trade.team_b_receives.map((asset) => `<li>${escapeHtml(asset)}</li>`).join("")}</ul></div></div>${trade.notes ? `<p class="muted">${escapeHtml(trade.notes)}</p>` : ""}</article>`).join("")}</div>`;
}

async function updateRosterResults(route, base) {
  const container = document.querySelector("#roster-results");
  if (!container) return;
  const team = document.querySelector("#roster-team").value;
  const compare = document.querySelector("#compare-team").value;
  const snapshot = document.querySelector("#roster-week").value;
  const query = document.querySelector("#roster-search").value;
  const [current, futurePicks, transactions] = await Promise.all([
    fetchJson(seasonPath(route.year, "rosters.json")),
    fetchJson(sharedPath("future-picks.json")),
    fetchJson(sharedPath("transactions.json")),
  ]);
  let entries;
  if (snapshot === "current") {
    entries = Object.entries(current.teams).map(([code, value]) => [code, value.players]);
  } else {
    const week = await fetchJson(seasonPath(route.year, `weeks/week_${snapshot}.json`));
    entries = week.teams.map((value) => [value.abbrev, value.roster]);
  }
  const selected = team === "ALL" ? entries : entries.filter(([code]) => code === team);
  const compared = compare ? entries.filter(([code]) => code === compare) : [];
  const selectedName = team !== "ALL" ? current.teams[team]?.name : "";
  const taxi = team !== "ALL" ? current.teams[team]?.taxi || [] : [];
  const ownerTerms = [team, selectedName].filter(Boolean).map((value) => value.toLowerCase());
  const pickNotes = futurePicks.notes.filter((note) => team === "ALL" || ownerTerms.some((term) => note.toLowerCase().includes(term)));
  const activity = transactions.trades.filter((trade) => activityMatchesTeam(trade, team, selectedName)).slice(0, 8);
  container.innerHTML = `${compare ? `<section class="compare"><article class="card"><h2>${escapeHtml(team)} roster</h2>${rosterCards(selected, query)}</article><article class="card"><h2>${escapeHtml(compare)} roster</h2>${rosterCards(compared, query)}</article></section>` : `<section class="card"><h2>${team === "ALL" ? "All rostered players" : `${escapeHtml(team)} roster`}</h2>${rosterCards(selected, query)}</section>`}
    <section class="grid two" style="margin-top:1rem">
      <article class="card"><h2>Taxi & future picks</h2>${taxi.length ? rosterCards([[team, taxi]], query) : '<p class="muted">No taxi players recorded in this snapshot.</p>'}<ul class="asset-list">${pickNotes.slice(0, 12).map((note) => `<li>${escapeHtml(note)}</li>`).join("") || "<li>No matching future-pick notes.</li>"}</ul></article>
      <article class="card"><h2>Filtered activity</h2>${renderActivityItems(activity)}</article>
    </section>`;
}

async function renderStats(route) {
  const [players, teams] = await Promise.all([
    fetchJson(seasonPath(route.year, "player-stats.json")),
    fetchJson(seasonPath(route.year, "team-stats.json")),
  ]);
  const positions = [...new Set(players.players.map((player) => player.position))];
  return pageHeader(route.year, "Player & Team Stats", "Player leaders and team scoring profiles from all recorded weekly scores.") +
    `<div class="toolbar"><label class="field">Position<select id="stats-position"><option value="">All</option>${positions.map((position) => `<option>${escapeHtml(position)}</option>`).join("")}</select></label></div>
    <section class="grid two"><article class="card flush"><div class="card-header"><h2>Player leaders</h2><span class="badge">Bench included</span></div><div id="player-leaders">${renderPlayerLeaders(players.players, "")}</div></article>
    <article class="card flush"><div class="card-header"><h2>Team stats</h2></div>${renderTeamStats(teams.teams)}</article></section>`;
}

function renderPlayerLeaders(players, position) {
  const filtered = players.filter((player) => !position || player.position === position).slice(0, 60);
  return `<div class="table-wrap"><table><thead><tr><th>Player</th><th>Pos</th><th>Team</th><th>Pts</th><th>PPG</th><th>High</th><th>Starts</th></tr></thead><tbody>${filtered.map((player) => `<tr><td>${escapeHtml(player.name)}<br><small class="muted">${escapeHtml(player.nfl_team)}</small></td><td>${escapeHtml(player.position)}</td><td>${escapeHtml(player.fantasy_teams.join(" → "))}</td><td><strong>${formatNumber(player.points)}</strong></td><td>${formatNumber(player.ppg)}</td><td>${formatNumber(player.high)}</td><td>${escapeHtml(player.starts)}</td></tr>`).join("")}</tbody></table></div>`;
}

function renderTeamStats(teams) {
  return `<div class="table-wrap"><table><thead><tr><th>Team</th><th>Record</th><th>PPG</th><th>PF</th><th>PA</th><th>Diff</th><th>High</th><th>Low</th><th>Median</th></tr></thead><tbody>${teams.map((team) => `<tr><td><span class="team-code">${escapeHtml(team.abbrev)}</span></td><td>${team.wins}-${team.losses}-${team.ties}</td><td>${formatNumber(team.average_points_for)}</td><td>${formatNumber(team.points_for)}</td><td>${formatNumber(team.points_against)}</td><td class="${Number(team.point_differential) >= 0 ? "positive" : "negative"}">${formatNumber(team.point_differential)}</td><td>${formatNumber(team.high_score)}</td><td>${formatNumber(team.low_score)}</td><td>${formatNumber(team.median_score)}</td></tr>`).join("")}</tbody></table></div>`;
}

async function renderTransactions(route) {
  const data = await fetchJson(sharedPath("transactions.json"));
  const teams = [...new Set(data.trades.flatMap((trade) => trade.teams))].sort();
  const yearTrades = data.trades.filter((trade) => Number(trade.date.slice(0, 4)) === route.year);
  return pageHeader(route.year, "Trade History", "Search recorded 2024–2025 transactions by team, asset, pick, or note.") +
    `<div class="toolbar"><label class="field">Team<select id="trade-team"><option value="">All teams</option>${teams.map((team) => `<option>${escapeHtml(team)}</option>`).join("")}</select></label><label class="field grow">Search<input id="trade-search" type="search" placeholder="Player, pick, or note"></label></div>
    <div id="transaction-results">${renderActivityItems(yearTrades)}</div>`;
}

async function renderDrafts(route) {
  const data = await fetchJson(sharedPath("drafts.json"));
  const boards = data.drafts;
  const years = [...new Set(boards.map((board) => board.year))].sort((a, b) => b - a);
  const initialYear = years.includes(route.year) ? route.year : years[0];
  const board = boards.find((item) => item.year === initialYear && item.type === "preseason") || boards.find((item) => item.year === initialYear);
  return pageHeader(route.year, "Draft Archives", "Preseason and waiver boards, including taxi rounds, picks, passes, and recorded drops.") +
    `<div class="toolbar"><label class="field">Draft year<select id="draft-year">${years.map((year) => `<option value="${year}" ${year === initialYear ? "selected" : ""}>${year}</option>`).join("")}</select></label><label class="field">Board<select id="draft-type"><option value="preseason" ${board.type === "preseason" ? "selected" : ""}>Preseason</option><option value="waiver" ${board.type === "waiver" ? "selected" : ""}>Waiver</option></select></label><label class="field grow">Filter<input id="draft-search" type="search" placeholder="Owner, selection, or drop"></label></div>
    <div id="draft-board">${renderDraftBoard(board, "")}</div>`;
}

function renderDraftBoard(board, query) {
  if (!board) return emptyState("Board unavailable", "No draft board matches these filters.");
  const needle = query.toLowerCase();
  const picks = board.picks.filter((pick) => !needle || `${pick.owner} ${pick.selection} ${pick.drop}`.toLowerCase().includes(needle));
  return `<section class="card flush"><div class="card-header"><h2>${escapeHtml(board.title)}</h2><span class="badge">${escapeHtml(board.date)}</span></div>${picks.map((pick) => `<div class="draft-pick ${pick.taxi ? "taxi" : ""}"><span>${pick.taxi ? `T${pick.round}.${pick.slot}` : `#${pick.overall}`}</span><strong>${escapeHtml(pick.owner)}</strong><span>${escapeHtml(pick.selection || "—")}</span><span class="muted">Drop: ${escapeHtml(pick.drop || "—")}</span></div>`).join("") || '<p class="muted" style="padding:1rem">No picks match.</p>'}</section>`;
}

async function renderHistory(route) {
  const data = await fetchJson(sharedPath("history.json"));
  const seasons = [...data.seasons].sort((a, b) => b.year - a.year);
  return pageHeader(route.year, "League History", "Oakland Bowl finishes, Jamboree winners, season notes, and historical owner summaries since 1988.") +
    `<div class="toolbar"><label class="field grow">Find a season or owner<input id="history-search" type="search" placeholder="Year, champion, owner, or note"></label></div>
    <section class="grid two"><div id="history-results" class="history-list">${renderHistoryItems(seasons, "")}</div><article class="card flush"><div class="card-header"><h2>Owner summaries</h2></div><div class="table-wrap"><table><thead><tr><th>Owner label</th><th>Seasons</th><th>Titles</th><th>W-L-T</th><th>PF</th></tr></thead><tbody>${data.owners.map((owner) => `<tr><td>${escapeHtml(owner.owner)}</td><td>${escapeHtml(owner.seasons)}</td><td>${formatNumber(owner.titles)}</td><td>${formatNumber(owner.wins, 1)}-${formatNumber(owner.losses, 1)}-${formatNumber(owner.ties, 1)}</td><td>${formatNumber(owner.points_for)}</td></tr>`).join("")}</tbody></table></div></article></section>`;
}

function renderHistoryItems(seasons, query) {
  const needle = query.toLowerCase();
  const filtered = seasons.filter((season) => !needle || JSON.stringify(season).toLowerCase().includes(needle));
  return filtered.map((season) => `<article class="history-item"><header><strong>${escapeHtml(season.year)} · ${escapeHtml(season.first)}</strong><span class="badge gold">Champion</span></header><p>${escapeHtml(season.first)} over ${escapeHtml(season.second)} · ${escapeHtml(season.championship_score)}</p><p class="muted">3rd ${escapeHtml(season.third)} · 4th ${escapeHtml(season.fourth)}${season.jamboree ? ` · Jamboree ${escapeHtml(season.jamboree)}` : ""}</p>${season.comments ? `<small class="muted">${escapeHtml(season.comments)}</small>` : ""}</article>`).join("") || emptyState("No history matches", "Try a year, owner, or broader phrase.");
}

async function renderRules(route) {
  const data = await fetchJson(sharedPath("rules.json"));
  return pageHeader(route.year, data.title, "The preserved OPFL scoring and league-format rules used by the JSON scorer.") + data.sections.map((section, index) => `<details class="rule-section" ${index === 0 ? "open" : ""}><summary>${escapeHtml(section.title)}</summary><ul>${section.rules.map((rule) => `<li>${escapeHtml(rule)}</li>`).join("")}</ul></details>`).join("");
}

async function renderBanners(route) {
  const data = await fetchJson(sharedPath("banners.json"));
  const safePath = (value) => /^images\/banners\/\d{4}\.png$/.test(value) ? value : "";
  return pageHeader(route.year, "Banner Room", "Every preserved OPFL championship banner from 1988 through 2024.", `<span class="badge gold">${data.banners.length} banners</span>`) +
    `<div class="banner-grid">${data.banners.map((banner) => `<figure class="banner-card"><img src="${safePath(banner.image)}" alt="${escapeHtml(`${banner.year} OPFL championship banner`)}" loading="lazy" width="400" height="300"><figcaption><strong>${escapeHtml(banner.year)}</strong><br>${escapeHtml(banner.champion)}</figcaption></figure>`).join("")}</div>`;
}

function bindViewEvents(route, base) {
  document.querySelector("#matchup-week")?.addEventListener("change", (event) => { window.location.hash = routeHref(route.year, "matchups", Number(event.target.value)); });
  document.querySelector("#schedule-team")?.addEventListener("change", async (event) => {
    const schedule = await fetchJson(seasonPath(route.year, "schedule.json"));
    const weekDocuments = await Promise.all(schedule.weeks.map((item) => fetchJson(seasonPath(route.year, `weeks/week_${item.week}.json`)).catch(() => null)));
    const scores = Object.fromEntries(weekDocuments.filter(Boolean).map((week) => [week.week, Object.fromEntries(week.teams.map((team) => [team.abbrev, team.total_score]))]));
    document.querySelector("#schedule-grid").innerHTML = renderScheduleGrid(schedule.weeks, scores, event.target.value);
  });
  ["#roster-team", "#roster-week", "#compare-team"].forEach((selector) => document.querySelector(selector)?.addEventListener("change", () => updateRosterResults(route, base)));
  document.querySelector("#roster-search")?.addEventListener("input", () => updateRosterResults(route, base));
  if (document.querySelector("#roster-results")) updateRosterResults(route, base);
  document.querySelector("#stats-position")?.addEventListener("change", async (event) => {
    const players = await fetchJson(seasonPath(route.year, "player-stats.json"));
    document.querySelector("#player-leaders").innerHTML = renderPlayerLeaders(players.players, event.target.value);
  });
  const updateTransactions = async () => {
    const data = await fetchJson(sharedPath("transactions.json"));
    const team = document.querySelector("#trade-team")?.value || "";
    const needle = (document.querySelector("#trade-search")?.value || "").toLowerCase();
    const trades = data.trades.filter((trade) => Number(trade.date.slice(0, 4)) === route.year && activityMatchesTeam(trade, team) && (!needle || JSON.stringify(trade).toLowerCase().includes(needle)));
    document.querySelector("#transaction-results").innerHTML = renderActivityItems(trades);
  };
  document.querySelector("#trade-team")?.addEventListener("change", updateTransactions);
  document.querySelector("#trade-search")?.addEventListener("input", updateTransactions);
  const updateDrafts = async () => {
    const data = await fetchJson(sharedPath("drafts.json"));
    const year = Number(document.querySelector("#draft-year").value);
    const type = document.querySelector("#draft-type").value;
    const query = document.querySelector("#draft-search").value;
    document.querySelector("#draft-board").innerHTML = renderDraftBoard(data.drafts.find((board) => board.year === year && board.type === type), query);
  };
  ["#draft-year", "#draft-type"].forEach((selector) => document.querySelector(selector)?.addEventListener("change", updateDrafts));
  document.querySelector("#draft-search")?.addEventListener("input", updateDrafts);
  document.querySelector("#history-search")?.addEventListener("input", async (event) => {
    const data = await fetchJson(sharedPath("history.json"));
    document.querySelector("#history-results").innerHTML = renderHistoryItems([...data.seasons].sort((a, b) => b.year - a.year), event.target.value);
  });
}

async function initialize() {
  menuToggle.addEventListener("click", () => {
    const open = navigation.classList.toggle("open");
    menuToggle.setAttribute("aria-expanded", String(open));
  });
  navigation.addEventListener("click", (event) => {
    const link = event.target.closest("a");
    if (link?.getAttribute("aria-disabled") === "true") {
      event.preventDefault();
      statusRegion.textContent = "That view is unavailable for this summary season.";
    }
  });
  seasonSelect.addEventListener("change", () => {
    const route = parseRoute();
    window.location.hash = routeHref(Number(seasonSelect.value), route.view, route.week);
  });
  window.addEventListener("hashchange", renderRoute);
  try {
    archiveIndex = await fetchJson("data/index.json");
    seasonSelect.innerHTML = [...archiveIndex.seasons].sort((a, b) => b.year - a.year).map((entry) => `<option value="${entry.year}">${entry.year} · ${entry.detail_level === "full" ? "Full" : "Summary"}</option>`).join("");
    if (!window.location.hash) history.replaceState(null, "", routeHref(archiveIndex.current_season, "home"));
    await renderRoute();
  } catch (error) {
    console.error(error);
    statusRegion.textContent = "Archive index unavailable.";
    app.innerHTML = `<section class="error-state"><strong>The OPFL archive could not start.</strong><p>${escapeHtml(error.message)}</p></section>`;
  }
}

initialize();
