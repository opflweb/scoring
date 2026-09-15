        let data = null;
        let currentWeek = 1;
        let currentTeam = null;

        // Team numbers used by the printed OPFL schedule. Overridden by
        // data.team_number_map when the export provides one.
        const DEFAULT_TEAM_NUMBER_MAP = {
            1: 'AND', 2: 'JOH', 3: 'W/B', 4: 'STL', 5: 'G/G', 6: 'J/M',
            7: 'KEV', 8: 'KAM', 9: 'K/D', 10: 'E/J', 11: 'ADA', 12: 'D/J'
        };

        function teamNumberMap() {
            return data?.team_number_map || DEFAULT_TEAM_NUMBER_MAP;
        }

        // Weekly pairings come straight from the workbook (data.schedule maps
        // a week number to a list of [abbrev, abbrev] pairs). Weeks that have
        // not been published yet simply return null.
        function getWeekMatchups(weekNum) {
            const schedule = data?.schedule || {};
            const pairs = schedule[String(weekNum)] || schedule[weekNum];
            if (!pairs || !pairs.length) return null;

            const numberMap = teamNumberMap();
            return pairs.map(pair => pair.map(entry => numberMap[entry] || numberMap[String(entry)] || entry));
        }

        // Historical championship data
        const CHAMPIONSHIPS = [
            { year: 2024, first: "Kemp/A/M", second: "Wes", third: "John", fourth: "Jarrett/Matt", score: "88-57", jamboree: "Steve L. (157)" },
            { year: 2023, first: "Kemp/A/M", second: "Greg/Griffin", third: "Jarrett/Matt", fourth: "Kevin", score: "43-31", jamboree: "John (120)" },
            { year: 2022, first: "Jarrett/Matt & Wes", second: "Jarrett/Matt & Wes", third: "John & Kemp/A/M", fourth: "John & Kemp/A/M", score: "53-45", jamboree: "Kirk/David (122)", comment: "Historic tie declared" },
            { year: 2021, first: "Kemp/A/M", second: "Greg/Griffin", third: "Kevin", fourth: "Eric/Jeff", score: "41-40", jamboree: "Wes/Bill (149)" },
            { year: 2020, first: "Kemp/A/M", second: "Jarrett/Matt", third: "Kirk/David", fourth: "Steve M.", score: "67-57", jamboree: "Kevin (125)" },
            { year: 2019, first: "Andrew", second: "Kirk/David", third: "Jarrett/Matt", fourth: "Steve L.", score: "76-61", jamboree: "Adam (141)" },
            { year: 2018, first: "Adam", second: "Jarrett/Matt", third: "Steve L.", fourth: "Kirk/David", score: "63-35", jamboree: "Steve M. (117)" },
            { year: 2017, first: "Kirk/David", second: "Steve L./Wes", third: "Greg/Griffin", fourth: "Andrew", score: "52-51", jamboree: "Eric/Jeff (115)" },
            { year: 2016, first: "Kemp/A/M", second: "Steve L.", third: "Steve/Wes", fourth: "Eric/Jeff", score: "61-33", jamboree: "Bill/Jill (115)" },
            { year: 2015, first: "Steve L.", second: "Adam", third: "Greg/Griffin", fourth: "Kirk/David", score: "65-32", jamboree: "Jarrett (126)" },
            { year: 2014, first: "Steve L.", second: "Kirk/David", third: "Steve/Wes", fourth: "John", score: "57-31", jamboree: "Greg/Griffin (111)" },
            { year: 2013, first: "Jarrett", second: "Steve L.", third: "Adam", fourth: "Steve/Wes", score: "48-42", jamboree: "Greg (124)" },
            { year: 2012, first: "Kevin", second: "Kirk/David", third: "John", fourth: "Steve/Wes", score: "67-62", jamboree: "Adam (147)" },
            { year: 2011, first: "Bill/Jill", second: "Kirk/David", third: "Greg", fourth: "Steve L.", score: "89-48", jamboree: "Kevin (124)" },
            { year: 2010, first: "Kirk/David", second: "John", third: "Steve L.", fourth: "Kevin", score: "62-43", jamboree: "Jarrett (145)" },
            { year: 2009, first: "Steve L.", second: "Bill", third: "Greg", fourth: "John", score: "57-57*", jamboree: "Jarrett (129)", comment: "Won on better record" },
            { year: 2008, first: "Steve L.", second: "Greg", third: "John", fourth: "Jarrett", score: "67-28", jamboree: "Eric S. (109)" },
            { year: 2007, first: "Eric S.", second: "Kevin", third: "Mike", fourth: "Kirk", score: "55-23", jamboree: "Jarrett (121)" },
            { year: 2006, first: "Kirk", second: "John", third: "Kevin", fourth: "Jarrett", score: "52-52*", jamboree: "Kemp (147)", comment: "Won on better record" },
            { year: 2005, first: "Jarrett", second: "Eric/Bret/Joe", third: "Kirk", fourth: "Steve/Wes", score: "48-34", jamboree: "Eric S. (132)" },
            { year: 2004, first: "Kreg", second: "John", third: "Adam/Nick/Bill", fourth: "Steve L.", score: "69-57", jamboree: "Eric H. (165)" },
            { year: 2003, first: "Kirk", second: "Steve/Wes", third: "John", fourth: "Steve L.", score: "55-46", jamboree: "Jarrett (139)" },
            { year: 2002, first: "Adam/Nick", second: "Greg", third: "Eric S.", fourth: "John", score: "28-28*", jamboree: "Kemp/Rams (115)", comment: "Won on better record" },
            { year: 2001, first: "Kemp/Rams", second: "John", third: "Kirk", fourth: "Steve L.", score: "65-59", jamboree: "" },
            { year: 2000, first: "John", second: "Steve/J", third: "Kemp/Rams", fourth: "Steve L.", score: "84-55", jamboree: "" },
            { year: 1999, first: "Eric S.", second: "Steve L.", third: "Steve/J", fourth: "Kemp/Rams", score: "86-51", jamboree: "" },
            { year: 1998, first: "Eric S.", second: "John", third: "Steve L.", fourth: "Eric H.", score: "47-41", jamboree: "" },
            { year: 1997, first: "Kemp", second: "Bill", third: "Chris", fourth: "Kreg/J", score: "27-24", jamboree: "" },
            { year: 1996, first: "Steve/J", second: "Eric S.", third: "John", fourth: "Chris", score: "39-16", jamboree: "" },
            { year: 1995, first: "Chris", second: "Steve L.", third: "Kemp", fourth: "John", score: "67-36", jamboree: "" },
            { year: 1994, first: "Bill", second: "John", third: "Eric H.", fourth: "Steve L.", score: "68-47", jamboree: "" },
            { year: 1993, first: "Bill", second: "Steve L.", third: "John", fourth: "Eric H.", score: "67-41", jamboree: "" },
            { year: 1992, first: "Steve L.", second: "Steve M.", third: "Kirk", fourth: "Bill", score: "38-35", jamboree: "" },
            { year: 1991, first: "Steve M.", second: "John", third: "Chris", fourth: "Nan/Kemp", score: "74-17", jamboree: "" },
            { year: 1990, first: "Bill", second: "Chris", third: "Kreg", fourth: "John", score: "25-22", jamboree: "" },
            { year: 1989, first: "Chris", second: "Eric H.", third: "Eric S.", fourth: "Nan/Kemp", score: "66-42", jamboree: "" },
            { year: 1988, first: "Bill", second: "Steve M.", third: "Eric H.", fourth: "Kreg", score: "45-34", jamboree: "" }
        ];

        // All-time records
        const ALL_TIME_RECORDS = [
            { owner: "Kemp", first: 7, second: 0, third: 2.5, fourth: 3.5 },
            { owner: "Steve L.", first: 5, second: 6, third: 3, fourth: 7 },
            { owner: "Bill/Wes", first: 5, second: 2, third: 0, fourth: 1 },
            { owner: "Kirk/David", first: 4, second: 4, third: 4, fourth: 3 },
            { owner: "Eric S.", first: 3, second: 1, third: 2, fourth: 2 },
            { owner: "Jarrett/Matt", first: 2.5, second: 2.5, third: 2, fourth: 3 },
            { owner: "Chris", first: 2, second: 1, third: 2, fourth: 1 },
            { owner: "Steve M.", first: 2, second: 4, third: 3, fourth: 4 },
            { owner: "Adam", first: 2, second: 1, third: 2, fourth: 0 },
            { owner: "John", first: 1, second: 7, third: 6.5, fourth: 5.5 },
            { owner: "Kevin", first: 1, second: 1, third: 2, fourth: 2 },
            { owner: "Kreg", first: 1, second: 0, third: 1, fourth: 2 },
            { owner: "Andrew", first: 1, second: 0, third: 0, fourth: 1 },
            { owner: "Wes", first: 0.5, second: 1.5, third: 0, fourth: 0 },
            { owner: "Greg/Griffin", first: 0, second: 4, third: 4, fourth: 0 },
            { owner: "Eric H.", first: 0, second: 2, third: 2, fourth: 2 }
        ];

        // Load data
        async function loadData() {
            try {
                const response = await fetch('data.json');
                data = await response.json();
                // Default to the most recent week that has data
                const availableWeeks = (data.weeks || []).map(w => w.week);
                currentWeek = availableWeeks.length > 0 ? Math.max(...availableWeeks) : 1;
                initApp();
            } catch (error) {
                console.error('Error loading data:', error);
                document.getElementById('matchups-container').innerHTML = 
                    '<div class="loading">Error loading data. Please refresh.</div>';
            }
        }

        function initApp() {
            // Season labels
            if (data.season) {
                document.getElementById('season-badge').textContent = `${data.season} SEASON`;
                document.getElementById('schedule-title').textContent = `${data.season} SEASON SCHEDULE & PLAYOFFS`;
                document.getElementById('rules-title').textContent = `${data.season} OPFL Scoring and Rules`;
            }

            // Update timestamp
            if (data.updated_at) {
                const date = new Date(data.updated_at);
                document.getElementById('updated-time').textContent = 
                    `Updated: ${date.toLocaleDateString()} ${date.toLocaleTimeString()}`;
            }

            // Navigation
            document.querySelectorAll('.nav-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
                    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
                    btn.classList.add('active');
                    document.getElementById(btn.dataset.section).classList.add('active');
                });
            });
            
            renderWeekSelector();
            renderMatchups();
            renderSchedule();
            renderStandings();
            renderStatsLeaders();
            renderTeamSelector();
            renderHistory();
            renderBanners();
        }

        function renderWeekSelector() {
            const container = document.getElementById('week-selector');
            const weeks = data.weeks || [];
            const regularSeasonWeeks = data.regular_season_weeks || 15;
            
            // Separate regular season and playoff weeks
            const regSeasonWeeks = weeks.filter(w => w.week <= regularSeasonWeeks);
            const playoffWeeks = weeks.filter(w => w.week > regularSeasonWeeks);
            
            let html = `<span class="week-label">WEEK</span>`;
            
            // Regular season weeks
            html += regSeasonWeeks.map(w => `
                <button class="week-btn ${w.week === currentWeek ? 'active' : ''}" 
                        data-week="${w.week}">${w.week}</button>
            `).join('');
            
            // Show playoff weeks after regular season is complete
            if (data.playoffs || playoffWeeks.length > 0) {
                html += `<span class="week-label" style="margin-left: 1rem;">PLAYOFFS</span>`;
                
                // Week 16
                html += `<button class="week-btn ${currentWeek === 16 ? 'active' : ''}" 
                        data-week="16" style="background: rgba(255, 215, 0, 0.1); border-color: var(--accent-gold);">16</button>`;
                
                // Week 17
                html += `<button class="week-btn ${currentWeek === 17 ? 'active' : ''}" 
                        data-week="17" style="background: rgba(255, 215, 0, 0.1); border-color: var(--accent-gold);">17</button>`;
            }
            
            container.innerHTML = html;

            container.querySelectorAll('.week-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    currentWeek = parseInt(btn.dataset.week);
                    container.querySelectorAll('.week-btn').forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                    renderMatchups();
                });
            });
        }

        function renderSchedule() {
            const container = document.getElementById('schedule-container');
            const currentWeekNum = data.current_week || 1;
            const weeks = data.weeks || [];
            const regularSeasonWeeks = data.regular_season_weeks || 15;
            const playoffs = data.playoffs;
            
            // Create lookup for week data
            const weekDataMap = {};
            weeks.forEach(w => {
                weekDataMap[w.week] = w;
            });

            // Build schedule for regular season (weeks 1-15)
            let html = '';
            for (let weekNum = 1; weekNum <= regularSeasonWeeks; weekNum++) {
                const schedule = getWeekMatchups(weekNum);
                const weekData = weekDataMap[weekNum];
                const isCurrent = weekNum === currentWeekNum;
                const isComplete = weekNum < currentWeekNum;
                const isUpcoming = weekNum > currentWeekNum;
                
                // Get team data for this week
                const teamsByAbbrev = {};
                if (weekData?.teams) {
                    weekData.teams.forEach(t => {
                        teamsByAbbrev[t.abbrev] = t;
                    });
                }

                // Determine badge
                let badgeClass = 'upcoming';
                let badgeText = 'Upcoming';
                if (isCurrent) {
                    badgeClass = 'current';
                    badgeText = 'Current';
                } else if (isComplete) {
                    badgeClass = 'complete';
                    badgeText = 'Complete';
                }

                // Weeks whose pairings have not been published yet are shown
                // as a placeholder rather than skipped entirely.
                if (!schedule) {
                    html += `
                        <div class="schedule-week ${isCurrent ? 'current' : ''}">
                            <div class="schedule-week-header">
                                <span class="schedule-week-title">Week ${weekNum}</span>
                                <span class="schedule-week-badge upcoming">Not yet set</span>
                            </div>
                            <div class="schedule-matchups">
                                <div class="schedule-matchup">
                                    <span class="schedule-no-score">Matchups not posted yet</span>
                                </div>
                            </div>
                        </div>
                    `;
                    continue;
                }

                // Build matchups HTML
                const matchupsHtml = schedule.map(([abbrev1, abbrev2]) => {
                    const team1 = teamsByAbbrev[abbrev1];
                    const team2 = teamsByAbbrev[abbrev2];
                    
                    const hasScores = team1 && team2 && (team1.total_score > 0 || team2.total_score > 0);
                    const score1 = team1?.total_score || 0;
                    const score2 = team2?.total_score || 0;
                    const team1Winning = score1 > score2;
                    const team2Winning = score2 > score1;
                    
                    // Get team names from standings if not in week data
                    const team1Name = team1?.name || data.standings?.find(s => s.abbrev === abbrev1)?.name || abbrev1;
                    const team2Name = team2?.name || data.standings?.find(s => s.abbrev === abbrev2)?.name || abbrev2;

                    if (hasScores || isComplete) {
                        return `
                            <div class="schedule-matchup">
                                <div class="schedule-team ${team1Winning ? 'winner' : team2Winning ? 'loser' : ''}">
                                    <span class="schedule-team-name">${team1Name}</span>
                                </div>
                                <div class="schedule-score-box">
                                    <span class="schedule-score ${team1Winning ? 'winning' : ''}">${score1}</span>
                                    <span class="schedule-vs">-</span>
                                    <span class="schedule-score ${team2Winning ? 'winning' : ''}">${score2}</span>
                                </div>
                                <div class="schedule-team away ${team2Winning ? 'winner' : team1Winning ? 'loser' : ''}">
                                    <span class="schedule-team-name">${team2Name}</span>
                                </div>
                            </div>
                        `;
                    } else {
                        return `
                            <div class="schedule-matchup">
                                <div class="schedule-team">
                                    <span class="schedule-team-name">${team1Name}</span>
                                </div>
                                <div class="schedule-score-box">
                                    <span class="schedule-no-score">vs</span>
                                </div>
                                <div class="schedule-team away">
                                    <span class="schedule-team-name">${team2Name}</span>
                                </div>
                            </div>
                        `;
                    }
                }).join('');

                html += `
                    <div class="schedule-week ${isCurrent ? 'current' : ''}">
                        <div class="schedule-week-header">
                            <span class="schedule-week-title">Week ${weekNum}</span>
                            <span class="schedule-week-badge ${badgeClass}">${badgeText}</span>
                        </div>
                        <div class="schedule-matchups">
                            ${matchupsHtml}
                        </div>
                    </div>
                `;
            }
            
            // Add playoff weeks (16-17) - show projected matchups from current standings
            if (playoffs) {
                html += renderPlayoffScheduleWeeks(playoffs, weekDataMap, currentWeekNum);
            }

            container.innerHTML = html;
        }
        
        function renderPlayoffScheduleWeeks(playoffs, weekDataMap, currentWeekNum) {
            const standings = data.standings || [];
            
            const getTeamName = (abbrev) => {
                if (!abbrev) return 'TBD';
                const team = standings.find(s => s.abbrev === abbrev);
                return team ? team.name : abbrev;
            };
            
            let html = '';
            
            // Week 16 - Semifinals
            const sf1 = playoffs?.week_16?.semifinal_1 || {};
            const sf2 = playoffs?.week_16?.semifinal_2 || {};
            const jamboreeTeams = playoffs?.jamboree_teams || [];
            const week16Data = weekDataMap[16];
            
            const isCurrent16 = currentWeekNum === 16;
            const isComplete16 = currentWeekNum > 16;
            let badge16Class = 'upcoming';
            let badge16Text = 'Upcoming';
            if (isCurrent16) { badge16Class = 'current'; badge16Text = 'Current'; }
            else if (isComplete16) { badge16Class = 'complete'; badge16Text = 'Complete'; }
            
            // Build semifinal matchups
            const renderPlayoffMatchup = (game, label) => {
                if (!game.higher_seed && !game.lower_seed) {
                    return `
                        <div class="schedule-matchup">
                            <div class="schedule-team"><span class="schedule-team-name">TBD</span></div>
                            <div class="schedule-score-box"><span class="schedule-no-score">vs</span></div>
                            <div class="schedule-team away"><span class="schedule-team-name">TBD</span></div>
                        </div>
                    `;
                }
                const team1Name = getTeamName(game.higher_seed);
                const team2Name = getTeamName(game.lower_seed);
                const score1 = game.higher_score || 0;
                const score2 = game.lower_score || 0;
                const hasScores = score1 > 0 || score2 > 0;
                const team1Winning = score1 > score2;
                const team2Winning = score2 > score1;
                
                if (hasScores || isComplete16) {
                    return `
                        <div class="schedule-matchup">
                            <div class="schedule-team ${team1Winning ? 'winner' : team2Winning ? 'loser' : ''}">
                                <span class="schedule-team-name">${team1Name}</span>
                                <span class="schedule-playoff-label">${label}</span>
                            </div>
                            <div class="schedule-score-box">
                                <span class="schedule-score ${team1Winning ? 'winning' : ''}">${score1}</span>
                                <span class="schedule-vs">-</span>
                                <span class="schedule-score ${team2Winning ? 'winning' : ''}">${score2}</span>
                            </div>
                            <div class="schedule-team away ${team2Winning ? 'winner' : team1Winning ? 'loser' : ''}">
                                <span class="schedule-team-name">${team2Name}</span>
                            </div>
                        </div>
                    `;
                }
                return `
                    <div class="schedule-matchup">
                        <div class="schedule-team">
                            <span class="schedule-team-name">${team1Name}</span>
                            <span class="schedule-playoff-label">${label}</span>
                        </div>
                        <div class="schedule-score-box"><span class="schedule-no-score">vs</span></div>
                        <div class="schedule-team away"><span class="schedule-team-name">${team2Name}</span></div>
                    </div>
                `;
            };
            
            html += `
                <div class="schedule-week playoff ${isCurrent16 ? 'current' : ''}">
                    <div class="schedule-week-header">
                        <span class="schedule-week-title">Week 16 - Semifinals</span>
                        <span class="schedule-week-badge ${badge16Class}">${badge16Text}</span>
                    </div>
                    <div class="schedule-matchups">
                        ${renderPlayoffMatchup(sf1, '(1) vs (4)')}
                        ${renderPlayoffMatchup(sf2, '(2) vs (3)')}
                        <div class="schedule-matchup" style="opacity: 0.7; border-top: 1px dashed var(--border); margin-top: 0.5rem; padding-top: 0.5rem;">
                            <div class="schedule-team"><span class="schedule-team-name" style="color: var(--accent-tertiary);">🎪 Jamboree: Seeds 5-12</span></div>
                        </div>
                    </div>
                </div>
            `;
            
            // Week 17 - Finals
            const champ = playoffs?.week_17?.championship || {};
            const third = playoffs?.week_17?.third_place || {};
            
            const isCurrent17 = currentWeekNum === 17;
            const isComplete17 = currentWeekNum > 17;
            let badge17Class = 'upcoming';
            let badge17Text = 'Upcoming';
            if (isCurrent17) { badge17Class = 'current'; badge17Text = 'Current'; }
            else if (isComplete17) { badge17Class = 'complete'; badge17Text = 'Complete'; }
            
            const renderFinalMatchup = (game, label, isChampionship = false) => {
                if (!game.team1 && !game.team2) {
                    return `
                        <div class="schedule-matchup">
                            <div class="schedule-team"><span class="schedule-team-name">TBD</span></div>
                            <div class="schedule-score-box"><span class="schedule-no-score">vs</span></div>
                            <div class="schedule-team away"><span class="schedule-team-name">TBD</span></div>
                        </div>
                    `;
                }
                const team1Name = getTeamName(game.team1);
                const team2Name = getTeamName(game.team2);
                const score1 = game.score1 || 0;
                const score2 = game.score2 || 0;
                const hasScores = score1 > 0 || score2 > 0;
                const team1Winning = score1 > score2;
                const team2Winning = score2 > score1;
                
                const labelStyle = isChampionship ? 'color: var(--accent-gold);' : '';
                
                if (hasScores || isComplete17) {
                    return `
                        <div class="schedule-matchup">
                            <div class="schedule-team ${team1Winning ? 'winner' : team2Winning ? 'loser' : ''}">
                                <span class="schedule-team-name">${team1Name}</span>
                                <span class="schedule-playoff-label" style="${labelStyle}">${label}</span>
                            </div>
                            <div class="schedule-score-box">
                                <span class="schedule-score ${team1Winning ? 'winning' : ''}">${score1}</span>
                                <span class="schedule-vs">-</span>
                                <span class="schedule-score ${team2Winning ? 'winning' : ''}">${score2}</span>
                            </div>
                            <div class="schedule-team away ${team2Winning ? 'winner' : team1Winning ? 'loser' : ''}">
                                <span class="schedule-team-name">${team2Name}</span>
                            </div>
                        </div>
                    `;
                }
                return `
                    <div class="schedule-matchup">
                        <div class="schedule-team">
                            <span class="schedule-team-name">${team1Name}</span>
                            <span class="schedule-playoff-label" style="${labelStyle}">${label}</span>
                        </div>
                        <div class="schedule-score-box"><span class="schedule-no-score">vs</span></div>
                        <div class="schedule-team away"><span class="schedule-team-name">${team2Name}</span></div>
                    </div>
                `;
            };
            
            html += `
                <div class="schedule-week playoff ${isCurrent17 ? 'current' : ''}">
                    <div class="schedule-week-header">
                        <span class="schedule-week-title">Week 17 - Finals</span>
                        <span class="schedule-week-badge ${badge17Class}">${badge17Text}</span>
                    </div>
                    <div class="schedule-matchups">
                        ${renderFinalMatchup(champ, '🏆 Oakland Bowl', true)}
                        ${renderFinalMatchup(third, '3rd Place')}
                        <div class="schedule-matchup" style="opacity: 0.7; border-top: 1px dashed var(--border); margin-top: 0.5rem; padding-top: 0.5rem;">
                            <div class="schedule-team"><span class="schedule-team-name" style="color: var(--accent-tertiary);">🎪 Jamboree continues</span></div>
                        </div>
                    </div>
                </div>
            `;
            
            return html;
        }

        function renderMatchups() {
            const container = document.getElementById('matchups-container');
            const playoffContainer = document.getElementById('playoff-container');
            const regularSeasonWeeks = data.regular_season_weeks || 15;
            
            // Check if this is a playoff week
            const isPlayoffWeek = currentWeek > regularSeasonWeeks;
            
            if (isPlayoffWeek) {
                // Show playoff bracket and Jamboree
                playoffContainer.style.display = 'block';
                container.style.display = 'none';
                renderPlayoffBracket();
                renderJamboree();
                return;
            }
            
            // Regular season - hide playoff container
            playoffContainer.style.display = 'none';
            container.style.display = 'block';
            
            const weekData = data.weeks?.find(w => w.week === currentWeek);
            const schedule = getWeekMatchups(currentWeek);
            
            if (!weekData || !weekData.teams || !schedule) {
                container.innerHTML = '<div class="loading">No matchup data for this week</div>';
                return;
            }

            // Create a lookup map for team data by abbreviation
            const teamsByAbbrev = {};
            weekData.teams.forEach(t => {
                teamsByAbbrev[t.abbrev] = t;
            });

            // Build matchups from schedule
            const matchups = schedule.map(([abbrev1, abbrev2]) => {
                return {
                    team1: teamsByAbbrev[abbrev1],
                    team2: teamsByAbbrev[abbrev2]
                };
            }).filter(m => m.team1 && m.team2);

            container.innerHTML = matchups.map((matchup, idx) => {
                const t1 = matchup.team1;
                const t2 = matchup.team2;
                const t1Winning = t1.total_score > t2.total_score;
                const t2Winning = t2.total_score > t1.total_score;

                return `
                    <div class="matchup-card">
                        <div class="matchup-header">
                            <div class="team">
                                <div class="team-name">${t1.name}</div>
                                <div class="team-owner">${t1.owner}</div>
                            </div>
                            <div class="vs-container">
                                <div class="score-display">
                                    <span class="score ${t1Winning ? 'winning' : t2Winning ? 'losing' : ''}">${t1.total_score}</span>
                                    <span class="score-divider">-</span>
                                    <span class="score ${t2Winning ? 'winning' : t1Winning ? 'losing' : ''}">${t2.total_score}</span>
                                </div>
                            </div>
                            <div class="team right">
                                <div class="team-name">${t2.name}</div>
                                <div class="team-owner">${t2.owner}</div>
                            </div>
                        </div>
                        <button class="expand-btn" onclick="toggleRoster(${idx})">View Rosters</button>
                        <div class="roster-panel" id="roster-${idx}">
                            <div class="roster-grid">
                                <div class="roster-column">
                                    <h4>${t1.name}</h4>
                                    ${renderRosterList(t1.roster)}
                                </div>
                                <div class="roster-column">
                                    <h4>${t2.name}</h4>
                                    ${renderRosterList(t2.roster)}
                                </div>
                            </div>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function renderRosterList(roster) {
            if (!roster) return '<div class="player-row">No roster data</div>';
            
            const starters = roster.filter(p => p.starter);
            const bench = roster.filter(p => !p.starter);
            
            return [...starters, ...bench].map(p => `
                <div class="player-row ${p.starter ? '' : 'bench'}">
                    <div class="player-info">
                        <span class="position-tag">${p.position}</span>
                        <span class="player-name">${p.name}</span>
                        <span class="player-team">${p.nfl_team}</span>
                    </div>
                    <span class="player-score">${p.score || 0}</span>
                </div>
            `).join('');
        }

        function toggleRoster(idx) {
            const panel = document.getElementById(`roster-${idx}`);
            panel.classList.toggle('expanded');
        }

        function renderPlayoffBracket() {
            const playoffs = data.playoffs;
            if (!playoffs) {
                document.getElementById('playoff-bracket').innerHTML = '<div class="loading">Playoff seeding not yet determined</div>';
                return;
            }
            
            const standings = data.standings || [];
            
            const getTeamName = (abbrev) => {
                const team = standings.find(s => s.abbrev === abbrev);
                return team ? team.name : abbrev;
            };
            
            const seeds = playoffs.seeds || {};
            const sf1 = playoffs.week_16?.semifinal_1 || {};
            const sf2 = playoffs.week_16?.semifinal_2 || {};
            const champ = playoffs.week_17?.championship || {};
            const third = playoffs.week_17?.third_place || {};
            
            const renderBracketTeam = (abbrev, score, isWinner, isLoser, showSeed = true) => {
                if (!abbrev) {
                    return `<div class="bracket-tbd">TBD</div>`;
                }
                const seed = seeds[abbrev] || '?';
                const winClass = isWinner ? 'winner' : (isLoser ? 'loser' : '');
                return `
                    <div class="bracket-team ${winClass}">
                        <div class="bracket-team-info">
                            ${showSeed ? `<span class="bracket-seed">${seed}</span>` : ''}
                            <span class="bracket-team-name">${getTeamName(abbrev)}</span>
                        </div>
                        <span class="bracket-score">${score || 0}</span>
                    </div>
                `;
            };
            
            const html = `
                <!-- Left: Semifinals -->
                <div class="bracket-round">
                    <div class="bracket-game">
                        <div class="bracket-game-header">Semifinal 1 • Week 16</div>
                        ${renderBracketTeam(sf1.higher_seed, sf1.higher_score, sf1.winner === sf1.higher_seed, sf1.loser === sf1.higher_seed)}
                        ${renderBracketTeam(sf1.lower_seed, sf1.lower_score, sf1.winner === sf1.lower_seed, sf1.loser === sf1.lower_seed)}
                    </div>
                    <div class="bracket-game">
                        <div class="bracket-game-header">Semifinal 2 • Week 16</div>
                        ${renderBracketTeam(sf2.higher_seed, sf2.higher_score, sf2.winner === sf2.higher_seed, sf2.loser === sf2.higher_seed)}
                        ${renderBracketTeam(sf2.lower_seed, sf2.lower_score, sf2.winner === sf2.lower_seed, sf2.loser === sf2.lower_seed)}
                    </div>
                </div>
                
                <!-- Center: Finals -->
                <div class="bracket-finals">
                    <div class="bracket-game championship">
                        <div class="bracket-game-header">🏆 Oakland Bowl • Week 17</div>
                        ${renderBracketTeam(champ.team1, champ.score1, champ.winner === champ.team1, champ.loser === champ.team1, false)}
                        ${renderBracketTeam(champ.team2, champ.score2, champ.winner === champ.team2, champ.loser === champ.team2, false)}
                    </div>
                    <div class="bracket-game third-place">
                        <div class="bracket-game-header">3rd Place Game • Week 17</div>
                        ${renderBracketTeam(third.team1, third.score1, third.winner === third.team1, third.loser === third.team1, false)}
                        ${renderBracketTeam(third.team2, third.score2, third.winner === third.team2, third.loser === third.team2, false)}
                    </div>
                </div>
                
                <!-- Right: Results summary -->
                <div class="bracket-round">
                    <div class="bracket-game">
                        <div class="bracket-game-header">Final Standings</div>
                        <div class="bracket-team ${champ.winner ? 'winner' : ''}">
                            <div class="bracket-team-info">
                                <span class="bracket-seed">1st</span>
                                <span class="bracket-team-name">${champ.winner ? getTeamName(champ.winner) : 'TBD'}</span>
                            </div>
                        </div>
                        <div class="bracket-team ${champ.loser ? 'loser' : ''}">
                            <div class="bracket-team-info">
                                <span class="bracket-seed">2nd</span>
                                <span class="bracket-team-name">${champ.loser ? getTeamName(champ.loser) : 'TBD'}</span>
                            </div>
                        </div>
                        <div class="bracket-team ${third.winner ? '' : ''}">
                            <div class="bracket-team-info">
                                <span class="bracket-seed">3rd</span>
                                <span class="bracket-team-name">${third.winner ? getTeamName(third.winner) : 'TBD'}</span>
                            </div>
                        </div>
                        <div class="bracket-team ${third.loser ? 'loser' : ''}">
                            <div class="bracket-team-info">
                                <span class="bracket-seed">4th</span>
                                <span class="bracket-team-name">${third.loser ? getTeamName(third.loser) : 'TBD'}</span>
                            </div>
                        </div>
                    </div>
                </div>
            `;
            
            document.getElementById('playoff-bracket').innerHTML = html;
        }

        function renderJamboree() {
            const playoffs = data.playoffs;
            const container = document.getElementById('jamboree-standings');
            
            if (!playoffs || !playoffs.jamboree || !playoffs.jamboree.standings) {
                container.innerHTML = '<div class="loading">Jamboree standings will appear once playoffs begin</div>';
                return;
            }
            
            const standings = playoffs.jamboree.standings;
            const winner = playoffs.jamboree.winner;
            
            if (standings.length === 0) {
                container.innerHTML = '<div class="loading">No Jamboree teams</div>';
                return;
            }
            
            container.innerHTML = `
                <table class="jamboree-table">
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>Team</th>
                            <th class="num">Week 16</th>
                            <th class="num">Week 17</th>
                            <th class="num">Total</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${standings.map((team, idx) => {
                            const isWinner = winner && team.abbrev === winner;
                            return `
                                <tr class="${isWinner ? 'winner' : ''}">
                                    <td><span class="jamboree-rank">${idx + 1}</span></td>
                                    <td>
                                        ${team.name}
                                        ${isWinner ? '<span class="jamboree-winner-badge">CHAMPION</span>' : ''}
                                    </td>
                                    <td class="num">${team.week_16_score}</td>
                                    <td class="num">${team.week_17_score}</td>
                                    <td class="num total-col">${team.total}</td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            `;
        }

        function renderStandings() {
            const container = document.getElementById('standings-container');
            let standings = data.standings || [];

            if (standings.length === 0) {
                container.innerHTML = '<div class="loading">No standings data yet</div>';
                return;
            }

            // Sort standings by: 1) rank_points (wins + top6), 2) wins, 3) points_for
            standings = [...standings].sort((a, b) => {
                // First by rank_points (descending)
                const rpDiff = (b.rank_points || 0) - (a.rank_points || 0);
                if (rpDiff !== 0) return rpDiff;
                // Then by wins (descending)
                const winsDiff = (b.wins || 0) - (a.wins || 0);
                if (winsDiff !== 0) return winsDiff;
                // Then by points_for (descending)
                return (b.points_for || 0) - (a.points_for || 0);
            });

            const through = data.standings_through_week || 0;
            const note = through
                ? `<div class="standings-note">Through Week ${through}${data.standings_in_progress ? ' &middot; week in progress' : ''}</div>`
                : '';

            container.innerHTML = note + `
                <table class="standings-table">
                    <thead>
                        <tr>
                            <th>Rank</th>
                            <th>Team</th>
                            <th class="num">Pts</th>
                            <th class="num">W-L-T</th>
                            <th class="num">Top 6</th>
                            <th class="num">PF</th>
                            <th class="num">PA</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${standings.map((team, idx) => {
                            const rank = idx + 1;
                            const isPlayoff = rank <= 4;
                            const isToilet = rank >= standings.length - 1;
                            const record = `${team.wins || 0}-${team.losses || 0}-${team.ties || 0}`;
                
                            return `
                                <tr>
                                    <td>
                                        <span class="rank ${isPlayoff ? 'playoffs' : isToilet ? 'toilet-bowl' : ''}">${rank}</span>
                                    </td>
                                    <td>
                                        <span class="team-name">${team.name}</span>
                                        <span class="team-code">${team.abbrev}</span>
                                        ${isPlayoff ? '<span class="playoff-label playoffs">PLAYOFF</span>' : ''}
                                        ${isToilet ? '<span class="playoff-label toilet">TOILET</span>' : ''}
                                    </td>
                                    <td class="num rank-points">${team.rank_points?.toFixed(1) || 0}</td>
                                    <td class="num record">${record}</td>
                                    <td class="num top-half">${team.top_half || 0}</td>
                                    <td class="num points-for">${team.points_for?.toFixed(1) || 0}</td>
                                    <td class="num points-against">${team.points_against?.toFixed(1) || 0}</td>
                                </tr>
                            `;
                        }).join('')}
                    </tbody>
                </table>
            `;
        }

        // Points Leaders
        //
        // Ported from the sibling QPFL site's getStatsLeaders/renderStatsLeaders.
        // OPFL's weeks[] is flatter than QPFL's (week.teams[] rather than
        // week.matchups[].team1/team2), which simplifies the aggregation loop;
        // everything else - the composite key, the per-position grouping, the
        // "top 5 unless a position is expanded" UI - carries over unchanged.
        const STATS_POSITIONS = ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC'];
        const STATS_POSITION_NAMES = {
            QB: 'Quarterbacks',
            RB: 'Running Backs',
            WR: 'Wide Receivers',
            TE: 'Tight Ends',
            K: 'Kickers',
            DF: 'Defenses',
            HC: 'Head Coaches',
        };

        let currentStatsPosition = 'ALL';
        let _statsLeadersCache = { dataRef: null, value: null };

        function getStatsLeaders() {
            if (!data) return {};

            // Memoized: leaders depend only on the current data object, and
            // recomputing means re-walking every week's every roster.
            if (_statsLeadersCache.dataRef === data) {
                return _statsLeadersCache.value;
            }

            // key: "name|nflTeam|position" -> aggregate. Position is part of the
            // key (not just name+team) because a head coach and a defense can
            // share an NFL team abbreviation.
            const playerStats = {};

            for (const week of data.weeks || []) {
                // An in-progress week's scores are still moving; counting it
                // would make the leaderboard jump around mid-Sunday.
                if (week.final === false) continue;

                for (const team of week.teams || []) {
                    for (const player of team.roster || []) {
                        if (!player.name || !player.position) continue;

                        const key = `${player.name}|${player.nfl_team || ''}|${player.position}`;
                        if (!playerStats[key]) {
                            playerStats[key] = {
                                name: player.name,
                                nfl_team: player.nfl_team || '',
                                position: player.position,
                                fantasy_team: team.abbrev,
                                total_points: 0,
                                weeks_played: 0,
                            };
                        }

                        playerStats[key].fantasy_team = team.abbrev;
                        if (player.score !== undefined && player.score !== null) {
                            playerStats[key].total_points += player.score;
                            if (player.score !== 0) {
                                playerStats[key].weeks_played++;
                            }
                        }
                    }
                }
            }

            const byPosition = {};
            for (const player of Object.values(playerStats)) {
                if (!byPosition[player.position]) byPosition[player.position] = [];
                byPosition[player.position].push(player);
            }
            for (const pos of Object.keys(byPosition)) {
                byPosition[pos].sort((a, b) => b.total_points - a.total_points);
            }

            _statsLeadersCache = { dataRef: data, value: byPosition };
            return byPosition;
        }

        function renderStatsLeaders() {
            const leaders = getStatsLeaders();

            const selector = document.getElementById('stats-position-selector');
            selector.innerHTML = `
                <button class="stats-pos-btn ${currentStatsPosition === 'ALL' ? 'active' : ''}"
                        role="tab" aria-selected="${currentStatsPosition === 'ALL'}" data-pos="ALL">All</button>
                ${STATS_POSITIONS.map(pos => `
                    <button class="stats-pos-btn ${currentStatsPosition === pos ? 'active' : ''}"
                            role="tab" aria-selected="${currentStatsPosition === pos}" data-pos="${pos}">${pos}</button>
                `).join('')}
            `;

            selector.querySelectorAll('.stats-pos-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    currentStatsPosition = btn.dataset.pos;
                    renderStatsLeaders();
                });
            });

            const container = document.getElementById('stats-leaders-container');
            const positionsToShow = currentStatsPosition === 'ALL' ? STATS_POSITIONS : [currentStatsPosition];
            container.classList.toggle('single-position', currentStatsPosition !== 'ALL');

            const anyLeaders = positionsToShow.some(pos => (leaders[pos] || []).length > 0);
            if (!anyLeaders) {
                container.innerHTML = '<div class="loading">No completed weeks yet</div>';
                return;
            }

            container.innerHTML = positionsToShow.map(pos => {
                const posLeaders = currentStatsPosition === 'ALL'
                    ? (leaders[pos] || []).slice(0, 5)
                    : (leaders[pos] || []);
                if (posLeaders.length === 0) return '';

                return `
                    <div class="stats-position-card">
                        <div class="stats-position-header">${STATS_POSITION_NAMES[pos] || pos}</div>
                        ${posLeaders.map((player, idx) => {
                            const rank = idx + 1;
                            const rankClass = rank <= 3 ? `rank-${rank}` : '';
                            return `
                                <div class="stats-leader-row ${rankClass}">
                                    <div class="stats-rank">${rank}</div>
                                    <div class="stats-player-info">
                                        <span class="stats-player-name">${player.name}</span>
                                        <div class="stats-player-meta">
                                            <span class="stats-nfl-team">${player.nfl_team}</span>
                                            <span class="stats-fantasy-team">• ${player.fantasy_team}</span>
                                        </div>
                                    </div>
                                    <div class="stats-points">${player.total_points.toFixed(1)}</div>
                                </div>
                            `;
                        }).join('')}
                        ${currentStatsPosition === 'ALL' && (leaders[pos] || []).length > 5 ? `
                            <button class="stats-view-all" data-pos="${pos}">View all ${STATS_POSITION_NAMES[pos]}</button>
                        ` : ''}
                    </div>
                `;
            }).join('');

            container.querySelectorAll('.stats-view-all').forEach(btn => {
                btn.addEventListener('click', () => {
                    currentStatsPosition = btn.dataset.pos;
                    renderStatsLeaders();
                });
            });
        }

        function renderTeamSelector() {
            const container = document.getElementById('team-selector');
            const teams = data.standings || [];

            if (teams.length === 0) return;

            container.innerHTML = teams.map(team => `
                <button class="team-btn" data-abbrev="${team.abbrev}">${team.name}</button>
            `).join('');
            
            container.querySelectorAll('.team-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    container.querySelectorAll('.team-btn').forEach(b => b.classList.remove('active'));
                    btn.classList.add('active');
                    currentTeam = btn.dataset.abbrev;
                    renderTeamPage();
                });
            });
            
            // Select first team by default
            if (teams.length > 0) {
                currentTeam = teams[0].abbrev;
                container.querySelector('.team-btn').classList.add('active');
                renderTeamPage();
            }
        }

        function renderTeamPage() {
            const container = document.getElementById('team-roster-container');
            const teamInfo = data.standings?.find(t => t.abbrev === currentTeam);
            
            if (!teamInfo) {
                container.innerHTML = '<div class="loading">Select a team</div>';
                return;
            }

            // Get roster across all weeks
            const weeks = data.weeks || [];
            const rosterByPlayer = {};

            weeks.forEach(week => {
                const teamData = week.teams?.find(t => t.abbrev === currentTeam);
                if (teamData?.roster) {
                teamData.roster.forEach(player => {
                    const key = `${player.position}-${player.name}`;
                        if (!rosterByPlayer[key]) {
                            rosterByPlayer[key] = {
                            name: player.name,
                            nfl_team: player.nfl_team,
                            position: player.position,
                            weeks: {}
                            };
                    }
                        rosterByPlayer[key].weeks[week.week] = {
                            score: player.score || 0,
                        starter: player.starter
                    };
                });
                }
            });
            
            // Group by position
            const positions = ['QB', 'RB', 'WR', 'TE', 'K', 'DF', 'HC'];
            const weekNums = weeks.map(w => w.week).sort((a, b) => a - b);
            
            let tableRows = '';
            positions.forEach(pos => {
                const playersInPos = Object.values(rosterByPlayer).filter(p => p.position === pos);
                if (playersInPos.length === 0) return;

                tableRows += `<tr class="position-group"><td colspan="${weekNums.length + 3}">${pos}</td></tr>`;
                
                playersInPos.forEach(player => {
                    // Calculate season total for this player (all weeks, regardless of starter status)
                    const seasonTotal = Object.values(player.weeks).reduce((sum, w) => sum + (w.score || 0), 0);
                    
                    tableRows += `
                        <tr>
                            <td><span class="player-name">${player.name}</span></td>
                            <td><span class="player-team">${player.nfl_team}</span></td>
                            ${weekNums.map(w => {
                                const weekData = player.weeks[w];
                                if (!weekData) return '<td class="week-score">-</td>';
                                const cls = weekData.starter ? 'starter' : 'bench';
                                return `<td class="week-score ${cls}">${weekData.score}</td>`;
                            }).join('')}
                            <td class="week-score season-total">${seasonTotal}</td>
                        </tr>
                    `;
                });
            });
            
            // Calculate totals
            const totals = weekNums.map(w => {
                const weekData = weeks.find(wk => wk.week === w);
                const teamData = weekData?.teams?.find(t => t.abbrev === currentTeam);
                return teamData?.total_score || 0;
            });
            const seasonGrandTotal = totals.reduce((sum, t) => sum + t, 0);

            tableRows += `
                <tr class="total-row">
                    <td colspan="2"><strong>TOTAL</strong></td>
                    ${totals.map(t => `<td class="week-score">${t}</td>`).join('')}
                    <td class="week-score season-total">${seasonGrandTotal}</td>
                </tr>
            `;
            
            // Taxi Squad section
            const taxiSquad = data.taxi_squads?.[currentTeam] || [];
            let taxiHtml = `
                <div class="taxi-section">
                    <h3>Taxi Squad (${taxiSquad.length}/3)</h3>
                    ${taxiSquad.length === 0 
                        ? '<div class="taxi-empty">No players on taxi squad</div>'
                        : taxiSquad.map(p => `
                            <div class="player-row">
                                <div class="player-info">
                                    <span class="position-tag">${p.position}</span>
                                    <span class="player-name">${p.name}</span>
                                    <span class="player-team">${p.nfl_team}</span>
                                    </div>
                            </div>
                        `).join('')
                    }
                        </div>
                    `;
            
            // Draft Picks section
            let picksHtml = '';
            if (data.draft_picks && data.draft_picks[currentTeam]) {
                const teamPicks = data.draft_picks[currentTeam];
                const seasons = Object.keys(teamPicks).sort();
                
                picksHtml = `
                    <div class="draft-picks-section">
                        <h3>Future Draft Picks</h3>
                        <div class="picks-grid">
                            ${seasons.map(season => `
                                <div class="picks-season">
                                    <div class="picks-season-header">${season}</div>
                                    ${Object.entries(teamPicks[season]).map(([draftType, picks]) => {
                                        if (!picks || picks.length === 0) return '';
                                        const typeLabel = draftType.replace('_', ' ').replace(/\b\w/g, c => c.toUpperCase());
                                        return `
                                            <div class="picks-draft-type">
                                                <div class="picks-type-label">${typeLabel}</div>
                                                <div class="picks-list">
                                                    ${picks.map(p => {
                                                        const fromLabel = p.own ? '' : ` <span class="pick-from">(${p.from})</span>`;
                                                        return `<span class="pick-item ${p.own ? 'own' : 'acquired'}">R${p.round}${fromLabel}</span>`;
                                                    }).join('')}
                                                </div>
                                            </div>
                                        `;
                                    }).join('')}
                                </div>
                            `).join('')}
                        </div>
                    </div>
                `;
            }
            
            container.innerHTML = `
                <div class="team-header">
                    <h2>${teamInfo.name}</h2>
                    <div class="owner">${teamInfo.owner}</div>
                </div>
                <div style="overflow-x: auto;">
                    <table class="roster-table">
                        <thead>
                            <tr>
                                <th>Player</th>
                                <th>Team</th>
                                ${weekNums.map(w => `<th class="week-col">W${w}</th>`).join('')}
                                <th class="week-col total-col">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRows}
                        </tbody>
                    </table>
                </div>
                ${taxiHtml}
                ${picksHtml}
            `;
        }

        function renderHistory() {
            // Render championships table
            const champBody = document.getElementById('championships-body');
            champBody.innerHTML = CHAMPIONSHIPS.map(c => `
                <tr>
                    <td class="year-cell">${c.year}</td>
                    <td class="champion">${c.first}</td>
                    <td>${c.second}</td>
                    <td>${c.third}</td>
                    <td>${c.fourth}</td>
                    <td class="score-cell">${c.score}</td>
                    <td class="jamboree-cell">${c.jamboree || '-'}</td>
                </tr>
            `).join('');

            // Render all-time records
            const recordsBody = document.getElementById('records-body');
            recordsBody.innerHTML = ALL_TIME_RECORDS.map(r => {
                const playoffs = r.first + r.second + r.third + r.fourth;
                // Approximate seasons (based on data)
                const seasons = getSeasonsForOwner(r.owner);
                const pct = seasons > 0 ? ((playoffs / seasons) * 100).toFixed(1) : 'N/A';
                return `
                    <tr>
                        <td><strong>${r.owner}</strong></td>
                        <td class="num champion">${r.first}</td>
                        <td class="num">${r.second}</td>
                        <td class="num">${r.third}</td>
                        <td class="num">${r.fourth}</td>
                        <td class="num">${playoffs}</td>
                        <td class="num">${pct}%</td>
                    </tr>
                `;
            }).join('');

            // Tab switching
            document.querySelectorAll('.history-tab').forEach(tab => {
                tab.addEventListener('click', () => {
                    document.querySelectorAll('.history-tab').forEach(t => t.classList.remove('active'));
                    document.querySelectorAll('.history-content').forEach(c => c.classList.add('hidden'));
                    tab.classList.add('active');
                    document.getElementById(`${tab.dataset.tab}-tab`).classList.remove('hidden');
                });
            });
        }

        // Helper function to estimate seasons for playoff percentage
        function getSeasonsForOwner(owner) {
            const seasonMap = {
                'Kemp': 37, 'Steve L.': 37, 'Bill/Wes': 26, 'Kirk/David': 35,
                'Eric S.': 21, 'Jarrett/Matt': 23, 'Chris': 12, 'Steve M.': 27,
                'Adam': 23, 'John': 37, 'Kevin': 16, 'Kreg': 17,
                'Andrew': 16, 'Wes': 6, 'Greg/Griffin': 23, 'Eric H.': 17
            };
            return seasonMap[owner] || 0;
        }

        function renderBanners() {
            const container = document.getElementById('banners-container');
            const bannerImages = data.banners || [];
            
            if (bannerImages.length === 0) {
                container.innerHTML = '<div class="loading">No banners available</div>';
                return;
            }
            
            // Banners are sorted by year descending (most recent first)
            // Filename format: 2024.png, 2023.png, etc.
            container.innerHTML = bannerImages.map(img => {
                const year = img.replace('.png', '');
                return `
                    <div class="banner-item">
                        <img src="images/banners/${img}" alt="OPFL ${year} Champion" loading="lazy">
                    </div>
                `;
            }).join('');
        }

        // Initialize
                    loadData();
