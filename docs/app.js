(function () {
  'use strict';

  const bridgeClient = createGasBridgeClient_();
  const state = {
    page: getInitialPage_(),
    participantId: '',
    participantOptions: [],
    newParticipantName: '',
    participantDashboard: null,
    adminDashboard: null,
    activeParticipantTab: 'entry',
    activePredictionGender: '男子',
    revealViewMode: 'user',
    selectedRevealParticipantId: '',
    selectedRevealEventId: '',
    busy: false,
    busyMessage: '',
  };
  const FOUR_PLACE_EVENT_NAMES = { '100m 自由形': true, '200m 自由形': true };

  function init() {
    render();
    if (!bridgeClient.isConfigured()) {
      flash('GAS のURLが未設定です。docs/config.js（Pages配信時はGitHub Variables）を確認してください。', true);
      return;
    }
    if (state.page === 'admin') {
      loadAdminDashboard();
    } else {
      loadParticipantOptions();
    }
  }

  function render() {
    const app = document.getElementById('app');
    app.innerHTML = '';

    const container = el('div', { className: 'container' });
    container.appendChild(renderHero());
    if (state.page === 'admin') {
      container.appendChild(renderAdminMainPanel());
    } else {
      if (state.participantDashboard) {
        container.appendChild(renderParticipantMainPanel());
      } else {
        const grid = el('div', { className: 'grid' });
        grid.appendChild(renderParticipantLoginPanel());
        container.appendChild(grid);
      }
    }
    app.appendChild(container);
    syncBusyOverlay();
  }

  function renderHero() {
    const title =
      (state.participantDashboard && state.participantDashboard.config && state.participantDashboard.config.appTitle) ||
      (state.adminDashboard && state.adminDashboard.config && state.adminDashboard.config.appTitle) ||
      '日本選手権 予想アプリ';
    return el('section', { className: 'hero' }, [
      el('h1', {}, [text(title)]),
      state.page === 'admin'
        ? el('p', {}, [text('管理用ページ（URL直打ち用）。公開設定・結果入力・集計確認を行います。')])
        : text(''),
    ]);
  }

  function renderParticipantLoginPanel() {
    const panel = el('section', { className: 'panel' });
    panel.appendChild(el('h2', {}, [text('名前を選択')]));
    const select = el('select', {
      value: state.participantId,
      onchange: function (e) { state.participantId = e.target.value; }
    }, [option('', '名前を選択してください')]);
    (state.participantOptions || []).forEach(function (p) {
      select.appendChild(option(p.participantId, p.displayName));
    });
    panel.appendChild(el('div', { className: 'row' }, [select]));
    panel.appendChild(el('div', { className: 'button-row' }, [
      el('button', { onclick: participantLogin }, [text('入力を開始')]),
    ]));
    return panel;
  }

  function renderParticipantMainPanel() {
    const panel = el('section', { className: 'panel' });
    panel.appendChild(renderParticipantView());
    return panel;
  }

  function renderParticipantView() {
    const root = el('div');
    const d = state.participantDashboard;
    const deadline = (d.config && d.config.predictionDeadlineIso) || '';
    const locked = isLocked(deadline);

    root.appendChild(el('h2', {}, [text('予想入力')]));
    root.appendChild(el('div', {}, [
      badge('名前: ' + d.participant.displayName),
      locked ? badge('締切後（編集不可）', true) : badge('締切前（編集可）'),
    ]));
    const tabs = el('div', { className: 'tabs' });
    const visibleTabs = [['entry', '予想入力']];
    if (d.revealEnabled) visibleTabs.push(['reveal', 'みんなの予想']);
    if (d.crowdForecastEnabled) visibleTabs.push(['crowd', '投票集計予想']);
    if (d.scoreboardEnabled) visibleTabs.push(['scoreboard', '結果比較']);
    if (!visibleTabs.some(function (t) { return t[0] === state.activeParticipantTab; })) {
      state.activeParticipantTab = 'entry';
    }
    visibleTabs.forEach(function (t) {
      tabs.appendChild(el('button', {
        className: 'tab-btn' + (state.activeParticipantTab === t[0] ? ' active' : ''),
        onclick: function () {
          setParticipantTab_(t[0]);
        }
      }, [text(t[1])]));
    });
    root.appendChild(tabs);

    if (state.activeParticipantTab === 'entry') root.appendChild(renderPredictionEditor(d, locked));
    if (state.activeParticipantTab === 'reveal') root.appendChild(renderRevealedPredictionsPanel(d));
    if (state.activeParticipantTab === 'crowd') root.appendChild(renderCrowdForecastPanel(d));
    if (state.activeParticipantTab === 'scoreboard') root.appendChild(renderParticipantScoreboardPanel(d));

    return root;
  }

  function renderPredictionEditor(dashboard, locked) {
    const box = el('div');
    const master = dashboard.masterData || { events: [], entries: [] };
    const entryByEvent = groupBy(master.entries, 'eventId');
    const myPredictions = dashboard.myPredictions || {};
    const grouped = { '男子': [], '女子': [], 'その他': [] };

    master.events.forEach(function (evt) {
      const k = evt.gender === '男子' || evt.gender === '女子' ? evt.gender : 'その他';
      grouped[k].push(evt);
    });

    box.appendChild(el('div', { className: 'button-row' }, [
      el('button', { onclick: savePredictions, disabled: locked }, [text('予想を保存')]),
      el('button', { className: 'secondary', onclick: participantLogin }, [text('再読み込み')]),
    ]));

    const genderKeys = ['男子', '女子', 'その他'].filter(function (gender) { return grouped[gender].length; });
    if (!genderKeys.includes(state.activePredictionGender)) {
      state.activePredictionGender = genderKeys[0] || '男子';
    }
    if (genderKeys.length > 1) {
      const genderTabs = el('div', { className: 'tabs' });
      genderKeys.forEach(function (gender) {
        genderTabs.appendChild(el('button', {
          className: 'tab-btn' + (state.activePredictionGender === gender ? ' active' : ''),
          onclick: function () {
            state.activePredictionGender = gender;
            render();
          }
        }, [text(gender)]));
      });
      box.appendChild(genderTabs);
    }

    (grouped[state.activePredictionGender] || []).forEach(function (evt) {
        const current = myPredictions[evt.eventId] || {};
        const entries = (entryByEvent[evt.eventId] || []).slice().sort(function (a, b) { return a.seedOrder - b.seedOrder; });
        const card = el('div', { className: 'event-card', 'data-event-id': evt.eventId });
        const searchInput = el('input', {
          type: 'text',
          className: 'entry-filter-input',
          placeholder: '選手名で絞り込み',
          disabled: locked
        });
        card.appendChild(el('div', { className: 'event-head' }, [
          el('div', { className: 'event-title' }, [text(evt.eventName)]),
        ]));
        card.appendChild(field('選手検索', searchInput));
        predictionFieldsForEvent(evt).forEach(function (fieldName, idx) {
          const sel = el('select', { disabled: locked, 'data-pick-field': fieldName }, []);
          populatePredictionSelectOptions(sel, entries, current[fieldName] || '', (idx + 1) + '位を選択', '');
          card.appendChild(el('div', { className: 'pick-grid' }, [
            el('span', {}, [text((idx + 1) + '位')]),
            sel
          ]));
        });
        searchInput.oninput = function () {
          const keyword = String(searchInput.value || '');
          card.querySelectorAll('[data-pick-field]').forEach(function (sel, idx) {
            const placeholder = (idx + 1) + '位を選択';
            populatePredictionSelectOptions(sel, entries, sel.value, placeholder, keyword);
          });
        };
        box.appendChild(card);
    });

    return box;
  }

  function renderRevealedPredictionsPanel(dashboard) {
    const wrap = el('div');
    if (!dashboard.visiblePredictions) {
      wrap.appendChild(el('div', { className: 'message' }, [text('読み込み中です...')]));
      return wrap;
    }

    const model = buildRevealViewModel(dashboard.visiblePredictions, dashboard.masterData);
    syncRevealSelections(model);

    const modeTabs = el('div', { className: 'tabs' }, [
      el('button', {
        className: 'tab-btn' + (state.revealViewMode === 'user' ? ' active' : ''),
        onclick: function () { state.revealViewMode = 'user'; render(); }
      }, [text('ユーザー別')]),
      el('button', {
        className: 'tab-btn' + (state.revealViewMode === 'event' ? ' active' : ''),
        onclick: function () { state.revealViewMode = 'event'; render(); }
      }, [text('種目別')]),
    ]);
    wrap.appendChild(modeTabs);

    if (state.revealViewMode === 'user') {
      const select = el('select', {
        value: state.selectedRevealParticipantId,
        onchange: function (e) {
          state.selectedRevealParticipantId = e.target.value;
          render();
        }
      });
      model.participants.forEach(function (p) {
        select.appendChild(option(p.participantId, p.participantName));
      });
      wrap.appendChild(field('ユーザー', select));
      wrap.appendChild(renderRevealByUser(model));
    } else {
      const select = el('select', {
        value: state.selectedRevealEventId,
        onchange: function (e) {
          state.selectedRevealEventId = e.target.value;
          render();
        }
      });
      model.events.forEach(function (evt) {
        select.appendChild(option(evt.eventId, eventDisplayLabelForReveal(evt)));
      });
      wrap.appendChild(field('種目', select));
      wrap.appendChild(renderRevealByEvent(model));
    }
    return wrap;
  }

  function renderRevealByUser(model) {
    const participantId = state.selectedRevealParticipantId;
    const table = el('table', { className: 'list-table reveal-table' });
    table.appendChild(el('colgroup', {}, [
      el('col', { className: 'reveal-col-first' }),
      el('col', { className: 'reveal-col-rank' }),
      el('col', { className: 'reveal-col-rank' }),
      el('col', { className: 'reveal-col-rank' }),
      el('col', { className: 'reveal-col-rank' }),
    ]));
    table.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', {}, [text('種目')]),
      el('th', {}, [text('1位')]),
      el('th', {}, [text('2位')]),
      el('th', {}, [text('3位')]),
      el('th', {}, [text('4位')]),
    ])]));
    const body = el('tbody');
    model.events.forEach(function (evt) {
      const row = model.byParticipant[participantId] && model.byParticipant[participantId][evt.eventId];
      body.appendChild(el('tr', {}, [
        el('td', {}, [text(eventDisplayLabelForReveal(evt))]),
        el('td', {}, [text(row ? row.pick1 : '')]),
        el('td', {}, [text(row ? row.pick2 : '')]),
        el('td', {}, [text(row ? row.pick3 : '')]),
        el('td', {}, [text(row ? row.pick4 : '')]),
      ]));
    });
    table.appendChild(body);
    return el('div', { className: 'table-scroll reveal-scroll' }, [table]);
  }

  function renderRevealByEvent(model) {
    const eventId = state.selectedRevealEventId;
    const table = el('table', { className: 'list-table reveal-table' });
    table.appendChild(el('colgroup', {}, [
      el('col', { className: 'reveal-col-first' }),
      el('col', { className: 'reveal-col-rank' }),
      el('col', { className: 'reveal-col-rank' }),
      el('col', { className: 'reveal-col-rank' }),
      el('col', { className: 'reveal-col-rank' }),
    ]));
    table.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', {}, [text('ユーザー')]),
      el('th', {}, [text('1位')]),
      el('th', {}, [text('2位')]),
      el('th', {}, [text('3位')]),
      el('th', {}, [text('4位')]),
    ])]));
    const body = el('tbody');
    model.participants.forEach(function (p) {
      const row = model.byParticipant[p.participantId] && model.byParticipant[p.participantId][eventId];
      body.appendChild(el('tr', {}, [
        el('td', {}, [text(p.participantName)]),
        el('td', {}, [text(row ? row.pick1 : '')]),
        el('td', {}, [text(row ? row.pick2 : '')]),
        el('td', {}, [text(row ? row.pick3 : '')]),
        el('td', {}, [text(row ? row.pick4 : '')]),
      ]));
    });
    table.appendChild(body);
    return el('div', { className: 'table-scroll reveal-scroll' }, [table]);
  }

  function renderParticipantScoreboardPanel(dashboard) {
    const wrap = el('div');
    if (dashboard.scoreboard) wrap.appendChild(renderScoreboard(dashboard.scoreboard));
    else wrap.appendChild(el('div', { className: 'message' }, [text('読み込み中です...')]));
    return wrap;
  }

  function renderCrowdForecastPanel(dashboard) {
    const wrap = el('div');
    if (dashboard.crowdForecast) wrap.appendChild(renderCrowdForecast(dashboard.crowdForecast));
    else wrap.appendChild(el('div', { className: 'message' }, [text('読み込み中です...')]));
    return wrap;
  }

  function renderAdminMainPanel() {
    const panel = el('section', { className: 'panel' });
    if (!state.adminDashboard) {
      panel.appendChild(el('div', { className: 'message' }, [text('管理データを読み込み中です。')]));
      return panel;
    }
    panel.appendChild(renderAdminView());
    return panel;
  }

  function renderAdminView() {
    const d = state.adminDashboard;
    const root = el('div');
    root.appendChild(el('h2', {}, [text('管理設定 / 集計')]));

    const sections = el('div', { className: 'two-col' });
    sections.appendChild(renderAdminConfigPanel(d));
    sections.appendChild(renderAdminParticipantsPanel(d));
    root.appendChild(sections);

    root.appendChild(renderAdminResultsPanel(d));
    root.appendChild(renderAdminScoreboardPanel(d));
    root.appendChild(renderAdminCrowdForecastPanel(d));
    root.appendChild(renderAdminRevealPanel(d));
    return root;
  }

  function renderAdminConfigPanel(d) {
    const panel = el('div', { className: 'panel' });
    panel.appendChild(el('h3', {}, [text('公開設定・締切')]));

    panel.appendChild(field('アプリ名', el('input', {
      type: 'text',
      id: 'admin-app-title',
      value: (d.config && d.config.appTitle) || '日本選手権 予想アプリ'
    })));
    panel.appendChild(field('予想締切 (ISO形式/JST推奨)', el('input', {
      type: 'text',
      id: 'admin-deadline',
      placeholder: '例: 2026-03-13T23:59:59+09:00',
      value: (d.config && d.config.predictionDeadlineIso) || ''
    })));
    panel.appendChild(field('他参加者の予想を公開', el('input', {
      type: 'checkbox',
      id: 'admin-reveal',
      checked: String(d.config && d.config.revealPredictions) === 'true'
    })));
    panel.appendChild(field('結果比較を公開', el('input', {
      type: 'checkbox',
      id: 'admin-scoreboard',
      checked: String(d.config && d.config.publishScoreboard) === 'true'
    })));
    panel.appendChild(field('投票集計予想を公開', el('input', {
      type: 'checkbox',
      id: 'admin-crowd-forecast',
      checked: String(d.config && d.config.publishCrowdForecast) === 'true'
    })));
    panel.appendChild(el('div', { className: 'button-row' }, [
      el('button', { className: 'warn', onclick: saveAdminConfig }, [text('設定を保存')]),
    ]));
    return panel;
  }

  function renderAdminResultsPanel(d) {
    const panel = el('div', { className: 'panel' });
    panel.appendChild(el('h3', {}, [text('大会結果入力（1〜3位 / 一部4位）')]));
    const master = d.masterData || { events: [], entries: [] };
    const results = d.results || {};
    const entryByEvent = groupBy(master.entries || [], 'eventId');
    master.events.forEach(function (evt) {
      const curr = results[evt.eventId] || {};
      const card = el('div', { className: 'event-card', 'data-result-event-id': evt.eventId });
      card.appendChild(el('div', { className: 'event-head' }, [
        el('div', { className: 'event-title' }, [text((evt.gender ? evt.gender + ' ' : '') + evt.eventName)]),
        el('div', { className: 'event-sub' }, [text(evt.eventId)]),
      ]));
      resultFieldsForEvent(evt).forEach(function (field, idx) {
        const sel = el('select', { 'data-result-field': field }, [option('', (idx + 1) + '位')]);
        (entryByEvent[evt.eventId] || []).forEach(function (ent) {
          sel.appendChild(option(ent.entryId, formatEntryLabel(ent)));
        });
        sel.value = curr[field] || '';
        card.appendChild(el('div', { className: 'pick-grid' }, [
          el('span', {}, [text((idx + 1) + '位')]),
          sel
        ]));
      });
      panel.appendChild(card);
    });
    panel.appendChild(el('div', { className: 'button-row' }, [
      el('button', { className: 'warn', onclick: saveResults }, [text('結果を保存')]),
    ]));
    return panel;
  }

  function renderAdminParticipantsPanel(d) {
    const panel = el('div', { className: 'panel' });
    panel.appendChild(el('h3', {}, [text('参加者一覧')]));
    panel.appendChild(field('参加者名を追加', el('input', {
      type: 'text',
      id: 'new-participant-name',
      value: state.newParticipantName,
      placeholder: '例: Aさん',
      oninput: function (e) { state.newParticipantName = e.target.value; }
    })));
    panel.appendChild(el('div', { className: 'button-row' }, [
      el('button', { onclick: addParticipant }, [text('参加者を追加')]),
    ]));

    const table = el('table', { className: 'list-table' });
    table.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', {}, [text('名前')]),
      el('th', {}, [text('登録日時')]),
      el('th', {}, [text('操作')]),
    ])]));

    const body = el('tbody');
    (d.participants || []).forEach(function (p) {
      body.appendChild(el('tr', {}, [
        el('td', {}, [text(p.displayName)]),
        el('td', {}, [text(p.createdAt || '')]),
        el('td', {}, [
          el('button', {
            className: 'warn',
            onclick: function () { deleteParticipant(p); }
          }, [text('削除')])
        ]),
      ]));
    });
    table.appendChild(body);
    panel.appendChild(table);
    return panel;
  }

  function renderAdminScoreboardPanel(d) {
    const panel = el('div', { className: 'panel' });
    panel.appendChild(el('h3', {}, [text('管理者用 集計プレビュー')]));
    panel.appendChild(el('p', { className: 'muted' }, [text('公開前でも確認できます。')]));
    panel.appendChild(renderScoreboard(d.scoreboard || { ranking: [], completedEvents: 0, scoringRule: '' }));
    return panel;
  }

  function renderAdminCrowdForecastPanel(d) {
    const panel = el('div', { className: 'panel' });
    panel.appendChild(el('h3', {}, [text('投票集計予想（管理者確認）')]));
    panel.appendChild(renderCrowdForecast(d.crowdForecast || { events: [], rule: '' }));
    return panel;
  }

  function renderAdminRevealPanel(d) {
    const panel = el('div', { className: 'panel' });
    panel.appendChild(el('h3', {}, [text('全員の予想（管理者確認）')]));
    panel.appendChild(renderPredictionListTable(d.visiblePredictions || [], d.masterData));
    return panel;
  }

  function renderPredictionListTable(list, masterData) {
    const master = masterData || { events: [], entries: [] };
    const eventMap = toMap(master.events || [], 'eventId');
    const entryMap = toMap(master.entries || [], 'entryId');
    const table = el('table', { className: 'list-table' });
    table.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', {}, [text('参加者')]),
      el('th', {}, [text('種目')]),
      el('th', {}, [text('1位')]),
      el('th', {}, [text('2位')]),
      el('th', {}, [text('3位')]),
      el('th', {}, [text('4位')]),
      el('th', {}, [text('更新')]),
    ])]));
    const body = el('tbody');
    (list || []).forEach(function (row) {
      body.appendChild(el('tr', {}, [
        el('td', {}, [text(row.participantName || row.participantId)]),
        el('td', {}, [text(eventMap[row.eventId] ? eventMap[row.eventId].eventName : row.eventId)]),
        el('td', {}, [text(entryLabelById(entryMap, row.pick1EntryId))]),
        el('td', {}, [text(entryLabelById(entryMap, row.pick2EntryId))]),
        el('td', {}, [text(entryLabelById(entryMap, row.pick3EntryId))]),
        el('td', {}, [text(entryLabelById(entryMap, row.pick4EntryId))]),
        el('td', {}, [text(row.updatedAt || '')]),
      ]));
    });
    table.appendChild(body);
    return table;
  }

  function renderScoreboard(scoreboard) {
    const wrap = el('div');
    wrap.appendChild(el('div', { className: 'message' }, [
      text('採点: ' + (scoreboard.scoringRule || 'exact=3pt, top3/4=1pt') + ' / 結果入力済み種目数: ' + (scoreboard.completedEvents || 0))
    ]));
    const table = el('table', { className: 'list-table score-table' });
    table.appendChild(el('colgroup', {}, [
      el('col', { className: 'score-col-rank' }),
      el('col', { className: 'score-col-name' }),
      el('col', { className: 'score-col-num' }),
      el('col', { className: 'score-col-num' }),
      el('col', { className: 'score-col-num' }),
    ]));
    table.appendChild(el('thead', {}, [el('tr', {}, [
      el('th', {}, [text('順位')]),
      el('th', {}, [text('参加者')]),
      el('th', {}, [text('得点')]),
      el('th', {}, [text('完全一致')]),
      el('th', {}, [text('TOP3/4的中')]),
    ])]));
    const body = el('tbody');
    (scoreboard.ranking || []).forEach(function (row, idx) {
      body.appendChild(el('tr', {}, [
        el('td', {}, [text(String(idx + 1))]),
        el('td', {}, [text(row.participantName)]),
        el('td', {}, [text(String(row.totalScore))]),
        el('td', {}, [text(String(row.exactHits))]),
        el('td', {}, [text(String(row.top3Hits))]),
      ]));
    });
    table.appendChild(body);
    wrap.appendChild(el('div', { className: 'table-scroll score-scroll' }, [table]));
    return wrap;
  }

  function renderCrowdForecast(crowdForecast) {
    const wrap = el('div');
    wrap.appendChild(el('div', { className: 'message' }, [
      text('集計ルール: ' + (crowdForecast.rule || ''))
    ]));

    (crowdForecast.events || []).forEach(function (evt) {
      const card = el('div', { className: 'event-card' });
      card.appendChild(el('div', { className: 'event-head' }, [
        el('div', { className: 'event-title' }, [text((evt.gender ? evt.gender + ' ' : '') + evt.eventName)]),
        el('div', { className: 'event-sub' }, [text('投票数: ' + String(evt.submissionCount || 0))]),
      ]));

      const table = el('table', { className: 'list-table crowd-table' });
      table.appendChild(el('colgroup', {}, [
        el('col', { className: 'crowd-col-rank' }),
        el('col', { className: 'crowd-col-name' }),
        el('col', { className: 'crowd-col-num' }),
        el('col', { className: 'crowd-col-num' }),
      ]));
      table.appendChild(el('thead', {}, [el('tr', {}, [
        el('th', {}, [text('順位')]),
        el('th', {}, [text('選手')]),
        el('th', {}, [text('点')]),
        el('th', {}, [text('票数')]),
      ])]));
      const body = el('tbody');
      if (!(evt.ranking || []).length) {
        body.appendChild(el('tr', {}, [
          el('td', { colspan: '4', className: 'muted' }, [text('まだ投票がありません')]),
        ]));
      } else {
        (evt.ranking || []).forEach(function (row) {
          body.appendChild(el('tr', {}, [
            el('td', {}, [text(String(row.predictedRank))]),
            el('td', {}, [
              el('div', {}, [text(String(row.athleteName || ''))]),
              row.team ? el('div', { className: 'muted crowd-team' }, [text(String(row.team))]) : text('')
            ]),
            el('td', {}, [text(String(row.totalPoints))]),
            el('td', {}, [text(String(row.appearances))]),
          ]));
        });
      }
      table.appendChild(body);
      card.appendChild(el('div', { className: 'table-scroll crowd-scroll' }, [table]));
      wrap.appendChild(card);
    });
    return wrap;
  }

  function participantLogin() {
    if (!state.participantId) return flash('参加者を選択してください。', true);
    runServer('participantLogin', [state.participantId], function (res) {
      state.participantDashboard = res;
      state.participantId = res.participant.participantId;
      render();
      flash('入力画面を読み込みました。', false);
    });
  }

  function loadParticipantOptions() {
    runServer('getParticipantOptions', [], function (res) {
      state.participantOptions = (res && res.participants) || [];
      if (state.participantId && !state.participantOptions.some(function (p) { return p.participantId === state.participantId; })) {
        state.participantId = '';
      }
      render();
    });
  }

  function loadAdminDashboard() {
    runServer('adminGetDashboard', [], function (res) {
      state.adminDashboard = res;
      render();
    });
  }

  function savePredictions() {
    if (!state.participantDashboard) return;
    const rows = [];
    document.querySelectorAll('[data-event-id]').forEach(function (card) {
      const vals = { eventId: card.getAttribute('data-event-id') };
      card.querySelectorAll('[data-pick-field]').forEach(function (sel) {
        vals[sel.getAttribute('data-pick-field')] = sel.value;
      });
      rows.push(vals);
    });
    runServer('saveParticipantPredictions', [state.participantId, rows], function (res) {
      state.participantDashboard = res;
      render();
      flash('予想を保存しました。', false);
    });
  }

  function loadVisiblePredictions() {
    runServer('getVisiblePredictions', [state.participantId], function (res) {
      state.participantDashboard.visiblePredictions = res.submissions || [];
      render();
      flash('全員の予想を取得しました。', false);
    });
  }

  function loadParticipantScoreboard() {
    runServer('getParticipantScoreboard', [state.participantId], function (res) {
      state.participantDashboard.scoreboard = res.scoreboard;
      render();
      flash('結果比較を取得しました。', false);
    });
  }

  function loadParticipantCrowdForecast() {
    runServer('getParticipantCrowdForecast', [state.participantId], function (res) {
      state.participantDashboard.crowdForecast = res.crowdForecast;
      render();
      flash('投票集計予想を取得しました。', false);
    });
  }

  function saveAdminConfig() {
    const patch = {
      appTitle: valueOf('admin-app-title'),
      predictionDeadlineIso: valueOf('admin-deadline'),
      revealPredictions: checkedOf('admin-reveal'),
      publishScoreboard: checkedOf('admin-scoreboard'),
      publishCrowdForecast: checkedOf('admin-crowd-forecast'),
    };
    runServer('adminUpdateConfig', [patch], function (res) {
      state.adminDashboard = res;
      render();
      flash('設定を保存しました。', false);
    });
  }

  function saveResults() {
    const rows = [];
    document.querySelectorAll('[data-result-event-id]').forEach(function (card) {
      const row = { eventId: card.getAttribute('data-result-event-id') };
      card.querySelectorAll('[data-result-field]').forEach(function (sel) {
        row[sel.getAttribute('data-result-field')] = sel.value;
      });
      rows.push(row);
    });
    runServer('adminSaveResults', [rows], function (res) {
      state.adminDashboard = res;
      render();
      flash('結果を保存しました。', false);
    });
  }

  function addParticipant() {
    const name = String(state.newParticipantName || '').trim();
    if (!name) return flash('参加者名を入力してください。', true);
    runServer('adminAddParticipant', [name], function (res) {
      state.newParticipantName = '';
      state.adminDashboard = res;
      render();
      flash('参加者を追加しました。', false);
    });
  }

  function deleteParticipant(participant) {
    if (!participant || !participant.participantId) return;
    if (!confirm('「' + (participant.displayName || participant.participantId) + '」を削除します。予想データも削除されます。よろしいですか？')) {
      return;
    }
    runServer('adminDeleteParticipant', [participant.participantId], function (res) {
      state.adminDashboard = res;
      render();
      flash('参加者を削除しました。', false);
    });
  }

  function renderBusyOverlay() {
    return el('div', { className: 'loading-overlay', id: 'busy-overlay' }, [
      el('div', { className: 'loading-card' }, [
        el('div', { className: 'spinner' }),
        el('div', { className: 'muted' }, [text(state.busyMessage || '処理中...')]),
      ]),
    ]);
  }

  function runServer(method, args, onSuccess) {
    setBusy(true, busyLabelForMethod(method));
    bridgeClient.call(method, args || [])
      .then(function (res) {
        setBusy(false, '');
        onSuccess(res);
      })
      .catch(function (err) {
        setBusy(false, '');
        flash((err && err.message) || String(err) || 'エラーが発生しました。', true);
      });
  }

  function setBusy(busy, message) {
    state.busy = !!busy;
    state.busyMessage = message || '';
    syncBusyOverlay();
  }

  function syncBusyOverlay() {
    const app = document.getElementById('app');
    if (!app) return;
    const existing = document.getElementById('busy-overlay');
    if (state.busy) {
      if (existing) {
        const msg = existing.querySelector('.muted');
        if (msg) msg.textContent = state.busyMessage || '処理中...';
      } else {
        app.appendChild(renderBusyOverlay());
      }
      return;
    }
    if (existing) existing.remove();
  }

  function busyLabelForMethod(method) {
    return {
      participantLogin: '入力フォームを読み込み中...',
      getParticipantOptions: '参加者一覧を読み込み中...',
      saveParticipantPredictions: '保存中...',
      getVisiblePredictions: 'みんなの予想を読み込み中...',
      getParticipantCrowdForecast: '投票集計予想を読み込み中...',
      getParticipantScoreboard: '結果比較を読み込み中...',
      adminGetDashboard: '管理画面を読み込み中...',
      adminUpdateConfig: '設定を保存中...',
      adminAddParticipant: '参加者を追加中...',
      adminSaveResults: '結果を保存中...',
      adminDeleteParticipant: '参加者を削除中...',
    }[method] || '処理中...';
  }

  function isLocked(deadlineIso) {
    if (!deadlineIso) return false;
    const t = new Date(deadlineIso).getTime();
    return Number.isFinite(t) && Date.now() > t;
  }

  function formatEntryLabel(ent) {
    const parts = [ent.athleteName];
    if (ent.team) parts.push('(' + ent.team + ')');
    if (ent.entryTime) parts.push(ent.entryTime);
    return parts.join(' ');
  }

  function entryLabelById(entryMap, entryId) {
    const ent = entryMap[entryId];
    return ent ? formatEntryLabel(ent) : '';
  }

  function groupBy(list, field) {
    const map = {};
    (list || []).forEach(function (item) {
      const key = item[field];
      if (!map[key]) map[key] = [];
      map[key].push(item);
    });
    return map;
  }

  function toMap(list, key) {
    const map = {};
    (list || []).forEach(function (item) { map[item[key]] = item; });
    return map;
  }

  function badge(content, locked) {
    return el('span', { className: 'badge' + (locked ? ' locked' : '') }, [text(content)]);
  }

  function field(labelText, inputNode) {
    return el('div', { className: 'row' }, [el('label', {}, [text(labelText)]), inputNode]);
  }

  function option(value, label) {
    return el('option', { value: value }, [text(label)]);
  }

  function text(str) {
    return document.createTextNode(str == null ? '' : String(str));
  }

  function el(tag, props, children) {
    const node = document.createElement(tag);
    const p = props || {};
    Object.keys(p).forEach(function (k) {
      const v = p[k];
      if (k === 'className') node.className = v;
      else if (k === 'checked') node.checked = !!v;
      else if (k === 'disabled') node.disabled = !!v;
      else if (k === 'value') node.value = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') node[k.toLowerCase()] = v;
      else node.setAttribute(k, v);
    });
    (children || []).forEach(function (child) { node.appendChild(child); });
    return node;
  }

  function valueOf(id) {
    const node = document.getElementById(id);
    return node ? node.value : '';
  }

  function checkedOf(id) {
    const node = document.getElementById(id);
    return !!(node && node.checked);
  }

  function flash(msg, isError) {
    const existing = document.getElementById('global-message');
    if (existing) existing.remove();
    const app = document.getElementById('app');
    const box = el('div', { id: 'global-message', className: 'container' }, [
      el('div', { className: 'message ' + (isError ? 'error' : 'success') }, [text(msg)])
    ]);
    app.insertBefore(box, app.firstChild);
    setTimeout(function () {
      const latest = document.getElementById('global-message');
      if (latest) latest.remove();
    }, 4000);
  }

  function placementCountForEvent(evt) {
    return evt && FOUR_PLACE_EVENT_NAMES[String(evt.eventName || '').trim()] ? 4 : 3;
  }

  function predictionFieldsForEvent(evt) {
    return ['pick1EntryId', 'pick2EntryId', 'pick3EntryId', 'pick4EntryId'].slice(0, placementCountForEvent(evt));
  }

  function resultFieldsForEvent(evt) {
    return ['firstEntryId', 'secondEntryId', 'thirdEntryId', 'fourthEntryId'].slice(0, placementCountForEvent(evt));
  }

  function populatePredictionSelectOptions(selectNode, entries, selectedValue, placeholder, keyword) {
    const q = String(keyword || '').trim().toLowerCase();
    const filtered = !q ? (entries || []) : (entries || []).filter(function (ent) {
      const hay = (formatEntryLabel(ent) + ' ' + (ent.athleteName || '')).toLowerCase();
      return hay.indexOf(q) !== -1;
    });
    const selectedEntry = (entries || []).find(function (ent) { return ent.entryId === selectedValue; }) || null;

    selectNode.innerHTML = '';
    selectNode.appendChild(option('', placeholder));
    filtered.forEach(function (ent) {
      selectNode.appendChild(option(ent.entryId, formatEntryLabel(ent)));
    });
    if (selectedEntry && !filtered.some(function (ent) { return ent.entryId === selectedEntry.entryId; })) {
      selectNode.appendChild(option(selectedEntry.entryId, '（選択中）' + formatEntryLabel(selectedEntry)));
    }
    selectNode.value = selectedValue || '';
  }

  function buildRevealViewModel(list, masterData) {
    const master = masterData || { events: [], entries: [] };
    const eventMap = toMap(master.events || [], 'eventId');
    const entryMap = toMap(master.entries || [], 'entryId');
    const participantMap = {};
    const byParticipant = {};

    (list || []).forEach(function (row) {
      participantMap[row.participantId] = row.participantName || row.participantId;
      if (!byParticipant[row.participantId]) byParticipant[row.participantId] = {};
      byParticipant[row.participantId][row.eventId] = {
        pick1: entryLabelById(entryMap, row.pick1EntryId),
        pick2: entryLabelById(entryMap, row.pick2EntryId),
        pick3: entryLabelById(entryMap, row.pick3EntryId),
        pick4: entryLabelById(entryMap, row.pick4EntryId),
      };
    });

    const participants = Object.keys(participantMap)
      .map(function (participantId) {
        return { participantId: participantId, participantName: participantMap[participantId] };
      })
      .sort(function (a, b) { return a.participantName.localeCompare(b.participantName, 'ja'); });

    const events = (master.events || []).slice().sort(function (a, b) { return (a.sortOrder || 0) - (b.sortOrder || 0); }).map(function (evt) {
      return {
        eventId: evt.eventId,
        gender: evt.gender,
        eventName: evt.eventName,
      };
    });

    return { participants: participants, events: events, byParticipant: byParticipant, eventMap: eventMap };
  }

  function syncRevealSelections(model) {
    if (!model.participants.length || !model.events.length) return;
    if (!model.participants.some(function (p) { return p.participantId === state.selectedRevealParticipantId; })) {
      state.selectedRevealParticipantId = model.participants[0].participantId;
    }
    if (!model.events.some(function (evt) { return evt.eventId === state.selectedRevealEventId; })) {
      state.selectedRevealEventId = model.events[0].eventId;
    }
  }

  function eventDisplayLabel(evt) {
    return [(evt && evt.gender) || '', (evt && evt.eventName) || ''].join(' ').trim();
  }

  function setParticipantTab_(tabId) {
    state.activeParticipantTab = tabId;
    render();
    if (!state.participantDashboard) return;
    if (tabId === 'reveal' && !state.participantDashboard.visiblePredictions) {
      loadVisiblePredictions();
    }
    if (tabId === 'crowd' && !state.participantDashboard.crowdForecast) {
      loadParticipantCrowdForecast();
    }
    if (tabId === 'scoreboard' && !state.participantDashboard.scoreboard) {
      loadParticipantScoreboard();
    }
  }

  function eventDisplayLabelForReveal(evt) {
    const gender = (evt && evt.gender) || '';
    const eventName = abbreviateRevealEventName_((evt && evt.eventName) || '');
    return [gender, eventName].join(' ').trim();
  }

  function abbreviateRevealEventName_(name) {
    return String(name || '')
      .replace(/\s*個人メドレー$/, ' IM')
      .replace(/\s*バタフライ$/, ' Fly')
      .replace(/\s*平泳ぎ$/, ' Br')
      .replace(/\s*背泳ぎ$/, ' Ba')
      .replace(/\s*自由形$/, ' Fr')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function getInitialPage_() {
    const params = new URLSearchParams(window.location.search);
    return params.get('page') === 'admin' ? 'admin' : 'participant';
  }

  function createGasBridgeClient_() {
    let iframe = null;
    let bridgePostOrigin = '*';
    let bridgeReadyOrigin = '*';
    let bridgeMessageTargetWindow = null;
    let reqSeq = 1;
    const pending = {};
    let readySettled = false;
    let readyTimer = null;
    let readyResolve;
    let readyReject;
    let ready = new Promise(function (resolve, reject) {
      readyResolve = resolve;
      readyReject = reject;
    });

    window.addEventListener('message', function (event) {
      const msg = event.data || {};
      if (msg.type === 'gas-bridge-ready') {
        if (event.origin) bridgeReadyOrigin = event.origin;
        if (event.source && typeof event.source.postMessage === 'function') {
          bridgeMessageTargetWindow = event.source;
        }
        readySettled = true;
        if (readyTimer) {
          clearTimeout(readyTimer);
          readyTimer = null;
        }
        readyResolve();
        return;
      }
      if (msg.type !== 'gas-rpc-response') return;
      const item = pending[msg.id];
      if (!item) return;
      delete pending[msg.id];
      if (msg.ok) item.resolve(msg.result);
      else item.reject(new Error(msg.error || 'Unknown error'));
    });

    function isConfigured() {
      return !!resolveGasWebAppUrl_();
    }

    function ensureIframe() {
      if (iframe) return;
      const gasUrl = resolveGasWebAppUrl_();
      if (!gasUrl) throw new Error('GAS URL が未設定です。');
      const bridgeUrl = buildBridgeUrl_(gasUrl);
      try {
        bridgePostOrigin = new URL(bridgeUrl).origin;
      } catch (err) {
        bridgePostOrigin = '*';
      }
      iframe = document.createElement('iframe');
      iframe.src = bridgeUrl;
      iframe.title = 'gas-bridge';
      iframe.style.position = 'absolute';
      iframe.style.width = '0';
      iframe.style.height = '0';
      iframe.style.border = '0';
      iframe.style.opacity = '0';
      iframe.style.pointerEvents = 'none';
      iframe.setAttribute('aria-hidden', 'true');
      document.body.appendChild(iframe);
      readyTimer = setTimeout(function () {
        if (readySettled) return;
        readySettled = true;
        readyReject(new Error('GAS bridge の初期化に失敗しました。GAS を再デプロイ済みか、Bridge.html が追加されているか、GAS URL が /exec か確認してください。'));
      }, 10000);
    }

    function call(method, args) {
      return new Promise(function (resolve, reject) {
        try {
          ensureIframe();
        } catch (err) {
          reject(err);
          return;
        }
        ready.then(function () {
          const id = 'rpc_' + (reqSeq++);
          pending[id] = { resolve: resolve, reject: reject };
          const targetWin = bridgeMessageTargetWindow || (iframe && iframe.contentWindow);
          if (!targetWin || typeof targetWin.postMessage !== 'function') {
            delete pending[id];
            reject(new Error('GAS bridge の送信先ウィンドウを取得できません。'));
            return;
          }
          targetWin.postMessage({
            type: 'gas-rpc-request',
            id: id,
            method: method,
            args: args || [],
          }, bridgeReadyOrigin || bridgePostOrigin || '*');
          setTimeout(function () {
            if (!pending[id]) return;
            delete pending[id];
            reject(new Error('GAS応答がタイムアウトしました。'));
          }, 30000);
        }).catch(reject);
      });
    }

    return { call: call, isConfigured: isConfigured };
  }

  function resolveGasWebAppUrl_() {
    const configured = window.APP_CONFIG && String(window.APP_CONFIG.gasWebAppUrl || '').trim();
    if (configured) return configured;
    return '';
  }

  function buildBridgeUrl_(gasWebAppUrl) {
    const url = new URL(gasWebAppUrl);
    url.searchParams.set('page', 'bridge');
    return url.toString();
  }

  window.addEventListener('load', init);
})();
