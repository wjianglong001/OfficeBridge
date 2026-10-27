/* global OfficeBridgeContract, OfficeBridgeExecutor */
(function () {
  'use strict';
  var host = OfficeBridgeExecutor,
    C = OfficeBridgeContract;
  var el = function (id) {
    return document.getElementById(id);
  };
  var feed = el('messages'),
    empty = el('empty-state'),
    prompt = el('analysis-prompt');
  var controller = null,
    token = '',
    ready = false,
    checking = false,
    opening = false,
    applying = false,
    composing = false;
  var pending = null,
    canCapture = false,
    metadata = null,
    conversationId = '',
    operationFeedback = '',
    lastApplied = null,
    conversationLost = false,
    connectionRetry = null,
    connectionCancel = null;
  var selectedAgent = 'codex';
  var lastRangeContext = null,
    questionRecord = null;
  var queuedContinuation = null;
  var applyController = null;
  function revokeContinuation(record) {
    if (!record || !record.continuation || !record.continuation.token) return;
    fetch('/api/v1/continuations/cancel', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-OfficeBridge-Token': token,
      },
      body: JSON.stringify({
        token: record.continuation.token,
        conversationId: record.conversationId,
      }),
    }).catch(function () {});
  }
  function agentName(agentId) {
    return (
      { codex: 'Codex', workbuddy: 'WorkBuddy', doubao: '豆包' }[
        agentId || selectedAgent
      ] || 'Agent'
    );
  }
  function setAgentStatus(kind, text, agentId) {
    var status = el('agent-status');
    status.setAttribute('data-state', kind);
    status.setAttribute(
      'aria-busy',
      String(['connecting', 'switching', 'opening'].indexOf(kind) >= 0),
    );
    status.title = text;
    status.setAttribute('aria-label', text);
    el('agent-name').textContent = agentName(agentId);
    el('agent-state').textContent = {
      connecting: '连接中',
      switching: '切换中',
      opening: '准备中',
      ready: '已连接',
      error: '未连接',
    }[kind];
    el('bridge-status').textContent = text;
  }
  function connectionNotice(kind, text, hint, retryLabel, retry) {
    connectionCancel = null;
    el('cancel-connection').hidden = true;
    el('connection-notice').hidden = !text;
    el('connection-notice').setAttribute('data-state', kind);
    el('connection-text').textContent = text || '';
    el('connection-hint').textContent = hint || '';
    el('retry-connection').hidden = !retry;
    el('retry-connection').textContent = retryLabel || '重试';
    connectionRetry = retry || null;
  }
  async function prepareDoubaoConnection(restartToken) {
    var response = await fetch('/api/v1/agents/doubao/connect', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-OfficeBridge-Token': token,
      },
      body: JSON.stringify({ restartToken: restartToken }),
    });
    var result = await response.json();
    if (!response.ok) {
      var failure = new Error(result.message || '连接豆包失败，请重试。');
      failure.connectionChanged = result.changed === true;
      throw failure;
    }
    if (result.status === 'confirmation_required') {
      var error = new Error(result.message);
      error.doubaoConfirmation = result;
      throw error;
    }
    if (result.status !== 'ready')
      throw new Error('无法确认豆包连接状态，请重试。');
    return result;
  }
  function showDoubaoConfirmation(error, retry) {
    var confirmation = error.doubaoConfirmation;
    if (!confirmation) return false;
    el('agent-selector').value = selectedAgent;
    setAgentStatus(
      ready ? 'ready' : 'error',
      agentName() + (ready ? ' 已连接' : ' 未连接'),
    );
    connectionNotice(
      'confirm',
      confirmation.message,
      confirmation.hint,
      '重启并连接',
      function () {
        retry(confirmation.restartToken);
      },
    );
    el('cancel-connection').hidden = false;
    connectionCancel = function () {
      if (ready) connected();
      else
        connectionNotice(
          'error',
          '已取消连接豆包',
          '豆包保持运行。',
          '重新连接',
          function () {
            retry();
          },
        );
    };
    return true;
  }
  function connected() {
    setAgentStatus('ready', agentName() + ' 已连接');
    connectionNotice('ready', '');
  }
  function scrollBottom() {
    feed.scrollTop = feed.scrollHeight;
  }
  function state(message, text, kind) {
    message.state.textContent = text;
    message.state.className = 'message-state' + (kind ? ' ' + kind : '');
  }
  function addMessage(role, text, target) {
    empty.hidden = true;
    var node = document.createElement('article');
    node.className = 'message ' + role;
    var body = document.createElement('div');
    body.className = 'message-body';
    body.textContent = text;
    node.appendChild(body);
    if (target) {
      var label = document.createElement('div');
      label.className = 'message-scope';
      label.textContent = target;
      node.appendChild(label);
    }
    var status = document.createElement('div');
    status.className = 'message-state';
    status.setAttribute('role', 'status');
    node.appendChild(status);
    feed.appendChild(node);
    scrollBottom();
    return { node: node, body: body, state: status, scope: label };
  }
  function controls() {
    var busy =
      controller !== null ||
      applying ||
      opening ||
      checking ||
      !!queuedContinuation;
    el('analyze').disabled =
      busy ||
      !ready ||
      !token ||
      !conversationId ||
      !host.available() ||
      !canCapture ||
      !prompt.value.trim();
    el('analyze').hidden =
      controller !== null || applying || !!queuedContinuation;
    el('cancel-analysis').hidden =
      controller === null && !applying && !queuedContinuation;
    prompt.disabled = false;
    el('analyze').title =
      opening || checking
        ? el('connection-text').textContent
        : !ready
          ? '连接 Agent 后发送'
          : '发送（Enter）';
    el('retry-connection').disabled = busy;
    el('cancel-connection').disabled = busy;
    el('analysis-form').setAttribute('aria-busy', String(busy));
    el('check-agent').disabled = busy || checking;
    el('new-chat').disabled = busy || !token;
    el('agent-selector').disabled = busy || !token;
    el('undo-changes').disabled = busy || !host.canUndo();
    el('recovery-record').disabled = busy || !host.recovery();
    if (pending)
      pending.actions.querySelectorAll('button').forEach(function (button) {
        button.disabled = busy;
      });
  }
  function refreshRange() {
    if (controller || applying) return;
    try {
      metadata = host.describe('context');
      canCapture = true;
      var r = metadata.ranges[0];
      el('selection-summary').textContent = metadata.hasSelection
        ? r.sheetName + '!' + r.address
        : '未选择数据 · ' + r.sheetName;
      el('selection-context').title =
        metadata.workbookName +
        ' / ' +
        (metadata.hasSelection
          ? r.sheetName +
            '!' +
            r.address +
            ' · ' +
            metadata.cellCount +
            ' 个单元格，随消息提供'
          : '由 AI 按需读取');
    } catch (error) {
      metadata = null;
      canCapture = false;
      el('selection-summary').textContent = host.available()
        ? '请打开工作表'
        : '请在 WPS 中打开';
      el('selection-context').title = error.message;
    }
    el('selection-context').setAttribute('data-invalid', String(!canCapture));
    controls();
  }
  async function check(restartToken) {
    if (controller || applying || opening || checking) return;
    if (conversationLost) {
      markConversationLost();
      return;
    }
    checking = true;
    ready = false;
    token = '';
    var text = restartToken
      ? '正在重启并连接豆包…'
      : '正在连接 ' + agentName() + '…';
    setAgentStatus('connecting', text);
    connectionNotice('working', text, '可先输入，连接完成后发送。');
    controls();
    try {
      var session = await fetch('/api/v1/session');
      if (!session.ok)
        throw new Error('本地服务未连接，请启动 OfficeBridge 后重新连接。');
      token = (await session.json()).token;
      if (selectedAgent === 'doubao') {
        var setup = await prepareDoubaoConnection(restartToken);
        if (setup.changed && conversationId) {
          markConversationLost(
            '豆包已重新启动，原会话连接已失效。请开启新对话，输入内容会保留。',
          );
          return;
        }
      }
      var response = await fetch('/api/v1/agents');
      if (!response.ok) throw new Error('无法检测 Agent，请重试。');
      var agent = (await response.json()).agents.filter(function (item) {
        return item.id === selectedAgent;
      })[0];
      if (!agent || agent.status !== 'ready')
        throw new Error((agent && agent.reason) || agentName() + ' 暂不可用');
      if (!conversationId) await openConversation();
      ready = true;
      connected();
    } catch (error) {
      ready = false;
      if (showDoubaoConfirmation(error, check)) return;
      if (error.connectionChanged && conversationId) {
        markConversationLost(
          error.message + ' 原会话连接已失效，请开启新对话。',
        );
        return;
      }
      setAgentStatus('error', error.message);
      connectionNotice(
        'error',
        '连接 ' + agentName() + ' 失败',
        error.message,
        '重试',
        check,
      );
    } finally {
      checking = false;
      controls();
    }
  }
  async function openConversation(agentId) {
    agentId = agentId || selectedAgent;
    var response = await fetch('/api/v1/conversations', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'X-OfficeBridge-Token': token,
      },
      body: JSON.stringify({
        id: conversationId || undefined,
        agentId: agentId,
      }),
    });
    var result = await response.json();
    if (!response.ok || !result.id)
      throw new Error(result.message || '无法创建 Agent 会话。');
    conversationId = result.id;
    selectedAgent = result.agentId || agentId;
    conversationLost = false;
    el('agent-selector').value = selectedAgent;
  }
  function expirePending(reason) {
    if (!pending) return;
    if (reason) revokeContinuation(pending);
    if (reason) {
      state(pending.message, reason);
      operationFeedback = '方案 ' + pending.plan.snapshotId + '：' + reason;
    }
    pending.actions.remove();
    pending = null;
  }
  function formatValue(value) {
    return value === null ? '空白' : String(value);
  }
  function addDetail(container, text) {
    var item = document.createElement('div');
    item.className = 'change-item';
    item.textContent = text;
    container.appendChild(item);
  }
  function sheetActionDescription(action, range, execution) {
    if (action.type === 'rename')
      return (
        '工作表「' +
        range.sheetName +
        '」 → 「' +
        action.name +
        '」\n保留内容，原生更新引用。'
      );
    if (action.type === 'filter')
      return (
        range.sheetName +
        '!' +
        range.address +
        '\n' +
        (!action.conditions.length
          ? '清除当前范围的筛选。'
          : '首行为表头，以下条件同时满足：\n' +
            action.conditions
              .map(function (condition) {
                return (
                  '选区第 ' +
                  (condition.column + 1) +
                  ' 列 ' +
                  {
                    eq: '等于',
                    ne: '不等于',
                    gt: '大于',
                    gte: '大于等于',
                    lt: '小于',
                    lte: '小于等于',
                    contains: '包含',
                  }[condition.operator] +
                  ' ' +
                  formatValue(condition.value)
                );
              })
              .join('\n')) +
        '\n仅改变可见记录，不删除或改写数据。'
      );
    var summaryText = '';
    if (action.aggregation === 'sum') {
      summaryText =
        '\n按分类求和，分类首尾空白与数字文本在计算时处理。\n本次汇总快照，源数据变化后请重新生成。';
      if (range.values) {
        var summary = execution.summary;
        summaryText +=
          '\n' +
          summary.sourceRows +
          ' 条明细 → ' +
          summary.categories.length +
          ' 个分类；转换 ' +
          summary.convertedTextCells +
          ' 个数字文本；略过 ' +
          summary.ignoredBlankRows +
          ' 个空行。\n' +
          summary.categories
            .map(function (category, index) {
              return (
                category +
                '：' +
                summary.series
                  .map(function (series) {
                    return series.name + ' ' + series.values[index];
                  })
                  .join('；')
              );
            })
            .join('\n');
      }
    }
    return (
      '在「' +
      range.sheetName +
      '」新增' +
      { line: '折线图', column: '柱形图', bar: '条形图', pie: '饼图' }[
        action.chartType
      ] +
      '：' +
      action.title +
      '\n数据来源：' +
      range.address +
      '（首行为表头）\n分类：选区第 ' +
      (action.categoryColumn + 1) +
      ' 列；数值：选区第 ' +
      action.valueColumns
        .map(function (column) {
          return column + 1;
        })
        .join('、') +
      ' 列。原数据保留。' +
      summaryText
    );
  }
  function renderPlan(plan, snapshot, message, continuation, execution) {
    if (
      !execution ||
      execution.valid !== true ||
      execution.planKey !== JSON.stringify(plan) ||
      execution.snapshotId !== snapshot.snapshotId
    )
      throw new Error('修改方案无效，未执行任何修改。请重新发送。');
    var columnChanges = plan.columnChanges || [];
    var rowChanges = plan.rowChanges || [],
      sheetActions = plan.sheetActions || [];
    if (
      !plan.edits.length &&
      !plan.newSheets.length &&
      !columnChanges.length &&
      !rowChanges.length &&
      !sheetActions.length
    )
      return;
    var details = document.createElement('details');
    details.className = 'change-details';
    var summary = document.createElement('summary');
    var summaries = [];
    if (plan.edits.length) summaries.push(plan.edits.length + ' 处修改');
    if (plan.newSheets.length)
      summaries.push('新增 ' + plan.newSheets.length + ' 张表');
    columnChanges.forEach(function (change) {
      summaries.push(
        (change.action === 'insert' ? '新增 ' : '删除 ') + change.count + ' 列',
      );
    });
    rowChanges.forEach(function (change) {
      summaries.push(
        (change.action === 'insert' ? '新增 ' : '删除 ') + change.count + ' 行',
      );
    });
    sheetActions.forEach(function (action) {
      summaries.push(
        {
          rename: '重命名工作表',
          filter:
            action.conditions && !action.conditions.length
              ? '清除筛选'
              : '筛选记录',
          chart: '新增图表',
        }[action.type],
      );
    });
    summary.textContent = summaries.join('，');
    details.appendChild(summary);
    var list = document.createElement('div');
    list.className = 'change-list';
    details.appendChild(list);
    var targets = {};
    plan.edits
      .concat(
        plan.columnChanges || [],
        plan.rowChanges || [],
        plan.sheetActions || [],
      )
      .forEach(function (e) {
        targets[e.rangeId] = true;
      });
    snapshot.ranges
      .filter(function (r) {
        return targets[r.id];
      })
      .forEach(function (r) {
        var cells = plan.edits.filter(function (e) {
          return e.rangeId === r.id;
        });
        addDetail(
          list,
          '修改目标：' +
            r.sheetName +
            '!' +
            (cells.length
              ? cells
                  .slice(0, 20)
                  .map(function (e) {
                    return C.cellAddress(
                      r.startRow + e.row,
                      r.startColumn + e.column,
                    );
                  })
                  .join('、') +
                (cells.length > 20 ? ' 等 ' + cells.length + ' 格' : '')
              : r.address),
        );
      });
    if (
      /\.csv$/i.test(snapshot.workbookName) &&
      sheetActions.some(function (action) {
        return action.type === 'chart';
      })
    )
      addDetail(
        list,
        'CSV 无法保存图表。应用后请另存为 .xlsx 文件以保留图表。',
      );
    rowChanges.forEach(function (change) {
      var range = snapshot.ranges.filter(function (item) {
          return item.id === change.rangeId;
        })[0],
        start = range.startRow + change.row;
      addDetail(
        list,
        range.sheetName +
          ' · ' +
          (change.action === 'insert'
            ? '在第 ' + start + ' 行前新增 ' + change.count + ' 行'
            : '删除第 ' + start + '–' + (start + change.count - 1) + ' 行') +
          '\n整行操作，影响这些行的所有列。可恢复本次修改。',
      );
      if (change.values)
        addDetail(
          list,
          '新增行内容：\n' +
            change.values
              .slice(0, 10)
              .map(function (row) {
                return row.map(formatValue).join(' | ');
              })
              .join('\n') +
            (change.values.length > 10 ? '\n仅预览前 10 行' : ''),
        );
    });
    sheetActions.forEach(function (action) {
      var range = snapshot.ranges.filter(function (item) {
        return item.id === action.rangeId;
      })[0];
      addDetail(list, sheetActionDescription(action, range, execution));
    });
    columnChanges.forEach(function (change) {
      var range = snapshot.ranges.filter(function (item) {
        return item.id === change.rangeId;
      })[0];
      var column = range.startColumn + change.column;
      var first = C.cellAddress(1, column).replace(/1$/, ''),
        last = C.cellAddress(1, column + change.count - 1).replace(/1$/, '');
      addDetail(
        list,
        range.sheetName +
          ' · ' +
          (change.action === 'insert'
            ? '在 ' + first + ' 列前新增 ' + change.count + ' 列'
            : '删除 ' + first + (first !== last ? '–' + last : '') + ' 列') +
          '\n整列操作，影响工作表所有行。' +
          (change.action === 'delete'
            ? '这些列的数据将被移除，可恢复本次修改。'
            : ''),
      );
      if (change.values)
        addDetail(
          list,
          '新增列内容（从第 ' +
            range.startRow +
            ' 行开始）：\n' +
            change.values
              .slice(0, 10)
              .map(function (row) {
                return row.map(formatValue).join(' | ');
              })
              .join('\n') +
            (change.values.length > 10 ? '\n仅预览前 10 行' : ''),
        );
    });
    plan.edits.slice(0, 100).forEach(function (edit) {
      var range = snapshot.ranges.filter(function (item) {
        return item.id === edit.rangeId;
      })[0];
      var before =
        range.formulas[edit.row][edit.column] ||
        formatValue(range.values[edit.row][edit.column]);
      var after = Object.prototype.hasOwnProperty.call(edit, 'formula')
        ? edit.formula
        : Object.prototype.hasOwnProperty.call(edit, 'value')
          ? formatValue(edit.value)
          : before;
      var formats = [];
      if (edit.format) {
        if (edit.format.numberFormat !== undefined)
          formats.push('数字格式：' + edit.format.numberFormat);
        if (edit.format.bold !== undefined)
          formats.push(edit.format.bold ? '加粗' : '取消加粗');
        if (edit.format.fillColor !== undefined) formats.push('设置填充颜色');
      }
      addDetail(
        list,
        range.sheetName +
          '!' +
          C.cellAddress(
            range.startRow + edit.row,
            range.startColumn + edit.column,
          ) +
          '\n' +
          before +
          ' → ' +
          after +
          (formats.length ? '\n' + formats.join('；') : ''),
      );
    });
    if (plan.edits.length > 100)
      addDetail(
        list,
        '仅展示前 100 项，应用包含全部 ' + plan.edits.length + ' 项。',
      );
    plan.newSheets.forEach(function (sheet) {
      addDetail(
        list,
        '新增「' +
          sheet.name +
          '」· ' +
          sheet.values.length +
          ' 行 × ' +
          sheet.values[0].length +
          ' 列\n' +
          sheet.values
            .slice(0, 10)
            .map(function (row) {
              return row.map(formatValue).join(' | ');
            })
            .join('\n') +
          (sheet.values.length > 10 ? '\n仅预览前 10 行' : ''),
      );
    });
    message.node.appendChild(details);
    var actions = document.createElement('div');
    actions.className = 'message-actions';
    var apply = document.createElement('button');
    apply.type = 'button';
    apply.className = 'primary';
    apply.textContent = '应用';
    apply.setAttribute('data-action', 'apply');
    var discard = document.createElement('button');
    discard.type = 'button';
    discard.textContent = '放弃';
    discard.setAttribute('data-action', 'discard');
    actions.appendChild(apply);
    actions.appendChild(discard);
    message.node.appendChild(actions);
    var record = {
      plan: plan,
      message: message,
      actions: actions,
      continuation: C.isResumeOffer(continuation) ? continuation : null,
      conversationId: conversationId,
      agentId: selectedAgent,
      stopped: false,
    };
    pending = record;
    apply.addEventListener('click', async function () {
      if (pending !== record || controller || applying) return;
      applying = true;
      applyController = new AbortController();
      controls();
      state(message, '正在应用…', 'working');
      await new Promise(function (resolve) {
        setTimeout(resolve, 0);
      });
      var applied = false,
        result;
      try {
        if (record.stopped) throw new Error('已停止，未应用修改。');
        result = await OfficeBridgeExecutor.apply(record.plan, {
          signal: applyController.signal,
          onProgress: function (done, total) {
            state(
              message,
              '正在核对图表数据… ' + done + ' / ' + total,
              'working',
            );
          },
          onStage: function (text) {
            state(message, text, 'working');
          },
        });
        applied = true;
        state(
          message,
          (result.columnChanges && result.columnChanges.length
            ? '已完成整列操作'
            : result.rowChanges && result.rowChanges.length
              ? '已完成整行操作'
              : result.sheetActions && result.sheetActions.length
                ? {
                    rename: '已重命名工作表',
                    filter: '已更新筛选',
                    chart: '已新增图表',
                  }[result.sheetActions[0].type]
                : '已应用 ' + result.changedCells + ' 处修改') +
            (result.newSheets.length
              ? '，新增「' + result.newSheets.join('、') + '」'
              : '') +
            (result.columnChanges || [])
              .map(function (change) {
                return (
                  '，' +
                  (change.action === 'insert' ? '新增 ' : '删除 ') +
                  change.count +
                  ' 列'
                );
              })
              .join('') +
            (result.rowChanges || [])
              .map(function (change) {
                return (
                  '，' +
                  (change.action === 'insert' ? '新增 ' : '删除 ') +
                  change.count +
                  ' 行'
                );
              })
              .join('') +
            ' · 未保存',
        );
        lastApplied = message;
        operationFeedback =
          '方案 ' + record.plan.snapshotId + '：' + message.state.textContent;
        record.appliedStatus = message.state.textContent;
        if (
          record.plan.newSheets.length ||
          (record.plan.columnChanges || []).length ||
          (record.plan.rowChanges || []).length ||
          (record.plan.sheetActions || []).some(function (a) {
            return a.type === 'rename';
          })
        )
          lastRangeContext = null;
      } catch (error) {
        revokeContinuation(record);
        state(message, error.message, 'error');
        operationFeedback =
          '方案 ' + record.plan.snapshotId + ' 应用失败：' + error.message;
      } finally {
        expirePending();
        applying = false;
        applyController = null;
        refreshRange();
        controls();
        scrollBottom();
      }
      if (applied && !record.stopped && record.continuation) {
        if (!record.continuation.token) {
          state(
            message,
            message.state.textContent + ' · ' + record.continuation.reason,
          );
          return;
        }
        queuedContinuation = record;
        record.application = {
          status: 'succeeded',
          snapshotId: plan.snapshotId,
          workbookFullName: plan.workbookFullName,
          result: result,
        };
        state(
          message,
          message.state.textContent + ' · 正在继续后续步骤…',
          'working',
        );
        controls();
        await paintBeforeWork();
        if (queuedContinuation !== record || record.stopped) return;
        queuedContinuation = null;
        await send(null, record);
      }
    });
    discard.addEventListener('click', function () {
      if (pending !== record) return;
      expirePending('已放弃，表格未修改。');
      controls();
    });
    state(
      message,
      record.continuation && record.continuation.token
        ? '等待应用。应用成功后继续：' + record.continuation.nextStep
        : '等待应用，表格未修改。',
    );
    controls();
    scrollBottom();
  }
  async function send(event, automatic) {
    if (event) event.preventDefault();
    if (
      controller ||
      applying ||
      opening ||
      checking ||
      queuedContinuation ||
      !ready ||
      !token ||
      !conversationId ||
      (!automatic && !prompt.value.trim())
    )
      return;
    var text = automatic ? '继续已应用方案后的剩余步骤。' : prompt.value.trim(),
      requestedScope = 'context',
      snapshot;
    if (questionRecord) {
      questionRecord.node.remove();
      questionRecord = null;
    }
    expirePending('已开始新任务，上一个方案未应用。');
    var userMessage = automatic
      ? null
      : addMessage('user', text, '正在确认范围…');
    var message = addMessage('assistant', '');
    if (automatic) state(automatic.message, automatic.appliedStatus);
    state(message, '正在读取表格，准备本次任务…', 'working');
    controller = new AbortController();
    var active = controller,
      requestStarted = false,
      proposal = null,
      terminal = '',
      failureMessage = '',
      hostFailure = '',
      continuation = null;
    if (!automatic) prompt.value = '';
    resize();
    var completionNotice = '',
      question = null,
      references = [],
      rangeMap = {};
    var proposalSnapshot = null,
      hostQueue = Promise.resolve(),
      timings = {},
      submittedAt = 0;
    function showTimings() {
      message.state.title = Object.keys(timings)
        .map(function (k) {
          return k + '：' + timings[k] + ' ms';
        })
        .join('；');
    }
    async function hostRequest(item) {
      if (active.signal.aborted) return;
      var body = {
        runId: item.runId,
        callId: item.callId,
        snapshotId: item.snapshotId,
      };
      var toolFailure = '';
      try {
        if (item.snapshotId !== snapshot.snapshotId)
          throw new Error('宿主请求不属于本次任务。');
        state(message, '正在按需读取或计算表格…', 'working');
        await paintBeforeWork();
        if (active.signal.aborted) return;
        body.reply = await OfficeBridgeExecutor.query(
          item.snapshotId,
          item.name,
          item.args,
          { signal: active.signal },
        );
        if (body.reply.range) rangeMap[body.reply.range.id] = body.reply.range;
        if (item.name === 'officebridge_ask_user')
          question = C.validateQuestion(body.reply.result);
        if (
          [
            'officebridge_read_range',
            'officebridge_summarize_range',
            'officebridge_inspect_range',
          ].indexOf(item.name) >= 0
        ) {
          var r = rangeMap[item.args.rangeId];
          if (r) {
            var o = C.readOptions(
              r,
              item.name === 'officebridge_read_range'
                ? item.args
                : {
                    columns:
                      item.name === 'officebridge_summarize_range'
                        ? C.summaryColumns(r, item.args).columns
                        : undefined,
                  },
            );
            references.push(
              r.sheetName +
                '!' +
                r.address +
                ' · 第 ' +
                (r.startRow + o.rowOffset) +
                '–' +
                (r.startRow + o.rowOffset + o.rowCount - 1) +
                ' 行；' +
                o.columns
                  .slice(0, 8)
                  .map(function (c) {
                    return C.cellAddress(1, r.startColumn + c).replace(
                      /1$/,
                      '',
                    );
                  })
                  .join('、') +
                (o.columns.length > 8
                  ? ' 等 ' + o.columns.length + ' 列'
                  : ' 列') +
                '；' +
                o.fields
                  .map(function (f) {
                    return { values: '值', formulas: '公式', styles: '格式' }[
                      f
                    ];
                  })
                  .join('、'),
            );
          }
        }
        await paintBeforeWork();
      } catch (error) {
        body.error = error.message;
        toolFailure =
          (item.name === 'officebridge_create_chart'
            ? '创建图表失败：'
            : '表格操作失败：') + error.message;
      }
      if (active.signal.aborted) return;
      var response = await fetch('/api/v1/host-result', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-OfficeBridge-Token': token,
        },
        body: JSON.stringify(body),
        signal: active.signal,
      });
      if (!response.ok && response.status !== 410)
        throw new Error('宿主结果发送失败。');
      // Host query failures invalidate this task. Report the first cause and
      // cancel native generation instead of paying for retries on an unusable snapshot.
      if (toolFailure)
        throw new Error(toolFailure + ' 本轮已停止，未应用修改。');
    }
    controls();
    function receive(line) {
      if (!line.trim()) return;
      var item = JSON.parse(line);
      var follow = feed.scrollHeight - feed.scrollTop - feed.clientHeight < 80;
      if (item.type === 'text') {
        if (timings['Agent首条文本'] === undefined) {
          timings['Agent首条文本'] = Date.now() - submittedAt;
          showTimings();
        }
        message.body.textContent += item.text;
      } else if (item.type === 'host_request') {
        hostQueue = hostQueue
          .then(function () {
            return hostRequest(item);
          })
          .catch(function (error) {
            if (!active.signal.aborted) {
              hostFailure = error.message;
              active.abort();
            }
          });
      } else if (item.type === 'activity')
        state(
          message,
          agentName() +
            (item.phase === 'thinking' ? ' 正在思考…' : ' 正在准备表格操作…'),
          'working',
        );
      else if (item.type === 'status') state(message, item.message, 'working');
      else if (item.type === 'proposal') {
        proposal = item.plan;
        proposalSnapshot = item.snapshot;
        continuation = item.continuation;
      } else if (item.type === 'complete') {
        terminal = 'complete';
        completionNotice = item.notice || '';
      } else if (item.type === 'error') {
        terminal = 'error';
        failureMessage = item.message;
        if (item.code === 'conversation_lost') markConversationLost();
      }
      if (follow) scrollBottom();
    }
    try {
      // Pin the click's target before yielding. A changed selection must not
      // silently redirect the user's request during the first paint.
      var target = host.describe(requestedScope);
      if (
        automatic &&
        (automatic.stopped ||
          automatic.conversationId !== conversationId ||
          automatic.agentId !== selectedAgent ||
          automatic.plan.workbookFullName !== target.workbookFullName)
      )
        throw new Error('文件或会话已变化，已暂停后续步骤。');
      if (userMessage) {
        userMessage.scope.textContent = target.hasSelection
          ? target.ranges
              .map(function (r) {
                return r.sheetName + '!' + r.address;
              })
              .join('、')
          : '未附带选区 · ' + target.ranges[0].sheetName;
        userMessage.scope.title = '随消息提供的数据；AI 会判断是否需要补充';
      }
      await paintBeforeWork();
      if (active.signal.aborted) throw new Error('任务已停止。');
      var readStart = Date.now();
      snapshot = await OfficeBridgeExecutor.beginRead(requestedScope, target, {
        prompt: text,
        previous: lastRangeContext,
        omitSelection: !!automatic,
        signal: active.signal,
      });
      if (userMessage && !snapshot.navigation.selectionProvided)
        userMessage.scope.textContent =
          '未附带选区数据 · ' + target.ranges[0].sheetName;
      snapshot.ranges.forEach(function (r) {
        rangeMap[r.id] = r;
      });
      timings['本地描述准备'] = Date.now() - readStart;
      showTimings();
      // Process queued stop clicks before any data goes to the Agent.
      await paintBeforeWork();
      if (active.signal.aborted) throw new Error('任务已停止。');
      state(message, '正在发送给 ' + agentName() + '…', 'working');
      requestStarted = true;
      submittedAt = Date.now();
      var response = await fetch('/api/v1/run', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-OfficeBridge-Token': token,
        },
        body: JSON.stringify({
          prompt: text,
          workbook: snapshot,
          conversationId: conversationId,
          operationFeedback: operationFeedback.slice(0, 4000),
          ...(automatic
            ? {
                resume: {
                  token: automatic.continuation.token,
                  application: automatic.application,
                },
              }
            : {}),
        }),
        signal: active.signal,
      });
      if (!response.ok) {
        var failure = await response.json();
        if (failure.error === 'conversation_lost') markConversationLost();
        throw new Error(failure.message || '请求失败，请重试。');
      }
      timings['本地服务接收'] = Date.now() - submittedAt;
      showTimings();
      state(message, '已发送，等待 ' + agentName() + ' 回复…', 'working');
      if (response.body && response.body.getReader) {
        var reader = response.body.getReader(),
          decoder = new TextDecoder(),
          buffer = '';
        while (true) {
          var chunk = await reader.read();
          if (chunk.done) {
            buffer += decoder.decode();
            break;
          }
          buffer += decoder.decode(chunk.value, { stream: true });
          var lines = buffer.split('\n');
          buffer = lines.pop();
          lines.forEach(receive);
        }
        if (buffer.trim()) receive(buffer);
      } else
        throw new Error('当前 WPS 浏览器不支持按需数据流，请更新 WPS 后重试。');
      if (active.signal.aborted) throw new Error('任务已停止。');
      if (terminal !== 'complete')
        throw new Error(failureMessage || '未收到完整结果，请重试。');
      state(message, completionNotice);
      await hostQueue;
      if (active.signal.aborted) throw new Error(hostFailure || '任务已停止。');
      if (question) {
        proposal = null;
        message.body.textContent = question.question;
        var choices = document.createElement('div');
        choices.className = 'message-actions scope-choices';
        var record = { node: choices };
        questionRecord = record;
        question.options.forEach(function (option) {
          var button = document.createElement('button');
          button.type = 'button';
          button.textContent = option;
          button.addEventListener('click', async function () {
            if (controller || applying || questionRecord !== record) return;
            var draft = prompt.value;
            prompt.value = option;
            resize();
            await send();
            if (draft && !prompt.value) {
              prompt.value = draft;
              resize();
              controls();
            }
          });
          choices.appendChild(button);
        });
        message.node.appendChild(choices);
        state(message, '请选择，也可以直接输入补充说明。');
      } else if (proposal) {
        var execution = await OfficeBridgeExecutor.prepare(
          proposalSnapshot,
          proposal,
          {
            signal: active.signal,
            onProgress: function (text) {
              state(message, text, 'working');
            },
          },
        );
        if (active.signal.aborted) throw new Error('任务已停止。');
        renderPlan(
          proposal,
          proposalSnapshot,
          message,
          continuation,
          execution,
        );
      }
      if (!question && !references.length && snapshot.inline)
        snapshot.inline.forEach(function (item) {
          var r = rangeMap[item.rangeId];
          if (r)
            references.push(
              r.sheetName + '!' + r.address + ' · 随问题提供的选区值',
            );
        });
      if (references.length) {
        var dataDetails = document.createElement('details');
        dataDetails.className = 'change-details data-sources';
        var dataSummary = document.createElement('summary');
        dataSummary.textContent = '参考数据范围';
        dataDetails.appendChild(dataSummary);
        var sources = document.createElement('div');
        sources.className = 'change-list';
        Array.from(new Set(references)).forEach(function (t) {
          addDetail(sources, t);
        });
        dataDetails.appendChild(sources);
        message.node.appendChild(dataDetails);
      }
      if (!message.body.textContent && !proposal)
        message.body.textContent = '已完成。';
    } catch (error) {
      if (automatic) revokeContinuation(automatic);
      var stopped = active.signal.aborted && !hostFailure;
      if (!stopped) active.abort();
      expirePending('本轮已停止或失败，未应用方案。');
      if (!requestStarted && userMessage) {
        state(userMessage, '未发送', stopped ? '' : 'error');
        if (!prompt.value) prompt.value = text;
        resize();
      }
      state(
        message,
        stopped
          ? requestStarted
            ? '已停止，未应用修改。'
            : '已取消读取，未发送给 Agent。'
          : (hostFailure || error.message) +
              (automatic ? ' 自动续办已暂停，上一步已应用的结果保留。' : ''),
        stopped ? '' : 'error',
      );
    } finally {
      if (snapshot && terminal === 'complete' && !active.signal.aborted)
        lastRangeContext = host.navigationContext() || lastRangeContext;
      if (snapshot)
        host.finishRead(
          snapshot.snapshotId,
          terminal === 'complete' && !!proposal && !active.signal.aborted,
        );
      controller = null;
      refreshRange();
      controls();
      scrollBottom();
      prompt.focus();
    }
  }
  function paintBeforeWork() {
    return new Promise(function (resolve) {
      var finished = false,
        timer = setTimeout(finish, 150);
      function finish() {
        if (finished) return;
        finished = true;
        clearTimeout(timer);
        resolve();
      }
      if (typeof requestAnimationFrame === 'function')
        requestAnimationFrame(function () {
          requestAnimationFrame(finish);
        });
    });
  }
  function resize() {
    prompt.style.height = 'auto';
    prompt.style.height =
      Math.min(160, Math.max(52, prompt.scrollHeight)) + 'px';
    controls();
  }
  el('analysis-form').addEventListener('submit', send);
  prompt.addEventListener('input', resize);
  prompt.addEventListener('compositionstart', function () {
    composing = true;
  });
  prompt.addEventListener('compositionend', function () {
    composing = false;
  });
  prompt.addEventListener('keydown', function (event) {
    if (
      event.key === 'Enter' &&
      !event.shiftKey &&
      !event.isComposing &&
      !composing &&
      event.keyCode !== 229
    ) {
      event.preventDefault();
      if (canCapture) send();
    }
  });
  el('cancel-analysis').addEventListener('click', function () {
    if (controller) controller.abort();
    if (applyController) applyController.abort();
    if (applying && pending) {
      pending.stopped = true;
      revokeContinuation(pending);
    }
    if (queuedContinuation) {
      var queued = queuedContinuation;
      queuedContinuation = null;
      queued.stopped = true;
      revokeContinuation(queued);
      state(queued.message, '已应用，已停止后续步骤。');
      controls();
    }
  });
  el('check-agent').addEventListener('click', function () {
    el('options-menu').open = false;
    check();
  });
  async function startNewConversation(agentId, restartToken) {
    if (controller || applying || opening || checking) return;
    agentId = agentId || selectedAgent;
    var switching = agentId !== selectedAgent;
    var text = restartToken
      ? '正在重启并连接豆包…'
      : switching
        ? '正在切换到 ' + agentName(agentId) + '…'
        : '正在准备 ' + agentName(agentId) + ' 新对话…';
    opening = true;
    setAgentStatus(switching ? 'switching' : 'opening', text, agentId);
    connectionNotice('working', text, '可先输入，准备完成后发送。');
    controls();
    el('options-menu').open = false;
    try {
      if (agentId === 'doubao') await prepareDoubaoConnection(restartToken);
      await openConversation(agentId);
      expirePending('已开启新对话，原方案未应用。');
      operationFeedback = '';
      lastApplied = null;
      lastRangeContext = null;
      if (questionRecord) {
        questionRecord.node.remove();
        questionRecord = null;
      }
      feed.textContent = '';
      empty.hidden = false;
      feed.appendChild(empty);
      ready = true;
      connected();
      resize();
      prompt.focus();
    } catch (error) {
      if (
        error.connectionChanged &&
        selectedAgent === 'doubao' &&
        conversationId
      ) {
        markConversationLost(
          error.message + ' 原会话连接已失效，请开启新对话。',
        );
        return;
      }
      if (
        showDoubaoConfirmation(error, function (confirmation) {
          startNewConversation(agentId, confirmation);
        })
      )
        return;
      el('agent-selector').value = selectedAgent;
      setAgentStatus(
        ready ? 'ready' : 'error',
        agentName() + (ready ? ' 已连接' : ' 会话不可用'),
      );
      connectionNotice(
        'error',
        switching ? '切换到 ' + agentName(agentId) + ' 失败' : '新建对话失败',
        error.message +
          (ready && conversationId
            ? ' 仍在使用 ' + agentName() + '，可以继续发送。'
            : ''),
        switching ? '重试切换' : '重试',
        function () {
          startNewConversation(agentId);
        },
      );
    } finally {
      opening = false;
      refreshRange();
      controls();
    }
  }
  el('new-chat').addEventListener('click', function () {
    startNewConversation(selectedAgent);
  });
  el('agent-selector').addEventListener('change', function () {
    startNewConversation(el('agent-selector').value);
  });
  el('retry-connection').addEventListener('click', function () {
    if (connectionRetry) connectionRetry();
  });
  el('cancel-connection').addEventListener('click', function () {
    if (connectionCancel) connectionCancel();
  });
  el('undo-changes').addEventListener('click', async function () {
    if (controller || applying) return;
    el('options-menu').open = false;
    expirePending('已请求恢复，原方案未应用。');
    var message = addMessage('assistant', '');
    state(message, '正在恢复…', 'working');
    applying = true;
    applyController = new AbortController();
    controls();
    await new Promise(function (resolve) {
      setTimeout(resolve, 0);
    });
    try {
      var result = await OfficeBridgeExecutor.undo({
        signal: applyController.signal,
      });
      message.body.textContent =
        result.sheetActions && result.sheetActions.length
          ? {
              rename: '已恢复工作表名称。',
              filter: '已恢复筛选和记录可见状态。',
              chart: '已移除本次新增的图表。',
            }[result.sheetActions[0].type]
          : result.rowChanges && result.rowChanges.length
            ? '已恢复本次整行操作。'
            : '已恢复 ' +
              result.restoredCells +
              ' 个单元格' +
              (result.removedSheets.length
                ? '，移除「' + result.removedSheets.join('、') + '」'
                : '') +
              '。';
      state(message, '未保存');
      if (lastApplied) state(lastApplied, '已恢复');
      lastApplied = null;
      operationFeedback = message.body.textContent;
      lastRangeContext = null;
    } catch (error) {
      state(message, error.message, 'error');
      operationFeedback = '恢复失败：' + error.message;
    } finally {
      applying = false;
      applyController = null;
      refreshRange();
      controls();
      scrollBottom();
    }
  });
  function markConversationLost(hint) {
    ready = false;
    conversationLost = true;
    setAgentStatus('error', '会话已失效，请开启新对话');
    connectionNotice(
      'error',
      agentName() + ' 会话已失效',
      hint || '请开启新对话后继续，输入内容会保留。',
      '新建对话',
      function () {
        startNewConversation(selectedAgent);
      },
    );
  }
  el('recovery-record').addEventListener('click', function () {
    el('options-menu').open = false;
    var record = host.recovery();
    if (!record) return;
    var message = addMessage(
      'assistant',
      record.incomplete
        ? '上次操作未完整恢复，请核对这些位置：'
        : '上次修改前的内容：',
    );
    var details = document.createElement('details');
    details.className = 'change-details';
    var summary = document.createElement('summary');
    summary.textContent =
      record.columnChanges && record.columnChanges.length
        ? '整列操作恢复记录'
        : record.rowChanges && record.rowChanges.length
          ? '整行操作恢复记录'
          : record.sheetActions && record.sheetActions.length
            ? '工作表操作恢复记录'
            : record.cellCount + ' 个单元格';
    details.appendChild(summary);
    var list = document.createElement('div');
    list.className = 'change-list';
    details.appendChild(list);
    (record.columnChanges || []).forEach(function (change) {
      addDetail(
        list,
        change.sheetName +
          ' · ' +
          (change.action === 'insert' ? '新增' : '删除') +
          ' ' +
          change.count +
          ' 列，起始列 ' +
          C.cellAddress(1, change.column).replace(/1$/, ''),
      );
    });
    (record.rowChanges || []).forEach(function (change) {
      addDetail(
        list,
        change.sheetName +
          ' · ' +
          (change.action === 'insert' ? '新增' : '删除') +
          ' ' +
          change.count +
          ' 行，起始行 ' +
          change.row,
      );
    });
    (record.sheetActions || []).forEach(function (record) {
      addDetail(
        list,
        sheetActionDescription(record.action, {
          sheetName: record.sheetName,
          address: record.range,
        }),
      );
    });
    record.cells.forEach(function (cell) {
      addDetail(
        list,
        cell.sheetName +
          '!' +
          cell.address +
          '\n' +
          (cell.before.formula || formatValue(cell.before.value)),
      );
    });
    if (record.newSheets.length)
      addDetail(list, '新增工作表：' + record.newSheets.join('、'));
    if (record.cellCount > 100) addDetail(list, '仅展示前 100 个位置。');
    message.node.appendChild(details);
    scrollBottom();
  });
  document.addEventListener('click', function (event) {
    document.querySelectorAll('details.popover[open]').forEach(function (menu) {
      if (!menu.contains(event.target)) menu.open = false;
    });
  });
  document.addEventListener('keydown', function (event) {
    if (event.key === 'Escape')
      document
        .querySelectorAll('details.popover[open]')
        .forEach(function (menu) {
          menu.open = false;
        });
  });
  refreshRange();
  check();
  var timer = setInterval(refreshRange, 800);
  globalThis.addEventListener('beforeunload', function () {
    clearInterval(timer);
    if (controller) controller.abort();
    if (conversationId && token)
      fetch('/api/v1/conversations/close', {
        method: 'POST',
        keepalive: true,
        headers: {
          'Content-Type': 'application/json',
          'X-OfficeBridge-Token': token,
        },
        body: JSON.stringify({ id: conversationId }),
      }).catch(function () {});
  });
})();
