/* global Application, wps, OfficeBridgeContract */
(function (global) {
  'use strict';
  var C = OfficeBridgeContract;
  var connectionId =
    'wps-task-' + Date.now() + '-' + Math.random().toString(36).slice(2);
  var captured = null;
  var history = null;
  var columnBaseline = null;
  var rowBaseline = null;
  var actionBaseline = null;
  var chartBaseline = null;
  var lastChartMetrics = null;
  var onDemand = null;
  var lastReadMetrics = null;
  var own = function (value, key) {
    return Object.prototype.hasOwnProperty.call(value, key);
  };
  function drain(steps) {
    try {
      var step = steps.next();
      while (!step.done)
        step = steps.next(
          step.value && step.value.operation
            ? C.executeComputation(step.value.operation, step.value.payload)
            : undefined,
        );
      return step.value;
    } catch (error) {
      try {
        steps.throw(error);
      } catch (_) {}
      throw error;
    }
  }
  async function schedule(steps, options, snapshotId) {
    options = options || {};
    var pinned = null;
    function guard() {
      if (options.signal && options.signal.aborted)
        throw new Error('任务已停止，未提交修改。');
      if (pinned && onDemand !== pinned)
        throw new Error('执行任务已失效，请重新发送。');
      if (
        snapshotId &&
        (!onDemand || onDemand.descriptor.snapshotId !== snapshotId)
      )
        throw new Error('执行任务已失效，请重新发送。');
      if (pinned && onDemand) currentBook(onDemand.descriptor.workbookFullName);
    }
    var step;
    try {
      guard();
      step = steps.next();
      pinned = onDemand;
      while (!step.done) {
        guard();
        var value;
        if (step.value && step.value.operation)
          value = options.compute
            ? await options.compute(step.value.operation, step.value.payload)
            : C.executeComputation(step.value.operation, step.value.payload);
        else
          await new Promise(function (resolve) {
            setTimeout(resolve, 0);
          });
        guard();
        step = steps.next(value);
      }
      return step.value;
    } catch (error) {
      try {
        steps.throw(error);
      } catch (_) {}
      if (pinned && onDemand === pinned) {
        onDemand.error = true;
        captured = null;
        columnBaseline = rowBaseline = actionBaseline = chartBaseline = null;
      }
      throw error;
    }
  }
  function app() {
    var host =
      typeof Application !== 'undefined'
        ? Application
        : typeof wps !== 'undefined'
          ? wps
          : null;
    if (!host) throw new Error('请从 WPS 的 OfficeBridge 选项卡打开侧栏。');
    return host;
  }
  function identity(book) {
    if (!book) throw new Error('请先打开工作簿。');
    return String(book.FullName || book.Name);
  }
  function bounds(range) {
    if (!range || !range.Rows || !range.Columns)
      throw new Error('请先选中单元格。');
    var result = {
      startRow: Number(range.Row),
      startColumn: Number(range.Column),
      rowCount: Number(range.Rows.Count),
      columnCount: Number(range.Columns.Count),
    };
    ['startRow', 'startColumn', 'rowCount', 'columnCount'].forEach(
      function (key) {
        if (!Number.isSafeInteger(result[key]) || result[key] < 1)
          throw new Error('表格范围无效。');
      },
    );
    return result;
  }
  function specs(scope) {
    var host = app(),
      book = host.ActiveWorkbook,
      sheet = host.ActiveSheet;
    identity(book);
    if (
      ['selection', 'rows', 'columns', 'sheet', 'workbook', 'context'].indexOf(
        scope,
      ) < 0
    )
      throw new Error('处理范围无效。');
    var sheetNames = [],
      ranges = [];
    var worksheets = book.Worksheets;
    if (!worksheets || worksheets.Count > 500)
      throw new Error('当前工作簿的工作表数量不受支持。');
    for (var index = 1; index <= worksheets.Count; index++)
      sheetNames.push(String(worksheets.Item(index).Name));
    function add(target, range) {
      ranges.push(
        Object.assign(
          { id: 'r' + (ranges.length + 1), sheetName: String(target.Name) },
          range,
        ),
      );
    }
    if (scope === 'workbook') {
      if (worksheets.Count > C.TABLE_LIMITS.ranges)
        throw new Error('工作簿超过 50 张工作表，请选择具体工作表。');
      for (var item = 1; item <= worksheets.Count; item++) {
        var current = worksheets.Item(item);
        add(current, bounds(current.UsedRange));
      }
    } else {
      if (!sheet || !sheet.Cells) throw new Error('请切换到数据工作表。');
      var used;
      if (scope === 'sheet') add(sheet, bounds(sheet.UsedRange));
      else {
        var selection = host.Selection;
        var selectionAvailable = !!(
          selection &&
          selection.Rows &&
          selection.Columns
        );
        if (scope === 'context' && !selectionAvailable)
          selection =
            host.ActiveCell && host.ActiveCell.Rows
              ? host.ActiveCell
              : {
                  Row: 1,
                  Column: 1,
                  Rows: { Count: 1 },
                  Columns: { Count: 1 },
                };
        if (selection && selection.Areas && selection.Areas.Count > 1)
          throw new Error('当前支持一块连续区域，请重新选择。');
        var selected = bounds(selection);
        var fullRows = selected.columnCount === Number(sheet.Columns.Count);
        var fullColumns = selected.rowCount === Number(sheet.Rows.Count);
        if (
          scope === 'rows' ||
          (['selection', 'context'].includes(scope) && fullRows)
        ) {
          used = bounds(sheet.UsedRange);
          selected.startColumn = used.startColumn;
          selected.columnCount = used.columnCount;
        }
        if (
          scope === 'columns' ||
          (['selection', 'context'].includes(scope) && fullColumns)
        ) {
          used = used || bounds(sheet.UsedRange);
          selected.startRow = used.startRow;
          selected.rowCount = used.rowCount;
        }
        add(sheet, selected);
      }
    }
    ranges.forEach(function (range) {
      range.address = C.rangeAddress(
        range.startRow,
        range.startColumn,
        range.rowCount,
        range.columnCount,
      );
    });
    return {
      book: book,
      workbookName: String(book.Name),
      workbookFullName: identity(book),
      scope: scope,
      sheetNames: sheetNames,
      ranges: ranges,
      ...(scope === 'context'
        ? { selectionAvailable: selectionAvailable }
        : {}),
    };
  }
  function describe(scope) {
    var data = specs(scope === 'auto' ? 'selection' : scope);
    return {
      workbookName: data.workbookName,
      workbookFullName: data.workbookFullName,
      scope: scope,
      ranges: data.ranges,
      cellCount: data.ranges.reduce(function (total, range) {
        return total + range.rowCount * range.columnCount;
      }, 0),
      ...(scope === 'context' ? { hasSelection: data.selectionAvailable } : {}),
    };
  }
  function style(cell) {
    try {
      var font = cell.Font,
        interior = cell.Interior;
      var result = {
        numberFormat: cell.NumberFormat,
        bold: font.Bold,
        fillColor: interior.Color,
        fillPattern: interior.Pattern,
      };
      return typeof result.numberFormat === 'string' &&
        typeof result.bold === 'boolean' &&
        Number.isInteger(result.fillColor) &&
        Number.isInteger(result.fillPattern)
        ? result
        : null;
    } catch (_error) {
      return null;
    }
  }
  function blank(value) {
    return value === null || value === undefined || value === '';
  }
  function sameLiteralValue(actual, expected) {
    return actual === expected || (blank(actual) && blank(expected));
  }
  function state(cell) {
    var value = cell.Value2;
    if (value === undefined) value = null;
    if (!C.isCellValue(value))
      throw new Error('单元格类型不受支持或内容过大。');
    var hasFormula = cell.HasFormula,
      formula = hasFormula === false ? null : cell.Formula;
    formula =
      hasFormula !== false &&
      typeof formula === 'string' &&
      formula.charAt(0) === '='
        ? formula
        : null;
    // WPS may expose an empty literal as null, undefined or empty text.
    // Keep empty formula results distinct so recovery never erases a formula.
    if (formula === null && blank(value)) value = null;
    return {
      value: value,
      formula: formula,
      style: style(cell),
    };
  }
  function byteBudget(message) {
    var encoder = new TextEncoder(),
      used = 0;
    function charge(bytes) {
      used += bytes;
      if (used > C.TABLE_LIMITS.bytes) throw new Error(message);
    }
    // Count serialized values, omitting separators/property names where they
    // have not been assembled yet. This lower bound never rejects a valid
    // packet; the complete JSON is still checked at the end.
    return {
      add: function (value) {
        charge(encoder.encode(JSON.stringify(value)).length);
      },
      array: function (value) {
        charge(encoder.encode(JSON.stringify(value)).length - 2);
      },
    };
  }
  function captureTarget(data) {
    return JSON.stringify({
      workbookFullName: data.workbookFullName,
      scope: data.scope,
      ranges: data.ranges,
      hasSelection:
        data.selectionAvailable === undefined
          ? data.hasSelection
          : data.selectionAvailable,
    });
  }
  function capture(scope, expected) {
    chartBaseline = null;
    onDemand = null;
    captured = null;
    columnBaseline = rowBaseline = actionBaseline = null;
    var data = specs(scope),
      budget = byteBudget(
        '完整范围超过 512 KB，请缩小范围。未读取或发送部分数据。',
      );
    if (expected && captureTarget(data) !== captureTarget(expected))
      throw new Error('读取前文件或范围已变化，请确认当前选择后重新发送。');
    var observed = Object.create(null);
    var snapshot = {
      host: 'wps',
      connectionId: connectionId,
      snapshotId:
        'snapshot-' + Date.now() + '-' + Math.random().toString(36).slice(2),
      workbookName: data.workbookName,
      workbookFullName: data.workbookFullName,
      scope: scope,
      sheetNames: data.sheetNames,
      capturedAt: new Date().toISOString(),
      ranges: data.ranges.map(function (range) {
        var sheet = data.book.Worksheets.Item(range.sheetName),
          cells = sheet.Cells,
          selected = observed[range.sheetName] || Object.create(null);
        observed[range.sheetName] = selected;
        var values = [],
          formulas = [],
          styles = [];
        for (var row = 0; row < range.rowCount; row++) {
          var rowValues = [],
            rowFormulas = [],
            rowStyles = [];
          for (var column = 0; column < range.columnCount; column++) {
            var current = state(
              cells.Item(range.startRow + row, range.startColumn + column),
            );
            selected[
              range.startRow + row + ':' + (range.startColumn + column)
            ] = current;
            budget.add(current.value);
            budget.add(current.formula);
            budget.add(current.style);
            rowValues.push(current.value);
            rowFormulas.push(current.formula);
            rowStyles.push(current.style);
          }
          values.push(rowValues);
          formulas.push(rowFormulas);
          styles.push(rowStyles);
        }
        return Object.assign({}, range, {
          values: values,
          formulas: formulas,
          styles: styles,
        });
      }),
    };
    if (!C.isWorkbookSnapshot(snapshot))
      throw new Error(
        '完整范围超过 512 KB，或宿主返回了无效数据。请缩小范围。',
      );
    columnBaseline = null;
    try {
      columnBaseline = captureColumnState(
        data.book,
        data.ranges,
        undefined,
        null,
        observed,
      );
      snapshot.columnOperations = { available: true };
    } catch (error) {
      snapshot.columnOperations = { available: false, reason: error.message };
    }
    rowBaseline = null;
    actionBaseline = null;
    try {
      rowBaseline = captureColumnState(
        data.book,
        data.ranges,
        'rows',
        columnBaseline,
        observed,
      );
      snapshot.rowOperations = { available: true };
    } catch (error) {
      snapshot.rowOperations = { available: false, reason: error.message };
    }
    try {
      actionBaseline = captureColumnState(
        data.book,
        data.ranges,
        'action',
        columnBaseline,
        observed,
      );
      snapshot.sheetOperations = { available: true };
    } catch (error) {
      snapshot.sheetOperations = { available: false, reason: error.message };
    }
    if (!C.isWorkbookSnapshot(snapshot)) {
      columnBaseline = null;
      throw new Error('完整范围及操作信息超过 512 KB，请缩小范围。');
    }
    // Keep a private copy; model/browser response objects cannot alter the baseline.
    captured = JSON.parse(JSON.stringify(snapshot));
    return snapshot;
  }
  function positiveMention(text, term) {
    var lower = text.toLowerCase(),
      needle = term.toLowerCase();
    for (
      var at = lower.indexOf(needle);
      at >= 0;
      at = lower.indexOf(needle, at + needle.length)
    ) {
      if (
        /^[a-z0-9_]+$/.test(needle) &&
        (/[a-z0-9_]/.test(lower.charAt(at - 1)) ||
          /[a-z0-9_]/.test(lower.charAt(at + needle.length)))
      )
        continue;
      var prefix = text
        .slice(0, at)
        .split(/[,，。；;\n]/)
        .pop();
      if (
        /(?:不要|禁止|不得|别|排除|除去|不读取|不处理|不涉及|除了)[^,，。；;\n]*$/.test(
          prefix,
        )
      )
        continue;
      if (
        term.length === 1 &&
        !['“' + term + '”', '"' + term + '"', term + '表', '表' + term].some(
          function (s) {
            return text.indexOf(s) >= 0;
          },
        )
      )
        continue;
      return true;
    }
    return false;
  }
  function beginRead(scope, expected, options) {
    return drain(beginReadSteps(scope, expected, options));
  }
  function* beginReadSteps(scope, expected, options) {
    chartBaseline = null;
    options = options || {};
    captured = null;
    columnBaseline = rowBaseline = actionBaseline = null;
    onDemand = null;
    var context = scope === 'context',
      auto = scope === 'auto' || context,
      data = specs(context ? 'context' : auto ? 'selection' : scope);
    if (options.locked) {
      if (
        !expected ||
        expected.workbookFullName !== data.workbookFullName ||
        expected.ranges.length !== 1 ||
        !data.sheetNames.includes(expected.ranges[0].sheetName)
      )
        throw new Error('锁定的文件或工作表已变化，请重新锁定。');
      if (!C.validRange(expected.ranges[0])) throw new Error('锁定范围无效。');
      data.ranges = JSON.parse(JSON.stringify(expected.ranges));
    }
    if (auto && !context) data.scope = 'auto';
    if (expected && captureTarget(data) !== captureTarget(expected))
      throw new Error('读取前文件或范围已变化，请确认当前选择后重新发送。');
    var descriptor = {
      host: 'wps',
      readMode: 'on-demand',
      connectionId: connectionId,
      snapshotId:
        'query-' + Date.now() + '-' + Math.random().toString(36).slice(2),
      workbookName: data.workbookName,
      workbookFullName: data.workbookFullName,
      scope: auto ? 'workbook' : scope,
      capturedAt: new Date().toISOString(),
      sheetNames: data.sheetNames,
      ranges: data.ranges,
      columnOperations: { available: true },
      rowOperations: { available: true },
      sheetOperations: { available: true },
    };
    if (auto || options.locked) {
      var anchor = JSON.parse(JSON.stringify(data.ranges[0])),
        allowed = [anchor.sheetName],
        prompt = String(options.prompt || '');
      if (auto) {
        var all = (
          prompt.match(
            /整个(?:工作簿|文档|文件)|所有工作表|全部工作表|各个工作表/g,
          ) || []
        ).some(function (term) {
          return positiveMention(prompt, term);
        });
        data.sheetNames.forEach(function (name) {
          var named = positiveMention(prompt, name);
          if ((all || named) && allowed.indexOf(name) < 0) allowed.push(name);
        });
        if (context) allowed = data.sheetNames.slice();
      }
      descriptor.navigation = {
        mode: context ? 'context' : auto ? 'auto' : 'locked',
        activeSheet: anchor.sheetName,
        allowedSheets: allowed,
        anchor: anchor,
      };
      var previous = options.previous;
      if (
        auto &&
        previous &&
        previous.workbookFullName === data.workbookFullName &&
        JSON.stringify(previous.anchor) === JSON.stringify(anchor) &&
        C.validRange(previous.range) &&
        data.sheetNames.includes(previous.range.sheetName)
      ) {
        if (allowed.indexOf(previous.range.sheetName) < 0)
          allowed.push(previous.range.sheetName);
        descriptor.navigation.previousRange = previous.range;
      }
    }
    var common = '';
    if (data.book.ProtectStructure) common = '工作簿结构受保护。';
    if (app().Workbooks && app().Workbooks.Count > 1)
      common = '为完整恢复公式引用，请先只保留目标工作簿打开，再发送结构操作。';
    var reason = common;
    try {
      if (common) throw new Error(common);
      data.ranges.forEach(function (r) {
        assertColumnSheet(data.book.Worksheets.Item(r.sheetName));
      });
    } catch (error) {
      reason = error.message;
    }
    if (reason) {
      descriptor.columnOperations = { available: false, reason: reason };
      descriptor.rowOperations = { available: false, reason: reason };
    }
    if (
      common ||
      data.ranges.some(function (r) {
        return data.book.Worksheets.Item(r.sheetName).ProtectContents;
      })
    )
      descriptor.sheetOperations = {
        available: false,
        reason: common || '工作表已受保护。',
      };
    onDemand = {
      descriptor: JSON.parse(JSON.stringify(descriptor)),
      reads: Object.create(null),
      cache: Object.create(null),
      loaded: [],
      task: null,
      error: false,
      calls: 0,
      budget: byteBudget('本轮读取内容超过 512 KB，请缩小范围或开启新任务。'),
      metrics: {
        valueReads: 0,
        formulaReads: 0,
        styleReads: 0,
        baselineCaptures: 0,
      },
      nextRange: 1,
      lastResolved: null,
      question: false,
    };
    var total = data.ranges.reduce(function (n, r) {
      return n + r.rowCount * r.columnCount;
    }, 0);
    if (
      (context && data.selectionAvailable && !options.omitSelection) ||
      (total <= 16 && !context)
    ) {
      descriptor.inline = [];
      for (
        var inlineIndex = 0;
        inlineIndex < data.ranges.length;
        inlineIndex++
      ) {
        var inlineRange = data.ranges[inlineIndex],
          inlineData = yield* queryReadSteps(inlineRange, {
            fields: ['values'],
          });
        descriptor.inline.push({
          rangeId: inlineRange.id,
          values: inlineData.values,
        });
      }
    }
    if (context) {
      if (
        descriptor.inline &&
        !descriptor.inline.some(function (item) {
          return item.values.some(function (row) {
            return row.some(function (v) {
              return v !== null && v !== '';
            });
          });
        })
      )
        delete descriptor.inline;
      descriptor.navigation.selectionProvided = !!descriptor.inline;
      onDemand.descriptor.navigation.selectionProvided =
        descriptor.navigation.selectionProvided;
    }
    if (!C.isWorkbookDescriptor(descriptor)) {
      onDemand = null;
      throw new Error('文件描述无效或超过体积限制。');
    }
    return descriptor;
  }
  function queryCell(range, row, column, field, cells) {
    var key = JSON.stringify([range.sheetName, row, column]),
      record = onDemand.reads[key],
      cell;
    if (record && own(record, field)) return record[field];
    cell = (
      cells ||
      currentBook(onDemand.descriptor.workbookFullName).Worksheets.Item(
        range.sheetName,
      ).Cells
    ).Item(row, column);
    var value;
    if (field === 'values') {
      value = cell.Value2;
      if (value === undefined) value = null;
      if (!C.isCellValue(value))
        throw new Error('单元格类型不受支持或内容过大。');
      onDemand.metrics.valueReads++;
    } else if (field === 'formulas') {
      var has = cell.HasFormula;
      value = has === false ? null : cell.Formula;
      value =
        typeof value === 'string' && value.charAt(0) === '=' ? value : null;
      onDemand.metrics.formulaReads++;
    } else {
      value = style(cell);
      onDemand.metrics.styleReads++;
    }
    if (!record)
      record = onDemand.reads[key] = {
        sheetName: range.sheetName,
        row: row,
        column: column,
      };
    onDemand.budget.add(value);
    record[field] = value;
    return value;
  }
  function queryRead(range, params) {
    return drain(queryReadSteps(range, params));
  }
  function* queryReadSteps(range, params) {
    var o = C.readOptions(range, params),
      budget = byteBudget('读取内容超过 512 KB，请缩小查询范围。'),
      sheet = currentBook(onDemand.descriptor.workbookFullName).Worksheets.Item(
        range.sheetName,
      ),
      cells = sheet.Cells,
      result = {
        id: range.id,
        rangeId: range.id,
        sheetName: range.sheetName,
        address: range.address,
        startRow: range.startRow + o.rowOffset,
        startColumn: range.startColumn,
        rowOffset: o.rowOffset,
        rowCount: o.rowCount,
        columnCount: o.columns.length,
        columnIndices: o.columns,
        sourceRange: C.readSourceRange(range),
      };
    if (o.fields.indexOf('values') >= 0)
      yield* primeQueryValuesSteps(sheet, range, o);
    var slice = Date.now(),
      visited = 0;
    for (var fieldIndex = 0; fieldIndex < o.fields.length; fieldIndex++) {
      var field = o.fields[fieldIndex];
      var rows = [];
      for (var r = o.rowOffset; r < o.rowOffset + o.rowCount; r++) {
        var row = [];
        for (
          var columnIndex = 0;
          columnIndex < o.columns.length;
          columnIndex++
        ) {
          var c = o.columns[columnIndex];
          var value = queryCell(
            range,
            range.startRow + r,
            range.startColumn + c,
            field,
            cells,
          );
          budget.add(value);
          row.push(value);
          if (++visited % 64 === 0 && Date.now() - slice >= 8) {
            yield null;
            slice = Date.now();
          }
        }
        rows.push(row);
      }
      result[field] = rows;
    }
    return result;
  }
  function* primeQueryValuesSteps(sheet, range, options) {
    if (
      onDemand.bulkSupported === false ||
      options.rowCount * options.columns.length <= 16 ||
      typeof sheet.Range !== 'function'
    )
      return;
    var columns = options.columns.slice().sort(function (a, b) {
        return a - b;
      }),
      groups = [];
    columns.forEach(function (c) {
      var last = groups[groups.length - 1];
      if (last && last.length < 256 && last[last.length - 1] + 1 === c)
        last.push(c);
      else groups.push([c]);
    });
    for (var g = 0; g < groups.length; g++) {
      var group = groups[g],
        height = Math.max(1, Math.floor(256 / group.length));
      for (
        var offset = options.rowOffset;
        offset < options.rowOffset + options.rowCount;
        offset += height
      ) {
        var count = Math.min(
            height,
            options.rowOffset + options.rowCount - offset,
          ),
          missing = false;
        for (var rr = 0; rr < count && !missing; rr++)
          for (var cc = 0; cc < group.length; cc++) {
            var key = JSON.stringify([
                range.sheetName,
                range.startRow + offset + rr,
                range.startColumn + group[cc],
              ]),
              old = onDemand.reads[key];
            if (!old || !own(old, 'values')) {
              missing = true;
              break;
            }
          }
        yield null;
        if (!missing) continue;
        var matrix;
        try {
          matrix = sheet.Range(
            C.rangeAddress(
              range.startRow + offset,
              range.startColumn + group[0],
              count,
              group.length,
            ),
          ).Value2;
        } catch (_error) {
          onDemand.bulkSupported = false;
          return;
        }
        if (
          !Array.isArray(matrix) ||
          matrix.length !== count ||
          matrix.some(function (row) {
            return !Array.isArray(row) || row.length !== group.length;
          })
        ) {
          onDemand.bulkSupported = false;
          return;
        }
        onDemand.metrics.bulkReads = (onDemand.metrics.bulkReads || 0) + 1;
        for (var r = 0; r < count; r++)
          for (var c = 0; c < group.length; c++) {
            var row = range.startRow + offset + r,
              column = range.startColumn + group[c],
              id = JSON.stringify([range.sheetName, row, column]),
              record = onDemand.reads[id],
              value = matrix[r][c];
            if (value === undefined) value = null;
            if (!C.isCellValue(value))
              throw new Error('单元格类型不受支持或内容过大。');
            onDemand.metrics.valueReads++;
            if (record && own(record, 'values')) {
              if (!sameLiteralValue(record.values, value))
                throw new Error('读取过程中数据已变化，请重新发送。');
            } else {
              onDemand.budget.add(value);
              if (!record)
                record = onDemand.reads[id] = {
                  sheetName: range.sheetName,
                  row: row,
                  column: column,
                };
              record.values = value;
            }
          }
      }
    }
  }
  function assertQueryStable() {
    return drain(queryStableSteps());
  }
  function* queryStableSteps() {
    var descriptor = onDemand.descriptor,
      book = currentBook(descriptor.workbookFullName),
      names = [],
      collections = Object.create(null),
      worksheets = book.Worksheets,
      count = worksheets.Count;
    for (var n = 1; n <= count; n++)
      names.push(String(worksheets.Item(n).Name));
    if (JSON.stringify(names) !== JSON.stringify(descriptor.sheetNames))
      throw new Error('工作表结构已变化，请重新发送。');
    var keys = Object.keys(onDemand.reads),
      slice = Date.now();
    for (var keyIndex = 0; keyIndex < keys.length; keyIndex++) {
      var key = keys[keyIndex];
      var record = onDemand.reads[key],
        cells =
          collections[record.sheetName] ||
          (collections[record.sheetName] = worksheets.Item(
            record.sheetName,
          ).Cells),
        cell = cells.Item(record.row, record.column);
      if (own(record, 'values')) {
        onDemand.metrics.valueReads++;
        if (!sameLiteralValue(cell.Value2, record.values))
          throw new Error('本轮已读取的数据已变化，请重新发送。');
      }
      if (own(record, 'formulas')) {
        var has = cell.HasFormula,
          formula = has === false ? null : cell.Formula;
        formula =
          typeof formula === 'string' && formula.charAt(0) === '='
            ? formula
            : null;
        onDemand.metrics.formulaReads++;
        if (formula !== record.formulas)
          throw new Error('本轮已读取的公式已变化，请重新发送。');
      }
      if (own(record, 'styles')) {
        onDemand.metrics.styleReads++;
        if (JSON.stringify(style(cell)) !== JSON.stringify(record.styles))
          throw new Error('本轮已读取的格式已变化，请重新发送。');
      }
      if (keyIndex % 64 === 63 && Date.now() - slice >= 8) {
        yield null;
        slice = Date.now();
      }
    }
  }
  function* queryLoadSteps(range, tiny, chart) {
    if (onDemand.descriptor.navigation) onDemand.lastResolved = range;
    if (
      onDemand.loaded.some(function (r) {
        return r.sheetName === range.sheetName && r.id !== range.id;
      })
    )
      throw new Error(
        '同一工作表每轮只支持一个修改区域，请定位完整目标后再修改，或分轮处理。',
      );
    var existing = onDemand.loaded.filter(function (r) {
      return r.id === range.id;
    })[0];
    if (
      existing &&
      ((!tiny &&
        existing.rowCount === range.rowCount &&
        existing.columnCount === range.columnCount) ||
        tiny)
    )
      return;
    var target = Object.assign({}, range);
    if (tiny) {
      target.rowCount = target.columnCount = 1;
      target.address = C.rangeAddress(
        target.startRow,
        target.startColumn,
        1,
        1,
      );
    }
    var data = yield* queryReadSteps(target, {
      fields: chart ? ['values', 'formulas'] : ['values', 'formulas', 'styles'],
    });
    target.values = data.values;
    target.formulas = data.formulas;
    target.styles =
      data.styles ||
      data.values.map(function (row) {
        return row.map(function () {
          return null;
        });
      });
    if (existing) onDemand.loaded[onDemand.loaded.indexOf(existing)] = target;
    else onDemand.loaded.push(target);
  }
  function executeQuery(snapshotId, name, params) {
    return drain(querySteps(snapshotId, name, params));
  }
  function* querySteps(snapshotId, name, params) {
    if (
      !onDemand ||
      onDemand.descriptor.snapshotId !== snapshotId ||
      onDemand.error
    )
      throw new Error('读取任务已失效，请重新发送。');
    var definition = C.tableTools.filter(function (t) {
      return t.name === name;
    })[0];
    if (
      !definition ||
      !params ||
      typeof params !== 'object' ||
      Array.isArray(params) ||
      Object.keys(params).some(function (k) {
        return !own(definition.inputSchema.properties, k);
      }) ||
      definition.inputSchema.required.some(function (k) {
        return !own(params, k);
      })
    )
      throw new Error('工具或参数不受支持。');
    try {
      if (++onDemand.calls > 32)
        throw new Error('本轮宿主工具调用超过 32 次，请拆分任务。');
      yield* queryStableSteps();
      if (name === 'officebridge_resolve_range')
        return resolveQueryRange(params);
      if (name === 'officebridge_ask_user') {
        var question = C.validateQuestion(params);
        onDemand.question = true;
        onDemand.task = null;
        captured = null;
        columnBaseline = rowBaseline = actionBaseline = null;
        return {
          result: {
            awaitingUser: true,
            question: question.question,
            options: question.options,
          },
        };
      }
      var descriptor = onDemand.descriptor,
        range = descriptor.ranges.filter(function (r) {
          return r.id === params.rangeId;
        })[0],
        result,
        key = JSON.stringify([name, params]),
        write = false;
      if (own(onDemand.cache, key))
        return JSON.parse(JSON.stringify({ result: onDemand.cache[key] }));
      if (name === 'officebridge_read_range') {
        if (!range) throw new Error('范围不存在。');
        result = yield* queryReadSteps(range, params);
      } else if (name === 'officebridge_summarize_range') {
        if (!range) throw new Error('范围不存在。');
        var options = C.summaryColumns(range, params),
          data = yield* queryReadSteps(range, {
            columns: options.columns,
            fields: ['values'],
          });
        result = yield {
          operation: 'summary',
          payload: {
            range: range,
            params: params,
            values: data.values,
            columns: options.columns,
          },
        };
      } else if (name === 'officebridge_inspect_range') {
        if (
          !range ||
          (params.headerRow !== undefined &&
            typeof params.headerRow !== 'boolean')
        )
          throw new Error('范围或表头参数无效。');
        var inspected = yield* queryReadSteps(range, { fields: ['values'] });
        result = yield {
          operation: 'inspect',
          payload: {
            range: Object.assign({}, range, { values: inspected.values }),
            headerRow: params.headerRow === undefined ? true : params.headerRow,
          },
        };
      } else {
        if (onDemand.question)
          throw new Error('本轮正在等待用户选择，不能生成修改方案。');
        var ids =
          name === 'officebridge_propose_edits'
            ? Array.isArray(params.edits)
              ? params.edits.map(function (e) {
                  return e.rangeId;
                })
              : []
            : range
              ? [range.id]
              : [];
        if (name === 'officebridge_propose_result_sheet')
          yield* queryLoadSteps(descriptor.ranges[0], true);
        else
          for (var idIndex = 0; idIndex < ids.length; idIndex++) {
            var id = ids[idIndex];
            var selected = descriptor.ranges.filter(function (r) {
              return r.id === id;
            })[0];
            if (!selected) throw new Error('范围不存在。');
            yield* queryLoadSteps(
              selected,
              false,
              name === 'officebridge_create_chart',
            );
          }
        var snapshot = Object.assign({}, descriptor, {
          ranges: onDemand.loaded,
        });
        delete snapshot.readMode;
        yield { operation: 'snapshot', payload: snapshot };
        if (
          name.indexOf('_columns') >= 0 ||
          name.indexOf('_rows') >= 0 ||
          ['officebridge_filter_range', 'officebridge_rename_sheet'].indexOf(
            name,
          ) >= 0
        ) {
          var mode =
            name.indexOf('_rows') >= 0
              ? 'rows'
              : name.indexOf('_columns') >= 0
                ? undefined
                : 'action';
          var baseline = yield* captureColumnSteps(
            currentBook(descriptor.workbookFullName),
            onDemand.loaded,
            mode,
          );
          onDemand.metrics.baselineCaptures++;
          if (mode === 'rows') rowBaseline = baseline;
          else if (mode === 'action') actionBaseline = baseline;
          else columnBaseline = baseline;
        }
        if (name === 'officebridge_create_chart') {
          chartBaseline = prepareChartBaseline(snapshot, range);
          onDemand.metrics.baselineCaptures++;
        }
        var computed = yield {
          operation: 'tool',
          payload: {
            snapshot: snapshot,
            name: name,
            params: params,
            previous: onDemand.task,
          },
        };
        result = computed.result;
        var plan = (onDemand.task = computed.plan);
        if (plan) {
          captured = JSON.parse(JSON.stringify(snapshot));
          write = true;
        }
      }
      if (!write) onDemand.cache[key] = JSON.parse(JSON.stringify(result));
      return JSON.parse(
        JSON.stringify(
          write
            ? {
                result: result,
                plan: onDemand.task,
                snapshot: captured,
              }
            : { result: result },
        ),
      );
    } catch (error) {
      onDemand.error = true;
      captured = null;
      columnBaseline = rowBaseline = actionBaseline = null;
      chartBaseline = null;
      throw error;
    }
  }
  function resolveQueryRange(params) {
    var d = onDemand.descriptor,
      nav = d.navigation;
    if (!nav) throw new Error('此兼容任务不支持自动定位，请使用已提供的范围。');
    var target = params.target,
      name = params.sheetName || nav.activeSheet,
      anchor = nav.anchor,
      source = target,
      b;
    if (nav.allowedSheets.indexOf(name) < 0)
      throw new Error(
        '此工作表未获本轮指令授权，请让用户明确指定工作表；锁定模式可先解除锁定。',
      );
    var sheet = currentBook(d.workbookFullName).Worksheets.Item(name);
    if (target === 'previous') {
      if (
        !nav.previousRange ||
        (params.sheetName && name !== nav.previousRange.sheetName)
      )
        throw new Error('没有可延续的本轮参考范围，请重新定位。');
      b = Object.assign({}, nav.previousRange);
      name = b.sheetName;
    } else if (target === 'selection') {
      if (name !== anchor.sheetName)
        throw new Error('选区只属于用户指向的工作表。');
      b = Object.assign({}, anchor);
    } else if (target === 'custom') {
      b = {
        startRow: params.startRow,
        startColumn: params.startColumn,
        rowCount: params.rowCount,
        columnCount: params.columnCount,
      };
    } else if (target === 'region') {
      if (name !== anchor.sheetName)
        throw new Error('其他工作表请先定位 sheet，再按表头使用 custom。');
      try {
        b = bounds(
          sheet.Cells.Item(anchor.startRow, anchor.startColumn).CurrentRegion,
        );
      } catch (_error) {
        b = bounds(sheet.UsedRange);
        source = 'used-bounds-fallback';
      }
    } else if (target === 'sheet' || target === 'column' || target === 'row') {
      var used = bounds(sheet.UsedRange);
      b = Object.assign({}, used);
      if (target !== 'sheet') {
        if (name !== anchor.sheetName)
          throw new Error('行列指向只属于当前工作表；其他表请指定 custom。');
        if (target === 'column') {
          b.startColumn = anchor.startColumn;
          b.columnCount = anchor.columnCount;
        } else {
          b.startRow = anchor.startRow;
          b.rowCount = anchor.rowCount;
        }
      }
    } else throw new Error('定位目标无效。');
    if (
      ![b.startRow, b.startColumn, b.rowCount, b.columnCount].every(
        function (n) {
          return Number.isSafeInteger(n) && n > 0;
        },
      ) ||
      b.startRow + b.rowCount - 1 > 1048576 ||
      b.startColumn + b.columnCount - 1 > 16384
    )
      throw new Error('定位坐标无效或超过宿主边界。');
    var range = Object.assign({}, b, {
      id: 'q' + onDemand.nextRange,
      sheetName: name,
    });
    range.address = C.rangeAddress(
      range.startRow,
      range.startColumn,
      range.rowCount,
      range.columnCount,
    );
    if (!C.authorizedRange(d, range))
      throw new Error('定位范围越界；锁定模式下读取和修改均不能超出锁定区域。');
    var existing = d.ranges.filter(function (r) {
      return r.sheetName === range.sheetName && r.address === range.address;
    })[0];
    if (existing) range = existing;
    else {
      if (d.ranges.length >= C.TABLE_LIMITS.ranges)
        throw new Error('本轮定位范围超过 50 个。');
      d.ranges.push(range);
      onDemand.nextRange++;
    }
    onDemand.lastResolved = range;
    return {
      range: range,
      result: Object.assign({}, range, {
        rangeId: range.id,
        source: source,
        note:
          source === 'used-bounds-fallback'
            ? '宿主未提供连续区域；这是已用边界，可能含多块数据或残留格式，请读取相关表头确认，必要时询问用户。'
            : target === 'region'
              ? '连续区域仅为候选边界，请结合表头和用户要求确认完整性。'
              : '已定位；单元格值尚未读取。',
      }),
    };
  }
  function currentBook(fullName) {
    var book = app().ActiveWorkbook;
    if (identity(book) !== fullName)
      throw new Error(
        '当前工作簿已切换或重命名，请回到原文件，或重新发送任务。',
      );
    return book;
  }
  function same(actual, expected) {
    return (
      actual.value === expected.value &&
      actual.formula === expected.formula &&
      (!expected.style ||
        JSON.stringify(actual.style) === JSON.stringify(expected.style))
    );
  }
  function assertWritable(sheet, cell, knownStyle, protectionChecked) {
    if (!protectionChecked && sheet.ProtectContents)
      throw new Error('工作表已受保护，不能修改。');
    if (cell.MergeCells || cell.HasArray)
      throw new Error('首版不修改合并单元格或数组公式，请选择普通单元格。');
    if (!(arguments.length > 2 ? knownStyle : style(cell)))
      throw new Error('无法完整读取单元格格式，不能保证恢复，已停止修改。');
  }
  function setValue(cell, value, numberFormat) {
    if (blank(value)) {
      // Clear contents without changing formatting or leaving an empty text value.
      if (typeof cell.ClearContents === 'function') cell.ClearContents();
      else cell.Value2 = null;
    } else {
      // Assign text as text, rather than allowing Excel/WPS to interpret it as a formula.
      if (typeof value === 'string') cell.NumberFormat = '@';
      cell.Value2 = value;
    }
    cell.NumberFormat = numberFormat;
  }
  function restoreCell(cell, before) {
    if (!before.style) throw new Error('缺少恢复格式。');
    cell.NumberFormat =
      before.style.numberFormat === '@' && before.formula
        ? 'General'
        : before.style.numberFormat;
    if (before.formula) cell.Formula = before.formula;
    else setValue(cell, before.value, before.style.numberFormat);
    cell.NumberFormat = before.style.numberFormat;
    cell.Font.Bold = before.style.bold;
    cell.Interior.Color = before.style.fillColor;
    cell.Interior.Pattern = before.style.fillPattern;
  }
  function removeSheet(sheet) {
    var host = app(),
      alerts = host.DisplayAlerts,
      count = host.ActiveWorkbook.Worksheets.Count;
    try {
      host.DisplayAlerts = false;
      sheet.Delete();
      if (host.ActiveWorkbook.Worksheets.Count !== count - 1)
        throw new Error('工作表删除未完成。');
    } finally {
      host.DisplayAlerts = alerts;
    }
  }
  function readCreated(sheet, rows, columns) {
    var states = [];
    for (var row = 1; row <= rows; row++)
      for (var column = 1; column <= columns; column++)
        states.push(state(sheet.Cells.Item(row, column)));
    return { bounds: bounds(sheet.UsedRange), states: states };
  }
  function columnRange(sheet, column, count) {
    var first = C.cellAddress(1, column).replace(/1$/, ''),
      last = C.cellAddress(1, column + count - 1).replace(/1$/, '');
    if (typeof sheet.Range !== 'function')
      throw new Error('当前 WPS 不提供整列操作接口。');
    var range = sheet.Range(first + ':' + last);
    return range.EntireColumn || range;
  }
  function rowRange(sheet, row, count) {
    var range = sheet.Range(row + ':' + (row + count - 1));
    return range.EntireRow || range;
  }
  function extraFormat(cell, restore) {
    var result = { cell: {}, font: {}, borders: [] };
    var fields = [
      'HorizontalAlignment',
      'VerticalAlignment',
      'WrapText',
      'Orientation',
      'IndentLevel',
      'ShrinkToFit',
      'Locked',
      'FormulaHidden',
    ];
    var fonts = [
      'Name',
      'Size',
      'Color',
      'Italic',
      'Underline',
      'Strikethrough',
      'Superscript',
      'Subscript',
    ];
    function copy(target, names, values) {
      names.forEach(function (name) {
        if (restore) {
          if (own(values, name)) target[name] = values[name];
        } else {
          var value = target[name];
          if (
            typeof value === 'string' ||
            typeof value === 'boolean' ||
            (typeof value === 'number' && Number.isFinite(value))
          )
            values[name] = value;
        }
      });
    }
    copy(cell, fields, restore ? restore.cell : result.cell);
    copy(cell.Font, fonts, restore ? restore.font : result.font);
    if (restore)
      restore.borders.forEach(function (border) {
        var target = cell.Borders.Item(border.index);
        target.LineStyle = border.lineStyle;
        target.Weight = border.weight;
        target.Color = border.color;
      });
    else {
      var borders = cell.Borders;
      if (borders && typeof borders.Item === 'function')
        [5, 6, 7, 8, 9, 10].forEach(function (index) {
          var border = borders.Item(index),
            lineStyle = border.LineStyle,
            weight = border.Weight,
            color = border.Color;
          if (
            typeof lineStyle !== 'number' ||
            typeof weight !== 'number' ||
            typeof color !== 'number'
          )
            throw new Error('无法读取边框格式，不能恢复整列删除。');
          result.borders.push({
            index: index,
            lineStyle: lineStyle,
            weight: weight,
            color: color,
          });
        });
    }
    return result;
  }
  function nonemptyCollection(value) {
    return value && typeof value.Count === 'number' && value.Count > 0;
  }
  function chartObjects(sheet) {
    var charts =
      typeof sheet.ChartObjects === 'function'
        ? sheet.ChartObjects()
        : sheet.ChartObjects;
    if (!charts || typeof charts.Count !== 'number')
      throw new Error('当前宿主无法读取图表集合。');
    return charts;
  }
  function optionalProperty(object, path) {
    try {
      var value = object;
      path.split('.').forEach(function (part) {
        value = value[part];
      });
      return typeof value === 'string' ||
        typeof value === 'boolean' ||
        (typeof value === 'number' && Number.isFinite(value))
        ? value
        : null;
    } catch (_) {
      return null;
    }
  }
  function chartState(object) {
    var chart = object.Chart,
      collection =
        typeof chart.SeriesCollection === 'function'
          ? chart.SeriesCollection()
          : chart.SeriesCollection,
      record = {
        name: String(object.Name),
        left: object.Left,
        top: object.Top,
        width: object.Width,
        height: object.Height,
        type: chart.ChartType,
        hasTitle: chart.HasTitle,
        title: chart.HasTitle ? String(chart.ChartTitle.Text) : '',
        series: [],
        appearance: [],
      };
    if (
      !collection ||
      collection.Count > 20 ||
      !record.name ||
      ![
        record.left,
        record.top,
        record.width,
        record.height,
        record.type,
      ].every(function (value) {
        return typeof value === 'number' && Number.isFinite(value);
      })
    )
      throw new Error('无法完整读取图表状态。');
    [
      'ChartStyle',
      'HasLegend',
      'Legend.Position',
      'ChartArea.Interior.Color',
      'PlotArea.Interior.Color',
    ].forEach(function (path) {
      record.appearance.push(optionalProperty(chart, path));
    });
    for (var index = 1; index <= collection.Count; index++) {
      var series = collection.Item(index),
        item = { formula: series.Formula, appearance: [] };
      if (typeof item.formula !== 'string' || item.formula.length > 10000)
        throw new Error('无法读取图表数据系列引用。');
      [
        'ChartType',
        'AxisGroup',
        'MarkerStyle',
        'MarkerSize',
        'Interior.Color',
        'Format.Fill.ForeColor.RGB',
        'Format.Line.ForeColor.RGB',
        'Format.Line.Weight',
        'Format.Line.DashStyle',
      ].forEach(function (path) {
        item.appearance.push(optionalProperty(series, path));
      });
      record.series.push(item);
    }
    return record;
  }
  function readCharts(sheet) {
    var charts = chartObjects(sheet),
      result = [];
    if (
      charts.Count > 20 ||
      (sheet.Shapes && sheet.Shapes.Count !== charts.Count)
    )
      throw new Error('工作表包含非图表图形或图表过多，无法完整核对。');
    for (var index = 1; index <= charts.Count; index++)
      result.push(chartState(charts.Item(index)));
    return result;
  }
  function readFilter(sheet) {
    if (sheet.AutoFilterMode === false) return null;
    if (sheet.AutoFilterMode !== true || !sheet.AutoFilter)
      throw new Error('无法读取原表筛选状态。');
    var filter = sheet.AutoFilter,
      frame = bounds(filter.Range),
      conditions = [];
    if (!filter.Filters || filter.Filters.Count !== frame.columnCount)
      throw new Error('筛选字段状态不完整。');
    for (var index = 1; index <= filter.Filters.Count; index++) {
      var item = filter.Filters.Item(index),
        record = { on: !!item.On };
      if (record.on) {
        record.criteria1 = item.Criteria1;
        record.operator = item.Operator || 0;
        if ([0, 1, 2, 7].indexOf(record.operator) < 0)
          throw new Error('颜色、图标、动态日期或高级筛选暂不能安全恢复。');
        try {
          var second = item.Criteria2;
          if (second !== undefined && second !== null)
            record.criteria2 = second;
        } catch (_) {}
        [record.criteria1, record.criteria2].forEach(function (criterion) {
          if (criterion === undefined) return;
          var values = Array.isArray(criterion) ? criterion : [criterion];
          if (
            values.length > 100 ||
            !values.every(function (value) {
              return (
                value === null ||
                typeof value === 'string' ||
                typeof value === 'boolean' ||
                (typeof value === 'number' && Number.isFinite(value))
              );
            })
          )
            throw new Error('原筛选条件无法完整读取。');
        });
      }
      conditions.push(record);
    }
    return {
      address: C.rangeAddress(
        frame.startRow,
        frame.startColumn,
        frame.rowCount,
        frame.columnCount,
      ),
      conditions: conditions,
    };
  }
  function restoreFilter(sheet, record) {
    sheet.AutoFilterMode = false;
    if (record.filter) {
      var target = sheet.Range(record.filter.address);
      target.AutoFilter();
      record.filter.conditions.forEach(function (item, index) {
        if (item.on)
          target.AutoFilter(
            index + 1,
            item.criteria1,
            item.operator,
            item.criteria2,
          );
      });
    }
    record.rows.forEach(function (item) {
      rowRange(sheet, item.row, 1).Hidden = item.hidden;
    });
  }
  function criterion(condition) {
    var operators = {
        eq: '=',
        ne: '<>',
        gt: '>',
        gte: '>=',
        lt: '<',
        lte: '<=',
      },
      literal = condition.value === null ? '' : String(condition.value),
      escaped = literal
        .replace(/~/g, '~~')
        .replace(/\*/g, '~*')
        .replace(/\?/g, '~?');
    return condition.operator === 'contains'
      ? '=*' + escaped + '*'
      : operators[condition.operator] + escaped;
  }
  function matchesFilter(value, condition) {
    var actual = value === null ? '' : String(value).toLowerCase(),
      expected =
        condition.value === null ? '' : String(condition.value).toLowerCase();
    if (condition.operator === 'eq') return actual === expected;
    if (condition.operator === 'ne') return actual !== expected;
    if (condition.operator === 'contains') return actual.indexOf(expected) >= 0;
    if (typeof value !== 'number') return false;
    if (condition.operator === 'gt') return value > condition.value;
    if (condition.operator === 'gte') return value >= condition.value;
    if (condition.operator === 'lt') return value < condition.value;
    return value <= condition.value;
  }
  function assertActionData(before, after, action, range) {
    if (before.sheets.length !== after.sheets.length)
      throw new Error('工作表数量发生异常变化。');
    before.sheets.forEach(function (record, index) {
      var next = after.sheets[index];
      if (
        next.name !==
        (action.type === 'rename' && record.name === range.sheetName
          ? action.name
          : record.name)
      )
        throw new Error('工作表名称发生异常变化。');
      if (
        JSON.stringify(record.frame) !== JSON.stringify(next.frame) ||
        record.cells.length !== next.cells.length
      )
        throw new Error('源数据范围发生异常变化。');
      record.cells.forEach(function (cell, i) {
        var other = next.cells[i];
        if (
          cell[0] !== other[0] ||
          cell[1] !== other[1] ||
          cell[2] !== other[2] ||
          JSON.stringify(record.formats[cell[4]]) !==
            JSON.stringify(next.formats[other[4]]) ||
          (action.type !== 'rename' && cell[3] !== other[3]) ||
          (cell[3] &&
            (!other[3] || (/#REF!/i.test(other[3]) && !/#REF!/i.test(cell[3]))))
        )
          throw new Error('操作后源数据、公式或格式回读不一致。');
      });
      if (JSON.stringify(record.columns) !== JSON.stringify(next.columns))
        throw new Error('列属性发生异常变化。');
      record.rows.forEach(function (row, i) {
        var other = next.rows[i],
          filterRow =
            action.type === 'filter' &&
            record.name === range.sheetName &&
            row.row > range.startRow &&
            row.row < range.startRow + range.rowCount;
        if (
          !other ||
          row.row !== other.row ||
          (row.height !== other.height &&
            !(
              filterRow &&
              ((row.hidden && row.height === 0) ||
                (other.hidden && other.height === 0))
            )) ||
          (!filterRow && row.hidden !== other.hidden)
        )
          throw new Error('行属性发生异常变化。');
      });
      if (
        !(action.type === 'filter' && record.name === range.sheetName) &&
        JSON.stringify(record.filter) !== JSON.stringify(next.filter)
      )
        throw new Error('其他筛选发生变化。');
      if (
        action.type !== 'rename' &&
        !(action.type === 'chart' && record.name === range.sheetName) &&
        JSON.stringify(record.charts) !== JSON.stringify(next.charts)
      )
        throw new Error('其他图表发生变化。');
      if (action.type === 'rename') {
        var oldCharts = JSON.parse(JSON.stringify(record.charts)),
          newCharts = JSON.parse(JSON.stringify(next.charts));
        [oldCharts, newCharts].forEach(function (charts) {
          charts.forEach(function (chart) {
            chart.series.forEach(function (series) {
              if (/#REF!/i.test(series.formula))
                throw new Error('图表引用发生错误。');
              delete series.formula;
            });
          });
        });
        if (JSON.stringify(oldCharts) !== JSON.stringify(newCharts))
          throw new Error('改名后图表属性异常变化。');
      }
    });
  }
  function normalizeSeries(formula) {
    return formula
      .replace(/\$/g, '')
      .replace(/'((?:[^']|'')+)'!/g, function (_, name) {
        return name.replace(/''/g, "'") + '!';
      })
      .replace(/([A-Z]+\d+):\1(?=[,)])/gi, '$1')
      .toLowerCase();
  }
  // Charts never write source cells. Keep data dependencies and chart objects,
  // rather than the full-workbook format/row/column restore used for deletion.
  function prepareChartBaseline(snapshot, range, charts) {
    var source = snapshot.ranges.filter(function (r) {
        return r.id === range.id;
      })[0],
      records = Object.create(null);
    source.values.forEach(function (row, r) {
      row.forEach(function (value, c) {
        var record = {
          sheetName: source.sheetName,
          row: source.startRow + r,
          column: source.startColumn + c,
          values: value,
          formulas: source.formulas[r][c],
        };
        records[JSON.stringify([record.sheetName, record.row, record.column])] =
          record;
      });
    });
    if (onDemand && onDemand.descriptor.snapshotId === snapshot.snapshotId)
      Object.keys(onDemand.reads).forEach(function (key) {
        records[key] = onDemand.reads[key];
      });
    return JSON.parse(
      JSON.stringify({
        mode: 'chart',
        snapshotId: snapshot.snapshotId,
        workbookFullName: snapshot.workbookFullName,
        sheetNames: snapshot.sheetNames,
        sheetName: source.sheetName,
        sources: Object.keys(records).map(function (k) {
          return records[k];
        }),
        charts:
          charts ||
          readCharts(
            currentBook(snapshot.workbookFullName).Worksheets.Item(
              source.sheetName,
            ),
          ),
      }),
    );
  }
  function chartTargets(base) {
    var book = currentBook(base.workbookFullName),
      names = [];
    for (var i = 1; i <= book.Worksheets.Count; i++)
      names.push(String(book.Worksheets.Item(i).Name));
    if (JSON.stringify(names) !== JSON.stringify(base.sheetNames))
      throw new Error('工作表结构已有变化，请重新发送。');
    var sheet = book.Worksheets.Item(base.sheetName);
    if (
      book.ProtectStructure ||
      sheet.ProtectContents ||
      nonemptyCollection(sheet.ListObjects) ||
      nonemptyCollection(
        typeof sheet.PivotTables === 'function'
          ? sheet.PivotTables()
          : sheet.PivotTables,
      )
    )
      throw new Error('保护、表格或透视表暂不支持此图表操作。');
    return book;
  }
  function chartBlocks(base) {
    var records = base.sources.slice().sort(function (a, b) {
        return (
          a.sheetName.localeCompare(b.sheetName) ||
          a.row - b.row ||
          a.column - b.column
        );
      }),
      segments = [],
      blocks = [];
    records.forEach(function (record) {
      var tail = segments[segments.length - 1];
      if (
        tail &&
        tail.sheetName === record.sheetName &&
        tail.row === record.row &&
        tail.column + tail.entries.length === record.column &&
        tail.entries.length < 64
      )
        tail.entries.push(record);
      else
        segments.push({
          sheetName: record.sheetName,
          row: record.row,
          column: record.column,
          entries: [record],
        });
    });
    segments.forEach(function (segment) {
      var tail = blocks[blocks.length - 1],
        width = segment.entries.length;
      if (
        tail &&
        tail.sheetName === segment.sheetName &&
        tail.column === segment.column &&
        tail.width === width &&
        tail.row + tail.rows.length === segment.row &&
        (tail.rows.length + 1) * width <= 256
      )
        tail.rows.push(segment.entries);
      else
        blocks.push({
          sheetName: segment.sheetName,
          row: segment.row,
          column: segment.column,
          width: width,
          rows: [segment.entries],
        });
    });
    return blocks;
  }
  function checkChartBlock(book, block) {
    var sheet = book.Worksheets.Item(block.sheetName),
      cells = sheet.Cells,
      matrix,
      formulaFree = false;
    try {
      var target = sheet.Range(
        C.rangeAddress(block.row, block.column, block.rows.length, block.width),
      );
      matrix = target.Value2;
      if (
        !Array.isArray(matrix) ||
        matrix.length !== block.rows.length ||
        matrix.some(function (row) {
          return !Array.isArray(row) || row.length !== block.width;
        })
      )
        matrix = null;
      formulaFree = target.HasFormula === false;
    } catch (_) {
      matrix = null;
    }
    if (matrix) lastChartMetrics.bulkReads++;
    block.rows.forEach(function (row, r) {
      row.forEach(function (record, c) {
        var cell, value;
        if (own(record, 'values')) {
          if (matrix) value = matrix[r][c];
          else {
            cell = cells.Item(record.row, record.column);
            value = cell.Value2;
          }
          if (!sameLiteralValue(value, record.values))
            throw new Error(
              '本轮已读取的数据已有变化，请重新发送；未覆盖后续修改。',
            );
        }
        if (own(record, 'formulas')) {
          var formula = null;
          if (!formulaFree) {
            cell = cell || cells.Item(record.row, record.column);
            formula = cell.HasFormula === false ? null : cell.Formula;
            formula =
              typeof formula === 'string' && formula.charAt(0) === '='
                ? formula
                : null;
          }
          if (formula !== record.formulas)
            throw new Error(
              '本轮已读取的公式已有变化，请重新发送；未覆盖后续修改。',
            );
        }
        if (own(record, 'styles')) {
          cell = cell || cells.Item(record.row, record.column);
          if (JSON.stringify(style(cell)) !== JSON.stringify(record.styles))
            throw new Error('已读取的格式已有变化，请重新发送。');
        }
        lastChartMetrics.checkedCells++;
      });
    });
  }
  function chartPrepared(plan) {
    if (!captured || !C.isEditPlan(captured, plan))
      throw new Error('修改方案已失效或越过所选范围，请重新发送任务。');
    var action = plan.sheetActions[0],
      range = captured.ranges.filter(function (r) {
        return r.id === action.rangeId;
      })[0];
    if (!chartBaseline || chartBaseline.snapshotId !== plan.snapshotId) {
      var old =
        actionBaseline &&
        actionBaseline.sheets.filter(function (s) {
          return s.name === range.sheetName;
        })[0];
      if (!old) throw new Error('图表预览已失效，请重新发送。');
      chartBaseline = prepareChartBaseline(captured, range, old.charts);
    }
    if (onDemand && onDemand.descriptor.snapshotId === plan.snapshotId) {
      var dependencies = Object.create(null);
      chartBaseline.sources.forEach(function (record) {
        dependencies[
          JSON.stringify([record.sheetName, record.row, record.column])
        ] = record;
      });
      Object.keys(onDemand.reads).forEach(function (key) {
        dependencies[key] = JSON.parse(JSON.stringify(onDemand.reads[key]));
      });
      chartBaseline.sources = Object.keys(dependencies).map(function (key) {
        return dependencies[key];
      });
    }
    return { base: chartBaseline, action: action, range: range };
  }
  function startChartMetrics() {
    lastChartMetrics = { checkedCells: 0, bulkReads: 0, yields: 0, applyMs: 0 };
  }
  function verifyChartSources(base, blocks) {
    var book = chartTargets(base);
    blocks.forEach(function (block) {
      checkChartBlock(book, block);
    });
    return book;
  }
  async function applyAsync(plan, options) {
    options = options || {};
    var original = captured,
      execution;
    if (options.signal && options.signal.aborted)
      throw new Error('任务已停止，未提交修改。');
    if (options.prepare) {
      execution = await options.prepare(captured, plan);
      if (captured !== original) throw new Error('方案已失效，请重新发送。');
    }
    if (options.signal && options.signal.aborted)
      throw new Error('任务已停止，未提交修改。');
    if (
      !plan ||
      !plan.sheetActions ||
      !plan.sheetActions.length ||
      plan.sheetActions[0].type !== 'chart'
    ) {
      if (
        onDemand &&
        plan &&
        onDemand.descriptor.snapshotId === plan.snapshotId
      )
        await schedule(queryStableSteps(), options, plan.snapshotId);
      return apply(plan);
    }
    startChartMetrics();
    var prepared = chartPrepared(plan),
      base = prepared.base,
      snapshot = captured,
      blocks = chartBlocks(base),
      completed = 0,
      slice = Date.now();
    if (
      execution &&
      execution.valid &&
      execution.planKey === JSON.stringify(plan) &&
      execution.snapshotId === plan.snapshotId
    )
      prepared.execution = execution;
    function guard() {
      if (options.signal && options.signal.aborted)
        throw new Error('已停止，未应用修改。');
      if (captured !== snapshot || chartBaseline !== base)
        throw new Error('图表预览已失效，请重新发送。');
      return chartTargets(base);
    }
    for (var i = 0; i < blocks.length; i++) {
      checkChartBlock(guard(), blocks[i]);
      completed += blocks[i].width * blocks[i].rows.length;
      if (Date.now() - slice >= 8 || i === blocks.length - 1) {
        if (options.onProgress)
          options.onProgress(completed, base.sources.length);
        lastChartMetrics.yields++;
        await new Promise(function (resolve) {
          setTimeout(resolve, 0);
        });
        guard();
        slice = Date.now();
      }
    }
    guard();
    // Final bulk verification and mutation share one JS turn. No stale data
    // checked before yielding may authorize a later write.
    return applyChart(plan, prepared, blocks);
  }
  function applyChart(plan, prepared, blocks) {
    var started = Date.now();
    if (!prepared) {
      startChartMetrics();
      prepared = chartPrepared(plan);
      blocks = chartBlocks(prepared.base);
    }
    var base = prepared.base,
      action = prepared.action,
      range = prepared.range,
      book = verifyChartSources(base, blocks),
      sheet = book.Worksheets.Item(base.sheetName),
      previous = history,
      created = null;
    if (JSON.stringify(readCharts(sheet)) !== JSON.stringify(base.charts))
      throw new Error('图表集合已有变化，请重新发送；未覆盖后续修改。');
    var transaction = {
        workbookFullName: plan.workbookFullName,
        cells: [],
        sheets: [],
        incomplete: false,
        sheetAction: {
          summary: {
            type: 'chart',
            sheetName: base.sheetName,
            action: action,
            range: range.address,
          },
          before: base,
          after: null,
        },
      },
      host = app(),
      updating = host.ScreenUpdating,
      paused = false;
    try {
      if (typeof updating === 'boolean') {
        host.ScreenUpdating = false;
        paused = true;
      }
      created = createChart(sheet, range, action, prepared.execution);
      created.configure();
      transaction.sheetAction.createdName = String(created.object.Name);
      var after = readCharts(sheet);
      if (
        after.length !== base.charts.length + 1 ||
        JSON.stringify(after.slice(0, -1)) !== JSON.stringify(base.charts)
      )
        throw new Error('图表集合回读不一致。');
      verifyChartSources(base, blocks);
      transaction.sheetAction.after = Object.assign({}, base, {
        charts: after,
      });
      history = transaction;
      captured = null;
      chartBaseline = null;
      return {
        changedCells: 0,
        newSheets: [],
        sheetActions: [transaction.sheetAction.summary],
        recoverable: true,
      };
    } catch (error) {
      try {
        if (created) created.object.Delete();
        if (JSON.stringify(readCharts(sheet)) !== JSON.stringify(base.charts))
          throw new Error('图表集合回读不一致。');
        verifyChartSources(base, blocks);
      } catch (restoreError) {
        transaction.incomplete = true;
        history = transaction;
        captured = null;
        chartBaseline = null;
        throw new Error(
          error.message +
            '；无法确认完整恢复：' +
            restoreError.message +
            '。请检查原表，自动恢复已禁用。',
        );
      }
      history = previous;
      captured = null;
      chartBaseline = null;
      throw new Error(error.message + '；本次操作已恢复，未保留变更。');
    } finally {
      if (paused) host.ScreenUpdating = updating;
      lastChartMetrics.applyMs = Date.now() - started;
    }
  }
  function undoChart(record) {
    startChartMetrics();
    var base = record.before,
      blocks = chartBlocks(base),
      book = verifyChartSources(base, blocks),
      sheet = book.Worksheets.Item(base.sheetName);
    if (
      JSON.stringify(readCharts(sheet)) !== JSON.stringify(record.after.charts)
    )
      throw new Error('图表已有变化，不能覆盖后续修改。');
    try {
      chartObjects(sheet).Item(record.createdName).Delete();
      if (JSON.stringify(readCharts(sheet)) !== JSON.stringify(base.charts))
        throw new Error('图表恢复回读不一致。');
      history = null;
      captured = null;
      chartBaseline = null;
      return {
        restoredCells: 0,
        removedSheets: [],
        sheetActions: [record.summary],
      };
    } catch (error) {
      history.incomplete = true;
      throw new Error('图表恢复未完整完成，请检查原表：' + error.message);
    }
  }
  function createChart(sheet, range, action, execution) {
    var summary =
      action.aggregation === 'sum'
        ? execution
          ? execution.summary
          : C.summarizeChart(range, action)
        : null;
    var target = sheet.Range(range.address),
      charts = chartObjects(sheet),
      object = charts.Add(
        typeof target.Left === 'number' && typeof target.Width === 'number'
          ? target.Left + target.Width + 16
          : 20,
        typeof target.Top === 'number' ? target.Top : 20,
        480,
        280,
      );
    // Store the object immediately so failures during configuration can remove
    // exactly our new chart without deleting an existing chart.
    return {
      object: object,
      configure: function () {
        object.Name =
          'OfficeBridge_' +
          Date.now() +
          '_' +
          Math.random().toString(36).slice(2, 8);
        var chart = object.Chart;
        chart.ChartType = { line: 4, column: 51, bar: 57, pie: 5 }[
          action.chartType
        ];
        chart.HasTitle = true;
        chart.ChartTitle.Text = action.title;
        var series =
          typeof chart.SeriesCollection === 'function'
            ? chart.SeriesCollection()
            : chart.SeriesCollection;
        while (series.Count) series.Item(1).Delete();
        function absolute(row, column) {
          return C.cellAddress(row, column).replace(
            /^([A-Z]+)(\d+)$/,
            '$$$1$$$2',
          );
        }
        action.valueColumns.forEach(function (column, index) {
          var item = series.NewSeries(),
            prefix = "'" + range.sheetName.replace(/'/g, "''") + "'!",
            categories =
              prefix +
              absolute(
                range.startRow + 1,
                range.startColumn + action.categoryColumn,
              ) +
              ':' +
              absolute(
                range.startRow + range.rowCount - 1,
                range.startColumn + action.categoryColumn,
              ),
            values =
              prefix +
              absolute(range.startRow + 1, range.startColumn + column) +
              ':' +
              absolute(
                range.startRow + range.rowCount - 1,
                range.startColumn + column,
              ),
            name = range.values[0][column],
            formula =
              '=SERIES("' +
              String(name === null ? '系列 ' + (index + 1) : name).replace(
                /"/g,
                '""',
              ) +
              '",' +
              categories +
              ',' +
              values +
              ',' +
              (index + 1) +
              ')';
          if (summary)
            formula = execution
              ? execution.formulas[index]
              : C.summarySeriesFormula(summary, index);
          item.Formula = formula;
          if (
            summary
              ? String(item.Formula) !== formula
              : normalizeSeries(item.Formula) !== normalizeSeries(formula)
          )
            throw new Error('图表数据系列引用回读不一致。');
        });
        var result = chartState(object);
        if (
          result.type !==
            { line: 4, column: 51, bar: 57, pie: 5 }[action.chartType] ||
          result.title !== action.title ||
          result.series.length !== action.valueColumns.length
        )
          throw new Error('图表类型、标题或系列回读不一致。');
      },
    };
  }
  function applySheetAction(plan) {
    if (!actionBaseline)
      throw new Error('工作表操作恢复基线已失效，请重新发送。');
    var book = currentBook(plan.workbookFullName),
      before = actionBaseline,
      action = plan.sheetActions[0],
      range = captured.ranges.filter(function (item) {
        return item.id === action.rangeId;
      })[0],
      sheet = book.Worksheets.Item(range.sheetName),
      previous = history,
      source = before.sheets.filter(function (item) {
        return item.name === range.sheetName;
      })[0],
      created = null,
      transaction = {
        workbookFullName: plan.workbookFullName,
        cells: [],
        sheets: [],
        incomplete: false,
        sheetAction: {
          summary: {
            type: action.type,
            sheetName: range.sheetName,
            action: action,
            range: range.address,
          },
          before: before,
          after: null,
        },
      };
    assertColumnState(book, before);
    if (action.type === 'filter') {
      if (source.filter && source.filter.address !== range.address)
        throw new Error('已有其他范围筛选，请先清除或选择相同完整范围。');
      if (
        !source.filter &&
        source.rows.some(function (row) {
          return (
            row.row > range.startRow &&
            row.row < range.startRow + range.rowCount &&
            row.hidden
          );
        })
      )
        throw new Error('选区含手动隐藏行，请先取消隐藏再筛选，避免混淆记录。');
    }
    try {
      if (action.type === 'rename') sheet.Name = action.name;
      else if (action.type === 'filter') {
        sheet.AutoFilterMode = false;
        if (action.conditions.length) {
          var target = sheet.Range(range.address);
          action.conditions.forEach(function (condition) {
            target.AutoFilter(condition.column + 1, criterion(condition));
          });
        }
        range.values.slice(1).forEach(function (row, index) {
          var visible =
            !action.conditions.length ||
            action.conditions.every(function (condition) {
              return matchesFilter(row[condition.column], condition);
            });
          if (rowRange(sheet, range.startRow + 1 + index, 1).Hidden === visible)
            throw new Error('筛选行可见性回读与条件不一致。');
        });
      } else {
        created = createChart(sheet, range, action);
        created.configure();
        transaction.sheetAction.createdName = String(created.object.Name);
      }
      var specs = before.specs.map(function (spec) {
        return Object.assign({}, spec, {
          sheetName:
            action.type === 'rename' && spec.sheetName === range.sheetName
              ? action.name
              : spec.sheetName,
        });
      });
      transaction.sheetAction.after = captureColumnState(book, specs, 'action');
      assertActionData(before, transaction.sheetAction.after, action, range);
      if (action.type === 'filter') {
        var updated = transaction.sheetAction.after.sheets.filter(
          function (item) {
            return item.name === range.sheetName;
          },
        )[0].filter;
        if (
          action.conditions.length
            ? !updated || updated.address !== range.address
            : updated !== null
        )
          throw new Error('原生筛选范围回读不一致。');
      }
      if (action.type === 'rename' && String(sheet.Name) !== action.name)
        throw new Error('工作表名称回读不一致。');
      if (action.type === 'chart') {
        var chartsAfter = transaction.sheetAction.after.sheets.filter(
          function (item) {
            return item.name === range.sheetName;
          },
        )[0].charts;
        if (
          chartsAfter.length !== source.charts.length + 1 ||
          JSON.stringify(chartsAfter.slice(0, -1)) !==
            JSON.stringify(source.charts)
        )
          throw new Error('图表集合回读不一致。');
      }
      history = transaction;
      captured = null;
      actionBaseline = null;
      return {
        changedCells: 0,
        newSheets: [],
        sheetActions: [transaction.sheetAction.summary],
        recoverable: true,
      };
    } catch (error) {
      try {
        if (action.type === 'rename' && String(sheet.Name) === action.name)
          sheet.Name = range.sheetName;
        else if (action.type === 'filter') restoreFilter(sheet, source);
        else if (created) created.object.Delete();
        assertColumnState(book, before);
      } catch (restoreError) {
        history = transaction;
        transaction.incomplete = true;
        captured = null;
        actionBaseline = null;
        throw new Error(
          error.message +
            '；无法确认完整恢复：' +
            restoreError.message +
            '。请检查原表，自动恢复已禁用。',
        );
      }
      history = previous;
      captured = null;
      actionBaseline = null;
      throw new Error(error.message + '；本次操作已恢复，未保留变更。');
    }
  }
  function assertColumnSheet(sheet) {
    if (
      sheet.ProtectContents ||
      nonemptyCollection(sheet.Shapes) ||
      nonemptyCollection(sheet.ListObjects) ||
      sheet.AutoFilterMode ||
      nonemptyCollection(
        typeof sheet.PivotTables === 'function'
          ? sheet.PivotTables()
          : sheet.PivotTables,
      )
    )
      throw new Error(
        '此工作簿含保护、图形、筛选、表格或透视表，暂不能安全增删整行或整列。',
      );
  }
  function captureColumnState(book, ranges, mode, reusable, observed) {
    return drain(captureColumnSteps(book, ranges, mode, reusable, observed));
  }
  function* captureColumnSteps(book, ranges, mode, reusable, observed) {
    if (app().Workbooks && app().Workbooks.Count > 1)
      throw new Error(
        '为完整恢复公式引用，请先只保留目标工作簿打开，再发送列操作。',
      );
    var userNames = false;
    if (mode !== 'action' && nonemptyCollection(book.Names)) {
      if (book.Names.Count > 100 || typeof book.Names.Item !== 'function')
        userNames = true;
      else
        for (var namedIndex = 1; namedIndex <= book.Names.Count; namedIndex++)
          if (
            !/_FilterDatabase$/i.test(String(book.Names.Item(namedIndex).Name))
          )
            userNames = true;
    }
    if (userNames || book.ProtectStructure)
      throw new Error(
        '工作簿含不支持的命名区域或结构保护，不能安全执行结构操作。',
      );
    if (book.Worksheets.Count > C.TABLE_LIMITS.ranges)
      throw new Error('列操作恢复基线超过 50 张表。');
    var result = {
      sheets: [],
      names: [],
      specs: ranges.map(function (range) {
        return {
          sheetName: range.sheetName,
          startRow: range.startRow,
          startColumn: range.startColumn,
          rowCount: range.rowCount,
          columnCount: range.columnCount,
        };
      }),
    };
    if (mode) result.mode = mode;
    if (mode === 'action') {
      result.definedNames = [];
      if (book.Names && book.Names.Count > 100)
        throw new Error('命名区域过多，无法完整核对。');
      for (var n = 1; book.Names && n <= book.Names.Count; n++) {
        var named = book.Names.Item(n);
        if (!/_FilterDatabase$/i.test(String(named.Name)))
          result.definedNames.push([
            String(named.Name),
            String(named.RefersTo),
          ]);
      }
    }
    // Capture metadata first, then enforce the byte budget incrementally while
    // reading. Large empty/format-heavy ranges cannot allocate an unbounded baseline.
    for (var index = 1; index <= book.Worksheets.Count; index++) {
      var sheet = book.Worksheets.Item(index),
        used = bounds(sheet.UsedRange),
        frame = Object.assign({}, used);
      if (mode !== 'action') assertColumnSheet(sheet);
      else if (
        sheet.ProtectContents ||
        nonemptyCollection(sheet.ListObjects) ||
        nonemptyCollection(
          typeof sheet.PivotTables === 'function'
            ? sheet.PivotTables()
            : sheet.PivotTables,
        )
      )
        throw new Error('保护、表格或透视表暂不支持结构操作。');
      ranges
        .filter(function (range) {
          return range.sheetName === String(sheet.Name);
        })
        .forEach(function (range) {
          var bottom = Math.max(
              frame.startRow + frame.rowCount,
              range.startRow + range.rowCount,
            ),
            right = Math.max(
              frame.startColumn + frame.columnCount,
              range.startColumn + range.columnCount,
            );
          frame.startRow = Math.min(frame.startRow, range.startRow);
          frame.startColumn = Math.min(frame.startColumn, range.startColumn);
          frame.rowCount = bottom - frame.startRow;
          frame.columnCount = right - frame.startColumn;
        });
      result.names.push(String(sheet.Name));
      result.sheets.push({
        name: String(sheet.Name),
        used: used,
        frame: frame,
        cells: [],
        formats: [],
        columns: [],
      });
    }
    var budget = byteBudget(
      '完整恢复基线超过 512 KB，请缩小工作簿范围。普通单元格工具仍可使用。',
    );
    budget.add(result);
    var slice = Date.now(),
      visited = 0;
    for (
      var recordIndex = 0;
      recordIndex < result.sheets.length;
      recordIndex++
    ) {
      var record = result.sheets[recordIndex];
      var sheet = book.Worksheets.Item(record.name),
        frame = record.frame,
        formatIds = Object.create(null);
      var cached =
        reusable &&
        reusable.sheets.filter(function (item) {
          return (
            item.name === record.name &&
            JSON.stringify(item.frame) === JSON.stringify(frame)
          );
        })[0];
      // All three capabilities are captured synchronously from the same host
      // snapshot. Reuse the private cell baseline rather than reading every COM
      // cell three times; apply/undo always recapture independently.
      if (cached) {
        record.cells = JSON.parse(JSON.stringify(cached.cells));
        record.formats = JSON.parse(JSON.stringify(cached.formats));
        record.columns = JSON.parse(JSON.stringify(cached.columns));
        budget.array(record.cells);
        budget.array(record.formats);
        budget.array(record.columns);
      } else {
        var cells = sheet.Cells,
          selected = observed && observed[record.name];
        for (var r = 0; r < frame.rowCount; r++)
          for (var c = 0; c < frame.columnCount; c++) {
            var row = frame.startRow + r,
              column = frame.startColumn + c,
              cell = cells.Item(row, column),
              current =
                (selected && selected[row + ':' + column]) || state(cell);
            assertWritable(sheet, cell, current.style, true);
            if (
              mode !== 'action' &&
              (cell.Comment ||
                nonemptyCollection(cell.FormatConditions) ||
                nonemptyCollection(cell.Hyperlinks) ||
                nonemptyCollection(cell.SparklineGroups))
            )
              throw new Error(
                '工作簿含批注、条件格式或超链接，暂不能安全恢复整列删除。',
              );
            try {
              var validation = mode !== 'action' && cell.Validation;
              if (validation && validation.Type > 0)
                throw new Error('column-validation');
            } catch (error) {
              if (error.message === 'column-validation')
                throw new Error('工作簿含数据验证，暂不能安全恢复整列删除。');
            }
            if (current.formula && /\[|\]/.test(current.formula))
              throw new Error('工作簿含外部引用，暂不能安全增删整列。');
            var format = {
                style: Object.assign({}, current.style),
                extra: extraFormat(cell),
              },
              key = JSON.stringify(format),
              formatIndex = own(formatIds, key) ? formatIds[key] : -1;
            if (formatIndex < 0) {
              formatIndex = record.formats.length;
              record.formats.push(format);
              budget.add(format);
              formatIds[key] = formatIndex;
            }
            var entry = [
              row,
              column,
              current.value,
              current.formula,
              formatIndex,
            ];
            budget.add(entry);
            record.cells.push(entry);
            if (++visited % 32 === 0 && Date.now() - slice >= 8) {
              yield null;
              slice = Date.now();
            }
          }
        var end = Math.min(
          16384,
          frame.startColumn + frame.columnCount - 1 + C.TABLE_LIMITS.columns,
        );
        // Width/visibility are independent of UsedRange values.
        for (var column = 1; column <= end; column++) {
          var target = columnRange(sheet, column, 1),
            width = target.ColumnWidth,
            hidden = target.Hidden;
          if (
            typeof width !== 'number' ||
            !Number.isFinite(width) ||
            typeof hidden !== 'boolean'
          )
            throw new Error('当前 WPS 无法读取列宽或隐藏状态，列操作不可用。');
          var properties = { column: column, width: width, hidden: hidden };
          budget.add(properties);
          record.columns.push(properties);
          if (++visited % 32 === 0 && Date.now() - slice >= 8) {
            yield null;
            slice = Date.now();
          }
        }
      }
      if (mode) {
        record.rows = [];
        var rowEnd = Math.min(
          1048576,
          frame.startRow +
            frame.rowCount -
            1 +
            (mode === 'rows' ? C.TABLE_LIMITS.rows : 0),
        );
        for (var row = frame.startRow; row <= rowEnd; row++) {
          var whole = rowRange(sheet, row, 1),
            height = whole.RowHeight,
            rowHidden = whole.Hidden;
          if (
            typeof height !== 'number' ||
            !Number.isFinite(height) ||
            typeof rowHidden !== 'boolean'
          )
            throw new Error('无法读取行高或隐藏状态，操作不可用。');
          var properties = { row: row, height: height, hidden: rowHidden };
          budget.add(properties);
          record.rows.push(properties);
          if (++visited % 32 === 0 && Date.now() - slice >= 8) {
            yield null;
            slice = Date.now();
          }
        }
      }
      if (mode === 'action') {
        record.filter = readFilter(sheet);
        record.charts = readCharts(sheet);
        budget.add(record.filter);
        budget.add(record.charts);
      }
    }
    if (
      new TextEncoder().encode(JSON.stringify(result)).length >
      C.TABLE_LIMITS.bytes
    )
      throw new Error('整列恢复基线超过 512 KB，请使用更小工作簿。');
    return result;
  }
  function assertColumnState(book, expected) {
    if (
      JSON.stringify(
        captureColumnState(book, expected.specs, expected.mode),
      ) !== JSON.stringify(expected)
    )
      throw new Error(
        '工作簿内容、公式、格式或行列属性已有变化，请重新发送；未覆盖后续修改。',
      );
  }
  function storedColumnCell(record, entry) {
    var format = record.formats[entry[4]];
    return {
      row: entry[0],
      column: entry[1],
      state: { value: entry[2], formula: entry[3], style: format.style },
      format: format.extra,
    };
  }
  function restoreColumnState(book, baseline) {
    baseline.sheets.forEach(function (record) {
      var sheet = book.Worksheets.Item(record.name);
      record.cells.forEach(function (entry) {
        var item = storedColumnCell(record, entry),
          cell = sheet.Cells.Item(item.row, item.column);
        restoreCell(cell, item.state);
        extraFormat(cell, item.format);
      });
      record.columns.forEach(function (item) {
        var column = columnRange(sheet, item.column, 1);
        // Some WPS builds report zero for a hidden column's saved width.
        // Inverse insertion/deletion preserves surviving hidden columns natively.
        if (!item.hidden || item.width > 0) column.ColumnWidth = item.width;
        column.Hidden = item.hidden;
      });
      (record.rows || []).forEach(function (item) {
        var target = rowRange(sheet, item.row, 1);
        if (!item.hidden || item.height > 0) target.RowHeight = item.height;
        target.Hidden = item.hidden;
      });
    });
    baseline.sheets.forEach(function (record) {
      var sheet = book.Worksheets.Item(record.name);
      if (typeof sheet.Calculate === 'function') sheet.Calculate();
    });
    baseline.sheets.forEach(function (record) {
      var sheet = book.Worksheets.Item(record.name);
      record.cells.forEach(function (entry) {
        var item = storedColumnCell(record, entry),
          cell = sheet.Cells.Item(item.row, item.column);
        if (
          !same(state(cell), item.state) ||
          JSON.stringify(extraFormat(cell)) !== JSON.stringify(item.format)
        )
          throw new Error('列操作恢复后的内容或格式回读不一致。');
      });
      record.columns.forEach(function (item) {
        var column = columnRange(sheet, item.column, 1);
        if (column.ColumnWidth !== item.width || column.Hidden !== item.hidden)
          throw new Error('列宽或隐藏状态恢复不一致。');
      });
      (record.rows || []).forEach(function (item) {
        var target = rowRange(sheet, item.row, 1);
        if (target.RowHeight !== item.height || target.Hidden !== item.hidden)
          throw new Error('行高或隐藏状态恢复不一致。');
      });
    });
  }
  function nativeColumns(book, summary, reverse) {
    var action = reverse
      ? summary.action === 'insert'
        ? 'delete'
        : 'insert'
      : summary.action;
    var target = (summary.row !== undefined ? rowRange : columnRange)(
      book.Worksheets.Item(summary.sheetName),
      summary.row !== undefined ? summary.row : summary.column,
      summary.count,
    );
    if (typeof target[action === 'insert' ? 'Insert' : 'Delete'] !== 'function')
      throw new Error(
        summary.row !== undefined
          ? '当前 WPS 不支持该整行操作。'
          : '当前 WPS 不支持该整列操作。',
      );
    if (action === 'insert') target.Insert();
    else target.Delete();
  }
  function applyColumns(plan, rows) {
    if (!(rows ? rowBaseline : columnBaseline))
      throw new Error('行列操作恢复基线已失效，请重新发送。');
    var book = currentBook(plan.workbookFullName),
      before = rows ? rowBaseline : columnBaseline,
      change = (rows ? plan.rowChanges : plan.columnChanges)[0],
      range = captured.ranges.filter(function (item) {
        return item.id === change.rangeId;
      })[0];
    assertColumnState(book, before);
    var summary = {
      sheetName: range.sheetName,
      action: change.action,
      count: change.count,
    };
    summary[rows ? 'row' : 'column'] = rows
      ? range.startRow + change.row
      : range.startColumn + change.column;
    var targetRecord = before.sheets.filter(function (item) {
      return item.name === summary.sheetName;
    })[0];
    if (
      change.action === 'delete' &&
      (rows ? targetRecord.rows : targetRecord.columns).some(function (item) {
        return (
          item[rows ? 'row' : 'column'] >= summary[rows ? 'row' : 'column'] &&
          item[rows ? 'row' : 'column'] <
            summary[rows ? 'row' : 'column'] + summary.count &&
          item.hidden &&
          item[rows ? 'height' : 'width'] <= 0
        );
      })
    )
      throw new Error(
        rows
          ? 'WPS 无法读取所删隐藏行的原行高，请先取消这些行的隐藏后重新发送。'
          : 'WPS 无法读取所删隐藏列的原列宽，请先取消这些列的隐藏后重新发送。',
      );
    if (
      change.action === 'insert' &&
      targetRecord.frame[rows ? 'startRow' : 'startColumn'] +
        targetRecord.frame[rows ? 'rowCount' : 'columnCount'] -
        1 +
        change.count >
        (rows ? 1048576 : 16384)
    )
      throw new Error('插入后将超出工作表行列边界，未修改。');
    var previous = history,
      executed = false,
      transaction = {
        workbookFullName: plan.workbookFullName,
        cells: [],
        sheets: [],
        columnChange: { summary: summary, before: before, after: null },
        incomplete: false,
      };
    try {
      nativeColumns(book, summary, false);
      executed = true;
      before.sheets.forEach(function (record) {
        var sheet = book.Worksheets.Item(record.name);
        if (typeof sheet.Calculate === 'function') sheet.Calculate();
      });
      // Check moved source literals and formats; native WPS updates formula references.
      before.sheets.forEach(function (record) {
        var sheet = book.Worksheets.Item(record.name);
        record.cells.forEach(function (entry) {
          var item = storedColumnCell(record, entry);
          var column = item.column,
            row = item.row,
            position = rows ? row : column;
          if (record.name === summary.sheetName) {
            if (
              change.action === 'delete' &&
              position >= summary[rows ? 'row' : 'column'] &&
              position < summary[rows ? 'row' : 'column'] + summary.count
            )
              return;
            if (position >= summary[rows ? 'row' : 'column'])
              position +=
                change.action === 'insert' ? summary.count : -summary.count;
            if (rows) row = position;
            else column = position;
          }
          var cell = sheet.Cells.Item(row, column),
            current = state(cell);
          if (item.state.formula) {
            if (
              !current.formula ||
              (/#REF!/i.test(current.formula) &&
                !/#REF!/i.test(item.state.formula)) ||
              (/^#/.test(String(current.value)) &&
                !/^#/.test(String(item.state.value)))
            )
              throw new Error('整列操作导致公式错误，已停止。');
          } else if (!same(current, item.state))
            throw new Error('整列移位后的数据或格式回读不一致。');
          if (JSON.stringify(extraFormat(cell)) !== JSON.stringify(item.format))
            throw new Error('整列移位后的单元格格式不一致。');
        });
      });
      if (change.action === 'insert' && change.values)
        change.values.forEach(function (row, r) {
          row.forEach(function (value, c) {
            var cell = book.Worksheets.Item(summary.sheetName).Cells.Item(
              rows ? summary.row + r : range.startRow + r,
              rows ? range.startColumn + c : summary.column + c,
            );
            setValue(cell, value, cell.NumberFormat);
            var actual = state(cell);
            if (
              !sameLiteralValue(actual.value, value) ||
              actual.formula !== null
            )
              throw new Error('新增列数据回读不一致。');
          });
        });
      before.sheets.forEach(function (record) {
        var sheet = book.Worksheets.Item(record.name);
        if (typeof sheet.Calculate === 'function') sheet.Calculate();
      });
      transaction.columnChange.after = captureColumnState(
        book,
        before.specs,
        before.mode,
      );
      history = transaction;
      captured = null;
      columnBaseline = null;
      var result = {
        changedCells: change.values
          ? rows
            ? change.count * range.columnCount
            : range.rowCount * change.count
          : 0,
        newSheets: [],
        recoverable: true,
      };
      result[rows ? 'rowChanges' : 'columnChanges'] = [summary];
      return result;
    } catch (error) {
      try {
        if (executed) {
          nativeColumns(book, summary, true);
          restoreColumnState(book, before);
        } else assertColumnState(book, before);
      } catch (restoreError) {
        transaction.incomplete = true;
        history = transaction;
        captured = null;
        columnBaseline = null;
        throw new Error(
          error.message +
            '；无法确认列结构已完整恢复：' +
            restoreError.message +
            '。请检查原表，自动恢复已禁用。',
        );
      }
      history = previous;
      captured = null;
      columnBaseline = null;
      throw new Error(
        error.message +
          (rows
            ? '；本次行操作已恢复，未保留变更。'
            : '；本次列操作已恢复，未保留变更。'),
      );
    }
  }
  function apply(plan) {
    if (
      plan &&
      plan.sheetActions &&
      plan.sheetActions.length &&
      plan.sheetActions[0].type === 'chart'
    )
      return applyChart(plan);
    if (onDemand && plan && onDemand.descriptor.snapshotId === plan.snapshotId)
      assertQueryStable();
    if (!captured || !C.isEditPlan(captured, plan))
      throw new Error('修改方案已失效或越过所选范围，请重新发送任务。');
    if (plan.columnChanges && plan.columnChanges.length)
      return applyColumns(plan);
    if (plan.rowChanges && plan.rowChanges.length)
      return applyColumns(plan, true);
    if (plan.sheetActions && plan.sheetActions.length)
      return applySheetAction(plan);
    if (!plan.edits.length && !plan.newSheets.length)
      throw new Error('没有需要应用的修改。');
    var book = currentBook(plan.workbookFullName),
      previous = history;
    var transaction = {
      workbookFullName: plan.workbookFullName,
      cells: [],
      sheets: [],
      incomplete: false,
    };
    var names = [];
    for (var index = 1; index <= book.Worksheets.Count; index++)
      names.push(String(book.Worksheets.Item(index).Name).toLowerCase());
    plan.newSheets.forEach(function (sheet) {
      if (names.indexOf(sheet.name.toLowerCase()) >= 0)
        throw new Error(
          '结果工作表“' + sheet.name + '”已存在，请重新生成方案。',
        );
    });
    plan.edits.forEach(function (edit) {
      var range = captured.ranges.filter(function (item) {
        return item.id === edit.rangeId;
      })[0];
      var sheet = book.Worksheets.Item(range.sheetName),
        row = range.startRow + edit.row,
        column = range.startColumn + edit.column;
      var cell = sheet.Cells.Item(row, column),
        before = state(cell);
      var expected = {
        value: range.values[edit.row][edit.column],
        formula: range.formulas[edit.row][edit.column],
        style: range.styles[edit.row][edit.column],
      };
      assertWritable(sheet, cell);
      if (!same(before, expected))
        throw new Error(
          range.sheetName +
            '!' +
            C.cellAddress(row, column) +
            ' 已被修改，请重新发送任务。',
        );
      transaction.cells.push({
        sheetName: range.sheetName,
        row: row,
        column: column,
        before: before,
        after: null,
        edit: edit,
      });
    });
    var applied = [];
    try {
      transaction.cells.forEach(function (item) {
        var cell = book.Worksheets.Item(item.sheetName).Cells.Item(
            item.row,
            item.column,
          ),
          edit = item.edit;
        // Formula recalculation may change dependent values, but never a literal or formula expression.
        var now = state(cell);
        if (
          now.formula !== item.before.formula ||
          (!now.formula && now.value !== item.before.value)
        )
          throw new Error('执行中目标内容发生变化。');
        applied.push(item);
        if (own(edit, 'formula')) {
          if (cell.NumberFormat === '@') cell.NumberFormat = 'General';
          cell.Formula = edit.formula;
        } else if (own(edit, 'value'))
          setValue(cell, edit.value, item.before.style.numberFormat);
        if (edit.format) {
          if (own(edit.format, 'numberFormat'))
            cell.NumberFormat = edit.format.numberFormat;
          if (own(edit.format, 'bold')) cell.Font.Bold = edit.format.bold;
          if (own(edit.format, 'fillColor')) {
            cell.Interior.Pattern = 1;
            cell.Interior.Color = edit.format.fillColor;
          }
        }
      });
      plan.newSheets.forEach(function (result) {
        var sheet = book.Worksheets.Add();
        var record = {
          name: String(sheet.Name),
          rows: result.values.length,
          columns: result.values[0].length,
          after: null,
          reference: sheet,
        };
        transaction.sheets.push(record);
        sheet.Name = result.name;
        record.name = result.name;
        result.values.forEach(function (row, r) {
          row.forEach(function (value, c) {
            var cell = sheet.Cells.Item(r + 1, c + 1),
              originalFormat = cell.NumberFormat;
            // A fresh sheet is already blank; a placeholder needs no write.
            if (!blank(value)) setValue(cell, value, originalFormat);
            var actual = state(cell);
            if (
              !sameLiteralValue(actual.value, value) ||
              actual.formula !== null
            )
              throw new Error('新增结果表的值回读不一致。');
          });
        });
      });
      var touched = {};
      transaction.cells.forEach(function (item) {
        touched[item.sheetName] = true;
      });
      Object.keys(touched).forEach(function (name) {
        var sheet = book.Worksheets.Item(name);
        if (typeof sheet.Calculate === 'function') sheet.Calculate();
      });
      transaction.cells.forEach(function (item) {
        var cell = book.Worksheets.Item(item.sheetName).Cells.Item(
            item.row,
            item.column,
          ),
          after = state(cell),
          edit = item.edit;
        if (
          own(edit, 'formula') &&
          (!after.formula ||
            after.formula.toUpperCase() !== edit.formula.toUpperCase() ||
            /^#/.test(String(cell.Text || after.value)))
        )
          throw new Error('公式未正确写入或计算返回错误。');
        if (
          own(edit, 'value') &&
          (!sameLiteralValue(after.value, edit.value) || after.formula !== null)
        )
          throw new Error('单元格值回读不一致。');
        if (edit.format)
          Object.keys(edit.format).forEach(function (key) {
            if (!after.style || after.style[key] !== edit.format[key])
              throw new Error('格式回读不一致。');
          });
        item.after = after;
      });
      transaction.sheets.forEach(function (record) {
        record.after = readCreated(
          record.reference,
          record.rows,
          record.columns,
        );
      });
      history = transaction;
      captured = null;
      return {
        changedCells: transaction.cells.length,
        newSheets: transaction.sheets.map(function (sheet) {
          return sheet.name;
        }),
        recoverable: true,
      };
    } catch (error) {
      var failures = [];
      transaction.sheets
        .slice()
        .reverse()
        .forEach(function (record) {
          try {
            removeSheet(record.reference);
          } catch (_error) {
            failures.push('工作表 ' + record.name);
          }
        });
      applied
        .slice()
        .reverse()
        .forEach(function (item) {
          try {
            restoreCell(
              book.Worksheets.Item(item.sheetName).Cells.Item(
                item.row,
                item.column,
              ),
              item.before,
            );
          } catch (_error) {
            failures.push(
              item.sheetName + '!' + C.cellAddress(item.row, item.column),
            );
          }
        });
      applied.forEach(function (item) {
        try {
          var sheet = book.Worksheets.Item(item.sheetName);
          if (typeof sheet.Calculate === 'function') sheet.Calculate();
          if (
            !same(state(sheet.Cells.Item(item.row, item.column)), item.before)
          )
            throw new Error('恢复回读不一致');
        } catch (_error) {
          failures.push(
            item.sheetName + '!' + C.cellAddress(item.row, item.column),
          );
        }
      });
      history = previous;
      captured = null;
      if (failures.length) {
        transaction.incomplete = true;
        history = transaction;
        throw new Error(
          error.message +
            '；自动恢复未完整完成：' +
            failures.join('、') +
            '。请检查原表，并查看恢复详情。',
        );
      }
      throw new Error(error.message + '；本次修改已恢复，未保留变更。');
    }
  }
  function undo() {
    if (!history) throw new Error('没有可恢复的修改。关闭侧栏后记录不保留。');
    if (history.incomplete)
      throw new Error(
        '上次操作存在未完整恢复的位置，请按错误信息检查原表；不能安全执行自动恢复。',
      );
    var book = currentBook(history.workbookFullName);
    if (history.sheetAction) {
      var actionRecord = history.sheetAction,
        summary = actionRecord.summary,
        action = summary.action;
      if (actionRecord.before.mode === 'chart') return undoChart(actionRecord);
      assertColumnState(book, actionRecord.after);
      try {
        var sheet = book.Worksheets.Item(
          action.type === 'rename' ? action.name : summary.sheetName,
        );
        if (action.type === 'rename') sheet.Name = summary.sheetName;
        else if (action.type === 'filter')
          restoreFilter(
            sheet,
            actionRecord.before.sheets.filter(function (item) {
              return item.name === summary.sheetName;
            })[0],
          );
        else chartObjects(sheet).Item(actionRecord.createdName).Delete();
        assertColumnState(book, actionRecord.before);
        history = null;
        captured = null;
        actionBaseline = null;
        return { restoredCells: 0, removedSheets: [], sheetActions: [summary] };
      } catch (error) {
        history.incomplete = true;
        throw new Error(
          '工作表操作恢复未完整完成，请检查原表：' + error.message,
        );
      }
    }
    if (history.columnChange) {
      var record = history.columnChange;
      assertColumnState(book, record.after);
      try {
        nativeColumns(book, record.summary, true);
        restoreColumnState(book, record.before);
        history = null;
        captured = null;
        columnBaseline = null;
        var result = {
          restoredCells: record.before.sheets.reduce(function (total, sheet) {
            return total + sheet.cells.length;
          }, 0),
          removedSheets: [],
        };
        result[
          record.summary.row !== undefined ? 'rowChanges' : 'columnChanges'
        ] = [record.summary];
        return result;
      } catch (error) {
        history.incomplete = true;
        throw new Error('列操作恢复未完整完成，请检查原表：' + error.message);
      }
    }
    history.cells.forEach(function (item) {
      var sheet = book.Worksheets.Item(item.sheetName),
        cell = sheet.Cells.Item(item.row, item.column);
      assertWritable(sheet, cell);
      if (!same(state(cell), item.after))
        throw new Error(
          item.sheetName +
            '!' +
            C.cellAddress(item.row, item.column) +
            ' 有后续修改，未覆盖。',
        );
    });
    history.sheets.forEach(function (record) {
      var sheet = book.Worksheets.Item(record.name);
      if (
        JSON.stringify(bounds(sheet.UsedRange)) !==
          JSON.stringify(record.after.bounds) ||
        JSON.stringify(readCreated(sheet, record.rows, record.columns)) !==
          JSON.stringify(record.after)
      )
        throw new Error('结果表“' + record.name + '”有后续修改，未删除。');
    });
    var record = history;
    try {
      record.cells
        .slice()
        .reverse()
        .forEach(function (item) {
          restoreCell(
            book.Worksheets.Item(item.sheetName).Cells.Item(
              item.row,
              item.column,
            ),
            item.before,
          );
        });
      record.sheets
        .slice()
        .reverse()
        .forEach(function (sheet) {
          removeSheet(book.Worksheets.Item(sheet.name));
        });
      record.cells.forEach(function (item) {
        var sheet = book.Worksheets.Item(item.sheetName);
        if (typeof sheet.Calculate === 'function') sheet.Calculate();
      });
      record.cells.forEach(function (item) {
        if (
          !same(
            state(
              book.Worksheets.Item(item.sheetName).Cells.Item(
                item.row,
                item.column,
              ),
            ),
            item.before,
          )
        )
          throw new Error('恢复后内容回读不一致。');
      });
      history = null;
      captured = null;
      return {
        restoredCells: record.cells.length,
        removedSheets: record.sheets.map(function (sheet) {
          return sheet.name;
        }),
      };
    } catch (error) {
      record.incomplete = true;
      throw new Error('恢复未完整完成，请检查原表：' + error.message);
    }
  }
  function recovery() {
    if (!history) return null;
    return {
      incomplete: history.incomplete,
      columnChanges:
        history.columnChange && history.columnChange.summary.row === undefined
          ? [history.columnChange.summary]
          : [],
      rowChanges:
        history.columnChange && history.columnChange.summary.row !== undefined
          ? [history.columnChange.summary]
          : [],
      sheetActions: history.sheetAction ? [history.sheetAction.summary] : [],
      cellCount: history.cells.length,
      cells: history.cells.slice(0, 100).map(function (item) {
        return {
          sheetName: item.sheetName,
          address: C.cellAddress(item.row, item.column),
          before: item.before,
        };
      }),
      newSheets: history.sheets.map(function (sheet) {
        return sheet.name;
      }),
    };
  }
  Object.assign(global.OfficeBridgeHost, {
    describe: describe,
    capture: capture,
    beginRead: beginRead,
    beginReadAsync: function (scope, expected, options) {
      return schedule(beginReadSteps(scope, expected, options), options);
    },
    executeQuery: executeQuery,
    executeQueryAsync: function (snapshotId, name, params, options) {
      return schedule(
        querySteps(snapshotId, name, params),
        options,
        snapshotId,
      );
    },
    navigationContext: function () {
      return onDemand && onDemand.descriptor.navigation && onDemand.lastResolved
        ? JSON.parse(
            JSON.stringify({
              workbookFullName: onDemand.descriptor.workbookFullName,
              anchor: onDemand.descriptor.navigation.anchor,
              range: onDemand.lastResolved,
            }),
          )
        : null;
    },
    finishRead: function (snapshotId, keepPlan) {
      if (onDemand && onDemand.descriptor.snapshotId === snapshotId) {
        lastReadMetrics = Object.assign({}, onDemand.metrics);
        if (!keepPlan) {
          onDemand = null;
          captured = null;
          columnBaseline = rowBaseline = actionBaseline = null;
          chartBaseline = null;
        }
      }
    },
    queryMetrics: function () {
      return onDemand
        ? Object.assign({}, onDemand.metrics)
        : lastReadMetrics && Object.assign({}, lastReadMetrics);
    },
    apply: apply,
    applyAsync: applyAsync,
    chartMetrics: function () {
      return lastChartMetrics && Object.assign({}, lastChartMetrics);
    },
    undo: undo,
    undoAsync: async function (options) {
      await new Promise(function (resolve) {
        setTimeout(resolve, 0);
      });
      if (options && options.signal && options.signal.aborted)
        throw new Error('已停止，未恢复修改。');
      return undo();
    },
    recovery: recovery,
    canUndo: function () {
      return !!history && !history.incomplete;
    },
  });
})(window);
