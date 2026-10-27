/* global Application, wps */
(function (global) {
  'use strict';
  var connectionId =
    'wps-' + Date.now() + '-' + Math.random().toString(36).slice(2);
  function app() {
    return typeof Application !== 'undefined'
      ? Application
      : typeof wps !== 'undefined'
        ? wps
        : null;
  }
  function readSelection() {
    var host = app();
    if (!host) throw new Error('请从 WPS 的 OfficeBridge 选项卡打开侧栏。');
    var workbook = host.ActiveWorkbook;
    var sheet = host.ActiveSheet;
    var selection = host.Selection;
    if (
      !workbook ||
      !sheet ||
      !selection ||
      !selection.Rows ||
      !selection.Columns
    )
      throw new Error('请先打开工作簿并框选单元格。');
    if (selection.Areas && selection.Areas.Count > 1)
      throw new Error('请框选一块连续区域后重试。');
    var rowCount = Number(selection.Rows.Count);
    var columnCount = Number(selection.Columns.Count);
    if (
      !Number.isSafeInteger(rowCount) ||
      !Number.isSafeInteger(columnCount) ||
      rowCount < 1 ||
      columnCount < 1 ||
      rowCount > 1048576 ||
      columnCount > 16384
    )
      throw new Error('选区无效。');
    var values = [];
    var readBytes = 0;
    for (var row = 0; row < rowCount; row++) {
      var cells = [];
      for (var column = 0; column < columnCount; column++) {
        var value = sheet.Cells.Item(
          selection.Row + row,
          selection.Column + column,
        ).Value2;
        if (value === undefined || value === null) value = null;
        if (
          typeof value !== 'string' &&
          typeof value !== 'number' &&
          typeof value !== 'boolean' &&
          value !== null
        )
          throw new Error('选区中存在不支持的单元格类型。');
        if (
          (typeof value === 'number' && !Number.isFinite(value)) ||
          (typeof value === 'string' && value.length > 10000)
        )
          throw new Error('单元格内容过大或数值无效，请缩小范围。');
        // The value bytes are a lower bound on the final JSON size. Stop
        // before reading the rest of an oversized selection, without sampling.
        readBytes += new Blob([JSON.stringify(value)]).size;
        if (readBytes > 512000)
          throw new Error('选区内容超过 512 KB，请缩小范围。');
        cells.push(value);
      }
      values.push(cells);
    }
    var snapshot = {
      host: 'wps',
      connectionId: connectionId,
      workbookName: String(workbook.Name),
      sheetName: String(sheet.Name),
      address: String(
        typeof selection.Address === 'function'
          ? selection.Address()
          : selection.Address,
      ),
      rowCount: rowCount,
      columnCount: columnCount,
      values: values,
      capturedAt: new Date().toISOString(),
    };
    if (new Blob([JSON.stringify(snapshot)]).size > 512000)
      throw new Error('选区内容超过 512 KB，请缩小范围。');
    return snapshot;
  }
  global.OfficeBridgeHost = {
    available: function () {
      return !!app();
    },
    readSelection: readSelection,
  };
})(window);
