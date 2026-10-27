/* global OfficeBridgeExecutor */
(function () {
  var status = document.getElementById('status');
  var button = document.getElementById('read');
  async function read() {
    button.disabled = true;
    var id;
    try {
      var target = OfficeBridgeExecutor.describe('selection');
      if (target.workbookName !== 'OfficeBridge-XLL-Probe.xlsx')
        throw new Error('请切换到 OfficeBridge-XLL-Probe.xlsx 临时测试表。');
      var selected = target.ranges[0];
      if (!selected || selected.rowCount > 20 || selected.columnCount > 10)
        throw new Error('验证仅支持最多 20 行、10 列的小选区。');
      var descriptor = await OfficeBridgeExecutor.beginRead(
        'selection',
        target,
        { omitSelection: true },
      );
      id = descriptor.snapshotId;
      var data = await OfficeBridgeExecutor.query(
        id,
        'officebridge_read_range',
        { rangeId: descriptor.ranges[0].id },
      );
      document.getElementById('range').textContent = selected.address;
      document.getElementById('values').textContent = JSON.stringify(
        data.result.values,
        null,
        2,
      );
      var session = await (await fetch('/probe/session')).json();
      var response = await fetch('/probe/read-result', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'X-OfficeBridge-Probe-Token': session.token,
        },
        body: JSON.stringify({
          address: selected.address,
          rows: selected.rowCount,
          columns: selected.columnCount,
        }),
      });
      if (!response.ok) throw new Error('验证回报未确认。');
      status.textContent = '已连接 WPS，选区读取成功';
    } catch (error) {
      status.textContent = (error && error.message) || '读取失败';
    } finally {
      if (id) OfficeBridgeExecutor.finishRead(id, false);
      button.disabled = false;
    }
  }
  button.addEventListener('click', read);
  read();
})();
