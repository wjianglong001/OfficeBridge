/* global Application, wps */
var officeBridgeNativePaneId = null;

function OfficeBridgeNativeLoad() {
  return true;
}

function OfficeBridgeNativeOpen() {
  try {
    var app = typeof Application !== 'undefined' ? Application : wps;
    if (officeBridgeNativePaneId !== null) {
      try {
        var existing = app.GetTaskPane(officeBridgeNativePaneId);
        if (existing) {
          existing.Visible = true;
          return 'opened';
        }
      } catch (_) {
        officeBridgeNativePaneId = null;
      }
    }
    var pane = app.CreateTaskPane(
      new URL('launcher.html', window.location.href).href,
    );
    officeBridgeNativePaneId = pane.ID;
    pane.Visible = true;
    return 'opened';
  } catch (_) {
    alert('无法打开办公桥，请重新启用加载项。');
    return 'failed';
  }
}
