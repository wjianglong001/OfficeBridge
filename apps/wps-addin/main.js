/* global wps, Application */
var officeBridgeTaskpaneId = null;

function OfficeBridgeLoad() {
  return true;
}

function OfficeBridgeOpen() {
  try {
    var app = typeof Application !== 'undefined' ? Application : wps;
    if (officeBridgeTaskpaneId !== null) {
      try {
        var existing = app.GetTaskPane(officeBridgeTaskpaneId);
        if (existing) {
          existing.Visible = true;
          return;
        }
      } catch (error) {
        officeBridgeTaskpaneId = null;
      }
    }
    var url = new URL('taskpane.html', window.location.href).href;
    var pane = app.CreateTaskPane(url);
    officeBridgeTaskpaneId = pane.ID;
    pane.Visible = true;
  } catch (error) {
    alert('无法打开 OfficeBridge：' + error.message);
  }
}
