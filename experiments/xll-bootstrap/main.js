/* global Application, wps */
var officeBridgeProbePane = null;
function OfficeBridgeProbeLoad() {
  return true;
}
function OfficeBridgeProbeOpen() {
  var app = typeof Application !== 'undefined' ? Application : wps;
  if (officeBridgeProbePane !== null) {
    try {
      app.GetTaskPane(officeBridgeProbePane).Visible = true;
      return 'opened';
    } catch (_) {
      officeBridgeProbePane = null;
    }
  }
  var pane = app.CreateTaskPane(
    new URL('taskpane.html', window.location.href).href,
  );
  officeBridgeProbePane = pane.ID;
  pane.Visible = true;
  return 'opened';
}
