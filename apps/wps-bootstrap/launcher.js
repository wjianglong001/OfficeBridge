/* global Application, wps */
(function () {
  'use strict';
  var ENTRY_URL = 'http://127.0.0.1:4318/wps/taskpane.html';
  var BOOTSTRAP_ORIGIN = 'http://127.0.0.1:43192';
  var active = null;
  var root = document.getElementById('launcher');
  var status = document.getElementById('status');
  var hint = document.getElementById('hint');
  var retry = document.getElementById('retry');
  var cancel = document.getElementById('cancel');

  function show(state, message, detail) {
    root.setAttribute('data-state', state);
    root.setAttribute('aria-busy', String(state === 'working'));
    status.textContent = message;
    hint.textContent = detail;
    retry.hidden = state === 'working';
    cancel.hidden = state !== 'working';
  }
  function failure(message) {
    var error = new Error(message);
    error.officeBridge = true;
    return error;
  }
  function current(run) {
    return active === run && !run.controller.signal.aborted;
  }
  function assertCurrent(run) {
    if (!current(run)) throw failure('已停止等待。');
  }
  function officeHost() {
    if (typeof wps !== 'undefined' && wps) return wps;
    if (typeof Application !== 'undefined' && Application) return Application;
    return null;
  }
  function finish(run) {
    if (active !== run) return;
    active = null;
    run.controller.abort();
  }
  function stop(visible) {
    if (!active) return;
    var previous = active;
    active = null;
    previous.controller.abort();
    if (visible) show('cancelled', '已取消等待', '需要时可点击重试。');
  }
  async function fetchJson(run, path, options, acceptStart, deadline) {
    assertCurrent(run);
    var controller = new AbortController();
    var timedOut = false;
    var timeout = deadline
      ? Math.min(3000, Math.max(1, deadline - Date.now()))
      : 3000;
    var abort = function () {
      controller.abort();
    };
    run.controller.signal.addEventListener('abort', abort, { once: true });
    var timer = setTimeout(function () {
      timedOut = true;
      controller.abort();
    }, timeout);
    try {
      var response = await fetch(
        path,
        Object.assign(
          {
            method: 'GET',
            cache: 'no-store',
            credentials: 'same-origin',
            redirect: 'error',
            signal: controller.signal,
          },
          options || {},
        ),
      );
      assertCurrent(run);
      if (timedOut) throw failure('连接启动组件超时，请重试。');
      if (deadline && Date.now() >= deadline)
        throw failure('办公桥启动等待超时，请检查后重试。');
      if (!response.ok && !(acceptStart && response.status === 409))
        throw failure('无法连接办公桥启动组件，请重新启用加载项后重试。');
      var result = await response.json();
      assertCurrent(run);
      if (timedOut) throw failure('连接启动组件超时，请重试。');
      if (deadline && Date.now() >= deadline)
        throw failure('办公桥启动等待超时，请检查后重试。');
      if (!result || typeof result !== 'object' || Array.isArray(result))
        throw failure('启动组件返回的信息不完整，请重试。');
      if (acceptStart && response.status === 409) {
        if (result.error === 'already_requested') return result;
        if (result.error === 'runtime_unavailable')
          throw failure('安装文件缺失或不匹配，请完整解压安装包。');
        if (result.error === 'port_conflict')
          throw failure('办公桥所需的本地端口被其他程序占用，请检查后重试。');
        throw failure('启动请求未确认，请重试。');
      }
      if (
        acceptStart &&
        !(
          (response.status === 202 && result.state === 'starting') ||
          (response.status === 200 &&
            (result.state === 'ready' || result.state === 'failed'))
        )
      )
        throw failure('启动请求未确认，请重试。');
      return result;
    } catch (error) {
      if (!current(run)) throw error;
      if (timedOut) throw failure('连接启动组件超时，请重试。');
      if (error && error.officeBridge === true) throw error;
      throw failure('无法连接办公桥启动组件，请重新启用加载项后重试。');
    } finally {
      clearTimeout(timer);
      run.controller.signal.removeEventListener('abort', abort);
    }
  }
  function pause(run, milliseconds) {
    assertCurrent(run);
    return new Promise(function (resolve, reject) {
      var timer;
      var cleanup = function () {
        clearTimeout(timer);
        run.controller.signal.removeEventListener('abort', abort);
      };
      var abort = function () {
        cleanup();
        reject(failure('已停止等待。'));
      };
      timer = setTimeout(function () {
        cleanup();
        resolve();
      }, milliseconds);
      run.controller.signal.addEventListener('abort', abort, { once: true });
    });
  }
  function inspect(run, data) {
    assertCurrent(run);
    if (data.state === 'ready') {
      if (data.entryUrl !== ENTRY_URL)
        throw failure('办公桥入口未通过核对，请重新启用加载项。');
      show('working', '正在打开办公桥…', '即将进入对话。');
      window.location.replace(ENTRY_URL);
      finish(run);
      return true;
    }
    if (data.state === 'failed') throw failure(nativeFailure(data));
    if (data.state === 'conflict')
      throw failure('办公桥所需的本地端口被其他程序占用，请检查后重试。');
    if (data.state !== 'offline' && data.state !== 'starting')
      throw failure('无法确认办公桥状态，请重试。');
    return false;
  }
  function validCode(value) {
    return (
      typeof value === 'number' &&
      Number.isInteger(value) &&
      value >= 0 &&
      value <= 4294967295
    );
  }
  function nativeFailure(data) {
    var details = [];
    if (validCode(data.errorCode) && data.errorCode !== 0)
      details.push('系统错误码 ' + data.errorCode);
    if (validCode(data.exitCode) && data.exitCode !== 259)
      details.push('进程退出码 ' + data.exitCode);
    return (
      '办公桥后台启动失败' +
      (details.length ? '（' + details.join('，') + '）' : '') +
      '，请检查后重试。'
    );
  }
  async function start(explicitRetry) {
    if (active) return;
    var run = { controller: new AbortController() };
    active = run;
    show('working', '正在连接办公桥…', '请稍候。');
    try {
      if (window.location.origin !== BOOTSTRAP_ORIGIN)
        throw failure('请从 WPS 加载项打开办公桥。');
      var office = officeHost();
      if (!office) throw failure('请在 WPS 表格中打开办公桥。');
      var initial = await fetchJson(run, '/native/bridge-status');
      var retryFailed = explicitRetry === true && initial.state === 'failed';
      if (!retryFailed && inspect(run, initial)) return;
      if (retryFailed || initial.state === 'offline') {
        var session = await fetchJson(run, '/probe/session');
        if (
          typeof session.token !== 'string' ||
          !session.token ||
          session.token.length > 256 ||
          /[\r\n]/.test(session.token)
        )
          throw failure('启动连接已失效，请重新启用加载项。');
        var launch = await fetchJson(
          run,
          '/native/start',
          {
            method: 'POST',
            headers: {
              'Content-Type': 'application/json',
              'X-OfficeBridge-Probe-Token': session.token,
            },
            body: JSON.stringify({ retry: explicitRetry === true }),
          },
          true,
        );
        assertCurrent(run);
        if (launch.error === 'already_requested')
          show('working', '正在等待办公桥…', '另一个窗口正在启动，请稍候。');
        else show('working', '正在启动办公桥…', '请稍候。');
      } else show('working', '正在等待办公桥…', '请稍候。');
      var deadline = Date.now() + 20000;
      while (current(run) && Date.now() < deadline) {
        var state = await fetchJson(
          run,
          '/native/bridge-status',
          null,
          false,
          deadline,
        );
        if (inspect(run, state)) return;
        await pause(run, Math.min(300, Math.max(1, deadline - Date.now())));
      }
      assertCurrent(run);
      throw failure('办公桥启动等待超时，请检查后重试。');
    } catch (error) {
      if (!current(run)) return;
      show(
        'error',
        '暂时无法打开办公桥',
        (error && error.officeBridge === true && error.message) ||
          '无法连接办公桥启动组件，请重新启用加载项后重试。',
      );
      finish(run);
    }
  }
  retry.addEventListener('click', function () {
    void start(true);
  });
  cancel.addEventListener('click', function () {
    stop(true);
  });
  window.addEventListener('pagehide', function () {
    stop(false);
  });
  void start(false);
})();
