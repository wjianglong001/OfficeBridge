/* global OfficeBridgeHost */
(function (global) {
  'use strict';
  var job = null,
    queue = Promise.resolve();
  function compute(operation, payload, options) {
    options = options || {};
    return new Promise(function (resolve, reject) {
      if (job) {
        reject(new Error('已有后台执行任务，请先停止。'));
        return;
      }
      if (options.signal && options.signal.aborted) {
        reject(new Error('任务已停止。'));
        return;
      }
      if (typeof Worker !== 'function') {
        reject(new Error('当前 WPS 不支持后台执行，请更新 WPS 后重试。'));
        return;
      }
      var worker,
        timer,
        settled = false;
      function finish(error, result) {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        if (options.signal) options.signal.removeEventListener('abort', abort);
        if (worker) worker.terminate();
        job = null;
        if (error) reject(error);
        else resolve(result);
      }
      function abort() {
        finish(new Error('任务已停止，未提交后续操作。'));
      }
      try {
        worker = new Worker('./execution-worker.js');
        job = { cancel: abort };
        worker.onmessage = function (event) {
          var item = event.data;
          if (item.type === 'progress') {
            if (options.onProgress) options.onProgress(item.message);
          } else if (item.type === 'result') {
            finish(null, item.result);
          } else finish(new Error(item.message || '后台方案计算失败。'));
        };
        worker.onerror = function (event) {
          event.preventDefault();
          finish(new Error('后台执行器不可用，请重新打开侧栏。'));
        };
        worker.onmessageerror = function () {
          finish(new Error('后台执行结果无法读取。'));
        };
        if (options.signal)
          options.signal.addEventListener('abort', abort, { once: true });
        timer = setTimeout(function () {
          finish(new Error('后台方案计算超时，未提交修改。'));
        }, 30000);
        worker.postMessage({ operation: operation, payload: payload });
      } catch (error) {
        finish(error);
      }
    });
  }
  function enqueue(work, options) {
    options = options || {};
    var next = queue.then(function () {
      if (options.signal && options.signal.aborted)
        throw new Error('任务已停止。');
      return work();
    });
    queue = next.catch(function () {});
    return next;
  }
  function hostOptions(options) {
    options = options || {};
    return Object.assign({}, options, {
      compute: function (operation, payload) {
        return compute(operation, payload, {
          signal: options.signal,
          onProgress: options.onStage,
        });
      },
      prepare: function (snapshot, plan) {
        return compute(
          'prepare',
          { snapshot: snapshot, plan: plan },
          { signal: options.signal, onProgress: options.onStage },
        );
      },
    });
  }
  global.OfficeBridgeExecutor = {
    available: function () {
      return OfficeBridgeHost.available();
    },
    describe: function (scope) {
      return OfficeBridgeHost.describe(scope);
    },
    canUndo: function () {
      return OfficeBridgeHost.canUndo();
    },
    recovery: function () {
      return OfficeBridgeHost.recovery();
    },
    navigationContext: function () {
      return OfficeBridgeHost.navigationContext();
    },
    finishRead: function (id, keep) {
      return OfficeBridgeHost.finishRead(id, keep);
    },
    prepare: function (snapshot, plan, options) {
      return enqueue(async function () {
        var result = await compute(
          'prepare',
          { snapshot: snapshot, plan: plan },
          options,
        );
        if (
          !result ||
          !result.valid ||
          result.snapshotId !== snapshot.snapshotId ||
          result.planKey !== JSON.stringify(plan)
        )
          throw new Error('后台返回了不匹配的方案。');
        return result;
      }, options);
    },
    beginRead: function (scope, target, options) {
      return enqueue(function () {
        return OfficeBridgeHost.beginReadAsync(
          scope,
          target,
          hostOptions(options),
        );
      }, options);
    },
    query: function (snapshotId, name, args, options) {
      return enqueue(function () {
        return OfficeBridgeHost.executeQueryAsync(
          snapshotId,
          name,
          args,
          hostOptions(options),
        );
      }, options);
    },
    apply: function (plan, options) {
      return enqueue(function () {
        return OfficeBridgeHost.applyAsync(plan, hostOptions(options));
      }, options);
    },
    undo: function (options) {
      return enqueue(function () {
        return OfficeBridgeHost.undoAsync(hostOptions(options));
      }, options);
    },
    stop: function () {
      if (job) job.cancel();
    },
  };
  global.addEventListener('pagehide', function () {
    if (job) job.cancel();
  });
})(window);
