// Native relay doubles only: all services use temporary loopback ports.
// Never connects to WPS port 58890, reads real workbooks or calls an Agent.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { spawn } from 'node:child_process';
import { once } from 'node:events';
import { createServer } from 'node:http';
import { createServer as reserveServer } from 'node:net';
const info = JSON.parse(
  await readFile(
    new URL('../../.officebridge/xll-native/build-info.json', import.meta.url),
    'utf8',
  ),
);
const headers = { 'X-OfficeBridge-Xll-Probe': info.token };
const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));
const own = {
  name: 'OfficeBridgeXllProbe',
  addonType: 'et',
  online: 'true',
  enable: 'enable_dev',
  url: 'http://127.0.0.1:43191/probe/',
  version: '5.0.0',
};
async function freePort() {
  const reservation = reserveServer();
  reservation.listen(0, '127.0.0.1');
  await once(reservation, 'listening');
  const port = reservation.address().port;
  await new Promise((resolve) => reservation.close(resolve));
  return port;
}
async function scenario(name, options = {}) {
  const {
    budget = 2000,
    approvalDelay = 0,
    response = 'OK',
    status = 200,
    stop = false,
    redirect = false,
    refused = false,
    registry = [],
    rawRegistry,
    registryStatus = 200,
    registryDelay = 0,
    expectedDeployments = 1,
    expectedInvocations = 1,
    expectedDeployment,
    expectedRegistration,
    expectedHttp,
    invocationResponse = 'OK',
    invocationStatus = 200,
    expectedRequests = expectedInvocations,
  } = options;
  let queries = 0,
    deployments = 0,
    invocations = 0,
    unexpected = 0,
    accepted;
  const arrived = new Promise((resolve) => {
    accepted = resolve;
  });
  const timers = new Set();
  const later = (ms, operation) => {
    const timer = setTimeout(() => {
      timers.delete(timer);
      operation();
    }, ms);
    timers.add(timer);
  };
  const server = createServer((request, reply) => {
    let body = '';
    request.on('data', (chunk) => {
      body += chunk;
    });
    request.on('end', () => {
      if (request.url === '/publishlist') {
        queries++;
        later(registryDelay, () => {
          if (!reply.destroyed) {
            reply.writeHead(registryStatus);
            reply.end(rawRegistry ?? JSON.stringify(registry));
          }
        });
      } else if (request.url === '/deployaddons/runParams') {
        deployments++;
        accepted();
        const payload = JSON.parse(
          Buffer.from(body, 'base64').toString('utf8'),
        );
        assert.equal(payload.name, own.name);
        assert.equal(payload.version, own.version);
        assert.equal(payload.url, own.url);
        if (refused) {
          request.socket.destroy();
          return;
        }
        later(approvalDelay, () => {
          if (!reply.destroyed) {
            reply.writeHead(
              status,
              redirect ? { Location: '/unexpected-redirect' } : {},
            );
            reply.end(response);
          }
        });
      } else if (request.url === '/et/runParams') {
        invocations++;
        assert.ok(body.startsWith('ksowebstartupet://'));
        const payload = JSON.parse(
          Buffer.from(
            body.slice('ksowebstartupet://'.length),
            'base64',
          ).toString('utf8'),
        );
        assert.equal(payload.name, own.name);
        assert.equal(payload.function, 'OfficeBridgeProbeOpen');
        reply.writeHead(invocationStatus);
        reply.end(invocationResponse);
      } else {
        unexpected++;
        reply.writeHead(500);
        reply.end();
      }
    });
  });
  server.listen(0, '127.0.0.1');
  await once(server, 'listening');
  const port = await freePort();
  const base = 'http://127.0.0.1:' + port;
  const child = spawn(
    info.testHost,
    [
      info.xll,
      '--relay',
      String(server.address().port),
      String(budget),
      String(port),
    ],
    { windowsHide: true, stdio: 'ignore' },
  );
  let finished = false;
  const exited = once(child, 'exit').then(([code]) => {
    finished = true;
    return code;
  });
  const health = async () => {
    const reply = await fetch(base + '/probe/health', {
      headers,
      signal: AbortSignal.timeout(500),
    });
    assert.equal(reply.status, 200);
    const state = await reply.json();
    assert.equal(state.pid, child.pid);
    return state;
  };
  async function until(predicate, timeout = 4000) {
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      if (finished) throw new Error(name + ': host exited early.');
      try {
        const state = await health();
        if (predicate(state)) return state;
      } catch (error) {
        if (error.code === 'ERR_ASSERTION') throw error;
      }
      await delay(50);
    }
    throw new Error(name + ': timed out waiting for expected native state.');
  }
  try {
    await until(
      (state) => state.registration !== 0 || state.activationError !== 0,
    );
    if (stop) {
      await Promise.race([
        arrived,
        delay(3000).then(() => {
          throw new Error(name + ': no relay POST');
        }),
      ]);
      const start = Date.now();
      assert.equal(
        (
          await fetch(base + '/probe/stop', {
            method: 'POST',
            headers,
            body: '',
          })
        ).status,
        200,
      );
      const code = await Promise.race([
        exited,
        delay(2500).then(() => {
          throw new Error(name + ': stop did not cancel request');
        }),
      ]);
      assert.equal(code, 0);
      assert.ok(Date.now() - start < 2500);
      await delay(approvalDelay + 100);
      assert.equal(invocations, 0);
    } else {
      const state = await until(
        (s) =>
          expectedInvocations ? s.invocation === 1 : s.activationError !== 0,
        Math.max(4000, approvalDelay + 3000),
      );
      if (expectedDeployment !== undefined)
        assert.equal(state.deployment, expectedDeployment);
      if (expectedRegistration !== undefined)
        assert.equal(state.registration, expectedRegistration);
      if (expectedHttp !== undefined)
        assert.equal(state.relayHttpStatus, expectedHttp);
      if (approvalDelay >= budget) await delay(approvalDelay + 150);
      assert.equal(invocations, expectedRequests);
    }
    assert.equal(queries, 1);
    assert.equal(
      deployments,
      expectedDeployments,
      name + ': never resend/overwrite deployment',
    );
    assert.equal(unexpected, 0, name + ': no redirects or unrelated requests');
    console.info('PASS: ' + name);
  } finally {
    if (!finished)
      await fetch(base + '/probe/stop', {
        method: 'POST',
        headers,
        body: '',
      }).catch(() => {});
    await Promise.race([exited, delay(3000)]);
    if (!finished) {
      child.kill();
      await exited;
    }
    for (const timer of timers) clearTimeout(timer);
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
  }
}
await scenario('reuse exact enabled registration; do not redeploy', {
  registry: [own],
  expectedDeployments: 0,
  expectedDeployment: 3,
  expectedRegistration: 2,
});
await scenario('other add-ins remain untouched; missing probe deploys once', {
  registry: [{ ...own, name: 'AnotherAddin' }],
});
await scenario('approval after 11 seconds still activates exactly once', {
  budget: 18000,
  approvalDelay: 11000,
});
await scenario('invocation failure is preserved after registration reuse', {
  registry: [own],
  expectedDeployments: 0,
  expectedInvocations: 0,
  invocationStatus: 406,
  expectedHttp: 406,
  expectedRequests: 1,
});
await scenario('invalid invocation body is never reported as requested', {
  registry: [own],
  expectedDeployments: 0,
  expectedInvocations: 0,
  invocationResponse: 'CANCEL',
  expectedHttp: 200,
  expectedRequests: 1,
});
await scenario('empty 200 response retains publisher success semantics', {
  response: '',
});
await scenario('declined deployment never activates or retries', {
  response: 'CANCEL',
  expectedInvocations: 0,
  expectedDeployment: -2,
});
await scenario(
  'HTTP406 stays visible; no success or retry inferred from installation',
  { status: 406, expectedInvocations: 0, expectedHttp: 406 },
);
await scenario('deadline aborts; late approval never activates', {
  budget: 300,
  approvalDelay: 800,
  expectedInvocations: 0,
});
await scenario(
  'stop during approval cancels promptly; late response never activates',
  { budget: 20000, approvalDelay: 700, stop: true },
);
await scenario('HTTP failure never activates', {
  status: 500,
  expectedInvocations: 0,
  expectedHttp: 500,
});
await scenario('redirect is refused', {
  status: 302,
  redirect: true,
  expectedInvocations: 0,
  expectedHttp: 302,
});
await scenario('connection loss never retries or activates', {
  refused: true,
  expectedInvocations: 0,
});
await scenario('embedded NUL cannot truncate a refusal to OK', {
  response: 'OK\0CANCEL',
  expectedInvocations: 0,
});
await scenario('embedded NUL cannot truncate a refusal to empty', {
  response: '\0CANCEL',
  expectedInvocations: 0,
});
await scenario('oversized relay response is rejected', {
  response: 'X'.repeat(3000),
  expectedInvocations: 0,
});
for (const [field, value] of [
  ['url', 'http://example.invalid/probe/'],
  ['version', '4.0.0'],
  ['addonType', 'wps'],
  ['online', 'false'],
  ['enable', 'false'],
  ['enable', 'unknown'],
])
  await scenario(
    'conflicting registration ' +
      field +
      '=' +
      value +
      ' never installs or opens',
    {
      registry: [{ ...own, [field]: value }],
      expectedDeployments: 0,
      expectedInvocations: 0,
      expectedRegistration: -1,
    },
  );
await scenario('duplicate own registrations are rejected', {
  registry: [own, own],
  expectedDeployments: 0,
  expectedInvocations: 0,
});
await scenario('malformed registry never becomes missing', {
  rawRegistry: '{',
  expectedDeployments: 0,
  expectedInvocations: 0,
});
await scenario('registry HTTP failure never installs or opens', {
  registryStatus: 406,
  expectedDeployments: 0,
  expectedInvocations: 0,
  expectedHttp: 406,
});
console.info(
  'PASS: native registration/deployment doubles on temporary ports; no real WPS registration, workbook access or model requests.',
);
