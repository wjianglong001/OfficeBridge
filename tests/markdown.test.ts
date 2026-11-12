import assert from 'node:assert/strict';
import test from 'node:test';
import { readFile } from 'node:fs/promises';
import { createContext, runInContext } from 'node:vm';
import { renderMarkdown } from '../apps/wps-addin/markdown.js';

test('Markdown renders Chinese, nested lists, tables, quotes and escaped code', () => {
  const html = renderMarkdown(
    '# 分析结果\n\n**中文🙂**、*重点*、~~旧值~~ 和 `=SUM(A1:A3)`\n下一行\n\n' +
      '1. 第一步\n   - 检查数据\n\n> 数据仅供参考\n\n---\n\n' +
      '| 项目 | 金额 |\n| --- | ---: |\n| 收入 | 120 |\n\n' +
      '```js\nconst 中文 = "<script>";\n```',
  );
  for (const expected of [
    '<h1>分析结果</h1>',
    '<strong>中文🙂</strong>',
    '<em>重点</em>',
    '<s>旧值</s>',
    '<code>=SUM(A1:A3)</code>',
    '<br>',
    '<ol>',
    '<ul>',
    '<blockquote>',
    '<hr>',
    'class="markdown-table"',
    '<table>',
    '<td class="markdown-align-right">120</td>',
    '&lt;script&gt;',
  ])
    assert.ok(html.includes(expected), expected);
  assert.doesNotMatch(html, /style=/);
});

test('Markdown cannot create executable HTML or load remote images', () => {
  const html = renderMarkdown(
    '<script>alert(1)</script>\n<img src=x onerror=alert(1)>\n\n' +
      '[脚本](javascript:alert%281%29) [编码](jav&#x61;script:alert%281%29) ' +
      '[文件](file:///C:/private.txt) [本地接口](/api/v1/session) ' +
      '[数据](data:text/html,hello) [相对](//example.com)\n\n' +
      '![替代文字](https://example.com/tracking.png)\n\n' +
      '[网站](https://example.com) [邮箱](mailto:test@example.com)',
  );
  assert.doesNotMatch(
    html,
    /<(script|img|iframe)|onerror="|href="(?:javascript|file|data|\/)/i,
  );
  assert.ok(html.includes('&lt;script&gt;'));
  assert.ok(html.includes('替代文字'));
  assert.ok(
    html.includes(
      'href="https://example.com" target="_blank" rel="noopener noreferrer"',
    ),
  );
  assert.ok(html.includes('href="mailto:test@example.com"'));
});

test('stream fragments rebuild from raw Markdown and incomplete fences remain readable', () => {
  let text = '';
  for (const chunk of ['## 分', '析\n\n**中', '文**\n\n```js\n', '<tag>\n']) {
    text += chunk;
    assert.doesNotThrow(() => renderMarkdown(text));
  }
  assert.ok(renderMarkdown(text).includes('<strong>中文</strong>'));
  assert.ok(renderMarkdown(text).includes('&lt;tag&gt;'));
  assert.equal(renderMarkdown(''), '');
});

test('offline WPS contract bundle includes Markdown renderer without Node globals', async () => {
  const context = createContext({});
  runInContext(
    await readFile(
      new URL('../apps/wps-addin/.generated/contract.js', import.meta.url),
      'utf8',
    ),
    context,
  );
  assert.equal(
    context.OfficeBridgeContract.renderMarkdown('**离线中文**'),
    '<p><strong>离线中文</strong></p>\n',
  );
});
