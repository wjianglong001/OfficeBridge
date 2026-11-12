import MarkdownIt from 'markdown-it';

// All HTML comes from this parser, never directly from an Agent response.
const markdown = new MarkdownIt({ html: false, breaks: true });
markdown.validateLink = (url) => /^(https?:\/\/|mailto:)/i.test(url);
markdown.renderer.rules.link_open = (
  tokens,
  index,
  options,
  _env,
  renderer,
) => {
  tokens[index]!.attrSet('target', '_blank');
  tokens[index]!.attrSet('rel', 'noopener noreferrer');
  return renderer.renderToken(tokens, index, options);
};
markdown.renderer.rules.image = (tokens, index) =>
  markdown.utils.escapeHtml(tokens[index]!.content);
markdown.renderer.rules.table_open = () =>
  '<div class="markdown-table" tabindex="0" role="region" aria-label="表格，可横向滚动"><table>\n';
markdown.renderer.rules.table_close = () => '</table></div>\n';
// The host CSP blocks inline styles, including markdown-it's table alignment.
for (const name of ['th_open', 'td_open']) {
  markdown.renderer.rules[name] = (tokens, index, options, _env, renderer) => {
    const token = tokens[index]!;
    const alignment = /^text-align:(left|center|right)$/.exec(
      String(token.attrGet('style') || ''),
    );
    token.attrs = token.attrs?.filter(([key]) => key !== 'style') || null;
    if (alignment) token.attrSet('class', 'markdown-align-' + alignment[1]);
    return renderer.renderToken(tokens, index, options);
  };
}
markdown.renderer.rules.fence = (tokens, index) =>
  '<pre tabindex="0" aria-label="代码块，可横向滚动"><code>' +
  markdown.utils.escapeHtml(tokens[index]!.content) +
  '</code></pre>\n';

export function renderMarkdown(source: string): string {
  return markdown.render(source);
}
