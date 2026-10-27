/** Interpret only known protocol fields; never echo native messages, URLs or credentials. */
export function workBuddyFailure(
  result: any,
  textCount = 0,
): { code: string; message: string } {
  const objects: any[] = [],
    messages: string[] = [];
  const visit = (value: any, depth = 0) => {
    if (depth > 3) return;
    if (typeof value === 'string') {
      if (value.length > 32768) return;
      messages.push(value);
      try {
        const parsed = JSON.parse(value);
        if (parsed && typeof parsed === 'object') visit(parsed, depth + 1);
      } catch {}
    } else if (value && typeof value === 'object' && !Array.isArray(value)) {
      objects.push(value);
      for (const key of ['data', 'error', 'errorMessage', 'message'])
        if (Object.hasOwn(value, key)) visit(value[key], depth + 1);
    }
  };
  visit(result);
  visit(result?._meta?.['codebuddy.ai/errorMessage']);
  const hasCategory = (...values: string[]) =>
    objects.some((o) => values.includes(o.category));
  const status = (...codes: number[]) =>
    objects.some(
      (o) =>
        codes.includes(o.statusCode) ||
        codes.includes(o.status) ||
        codes.includes(o.httpStatus),
    );
  const restricted = objects.some(
    (o) => o.code === 14003 || o.bizCode === 14003,
  );
  const matches = (pattern: RegExp) => messages.some((m) => pattern.test(m));
  if (result?.stopReason === 'cancelled')
    return { code: 'cancelled', message: '已停止，未发布修改方案。' };
  if (hasCategory('quota') || restricted)
    return {
      code: 'agent_quota',
      message:
        'WorkBuddy 服务端返回请求限制' +
        (status(429) || matches(/\b429\b/) ? '（HTTP 429）' : '') +
        (restricted ? '（服务码 14003）' : '') +
        '，本轮已停止。此错误不足以判断积分余额，请在客户端检查当前模型的限制；未自动重试。',
    };
  if (
    status(429) ||
    hasCategory('rate_limit', 'rate_limit_exceeded') ||
    matches(
      /\b(?:status(?: code)?|HTTP)\s*[:=]?\s*429\b|^\s*429\b|too many requests/i,
    )
  )
    return {
      code: 'agent_rate_limited',
      message:
        'WorkBuddy 请求受限（HTTP 429），本轮已停止。此错误不足以判断积分余额，请稍后检查客户端状态再继续；未自动重试。',
    };
  if (
    status(401) ||
    hasCategory('authentication', 'auth') ||
    matches(/\bunauthorized\b/i)
  )
    return {
      code: 'agent_auth',
      message: 'WorkBuddy 身份验证失败，请在客户端确认登录状态后重试。',
    };
  if (hasCategory('permission') || matches(/^Permission denied for tool/i))
    return {
      code: 'agent_permission',
      message:
        'WorkBuddy 拒绝了本轮工具权限，未发布修改方案。请检查客户端权限设置。',
    };
  if (
    hasCategory('content_policy') ||
    result?.finishReason === 'content_filter'
  )
    return {
      code: 'agent_refused',
      message: 'WorkBuddy 拒绝处理本轮内容，未发布修改方案。请调整要求后重试。',
    };
  if (result?.stopReason === 'max_tokens')
    return {
      code: 'agent_token_limit',
      message:
        'WorkBuddy 本轮输出达到长度限制，未发布未完成的修改方案。请缩小任务或分步处理。',
    };
  if (result?.stopReason === 'max_turn_requests')
    return {
      code: 'agent_turn_limit',
      message:
        'WorkBuddy 本轮达到调用次数上限，未发布未完成的修改方案。请分步处理。',
    };
  if (result?.stopReason === 'end_turn' && !textCount)
    return {
      code: 'agent_empty',
      message: 'WorkBuddy 本轮没有返回有效结果，未发布修改方案。',
    };
  return {
    code: 'agent_failed',
    message:
      'WorkBuddy 未完成本轮处理' +
      (result?.stopReason === 'refusal' ? '（客户端返回 refusal）' : '') +
      '，未发布修改方案。请在 WorkBuddy 查看本轮失败原因。',
  };
}
