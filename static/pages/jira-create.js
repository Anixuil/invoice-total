import { enhanceSelects } from './jira-select.js?v=20261029';

const API = '/api/jira-create';
const DEFAULT_JIRA_URL = 'http://jira.kingsware.cn';
const DEFAULT_LLM_URL = 'https://api.deepseek.com/v1';
const AUTH_EVENT = 'jira-creator-auth-changed';
const AUTH_KEY = 'jira-creator-auth-version';
const categories = { frontend: '前端', backend: '后端', c_algorithm: 'C 算法（探针）', unclassified: '待确认' };
const states = { queued: '等待处理', creating: '正在创建', success: '已创建', failed: '失败', unknown: '结果待核实', ready: '待复核', valid: '校验通过' };
const jobStates = { queued: '等待处理', running: '处理中', done: '已结束', failed: '失败', interrupted: '已中断' };
const esc = value => String(value ?? '').replace(/[&<>"']/g, char => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[char]);
const uid = () => globalThis.crypto?.randomUUID?.() || `r_${Date.now()}_${Math.random().toString(36).slice(2)}`;
const split = value => String(value || '').split(/[,，\n]/).map(part => part.trim()).filter(Boolean);
const dateText = value => new Date(value * 1000).toLocaleString('zh-CN');
const defaultsEmpty = () => ({ project_key: '', issuetype_id: '', priority_id: '', assignee: null, due_date: null, components: [], labels: [] });
const fields = [
  ['project_key', '项目 Key', 'text'], ['issuetype_id', '类型 ID', 'text'], ['priority_id', '优先级 ID', 'text'],
  ['assignee', '经办人账号', 'text'], ['due_date', '截止日期', 'date'], ['components', '组件 ID（逗号分隔）', 'text'], ['labels', '标签（逗号分隔）', 'text'],
];
const input = (name, label, value = '', type = 'text', extra = '') => `<label>${esc(label)}<input name="${esc(name)}" type="${type}" value="${esc(value)}" ${extra}></label>`;
const option = (value, label, selected) => `<option value="${esc(value)}" ${String(value) === String(selected) ? 'selected' : ''}>${esc(label)}</option>`;
const button = (action, label, id = '') => `<button class="ghost-button" type="button" data-action="${action}" data-id="${esc(id)}">${esc(label)}</button>`;

export function createPage(context) {
  const root = context.root;
  const $ = selector => root.querySelector(selector);
  const $$ = selector => [...root.querySelectorAll(selector)];
  const state = { user: null, csrf: '', connections: [], profiles: [], templates: [], rows: [], selected: new Set(),
    defaults: defaultsEmpty(), jobs: [], drafts: [], users: [], offset: 0, epoch: 0, pending: null, activeJob: null, historyJob: null, draftId: null, draftVersion: null };
  let choices = { projects: [], types: [], priorities: [], components: [], users: [] };
  let choiceVersion = 0, userVersion = 0, userSearchTimer = null;
  let destroyed = false;
  let dialogSave = null;
  let dialogVersion = 0;
  const busyKeys = new Set();
  const watching = new Map();
  let polling = false;
  const channel = typeof BroadcastChannel !== 'undefined' ? new BroadcastChannel('jira-creator-auth') : null;

  function notice(message, error = false) {
    if (destroyed) return;
    $('[data-notice]').textContent = message;
    $('[data-notice]').classList.toggle('is-error', error);
  }

  function changed(broadcast = true) {
    if (destroyed) return;
    state.epoch += 1;
    watching.clear();
    if (broadcast) {
      if (channel) channel.postMessage('changed');
      else try { localStorage.setItem(AUTH_KEY, uid()); } catch { /* Session checks also detect account changes. */ }
    }
    window.dispatchEvent(new Event(AUTH_EVENT));
  }

  async function api(path, options = {}) {
    const epoch = state.epoch;
    const headers = { ...(options.headers || {}) };
    if (state.csrf) headers['X-CSRF-Token'] = state.csrf;
    let body = options.body;
    if (body && !(body instanceof FormData)) { body = JSON.stringify(body); headers['Content-Type'] = 'application/json'; }
    const response = await context.fetch(API + path, { ...options, body, headers, cache: 'no-store', credentials: 'same-origin' });
    if (destroyed || epoch !== state.epoch) throw new DOMException('账号上下文已失效', 'AbortError');
    const accountId = response.headers.get('X-Jira-Creator-User');
    if (accountId && state.user && accountId !== state.user.id && !path.startsWith('/auth/')) {
      changed();
      throw new DOMException('账号已切换', 'AbortError');
    }
    let result;
    try { result = await response.json(); } catch { throw new Error(`服务响应异常 (${response.status})`); }
    if (destroyed || epoch !== state.epoch) throw new DOMException('账号上下文已失效', 'AbortError');
    if (!response.ok) {
      if (response.status === 401 && state.user) changed();
      if (response.status === 403 && result.detail === '请先修改初始密码' && state.user) changed();
      const error = new Error(typeof result.detail === 'string' ? result.detail : `操作失败 (${response.status})`);
      error.status = response.status;
      throw error;
    }
    return result;
  }

  async function busy(key, element, operation) {
    if (busyKeys.has(key)) return;
    busyKeys.add(key);
    const old = element?.textContent;
    if (element) { element.disabled = true; element.setAttribute('aria-busy', 'true'); element.textContent = '处理中…'; }
    try { return await operation(); }
    catch (error) { if (error.name !== 'AbortError') notice(error.message || '操作失败', true); }
    finally {
      busyKeys.delete(key);
      if (element && !destroyed) { element.disabled = false; element.removeAttribute('aria-busy'); element.textContent = old; }
    }
  }

  function showTab(name) {
    $$('[data-pane]').forEach(pane => { pane.hidden = pane.dataset.pane !== name; });
    $$('[data-tab]').forEach(tab => { if (tab.dataset.tab === name) tab.setAttribute('aria-current', 'page'); else tab.removeAttribute('aria-current'); });
  }

  function openDialog(title, html, save, submit = '保存') {
    dialogVersion += 1;
    dialogSave = save;
    $('[data-dialog]').classList.remove('jc-drawer');
    $('[data-dialog-title]').textContent = title;
    $('[data-dialog-body]').innerHTML = html;
    $('[data-dialog-error]').textContent = '';
    $('[data-dialog-submit]').textContent = submit;
    $('[data-dialog-submit]').disabled = false;
    $('[data-dialog]').showModal();
  }

  function closeDialog() {
    dialogVersion += 1;
    $('[data-dialog]').close();
    $('[data-dialog-body]').replaceChildren();
    dialogSave = null;
  }

  async function boot() {
    try {
      const response = await api('/auth/me');
      state.user = response.user; state.csrf = response.csrf;
      $('[data-account]').textContent = `${state.user.display_name} · ${state.user.username}${state.user.role === 'admin' ? ' / 管理员' : ''}`;
      $('[data-authenticated]').hidden = false;
      $$('[data-admin]').forEach(element => { element.hidden = state.user.role !== 'admin'; });
      await loadConfiguration(true);
      notice('配置只属于当前账号。AI 提炼后请逐条复核，再提交到 Jira。');
    } catch (error) {
      if (destroyed || error.name === 'AbortError') return;
      $('[data-login]').hidden = !!state.user;
      notice(error.message === '请登录后使用 Jira 批量建单' ? '请使用管理员分配的账号登录。' : error.message, error.message !== '请登录后使用 Jira 批量建单');
    }
  }

  function currentProfile() { return state.profiles.find(record => record.id === $('[data-profile]').value); }

  async function loadConfiguration(initial = false) {
    const [connections, profiles, templates] = await Promise.all([api('/connections'), api('/profiles'), api('/templates')]);
    state.connections = connections; state.profiles = profiles.items; state.templates = templates;
    const selected = initial ? profiles.default_id || profiles.items[0]?.id || '' : $('[data-profile]').value;
    $('[data-profile]').innerHTML = option('', '请选择方案', selected) + profiles.items.map(record => option(record.id, record.name + (record.id === profiles.default_id ? ' · 默认' : ''), selected)).join('');
    renderConfig(profiles.default_id);
    if (initial) applyProfile(); else updateTarget();
  }

  function updateTarget() {
    const profile = currentProfile();
    const jira = state.connections.find(record => record.id === profile?.data.jira_connection_id);
    const model = state.connections.find(record => record.id === profile?.data.llm_connection_id);
    $('[data-target]').textContent = jira ? `目标：${jira.data.base_url} · ${jira.data.username || '尚未填写账号'}` : '先建立 Jira 连接和建单方案，再选择使用的方案。';
    $('[data-model]').textContent = model ? `提炼时发送文字至 ${model.data.base_url} / ${model.data.model}` : '未配置模型，可手工新增事项并提交。';
  }

  function defaultsForm(defaults, prefix = '', lists = true) {
    return fields.map(([name, label, type]) => input(prefix + name, label,
      Array.isArray(defaults[name]) ? defaults[name].join(', ') : defaults[name] || '', type,
      lists && ['project_key', 'issuetype_id', 'priority_id', 'assignee'].includes(name) ? `list="jc-${name}"` : '')).join('');
  }

  function readDefaults(container, prefix = '') {
    const value = { ...defaultsEmpty() };
    fields.forEach(([key]) => {
      const control = container.querySelector(`[name="${prefix + key}"]`);
      if (key === 'components' && control?.multiple) { value[key] = [...control.selectedOptions].map(item => item.value); return; }
      const raw = control?.value.trim() || '';
      value[key] = ['components', 'labels'].includes(key) ? split(raw) : ['assignee', 'due_date'].includes(key) ? raw || null : raw;
    });
    return value;
  }

  function applyProfile() {
    state.defaults = { ...defaultsEmpty(), ...currentProfile()?.data.defaults };
    renderChoiceDefaults();
    updateTarget();
    void refreshChoices();
  }

  function renderChoiceDefaults() {
    const select = (name, label, entries, multiple = false) => {
      const values = multiple ? state.defaults[name] || [] : [state.defaults[name] || ''];
      const available = entries.map(([id, title]) => [String(id), title]);
      for (const value of values) if (value && !available.some(([id]) => id === String(value))) available.push([String(value), `${value}（待核验）`]);
      return `<label>${label}<select name="${name}" ${multiple ? 'multiple' : ''}>${multiple ? '' : '<option value="">请选择 / 不设置</option>'}${available.map(([id,title]) => `<option value="${esc(id)}" ${values.includes(id) ? 'selected' : ''}>${esc(title)}</option>`).join('')}</select></label>`;
    };
    $('[data-defaults]').innerHTML = select('project_key', '项目', choices.projects.map(item => [item.key, `${item.name} (${item.key})`])) +
      select('issuetype_id', '问题类型', choices.types.map(item => [item.id, item.name])) +
      select('priority_id', '优先级', choices.priorities.map(item => [item.id, item.name])) +
      `<div><label>搜索经办人<input data-user-search placeholder="输入姓名或账号搜索"></label>${select('assignee', '经办人', choices.users.map(item => [item.name, `${item.displayName} (${item.name})`]))}</div>` +
      input('due_date', '截止日期', state.defaults.due_date || '', 'date') +
      select('components', '组件（可多选）', choices.components.map(item => [item.id, item.name]), true) +
      input('labels', '标签（可自定义，逗号分隔）', (state.defaults.labels || []).join(', ')) +
      '<p class="jc-muted jc-choice-status" data-choice-status>选择项目后加载可用选项；组件可展开搜索并勾选多项。</p>';
  }

  async function refreshChoices() {
    const ticket = ++choiceVersion;
    userVersion += 1;
    const connectionId = currentProfile()?.data.jira_connection_id;
    const project = state.defaults.project_key || '', type = state.defaults.issuetype_id || '';
    choices = { projects: [], types: [], priorities: [], components: [], users: [] };
    renderChoiceDefaults();
    const status = $('[data-choice-status]');
    if (!connectionId) { status.textContent = '请先选择已关联 Jira 连接的方案。'; return; }
    status.textContent = '正在读取 Jira 可选项…';
    $('[data-defaults]').querySelectorAll('select').forEach(element => { element.disabled = true; });
    try {
      const metadata = await getMetadata(connectionId, project, type);
      if (destroyed || ticket !== choiceVersion) return;
      choices = { ...choices, ...metadata, priorities: metadata.priorities || [], components: metadata.components || [] };
      setLists(metadata);
      renderChoiceDefaults();
      $('[data-choice-status]').textContent = '已加载 Jira 选项；切换项目或类型会自动更新。';
      if (project) await searchUsers('', ticket);
    } catch (error) {
      if (destroyed || ticket !== choiceVersion) return;
      $('[data-choice-status]').textContent = error.message + '；可点击“读取项目与字段”重试。';
    } finally {
      if (!destroyed && ticket === choiceVersion) $('[data-defaults]').querySelectorAll('select').forEach(element => { element.disabled = false; });
    }
  }

  async function searchUsers(query, ticket = choiceVersion) {
    const request = ++userVersion;
    const connectionId = currentProfile()?.data.jira_connection_id, project = state.defaults.project_key;
    if (!connectionId || !project) return;
    try {
      const users = await api('/metadata/users?' + new URLSearchParams({ connection_id: connectionId, project, query }));
      if (destroyed || ticket !== choiceVersion || request !== userVersion) return;
      choices.users = users.filter(item => item.active);
      const select = $('[data-defaults] [name="assignee"]');
      const selected = state.defaults.assignee || '';
      select.innerHTML = option('', '不指定经办人', selected) + choices.users.map(item => option(item.name, `${item.displayName} (${item.name})`, selected)).join('');
      if (selected && !choices.users.some(item => item.name === selected)) select.insertAdjacentHTML('beforeend', option(selected, `${selected}（当前值，提交时核验）`, selected));
      $('[data-choice-status]').textContent = `已找到 ${choices.users.length} 位可分配经办人；可继续输入姓名或账号搜索。`;
    } catch (error) { if (!destroyed && ticket === choiceVersion && request === userVersion) $('[data-choice-status]').textContent = error.message; }
  }

  function renderConfig(defaultId) {
    $('[data-connections]').innerHTML = state.connections.map(record => `<article class="jc-card"><h3>${esc(record.name)} · ${record.data.kind === 'jira' ? 'Jira' : '模型'}</h3><p>${esc(record.data.base_url)}<br>${esc(record.data.kind === 'jira' ? record.data.username : record.data.model)} · ${record.has_secret ? '凭据已设置' : '待填写凭据'}</p><div class="jc-actions">${button('edit-connection', '编辑', record.id)}${button('check-connection', '检查连接', record.id)}${button('delete-connection', '删除', record.id)}</div></article>`).join('') || '<p class="jc-muted">暂无连接。添加自己的 Jira 和模型凭据，或先复制公共模板。</p>';
    $('[data-profiles]').innerHTML = state.profiles.map(record => `<article class="jc-card"><h3>${esc(record.name)}${record.id === defaultId ? ' · 默认' : ''}</h3><p>${esc(record.data.defaults.project_key || '待选择项目')} · ${Object.keys(record.data.members || {}).length} 位成员 · ${(record.data.mappings || []).length} 个字段映射</p><div class="jc-actions">${button('edit-profile', '编辑', record.id)}${button('default-profile', '设为默认', record.id)}${button('delete-profile', '删除', record.id)}</div></article>`).join('') || '<p class="jc-muted">暂无建单方案。</p>';
    $('[data-templates]').innerHTML = state.templates.map(record => `<article class="jc-card"><h3>${esc(record.name)}</h3><p>${esc(record.data.defaults.project_key || '未设置默认项目')} · ${esc(record.data.model || '未设置默认模型')}</p><div class="jc-actions">${button('copy-template', '复制到我的配置', record.id)}${state.user.role === 'admin' ? button('edit-template', '编辑', record.id) + button('delete-template', '删除', record.id) : ''}</div></article>`).join('') || '<p class="jc-muted">管理员尚未发布公共模板。</p>';
  }

  function connectionDialog(record, kind) {
    const data = record?.data || { kind, base_url: kind === 'jira' ? DEFAULT_JIRA_URL : DEFAULT_LLM_URL,
      username: '', model: kind === 'llm' ? 'deepseek-chat' : '', output_mode: 'json_object' };
    openDialog(record ? '编辑连接' : `新增 ${kind === 'jira' ? 'Jira' : '模型'}连接`, `<div class="jc-fields">${input('name', '连接名称', record?.name, 'text', 'required maxlength="100"')}${input('base_url', '服务地址', data.base_url, 'url', 'required placeholder="https://服务地址/v1"')}${data.kind === 'jira' ? input('username', 'Jira 用户名', data.username, 'text', 'required autocomplete="off"') : input('model', '模型名称', data.model, 'text', 'required')}${input('secret', data.kind === 'jira' ? 'Jira 密码' : 'API Key', '', 'password', `autocomplete="new-password" placeholder="${record?.has_secret ? '已设置，留空保留' : '请输入凭据'}"`)}${data.kind === 'llm' ? `<label>结构化输出模式<select name="output_mode">${option('json_object', 'JSON Object', data.output_mode)}${option('json_schema', 'JSON Schema', data.output_mode)}</select></label>` : ''}<label class="jc-check"><input type="checkbox" name="clear_secret">明确清除已保存凭据</label></div><p class="jc-muted">地址须在管理员部署白名单内。模型检查会发送一条简短请求。</p>`, async form => {
      const values = Object.fromEntries(new FormData(form));
      const body = { ...data, ...values, kind: data.kind, clear_secret: form.elements.clear_secret.checked, version: record?.version || null };
      await api('/connections' + (record ? '/' + record.id : ''), { method: record ? 'PUT' : 'POST', body });
      await loadConfiguration(); notice('连接已保存。');
    });
  }

  function memberText(members) { return Object.entries(members || {}).map(([name, account]) => `${name}=${account}`).join('\n'); }
  function readMembers(value) {
    const result = {};
    for (const line of value.split('\n').map(line => line.trim()).filter(Boolean)) {
      const position = line.indexOf('=');
      if (position < 1 || !line.slice(position + 1).trim()) throw new Error('人员字典每行填写“中文名=Jira账号”。');
      result[line.slice(0, position).trim()] = line.slice(position + 1).trim();
    }
    return result;
  }

  function mappingRow(mapping = {}) {
    return `<div class="jc-mapping">${input('mapping_project', '项目 Key', mapping.project_key)}${input('mapping_type', '类型 ID', mapping.issuetype_id)}${input('mapping_id', '字段 ID', mapping.field_id, 'text', 'list="jc-custom-fields" placeholder="customfield_…"')}${input('mapping_label', '显示名称', mapping.label)}<label>字段类型<select name="mapping_kind">${[['text','文本'],['number','数字'],['date','日期'],['user','单人'],['users','多人'],['option','单选'],['options','多选']].map(([value,label]) => option(value,label,mapping.kind || 'text')).join('')}</select></label>${input('mapping_default', '默认值 / ID', Array.isArray(mapping.default) ? mapping.default.join(',') : mapping.default ?? '')}<button type="button" class="ghost-button" data-remove-mapping aria-label="删除映射">×</button></div>`;
  }

  function readMappings(container) {
    return [...container.querySelectorAll('.jc-mapping')].map(row => {
      const get = name => row.querySelector(`[name="mapping_${name}"]`).value.trim();
      const kind = get('kind');
      let value = get('default');
      if (!value) value = null;
      else if (kind === 'number') { value = Number(value); if (!Number.isFinite(value)) throw new Error('字段默认值必须是数字。'); }
      else if (['users', 'options'].includes(kind)) value = split(value);
      return { project_key: get('project'), issuetype_id: get('type'), field_id: get('id'), label: get('label'), kind, default: value };
    });
  }

  function profileDialog(record, template = false) {
    const data = record?.data || { defaults: defaultsEmpty(), members: {}, mappings: [],
      jira_base_url: DEFAULT_JIRA_URL, llm_base_url: DEFAULT_LLM_URL, model: 'deepseek-chat', output_mode: 'json_object' };
    const connectionSelect = (kind, selected, label) => `<label>${label}<select name="${kind}_connection_id">${option('', '暂不关联', selected)}${state.connections.filter(item => item.data.kind === kind).map(item => option(item.id, item.name, selected)).join('')}</select></label>`;
    openDialog(template ? '公共模板 · 不含个人凭据' : '建单方案', `<div class="jc-fields">${input('name', '名称', record?.name || '', 'text', 'required')}${template ? input('jira_base_url', 'Jira 地址', data.jira_base_url || '', 'url') + input('llm_base_url', '模型 API 地址', data.llm_base_url || '', 'url') + input('model', '模型名称', data.model || '') + `<label>输出模式<select name="output_mode">${option('json_object','JSON Object',data.output_mode)}${option('json_schema','JSON Schema',data.output_mode)}</select></label>` : connectionSelect('jira', data.jira_connection_id, 'Jira 连接') + connectionSelect('llm', data.llm_connection_id, '模型连接')}${template ? connectionSelect('jira', '', '读取字段使用的个人连接（不发布）') : ''}</div><h3>默认建单字段</h3><div class="jc-fields" data-profile-defaults>${defaultsForm(data.defaults || defaultsEmpty(), 'default_')}</div><div class="jc-toolbar"><button type="button" class="ghost-button" data-profile-meta>读取当前项目与类型的字段</button><span class="jc-muted" data-meta-notice></span></div><label class="jc-full">人员字典（每行：中文名=真实 Jira 账号）<textarea name="members" rows="5" placeholder="姓名=账号">${esc(memberText(data.members))}</textarea></label><div class="jc-section-head"><h3>自定义字段映射</h3><button type="button" class="ghost-button" data-add-mapping>新增映射</button></div><p class="jc-muted">按项目和类型分别配置。选项默认值填写 Jira 选项 ID，多值以逗号分隔。</p><div class="jc-mapping-wrap" data-mappings>${(data.mappings || []).map(mappingRow).join('')}</div><datalist id="jc-custom-fields"></datalist>`, async form => {
      const formData = new FormData(form);
      const body = { name: formData.get('name'), defaults: readDefaults(form, 'default_'), members: readMembers(formData.get('members')), mappings: readMappings(form), version: record?.version || null };
      if (template) Object.assign(body, { jira_base_url: formData.get('jira_base_url'), llm_base_url: formData.get('llm_base_url'), model: formData.get('model'), output_mode: formData.get('output_mode') });
      else Object.assign(body, { jira_connection_id: formData.get('jira_connection_id') || null, llm_connection_id: formData.get('llm_connection_id') || null });
      const endpoint = template ? '/admin/templates' : '/profiles';
      await api(endpoint + (record ? '/' + record.id : ''), { method: record ? 'PUT' : 'POST', body });
      await loadConfiguration(); notice(template ? '公共模板已保存。' : '建单方案已保存；重新选择方案可载入新的默认值。');
    });
    $('[data-add-mapping]').addEventListener('click', () => {
      const defaults = readDefaults($('[data-dialog-form]'), 'default_');
      $('[data-mappings]').insertAdjacentHTML('beforeend', mappingRow({ project_key: defaults.project_key, issuetype_id: defaults.issuetype_id }));
    });
    $('[data-mappings]').addEventListener('click', event => event.target.closest('[data-remove-mapping]')?.closest('.jc-mapping').remove());
    $('[data-profile-meta]').addEventListener('click', event => busy('profile-meta', event.currentTarget, async () => {
      const form = $('[data-dialog-form]');
      const ticket = dialogVersion;
      const defaults = readDefaults(form, 'default_');
      const metadata = await getMetadata(form.elements.jira_connection_id.value, defaults.project_key, defaults.issuetype_id);
      if (ticket !== dialogVersion) return;
      setLists(metadata);
      $('[data-meta-notice]').textContent = `已读取 ${metadata.projects.length} 个项目、${metadata.types.length} 个类型、${Object.keys(metadata.fields).length} 个字段。`;
      $('#jc-custom-fields').innerHTML = Object.entries(metadata.fields).filter(([id]) => id.startsWith('customfield_')).map(([id, field]) => option(id, `${field.name}${field.required ? ' · 必填' : ''}`, '')).join('');
    }));
  }

  async function getMetadata(connectionId, project = '', issuetype = '') {
    if (!connectionId) throw new Error('请先选择并保存 Jira 连接。');
    return api('/metadata?' + new URLSearchParams({ connection_id: connectionId, project, issuetype }));
  }

  function setLists(metadata) {
    const values = { project_key: metadata.projects.map(item => [item.key, item.name]), issuetype_id: metadata.types.map(item => [item.id, item.name]), priority_id: (metadata.fields.priority?.allowedValues || []).map(item => [item.id, item.name]) };
    Object.entries(values).forEach(([key, list]) => { $(`#jc-${key}`).innerHTML = list.map(([value, name]) => option(value, name, '')).join(''); });
    const members = currentProfile()?.data.members || {};
    $('#jc-assignee').innerHTML = Object.entries(members).map(([name, value]) => option(value, name, '')).join('');
  }

  function addRow(value = {}) {
    if (state.rows.length >= 100) throw new Error('每批最多 100 条事项，请分批处理。');
    const row = { ...defaultsEmpty(), ...readDefaults($('[data-defaults]')), item_id: uid(), summary: '', description: '', category: 'unclassified', category_reason: '', business_module: '', reproduction_steps: [], actual_result: '', expected_result: '', acceptance: '', custom_fields: {}, sources: [], ...value };
    if (state.rows.some(existing => existing.item_id === row.item_id)) return null;
    state.rows.push(row); state.selected.add(row.item_id);
    return row;
  }

  function isLocked(row) { return ['queued', 'creating', 'success', 'unknown'].includes(row._state) || !!state.pending; }
  function editableInput(row, name, type = 'text') {
    return `<input data-field="${name}" type="${type}" value="${esc(row[name] || '')}" ${isLocked(row) ? 'disabled' : ''} ${['project_key', 'issuetype_id', 'priority_id', 'assignee'].includes(name) ? `list="jc-${name}"` : ''} aria-label="${esc(name === 'summary' ? '标题' : fields.find(field => field[0] === name)?.[1])}">`;
  }
  function renderRows() {
    $('[data-rows]').innerHTML = state.rows.map(row => `<tr data-row="${esc(row.item_id)}" class="${state.selected.has(row.item_id) ? 'is-selected' : ''}"><td><input type="checkbox" data-row-select ${state.selected.has(row.item_id) ? 'checked' : ''} ${isLocked(row) ? 'disabled' : ''} aria-label="选择事项"></td><td>${editableInput(row, 'project_key')}</td><td>${editableInput(row, 'issuetype_id')}</td><td>${editableInput(row, 'summary')}</td><td><select data-field="category" aria-label="问题分类" ${isLocked(row) ? 'disabled' : ''}>${Object.entries(categories).map(([key, label]) => option(key, label, row.category || 'unclassified')).join('')}</select></td><td>${editableInput(row, 'priority_id')}</td><td>${editableInput(row, 'assignee')}</td><td>${editableInput(row, 'due_date', 'date')}</td><td><span class="jc-row-status" data-row-status></span>${button('row-detail', '详情', row.item_id)}</td></tr>`).join('');
    state.rows.forEach(updateRowStatus);
    filterRows();
    $('[data-empty]').hidden = state.rows.length > 0;
    $('[data-row-count]').textContent = state.rows.length;
    selectionCount();
  }

  function inCategory(row) { return !$('[data-category-filter]').value || (row.category || 'unclassified') === $('[data-category-filter]').value; }

  function filterRows() {
    const selected = $('[data-category-filter]').value;
    $$('[data-row]').forEach(tr => { const row = state.rows.find(row => row.item_id === tr.dataset.row); tr.hidden = !!selected && (row.category || 'unclassified') !== selected; });
    $('[data-category-counts]').textContent = Object.entries(categories).map(([key, label]) => `${label} ${state.rows.filter(row => (row.category || 'unclassified') === key).length}`).join(' · ');
  }

  function updateRowStatus(row) {
    const tr = $(`[data-row="${row.item_id}"]`);
    if (!tr) return;
    const status = tr.querySelector('[data-row-status]');
    status.dataset.state = row._state || 'ready';
    status.replaceChildren();
    if (row._key && safeUrl(row._url)) {
      const link = document.createElement('a'); link.href = row._url; link.target = '_blank'; link.rel = 'noopener noreferrer'; link.textContent = row._key; status.append(link);
    } else status.textContent = states[row._state] || '待复核';
    if (row._error) status.append(document.createTextNode(' · ' + row._error));
    tr.querySelectorAll('input, select').forEach(element => { element.disabled = isLocked(row); });
  }

  function safeUrl(value) { try { return ['https:', 'http:'].includes(new URL(value).protocol); } catch { return false; } }
  function selectionCount() {
    const count = state.rows.filter(row => state.selected.has(row.item_id) && !isLocked(row) && inCategory(row)).length;
    $('[data-selection]').textContent = `已选 ${count} 条可提交事项`;
    const eligible = state.rows.filter(row => !isLocked(row) && inCategory(row));
    $('[data-select-all]').checked = !!eligible.length && eligible.every(row => state.selected.has(row.item_id));
  }

  function issueData(row) {
    return Object.fromEntries(Object.entries(row).filter(([key]) => !key.startsWith('_')));
  }

  async function rowDialog(row, onSave) {
    const profile = currentProfile();
    const applicable = (profile?.data.mappings || []).filter(mapping => mapping.project_key === row.project_key && mapping.issuetype_id === row.issuetype_id);
    const locked = isLocked(row);
    const custom = applicable.map(mapping => {
      const value = Object.hasOwn(row.custom_fields || {}, mapping.field_id) ? row.custom_fields[mapping.field_id] : mapping.default;
      return input('custom_' + mapping.field_id, mapping.label + (['options', 'users'].includes(mapping.kind) ? '（逗号分隔）' : ''), Array.isArray(value) ? value.join(',') : value ?? '', mapping.kind === 'date' ? 'date' : mapping.kind === 'number' ? 'number' : 'text', `data-custom="${mapping.field_id}" ${locked ? 'disabled' : ''} ${mapping.kind === 'number' ? 'step="any"' : ''}`);
    }).join('');
    openDialog('事项详情', `<p class="jc-muted">${esc(row.summary || '未填写标题')} · ${esc(row.project_key)} / ${esc(row.issuetype_id)}</p><label>详细描述 / 验收标准<textarea name="description" rows="7" maxlength="30000" ${locked ? 'disabled' : ''}>${esc(row.description)}</textarea></label><div class="jc-fields">${input('components', '组件 ID（逗号分隔）', (row.components || []).join(','), 'text', locked ? 'disabled' : '')}${input('labels', '标签（逗号分隔）', (row.labels || []).join(','), 'text', locked ? 'disabled' : '')}${custom}</div><div class="jc-fields">${input('business_module', '业务模块', row.business_module || '', 'text', locked ? 'disabled' : '')}<label>问题分类<select name="category" ${locked ? 'disabled' : ''}>${Object.entries(categories).map(([key,label]) => option(key,label,row.category || 'unclassified')).join('')}</select></label></div>${[['category_reason','分类依据'],['reproduction_steps','复现步骤（每行一步）'],['actual_result','实际结果'],['expected_result','预期结果'],['acceptance','验收标准']].map(([key,label]) => `<label>${label}<textarea name="${key}" rows="3" ${locked ? 'disabled' : ''}>${esc(Array.isArray(row[key]) ? row[key].join('\n') : row[key] || '')}</textarea></label>`).join('')}<p class="jc-muted">来源：${esc((row.sources || []).join('；') || '手工输入')}</p><p class="jc-muted" data-row-meta></p>${locked ? '<p>该事项已提交或结果待核实，请在历史记录中处理。</p>' : '<button type="button" class="ghost-button" data-row-meta-load>读取字段及可分配人员</button>'}`, async form => {
      if (locked) return;
      const values = new FormData(form);
      const customFields = { ...(row.custom_fields || {}) };
      for (const mapping of applicable) {
        const raw = String(values.get('custom_' + mapping.field_id) || '').trim();
        customFields[mapping.field_id] = !raw ? null : mapping.kind === 'number' ? Number(raw) : ['options', 'users'].includes(mapping.kind) ? split(raw) : raw;
      }
      Object.assign(row, { description: values.get('description'), components: split(values.get('components')), labels: split(values.get('labels')), custom_fields: customFields, _state: 'ready', _error: '' });
      Object.assign(row, { category: values.get('category'), business_module: values.get('business_module'), category_reason: values.get('category_reason'), reproduction_steps: String(values.get('reproduction_steps') || '').split('\n').map(value => value.trim()).filter(Boolean), actual_result: values.get('actual_result'), expected_result: values.get('expected_result'), acceptance: values.get('acceptance') });
      if (onSave) await onSave(row); else renderRows();
    }, locked ? '关闭' : '保存');
    $('[data-dialog]').classList.add('jc-drawer');
    $('[data-row-meta-load]')?.addEventListener('click', event => busy('row-metadata', event.currentTarget, async () => {
      const ticket = dialogVersion;
      const connectionId = profile?.data.jira_connection_id;
      const metadata = await getMetadata(connectionId, row.project_key, row.issuetype_id);
      if (ticket !== dialogVersion) return;
      const notes = Object.entries(metadata.fields).map(([id, field]) => `${field.name} (${id})${field.required ? ' 必填' : ''}${field.allowedValues ? '：' + field.allowedValues.map(value => `${value.id}=${value.name || value.value || ''}`).join('、') : ''}`);
      $('[data-row-meta]').textContent = notes.join('；');
      const users = await api('/metadata/users?' + new URLSearchParams({ connection_id: connectionId, project: row.project_key, query: row.assignee || '' }));
      if (ticket !== dialogVersion) return;
      $('#jc-assignee').innerHTML = users.map(value => option(value.name, value.displayName, '')).join('');
    }));
  }

  function bulkDialog() {
    const chosen = state.rows.filter(row => state.selected.has(row.item_id) && !isLocked(row) && inCategory(row));
    if (!chosen.length) throw new Error('请先选择可编辑事项。');
    openDialog(`批量修改 ${chosen.length} 条事项`, `<p class="jc-muted">勾选需要覆盖的字段。未勾选字段保持原值。</p><div class="jc-fields">${fields.map(([key, label, type]) => `<div><label class="jc-check"><input type="checkbox" name="apply_${key}">${esc(label)}</label>${input(key, '', Array.isArray(state.defaults[key]) ? state.defaults[key].join(',') : state.defaults[key] || '', type)}</div>`).join('')}</div>`, async form => {
      const values = readDefaults(form);
      chosen.forEach(row => {
        fields.forEach(([key]) => { if (form.elements['apply_' + key].checked) row[key] = values[key]; });
        row._state = 'ready'; row._error = '';
      });
      renderRows();
    });
  }

  function batchBody() {
    if (!currentProfile()) throw new Error('请先选择建单方案。');
    const issues = state.rows.filter(row => state.selected.has(row.item_id) && !isLocked(row) && inCategory(row)).map(issueData);
    if (!issues.length) throw new Error('请先选择待创建事项。');
    if (issues.some(row => !row.summary.trim())) throw new Error('请填写所有选中事项的标题。');
    return { profile_id: currentProfile().id, defaults: readDefaults($('[data-defaults]')), issues, idempotency_key: uid() };
  }

  async function submitBatch(body) {
    state.pending = body;
    state.rows.forEach(updateRowStatus);
    try {
      const result = await api('/batch-create', { method: 'POST', body });
      state.pending = null; state.activeJob = result.job_id;
      body.issues.forEach(issue => { const row = state.rows.find(row => row.item_id === issue.item_id); if (row) { row._state = 'queued'; updateRowStatus(row); } });
      watch(result.job_id, 'batch'); selectionCount();
      notice('创建任务已提交。页面关闭后任务继续，可在历史记录查看。');
    } catch (error) {
      if (error.status >= 400 && error.status < 500) {
        state.pending = null;
      } else if (error.name !== 'AbortError') {
        // Keep the exact body/key for retry after a lost response.
        error.message = '未能确认提交结果。再次点击提交将查询或重试同一提交。' + error.message;
      }
      throw error;
    } finally {
      if (!destroyed) { state.rows.forEach(updateRowStatus); selectionCount(); }
    }
  }

  function processingProgress(job) {
    const panel = $(`[data-processing="${job.kind}"]`);
    if (!panel) return;
    panel.hidden = false;
    const progress = job.result?.progress;
    const bar = panel.querySelector('progress');
    const text = panel.querySelector('span');
    if (job.status === 'done') {
      bar.value = 100;
      text.textContent = job.kind === 'source' ? `识别完成 · ${job.result.parts?.length || 0} 个证据片段 · ${job.result.errors?.length || 0} 项需核对` : `提炼完成 · ${job.result.issues?.length || 0} 条问题`;
    } else if (['failed', 'interrupted'].includes(job.status)) {
      bar.removeAttribute('value');
      text.textContent = job.error || '处理已中断，请重试';
    } else {
      if (progress?.percent != null) bar.value = progress.percent; else bar.removeAttribute('value');
      text.textContent = progress ? `${progress.detail}${progress.total ? ` · 当前阶段 ${progress.completed}/${progress.total}（${progress.percent}%）` : ''}` : '等待处理资源…';
    }
    panel.classList.toggle('is-error', ['failed', 'interrupted'].includes(job.status));
  }

  function watch(jobId, kind, options = {}) {
    if (['source', 'extract'].includes(kind)) processingProgress({ kind, status: 'queued', result: {} });
    if (kind !== 'history' || !watching.has(jobId)) watching.set(jobId, { kind, ...options });
    void poll();
  }

  async function poll() {
    if (polling || destroyed || !state.user || !context.active) return;
    polling = true;
    try {
      for (const [jobId, target] of watching) {
        const job = await api('/jobs/' + jobId);
        if (job.kind === 'batch') applyBatch(job);
        else if (target.kind !== 'history' || state.historyJob === jobId) processingProgress(job);
        if (state.historyJob === jobId) renderJob(job);
        if (['queued', 'running'].includes(job.status)) continue;
        watching.delete(jobId);
        if (target.kind === 'source') {
          if (job.status === 'done') {
            showSources(job.result);
            if (target.autoExtract && job.result.parts?.length) {
              if (currentProfile()?.data.llm_connection_id) await busy('extract:', $('[data-action="extract"]'), () => action('extract', $('[data-action="extract"]')));
              else notice('图文解析完成。请先选择关联模型的建单方案，再点击 AI 提炼问题。', true);
            }
          } else notice(job.error || '材料解析失败', true);
        } else if (target.kind === 'extract') {
          if (job.status === 'done') {
            const values = job.result.issues || [];
            if (state.pending || target.profileId !== currentProfile()?.id ||
                (target.mode === 'replace' && target.previousRows !== JSON.stringify(state.rows.map(issueData)))) {
              notice('提炼已完成，但工作台内容或方案已变化。结果保存在历史记录，请从历史载入。');
              continue;
            }
            if (target.mode === 'replace') { state.rows = []; state.selected.clear(); }
            if (state.rows.length + values.length > 100) {
              notice('提炼结果与现有行合计超过 100 条，请在历史记录打开结果后分批处理。', true);
            } else { values.forEach(addRow); renderRows(); notice(`提炼完成，共 ${values.length} 条。请复核账号、日期和描述。`); }
          } else notice(job.error || 'AI 提炼失败', true);
        }
      }
    } catch (error) { if (!destroyed && error.name !== 'AbortError') notice('进度读取失败，将继续重试：' + error.message, true); }
    finally { polling = false; }
  }

  function showSources(result) {
    const parts = result.parts || [];
    $('[data-source-summary]').textContent = `原始图文证据（${parts.length} 个片段，可展开核对；不是 Jira 提炼结果）`;
    $('[data-source-results]').innerHTML = parts.map((part, index) => `<details class="jc-source-part" ${part.needs_review ? 'open' : ''}><summary>${esc(part.source)} · ${esc(part.location)}${part.needs_review ? ' · OCR 待核对' : ''}</summary><textarea data-source-text="${index}" data-source-label="${esc(part.source + ' / ' + part.location)}">${esc(part.text)}</textarea></details>`).join('');
    notice((result.errors || []).join('；') || `已解析 ${parts.length} 个内容片段。请核对文字后点击 AI 提炼。`, !!result.errors?.length);
  }

  function rawText() {
    return [$('[data-raw]').value, ...$$('[data-source-text]').map(element => `[来源：${element.dataset.sourceLabel}]\n${element.value}`)].filter(Boolean).join('\n\n');
  }

  function applyBatch(job) {
    job.items.forEach(item => {
      const row = state.rows.find(row => row.item_id === item.item_id);
      if (!row) return;
      Object.assign(row, { _state: item.state, _error: item.error, _key: item.issue_key, _url: item.issue_url });
      if (['success', 'unknown'].includes(item.state)) state.selected.delete(row.item_id);
      updateRowStatus(row);
    });
    const finished = job.counts.success + job.counts.failed + job.counts.unknown;
    const panel = $('[data-progress]'); panel.hidden = false;
    panel.querySelector('progress').value = job.total ? finished / job.total * 100 : 0;
    panel.querySelector('span').textContent = `${jobStates[job.status]} · ${finished}/${job.total} · 成功 ${job.counts.success} · 失败 ${job.counts.failed} · 待核实 ${job.counts.unknown}`;
    selectionCount();
  }

  async function uploadFiles(files) {
    if (!files.length) return;
    if (files.length > 10 || files.some(file => file.size > 20 * 1024 * 1024) || files.reduce((sum, file) => sum + file.size, 0) > 50 * 1024 * 1024) throw new Error('每批最多 10 个文件，单文件 20 MB，总计 50 MB。');
    if ([...watching.values()].some(item => item.kind === 'source')) throw new Error('上一批材料仍在解析，请等待完成。');
    const body = new FormData(); files.forEach(file => body.append('files', file, file.name || '截图.png'));
    const result = await api('/sources', { method: 'POST', body });
    watch(result.job_id, 'source', { autoExtract: $('[data-auto-extract]').checked }); notice('材料已上传，正在本地解析文字。');
  }

  async function loadHistory() {
    const [drafts, jobs] = await Promise.all([api('/drafts'), api('/jobs?offset=' + state.offset)]);
    state.drafts = drafts; state.jobs = jobs;
    $('[data-drafts]').innerHTML = drafts.map(record => `<article class="jc-card"><h3>${esc(record.name)}</h3><p>${dateText(record.updated_at)} · ${record.data.issues.length} 条事项</p><div class="jc-actions">${button('restore-draft', '恢复', record.id)}${button('delete-draft', '删除', record.id)}</div></article>`).join('') || '<p class="jc-muted">没有已保存草稿。</p>';
    $('[data-jobs]').innerHTML = `<div class="jc-list-wrap"><table class="jc-list"><thead><tr><th>时间</th><th>任务</th><th>状态</th><th>操作</th></tr></thead><tbody>${jobs.map(job => `<tr><td>${dateText(job.created_at)}</td><td>${({ batch: '批量建单', extract: 'AI 提炼', source: '材料解析' })[job.kind]}</td><td>${jobStates[job.status]} ${esc(job.error)}</td><td>${button('view-job', '查看', job.id)}</td></tr>`).join('')}</tbody></table></div>`;
    $('[data-action="history-prev"]').disabled = state.offset === 0;
    $('[data-action="history-next"]').disabled = jobs.length < 30;
  }

  function renderJob(job) {
    const content = job.kind === 'batch' ? `<div class="jc-list-wrap"><table class="jc-list"><thead><tr><th>标题</th><th>状态</th><th>结果</th><th>处理</th></tr></thead><tbody>${job.items.map(item => `<tr><td>${esc(item.data.summary)}</td><td data-state="${item.state}">${states[item.state]}</td><td>${item.issue_key && safeUrl(item.issue_url) ? `<a href="${esc(item.issue_url)}" target="_blank" rel="noopener noreferrer">${esc(item.issue_key)}</a>` : esc(item.error)}</td><td>${!['running','queued'].includes(job.status) && ['failed','unknown'].includes(item.state) ? button('resolve-item', '核实 / 重试', item.item_id) : ''}</td></tr>`).join('')}</tbody></table></div>` : `<p>${esc(job.error || (job.status === 'done' ? '结果可载入工作台继续编辑。材料解析记录在 24 小时后清理。' : '正在处理…'))}</p>${job.status === 'done' ? button('load-job-result', '载入工作台', job.id) : ''}`;
    const configuration = job.configuration ? `<p class="jc-muted">${esc(job.configuration.name)} · 版本 ${job.configuration.version} · ${esc(job.configuration.jira_base_url)}</p>` : '';
    $('[data-job-detail]').innerHTML = `<section class="jc-result"><h2>${dateText(job.created_at)} · ${jobStates[job.status]}</h2>${configuration}${content}</section>`;
    state.viewedJob = job;
  }

  function restoreDraft(record) {
    if (state.pending) throw new Error('请先确认上一次提交结果。');
    if (state.rows.length && !confirm('恢复草稿会替换当前工作台内容，继续？')) return;
    $('[data-profile]').value = record.data.profile_id || ''; updateTarget();
    $('[data-raw]').value = record.data.raw_text || ''; $('[data-source-results]').replaceChildren();
    state.defaults = { ...defaultsEmpty(), ...record.data.defaults };
    renderChoiceDefaults();
    state.rows = record.data.issues.map(row => ({ ...row })); state.selected = new Set(state.rows.map(row => row.item_id));
    state.draftId = record.id; state.draftVersion = record.version;
    void refreshChoices();
    renderRows(); showTab('work'); notice('草稿已恢复。已创建事项会在提交时再次检查。');
  }

  async function loadUsers() {
    state.users = await api('/admin/users');
    $('[data-users]').innerHTML = `<div class="jc-list-wrap"><table class="jc-list"><thead><tr><th>账号</th><th>姓名</th><th>角色</th><th>状态</th><th>操作</th></tr></thead><tbody>${state.users.map(user => `<tr><td>${esc(user.username)}</td><td>${esc(user.display_name)}</td><td>${user.role === 'admin' ? '管理员' : '普通用户'}</td><td>${user.active ? '启用' : '停用'}${user.must_change ? ' · 待改密' : ''}</td><td>${button('edit-user', '编辑 / 重置密码', user.id)}</td></tr>`).join('')}</tbody></table></div>`;
  }

  function userDialog(user) {
    openDialog(user ? '编辑账号' : '创建账号', `<div class="jc-fields">${input('username', '账号', user?.username || '', 'text', user ? 'disabled' : 'required pattern="[A-Za-z0-9_.@-]{2,80}"')}${input('display_name', '显示名称', user?.display_name || '', 'text', 'required')}<label>角色<select name="role">${option('user','普通用户',user?.role || 'user')}${option('admin','管理员',user?.role)}</select></label>${input('password', user ? '重置密码（留空不修改）' : '初始密码', '', 'password', 'autocomplete="new-password"')}${user ? `<label class="jc-check"><input type="checkbox" name="active" ${user.active ? 'checked' : ''}>启用账号</label>` : ''}</div><p class="jc-muted">账号密码匹配即可登录。重置密码会撤销已有登录。</p>`, async form => {
      const values = Object.fromEntries(new FormData(form));
      const body = user ? { display_name: values.display_name, role: values.role, active: form.elements.active.checked, password: values.password || null } : values;
      await api('/admin/users' + (user ? '/' + user.id : ''), { method: user ? 'PUT' : 'POST', body });
      if (user?.id === state.user.id && (body.password || !body.active || body.role !== state.user.role)) { changed(); return; }
      await loadUsers(); notice('账号已保存。');
    });
  }

  function passwordDialog() {
    openDialog('修改平台密码', `<div class="jc-fields">${input('current_password','当前密码','','password','autocomplete="current-password"')}${input('password','新密码','','password','autocomplete="new-password"')}</div>`, async form => {
      await api('/auth/password', { method: 'POST', body: Object.fromEntries(new FormData(form)) }); changed();
    }, '修改并重新登录');
  }

  function resolveDialog(itemId) {
    const job = state.viewedJob;
    const item = job?.items.find(value => value.item_id === itemId);
    if (!item) return;
    const resolvedData = { ...defaultsEmpty(), ...job.configuration?.defaults, ...item.data };
    openDialog('核实创建结果', `<p>${esc(item.data.summary)}</p><p class="jc-muted">${esc(item.error)}</p><label>处理方式<select name="action"><option value="link">关联 Jira 中已存在的事项</option><option value="retry">确认未创建，重新提交</option></select></label>${input('issue_key','已有 Issue Key','','text','placeholder="PROJECT-123"')}<label class="jc-check"><input type="checkbox" name="confirmed_not_created">我已到 Jira 核实该事项尚未创建</label><details><summary>重试前修改事项</summary><div class="jc-fields">${input('summary','标题',item.data.summary)}${defaultsForm({ ...defaultsEmpty(), ...item.data }, 'retry_')}</div><label>描述<textarea name="description" rows="5">${esc(item.data.description)}</textarea></label><label>自定义字段（每行 field_id=值；多值逗号分隔）<textarea name="custom" rows="4">${esc(Object.entries(item.data.custom_fields || {}).map(([key,value]) => `${key}=${Array.isArray(value) ? value.join(',') : value ?? ''}`).join('\n'))}</textarea></label></details>`, async form => {
      const values = new FormData(form);
      const custom = {};
      for (const line of String(values.get('custom')).split('\n').filter(line => line.trim())) {
        const index = line.indexOf('='); if (index < 1) throw new Error('自定义字段格式应为 field_id=值。');
        const key = line.slice(0,index).trim(), value = line.slice(index + 1).trim();
        const mapping = job.configuration?.mappings.find(mapping => mapping.field_id === key && mapping.project_key === readDefaults(form, 'retry_').project_key && mapping.issuetype_id === readDefaults(form, 'retry_').issuetype_id);
        custom[key] = ['users', 'options'].includes(mapping?.kind) ? split(value) : mapping?.kind === 'number' && value ? Number(value) : value || null;
      }
      const issue = { ...item.data, ...readDefaults(form, 'retry_'), summary: values.get('summary'), description: values.get('description'), custom_fields: custom };
      await api(`/jobs/${job.id}/items/${itemId}/resolve`, { method: 'POST', body: { action: values.get('action'), issue_key: values.get('issue_key'), confirmed_not_created: form.elements.confirmed_not_created.checked, issue } });
      const latest = await api('/jobs/' + job.id); renderJob(latest); applyBatch(latest);
      if (['queued','running'].includes(latest.status)) watch(job.id, 'batch');
      notice('事项处理状态已更新。');
    }, '确认处理');
    fields.forEach(([key]) => {
      const element = $('[data-dialog-form]').elements['retry_' + key];
      element.value = Array.isArray(resolvedData[key]) ? resolvedData[key].join(',') : resolvedData[key] || '';
    });
  }

  async function action(name, element) {
    const id = element.dataset.id;
    if (name === 'logout') { await api('/auth/logout', { method: 'POST' }); changed(); }
    else if (name === 'clear-work') {
      if (state.pending) throw new Error('请先确认上一次提交结果，再清空工作台。');
      if (!confirm('清空当前材料和表格，开始新的批次？已提交任务仍会继续，可在历史记录查看。')) return;
      state.rows = []; state.selected.clear(); state.draftId = null; state.draftVersion = null; state.activeJob = null;
      watching.clear(); $$('[data-processing]').forEach(panel => { panel.hidden = true; }); $('[data-raw]').value = ''; $('[data-source-results]').replaceChildren(); $('[data-progress]').hidden = true;
      renderRows(); showTab('work'); notice('工作台已清空，可以开始新的批次。');
    }
    else if (name === 'password') passwordDialog();
    else if (name === 'new-jira' || name === 'new-llm') connectionDialog(null, name.slice(4));
    else if (name === 'edit-connection') connectionDialog(state.connections.find(record => record.id === id));
    else if (name === 'check-connection') { const result = await api(`/connections/${id}/check`, { method: 'POST' }); notice(result.message); }
    else if (name === 'delete-connection' || name === 'delete-profile' || name === 'delete-template') {
      if (!confirm('确定删除这项配置？已有批次记录会保留。')) return;
      const group = { 'delete-connection': 'connections', 'delete-profile': 'profiles', 'delete-template': 'admin/templates' }[name];
      await api(`/${group}/${id}`, { method: 'DELETE' }); await loadConfiguration(); notice('配置已删除。');
    }
    else if (name === 'new-profile') profileDialog();
    else if (name === 'edit-profile') profileDialog(state.profiles.find(record => record.id === id));
    else if (name === 'new-template') profileDialog(null, true);
    else if (name === 'edit-template') profileDialog(state.templates.find(record => record.id === id), true);
    else if (name === 'copy-template') { const result = await api(`/templates/${id}/copy`, { method: 'POST' }); await loadConfiguration(); notice(result.message); }
    else if (name === 'default-profile') { await api(`/profiles/${id}/default`, { method: 'POST' }); await loadConfiguration(); notice('默认方案已更新。'); }
    else if (name === 'metadata') { state.defaults = readDefaults($('[data-defaults]')); await refreshChoices(); }
    else if (name === 'add-row') { addRow(); renderRows(); }
    else if (name === 'invert') { state.rows.filter(row => !isLocked(row) && inCategory(row)).forEach(row => state.selected.has(row.item_id) ? state.selected.delete(row.item_id) : state.selected.add(row.item_id)); renderRows(); }
    else if (name === 'delete-rows') { state.rows = state.rows.filter(row => !state.selected.has(row.item_id) || isLocked(row) || !inCategory(row)); renderRows(); }
    else if (name === 'bulk') bulkDialog();
    else if (name === 'row-detail') await rowDialog(state.rows.find(row => row.item_id === id));
    else if (name === 'paste') {
      if (!navigator.clipboard?.read) throw new Error('当前浏览器不支持读取剪贴板，请在材料输入框中 Ctrl+V 粘贴截图，或选择图片文件。');
      const items = await navigator.clipboard.read(); const files = [];
      for (const item of items) for (const type of item.types) if (type.startsWith('image/')) { files.push(new File([await item.getType(type)], '截图.png', { type })); break; }
      if (!files.length) throw new Error('剪贴板没有图片。'); await uploadFiles(files);
    }
    else if (name === 'extract') {
      if (!currentProfile()) throw new Error('请先选择建单方案。');
      if ([...watching.values()].some(item => item.kind === 'extract')) throw new Error('上一次提炼仍在进行，请稍候。');
      const raw = rawText(); if (!raw.trim() || raw.length > 100000) throw new Error('请输入材料，且总文字不超过 10 万字符。');
      const mode = $('[data-extract-mode]').value;
      if (state.pending || (mode === 'replace' && state.rows.some(isLocked))) throw new Error('请先完成或核实当前创建任务，再替换事项。');
      if (mode === 'replace' && state.rows.length && !confirm('提炼完成后将替换当前表格，请先保存需要保留的草稿。继续？')) return;
      const profileId = currentProfile().id;
      const previousRows = JSON.stringify(state.rows.map(issueData));
      const result = await api('/extract', { method: 'POST', body: { profile_id: currentProfile().id, raw_text: raw, defaults: readDefaults($('[data-defaults]')) } });
      watch(result.job_id, 'extract', { mode, profileId, previousRows }); notice('正在调用所选模型提炼事项，结果不会自动创建 Jira。');
    }
    else if (name === 'validate') {
      const body = batchBody();
      const originals = new Map(body.issues.map(item => [item.item_id, JSON.stringify(item)]));
      const result = await api('/validate', { method: 'POST', body });
      result.items.forEach(item => { const row = state.rows.find(row => row.item_id === item.item_id); if (row && !isLocked(row) && originals.get(item.item_id) === JSON.stringify(issueData(row))) { row._state = item.valid ? 'valid' : 'failed'; row._error = item.error; updateRowStatus(row); } });
      notice(`校验完成：${result.items.filter(item => item.valid).length}/${result.items.length} 条通过。`);
    }
    else if (name === 'submit') {
      if (state.pending) { await submitBatch(state.pending); return; }
      const body = batchBody();
      const jira = state.connections.find(record => record.id === currentProfile().data.jira_connection_id);
      openDialog('确认创建 Jira', `<p>即将向 <strong>${esc(jira?.data.base_url || '未配置连接')}</strong> 创建 <strong>${body.issues.length}</strong> 条事项。</p><p>项目：${esc([...new Set(body.issues.map(item => item.project_key))].join('、'))}</p><p class="jc-muted">后端会再次检查字段。成功事项不会因其他行失败而撤回。</p>`, async () => { await submitBatch(body); }, '确认创建');
    }
    else if (name === 'save-draft') {
      openDialog('保存工作台草稿', input('name','草稿名称',`建单草稿 ${new Date().toLocaleDateString('zh-CN')}`,'text','required') + (state.draftId ? '<label class="jc-check"><input type="checkbox" name="overwrite" checked>更新当前恢复的草稿（取消勾选则另存新草稿）</label>' : ''), async form => {
        const overwrite = !!form.elements.overwrite?.checked && state.draftId;
        const saved = await api('/drafts' + (overwrite ? '/' + state.draftId : ''), { method: overwrite ? 'PUT' : 'POST', body: { name: form.elements.name.value, profile_id: currentProfile()?.id || null, raw_text: rawText(), defaults: readDefaults($('[data-defaults]')), issues: state.rows.map(issueData), version: overwrite ? state.draftVersion : null } });
        state.draftId = saved.id; state.draftVersion = saved.version; notice('草稿已保存到当前账号。');
      });
    }
    else if (name === 'refresh-history') await loadHistory();
    else if (name === 'history-prev' || name === 'history-next') { state.offset = Math.max(0, state.offset + (name === 'history-next' ? 30 : -30)); await loadHistory(); }
    else if (name === 'restore-draft') restoreDraft(state.drafts.find(record => record.id === id));
    else if (name === 'delete-draft') { if (!confirm('确定删除这份草稿？')) return; await api('/drafts/' + id, { method: 'DELETE' }); if (state.draftId === id) { state.draftId = null; state.draftVersion = null; } await loadHistory(); }
    else if (name === 'view-job') { state.historyJob = id; const job = await api('/jobs/' + id); renderJob(job); if (['queued','running'].includes(job.status)) watch(id, 'history'); }
    else if (name === 'load-job-result') {
      const job = state.viewedJob;
      if (job.kind === 'source') showSources(job.result);
      else { const values = job.result.issues || []; if (state.rows.length + values.length > 100) throw new Error('合计超过 100 条，请先保存并清理当前事项。'); values.forEach(addRow); renderRows(); }
      showTab('work');
    }
    else if (name === 'new-user') userDialog();
    else if (name === 'edit-user') userDialog(state.users.find(user => user.id === id));
    else if (name === 'resolve-item') resolveDialog(id);
  }

  function mount() {
    const searchableSelects = enhanceSelects(root);
    context.onDeactivate(() => searchableSelects.close());
    context.onDestroy(() => searchableSelects.destroy());
    ['project_key','issuetype_id','priority_id','assignee'].forEach(key => { const list = document.createElement('datalist'); list.id = 'jc-' + key; root.append(list); });
    renderChoiceDefaults();
    root.addEventListener('click', event => {
      const tab = event.target.closest('[data-tab]');
      if (tab) { showTab(tab.dataset.tab); if (tab.dataset.tab === 'history') void busy('history', null, loadHistory); if (tab.dataset.tab === 'admin') void busy('admin', null, loadUsers); return; }
      const element = event.target.closest('[data-action]');
      if (element) void busy(element.dataset.action + ':' + (element.dataset.id || ''), element, () => action(element.dataset.action, element));
    });
    $('[data-login-form]').addEventListener('submit', event => {
      event.preventDefault(); const form = event.currentTarget;
      void busy('login', form.querySelector('button'), async () => { await api('/auth/login', { method: 'POST', body: Object.fromEntries(new FormData(form)) }); form.reset(); changed(); });
    });
    $('[data-force-password]').addEventListener('submit', event => {
      event.preventDefault(); const form = event.currentTarget;
      void busy('password', form.querySelector('button'), async () => { await api('/auth/password', { method: 'POST', body: Object.fromEntries(new FormData(form)) }); form.reset(); changed(); });
    });
    $('[data-dialog-close]').addEventListener('click', closeDialog);
    $('[data-dialog-cancel]').addEventListener('click', closeDialog);
    $('[data-dialog]').addEventListener('cancel', event => { event.preventDefault(); closeDialog(); });
    $('[data-dialog-form]').addEventListener('submit', async event => {
      event.preventDefault(); if (!dialogSave || busyKeys.has('dialog-save')) return;
      const ticket = dialogVersion; const save = dialogSave; const form = event.currentTarget;
      const submit = $('[data-dialog-submit]'); submit.disabled = true; busyKeys.add('dialog-save');
      try { await save(form); if (!destroyed && ticket === dialogVersion) closeDialog(); }
      catch (error) { if (!destroyed && ticket === dialogVersion && error.name !== 'AbortError') $('[data-dialog-error]').textContent = error.message; }
      finally { busyKeys.delete('dialog-save'); if (!destroyed && ticket === dialogVersion) submit.disabled = false; }
    });
    $('[data-category-filter]').addEventListener('change', () => { filterRows(); selectionCount(); });
    $('[data-profile]').addEventListener('change', applyProfile);
    $('[data-defaults]').addEventListener('input', event => {
      state.defaults = readDefaults($('[data-defaults]'));
      if (event.target.matches('[data-user-search]')) {
        userVersion += 1;
        if (userSearchTimer) context.clearTimeout(userSearchTimer);
        const query = event.target.value;
        userSearchTimer = context.setTimeout(() => { void searchUsers(query); }, 300);
      }
    });
    $('[data-defaults]').addEventListener('change', event => {
      state.defaults = readDefaults($('[data-defaults]'));
      if (event.target.name === 'project_key') { state.defaults.issuetype_id = ''; state.defaults.priority_id = ''; state.defaults.components = []; state.defaults.assignee = null; void refreshChoices(); }
      else if (event.target.name === 'issuetype_id') { state.defaults.priority_id = ''; state.defaults.components = []; void refreshChoices(); }
    });
    $('[data-rows]').addEventListener('input', event => {
      const tr = event.target.closest('[data-row]'); if (!tr) return;
      const row = state.rows.find(row => row.item_id === tr.dataset.row); if (!row || isLocked(row)) return;
      if (event.target.matches('[data-row-select]')) {
        if (event.target.checked) state.selected.add(row.item_id); else state.selected.delete(row.item_id);
        tr.classList.toggle('is-selected', event.target.checked); selectionCount();
      } else if (event.target.dataset.field) {
        row[event.target.dataset.field] = event.target.value || (['assignee','due_date'].includes(event.target.dataset.field) ? null : '');
        row._state = 'ready'; row._error = ''; updateRowStatus(row); filterRows();
      }
    });
    $('[data-select-all]').addEventListener('change', event => { state.rows.filter(row => !isLocked(row) && inCategory(row)).forEach(row => event.target.checked ? state.selected.add(row.item_id) : state.selected.delete(row.item_id)); renderRows(); });
    $('[data-files]').addEventListener('change', event => { const files = [...event.target.files]; event.target.value = ''; void busy('upload', null, () => uploadFiles(files)); });
    $('[data-raw]').addEventListener('paste', event => { const files = [...(event.clipboardData?.files || [])]; if (files.length) { event.preventDefault(); void busy('upload', null, () => uploadFiles(files)); } });
    context.visualInterval(poll, 2000);
    context.visualInterval(async () => {
      if (!state.user) return;
      try { const session = await api('/auth/me'); if (session.user.id !== state.user?.id || session.user.must_change !== state.user?.must_change || session.csrf !== state.csrf) changed(); }
      catch (error) { if (error.name !== 'AbortError') notice(error.message, true); }
    }, 60000);
    const storageListener = event => { if (event.key === AUTH_KEY) changed(false); };
    window.addEventListener('storage', storageListener);
    if (channel) channel.onmessage = () => changed(false);
    const resume = async event => {
      if (!context.active || !state.user || (event.type === 'visibilitychange' && document.hidden)) return;
      try { const session = await api('/auth/me'); if (session.user.id !== state.user?.id || session.user.must_change !== state.user?.must_change || session.csrf !== state.csrf) changed(); }
      catch (error) { if (error.name !== 'AbortError') notice(error.message, true); }
    };
    document.addEventListener('visibilitychange', resume);
    window.addEventListener('pageshow', resume);
    context.onDeactivate(() => { if ($('[data-dialog]').open) closeDialog(); });
    context.onDestroy(() => {
      destroyed = true; state.epoch += 1; watching.clear(); channel?.close(); window.removeEventListener('storage', storageListener);
      document.removeEventListener('visibilitychange', resume); window.removeEventListener('pageshow', resume);
      state.rows = []; state.connections = []; state.profiles = []; state.drafts = []; state.pending = null; state.csrf = ''; state.user = null; state.viewedJob = null;
    });
    renderRows(); void boot();
  }
  return context.lifecycle(mount);
}
