/***************
 * Kotobuku HubSpot Sync (NEW PROJECT)
 ***************/

const HS = (() => {
  function getProps_() {
    const p = PropertiesService.getScriptProperties();
    const token = p.getProperty('HUBSPOT_TOKEN');
    const base = p.getProperty('HUBSPOT_BASE') || 'https://api.hubapi.com';
    if (!token) throw new Error('Script Properties に HUBSPOT_TOKEN が未設定です');
    return { token, base };
  }

  function request_(method, path, payload) {
    const { token, base } = getProps_();
    const url = base.replace(/\/$/, '') + path;
    const opt = {
      method,
      muteHttpExceptions: true,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
    };
    if (payload !== undefined) opt.payload = JSON.stringify(payload);

    const res = UrlFetchApp.fetch(url, opt);
    const code = res.getResponseCode();
    const text = res.getContentText();

    if (code < 200 || code >= 300) {
      throw new Error(`HubSpot API Error ${code}: ${text}`);
    }
    return text ? JSON.parse(text) : {};
  }

  return {
    getCompanyById: (id, properties) =>
      request_('get', `/crm/v3/objects/companies/${encodeURIComponent(id)}${properties?.length ? `?properties=${properties.map(encodeURIComponent).join('&properties=')}` : ''}`),
    searchCompanyByProperty: (propName, value, properties) =>
      request_('post', `/crm/v3/objects/companies/search`, {
        filterGroups: [{ filters: [{ propertyName: propName, operator: 'EQ', value: String(value) }] }],
        properties: properties || [],
        limit: 1,
      }),
    createCompany: (propertiesObj) =>
      request_('post', `/crm/v3/objects/companies`, { properties: propertiesObj }),
    updateCompany: (id, propertiesObj) =>
      request_('patch', `/crm/v3/objects/companies/${encodeURIComponent(id)}`, { properties: propertiesObj }),
  };
})();

/**
 * ✅ 疎通テスト（read）
 * 実行時に hs_object_id を聞く（入力ダイアログ）
 */
function test_read_company_by_id() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('HubSpot Company hs_object_id を入力してください（数値）', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;

  const id = r.getResponseText().trim();
  if (!id) return ui.alert('IDが空です');

  const data = HS.getCompanyById(id, ['name', 'company_id']);
  ui.alert(`OK\nid=${data.id}\nname=${data.properties?.name || ''}\ncompany_id=${data.properties?.company_id || ''}`);
}
