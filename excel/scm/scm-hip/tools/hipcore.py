# -*- coding: utf-8 -*-
# HIPオブジェクト：Excelの行(項目キー→文字列) と API JSON の相互変換（modHip の RowToObj / ObjToRow と同じ規則）
# 生成スクリプトのサンプル作成と往復テストに使う
import json

OPS_STR = ['is', 'is_not', 'contains']
OS_VENDORS = ['Microsoft', 'Apple', 'Google', 'Linux', 'Other']
PV_OPS = ['greater_equal', 'greater_than', 'is', 'is_not', 'less_equal', 'less_than', 'contains', 'within', 'not_within']
SEV_OPS = ['greater_equal', 'greater_than', 'is', 'is_not', 'less_equal', 'less_than']
HOST_STR = [('domain', 'hi_domain'), ('host_name', 'hi_hostname'), ('client_version', 'hi_clientver'),
            ('serial_number', 'hi_serial'), ('host_id', 'hi_hostid')]
# (カテゴリ, 接頭辞, インストール, 有効化(is_enabled), ベンダー)
VEND_CATS = [('anti_malware', 'am', True, False), ('disk_backup', 'db', True, False), ('disk_encryption', 'de', True, False),
             ('firewall', 'fw', True, True), ('patch_management', 'pm', True, True), ('data_loss_prevention', 'dlp', True, True)]

def yn(b): return 'yes' if b else 'no'
def isyes(s): return s == 'yes'

# ---------------- 複数行セル ----------------
def lines(s):
    return [x.strip() for x in s.replace('\r', '').split('\n') if x.strip() != '']

def parts(s):
    return [x.strip() for x in s.split('|')]

def neg(s):
    s = s.strip()
    return (True, s[1:].strip()) if s.startswith('!') else (False, s)

def vendor_text(vs):
    out = []
    for v in vs:
        t = v['name']
        if v.get('product'): t += ' | ' + ' ; '.join(v['product'])
        out.append(t)
    return '\n'.join(out)

def vendor_obj(s):
    res = []
    for ln in lines(s):
        p = parts(ln)
        d = {'name': p[0]}
        if len(p) > 1 and p[1] != '':
            d['product'] = [x.strip() for x in p[1].split(';') if x.strip() != '']
        res.append(d)
    return res

# ---------------- JSON → 行 ----------------
def time_cond(o, r, k, units):
    # {not_available:{}} | {within:{days:n}} | {not_within:{hours:n}}
    for op in ('not_available', 'within', 'not_within'):
        if op in o:
            r[k + '_op'] = op
            for u in units:
                if u in (o[op] or {}):
                    r[k + '_unit'] = u; r[k + '_n'] = str(o[op][u])

def obj_to_row(it):
    r = {'description': it.get('description', '')}
    hi = (it.get('host_info') or {}).get('criteria') or {}
    if 'os' in hi:
        c = hi['os'].get('contains', {})
        for v in OS_VENDORS:
            if v in c: r['hi_os_vendor'] = v; r['hi_os_value'] = c[v]
    if 'managed' in hi: r['hi_managed'] = yn(hi['managed'])
    for k, p in HOST_STR:
        if k in hi:
            for op in OPS_STR:
                if op in hi[k]: r[p + '_op'] = op; r[p + '_val'] = hi[k][op]
    for cat, p, _, _ in VEND_CATS:
        if cat not in it: continue
        o = it[cat]; c = o.get('criteria') or {}
        if 'is_installed' in c: r[p + '_installed'] = yn(c['is_installed'])
        if 'is_enabled' in c: r[p + '_enabled'] = c['is_enabled']
        if 'vendor' in o: r[p + '_vendor'] = vendor_text(o['vendor'])
        if 'exclude_vendor' in o: r[p + '_excl'] = yn(o['exclude_vendor'])
        if cat == 'anti_malware':
            if 'real_time_protection' in c: r['am_rtp'] = c['real_time_protection']
            if 'virdef_version' in c: time_cond(c['virdef_version'], r, 'am_virdef', ('days', 'versions'))
            if 'last_scan_time' in c: time_cond(c['last_scan_time'], r, 'am_scan', ('days', 'hours'))
            if 'product_version' in c:
                pv = c['product_version']
                for op in PV_OPS:
                    if op in pv:
                        r['am_pv_op'] = op
                        r['am_pv_val'] = str(pv[op]['versions']) if op in ('within', 'not_within') else pv[op]
        if cat == 'disk_backup' and 'last_backup_time' in c:
            time_cond(c['last_backup_time'], r, 'db_time', ('days', 'hours'))
        if cat == 'disk_encryption' and 'encrypted_locations' in c:
            out = []
            for l in c['encrypted_locations']:
                es = l.get('encryption_state') or {}
                t = l['name']
                if 'is_not' in es: t = '!' + t + ' | ' + es['is_not']
                elif 'is' in es: t += ' | ' + es['is']
                out.append(t)
            r['de_locations'] = '\n'.join(out)
        if cat == 'patch_management' and 'missing_patches' in c:
            mp = c['missing_patches']
            r['pm_check'] = mp.get('check', '')
            for op in SEV_OPS:
                if op in (mp.get('severity') or {}): r['pm_sev_op'] = op; r['pm_sev_val'] = str(mp['severity'][op])
            if mp.get('patches'): r['pm_patches'] = '\n'.join(mp['patches'])
    ce = (it.get('certificate') or {}).get('criteria') or {}
    if 'certificate_profile' in ce: r['ce_profile'] = ce['certificate_profile']
    if ce.get('certificate_attributes'):
        r['ce_attrs'] = '\n'.join(a['name'] + (' | ' + a['value'] if a.get('value', '') != '' else '') for a in ce['certificate_attributes'])
    nw = ((it.get('network_info') or {}).get('criteria') or {}).get('network')
    if nw:
        for op in ('is', 'is_not'):
            if op in nw:
                r['nw_op'] = op
                for t in ('wifi', 'mobile', 'ethernet', 'unknown'):
                    if t in nw[op]:
                        r['nw_type'] = t
                        x = nw[op][t] or {}
                        r['nw_val'] = x.get('ssid', x.get('carrier', ''))
    cc = (it.get('custom_checks') or {}).get('criteria') or {}
    if cc.get('process_list'):
        r['cc_process'] = '\n'.join(('' if p.get('running', True) else '!') + p['name'] for p in cc['process_list'])
    if cc.get('registry_key'):
        out = []
        for k in cc['registry_key']:
            if not k.get('registry_value') and k.get('default_value_data', '') == '':
                out.append(('!' if k.get('negate') else '') + k['name'])
            if k.get('default_value_data', '') != '':
                out.append(k['name'] + ' | (既定) | ' + k['default_value_data'])
            for v in k.get('registry_value') or []:
                out.append(k['name'] + ' | ' + ('!' if v.get('negate') else '') + v['name'] + (' | ' + v['value_data'] if v.get('value_data', '') != '' else ''))
        r['cc_registry'] = '\n'.join(out)
    if cc.get('plist'):
        out = []
        for k in cc['plist']:
            if not k.get('key'):
                out.append(('!' if k.get('negate') else '') + k['name'])
            for v in k.get('key') or []:
                out.append(k['name'] + ' | ' + ('!' if v.get('negate') else '') + v['name'] + (' | ' + v['value'] if v.get('value', '') != '' else ''))
        r['cc_plist'] = '\n'.join(out)
    return {k: v for k, v in r.items() if v != ''}

# ---------------- 行 → JSON ----------------
def has(r, *keys): return any(r.get(k, '') != '' for k in keys)

def time_obj(r, k):
    op = r.get(k + '_op', '')
    if op == 'not_available': return {'not_available': {}}
    return {op: {r[k + '_unit']: int(r[k + '_n'])}}

def row_to_obj(name, r, raw=None):
    o = {'name': name}
    if r.get('description', ''): o['description'] = r['description']
    hk = ['hi_os_vendor', 'hi_os_value', 'hi_managed'] + [p + s for _, p in HOST_STR for s in ('_op', '_val')]
    if has(r, *hk):
        c = {}
        if r.get('hi_os_vendor'):
            c['os'] = {'contains': {r['hi_os_vendor']: r.get('hi_os_value') or 'All'}}
        for k, p in HOST_STR:
            if r.get(p + '_op'): c[k] = {r[p + '_op']: r.get(p + '_val', '')}
        if r.get('hi_managed'): c['managed'] = isyes(r['hi_managed'])
        o['host_info'] = {'criteria': c}
    for cat, p, _, en in VEND_CATS:
        ks = [k for k in FIELD_KEYS if k.startswith(p + '_')]
        if not has(r, *ks): continue
        c = {}
        if r.get(p + '_installed'): c['is_installed'] = isyes(r[p + '_installed'])
        if en and r.get(p + '_enabled'): c['is_enabled'] = r[p + '_enabled']
        if cat == 'anti_malware':
            if r.get('am_rtp'): c['real_time_protection'] = r['am_rtp']
            if r.get('am_virdef_op'): c['virdef_version'] = time_obj(r, 'am_virdef')
            if r.get('am_scan_op'): c['last_scan_time'] = time_obj(r, 'am_scan')
            if r.get('am_pv_op'):
                op = r['am_pv_op']
                c['product_version'] = {op: {'versions': int(r['am_pv_val'])}} if op in ('within', 'not_within') else {op: r.get('am_pv_val', '')}
        if cat == 'disk_backup' and r.get('db_time_op'): c['last_backup_time'] = time_obj(r, 'db_time')
        if cat == 'disk_encryption' and r.get('de_locations'):
            locs = []
            for ln in lines(r['de_locations']):
                n_, s = neg(ln); p_ = parts(s)
                st = p_[1] if len(p_) > 1 and p_[1] != '' else 'encrypted'
                locs.append({'name': p_[0], 'encryption_state': {('is_not' if n_ else 'is'): st}})
            c['encrypted_locations'] = locs
        if cat == 'patch_management' and has(r, 'pm_check', 'pm_sev_op', 'pm_patches'):
            mp = {'check': r.get('pm_check') or 'has-any'}
            if r.get('pm_sev_op'): mp['severity'] = {r['pm_sev_op']: int(r['pm_sev_val'])}
            if r.get('pm_patches'): mp['patches'] = lines(r['pm_patches'])
            c['missing_patches'] = mp
        d = {}
        if c: d['criteria'] = c
        if r.get(p + '_vendor'): d['vendor'] = vendor_obj(r[p + '_vendor'])
        if r.get(p + '_excl'): d['exclude_vendor'] = isyes(r[p + '_excl'])
        o[cat] = d
    if has(r, 'ce_profile', 'ce_attrs'):
        c = {}
        if r.get('ce_profile'): c['certificate_profile'] = r['ce_profile']
        if r.get('ce_attrs'):
            at = []
            for ln in lines(r['ce_attrs']):
                p_ = parts(ln); a = {'name': p_[0]}
                if len(p_) > 1 and p_[1] != '': a['value'] = p_[1]
                at.append(a)
            c['certificate_attributes'] = at
        o['certificate'] = {'criteria': c}
    if has(r, 'nw_op', 'nw_type', 'nw_val'):
        t = r['nw_type']; x = {}
        if t == 'wifi' and r.get('nw_val'): x = {'ssid': r['nw_val']}
        if t == 'mobile' and r.get('nw_val'): x = {'carrier': r['nw_val']}
        o['network_info'] = {'criteria': {'network': {r.get('nw_op') or 'is': {t: x}}}}
    if has(r, 'cc_process', 'cc_registry', 'cc_plist'):
        c = {}
        if r.get('cc_process'):
            c['process_list'] = [{'name': neg(l)[1], 'running': not neg(l)[0]} for l in lines(r['cc_process'])]
        if r.get('cc_registry'):
            keys = {}
            for ln in lines(r['cc_registry']):
                p_ = parts(ln)
                n_, kname = neg(p_[0])
                k = keys.setdefault(kname, {'name': kname})
                if len(p_) == 1:
                    if n_: k['negate'] = True
                elif p_[1] == '(既定)':
                    k['default_value_data'] = p_[2] if len(p_) > 2 else ''
                else:
                    vn, vname = neg(p_[1]); v = {'name': vname}
                    if vn: v['negate'] = True
                    if len(p_) > 2 and p_[2] != '': v['value_data'] = p_[2]
                    k.setdefault('registry_value', []).append(v)
            c['registry_key'] = list(keys.values())
        if r.get('cc_plist'):
            keys = {}
            for ln in lines(r['cc_plist']):
                p_ = parts(ln)
                n_, kname = neg(p_[0])
                k = keys.setdefault(kname, {'name': kname})
                if len(p_) == 1:
                    if n_: k['negate'] = True
                else:
                    vn, vname = neg(p_[1]); v = {'name': vname}
                    if vn: v['negate'] = True
                    if len(p_) > 2 and p_[2] != '': v['value'] = p_[2]
                    k.setdefault('key', []).append(v)
            c['plist'] = list(keys.values())
        o['custom_checks'] = {'criteria': c}
    if raw and 'mobile_device' in raw:      # このツールで編集しない条件は既存の内容を引き継ぐ
        o['mobile_device'] = raw['mobile_device']
    return o

# ---------------- 差分確認用の表示（カテゴリ単位、キーを並べ替えて1行1値） ----------------
DIFF_CATS = ['host_info', 'anti_malware', 'disk_backup', 'disk_encryption', 'firewall', 'patch_management',
             'data_loss_prevention', 'certificate', 'network_info', 'custom_checks', 'mobile_device']

def flat(v, pre=''):
    out = []
    if isinstance(v, dict):
        if not v: out.append(pre + ': {}' if pre else '{}')
        for k in sorted(v.keys()):
            if k == 'criteria': out += flat(v[k], pre)
            else: out += flat(v[k], (pre + '.' if pre else '') + k)
    elif isinstance(v, list):
        for i, x in enumerate(v, 1):
            if isinstance(x, dict) and 'name' in x:
                rest = {k: y for k, y in x.items() if k != 'name'}
                p = f'{pre}[{x["name"]}]'
                out += flat(rest, p) if rest else [p]
            else:
                out.append(f'{pre}[{i}]: {fmt(x)}')
    else:
        out.append(f'{pre}: {fmt(v)}')
    return out

def fmt(x):
    if x is True: return 'true'
    if x is False: return 'false'
    return str(x)

def flat_text(v): return '\n'.join(flat(v)) if v is not None else ''

# 項目キー（③② D列以降の順。modHip と生成スクリプトで共通）
FIELD_KEYS = [
    'description',
    'hi_os_vendor', 'hi_os_value', 'hi_managed', 'hi_domain_op', 'hi_domain_val', 'hi_hostname_op', 'hi_hostname_val',
    'hi_clientver_op', 'hi_clientver_val', 'hi_serial_op', 'hi_serial_val', 'hi_hostid_op', 'hi_hostid_val',
    'am_installed', 'am_rtp', 'am_virdef_op', 'am_virdef_unit', 'am_virdef_n', 'am_scan_op', 'am_scan_unit', 'am_scan_n',
    'am_pv_op', 'am_pv_val', 'am_vendor', 'am_excl',
    'db_installed', 'db_time_op', 'db_time_unit', 'db_time_n', 'db_vendor', 'db_excl',
    'de_installed', 'de_locations', 'de_vendor', 'de_excl',
    'fw_installed', 'fw_enabled', 'fw_vendor', 'fw_excl',
    'pm_installed', 'pm_enabled', 'pm_check', 'pm_sev_op', 'pm_sev_val', 'pm_patches', 'pm_vendor', 'pm_excl',
    'dlp_installed', 'dlp_enabled', 'dlp_vendor', 'dlp_excl',
    'ce_profile', 'ce_attrs',
    'nw_op', 'nw_type', 'nw_val',
    'cc_process', 'cc_registry', 'cc_plist',
]
