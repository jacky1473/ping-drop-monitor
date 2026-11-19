#!/usr/bin/env python3
# ==============================================================================
# 24x7 Multi-IP Drop-Only Ping Monitor & Web GUI (Port 8084)
# Monitors multiple IPs on a single network interface concurrently
# Saves records per IP and generates:
#   - Individual report: /ping/report_<IP>.xls
#   - Master consolidated report: /ping/report.xls (separate sheet per target)
# NOTE: Only records drops / outages. Reachable pings are not logged.
# ==============================================================================

import os
import sys
import time
import subprocess
import sqlite3
import signal
import re
import json
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

from flask import Flask, render_template, jsonify, send_file, request
import xlwt

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
TARGETS_FILE = os.path.join(SCRIPT_DIR, "targets.json")
EXCEL_MASTER_PATH = os.path.join(SCRIPT_DIR, "report.xls")
DB_PATH = os.path.join(SCRIPT_DIR, "ping_history.db")
TEMPLATES_DIR = os.path.join(SCRIPT_DIR, "templates")

PING_INTERVAL = 10     # seconds between probe cycles
PACKETS_PER_PROBE = 3  # packets per probe
WEB_PORT = 8084

app = Flask(__name__, template_folder=TEMPLATES_DIR)

# Global Targets & Runtime State
_targets = []
_target_states = {}  # keyed by IP: {'probes', 'sent', 'lost', 'last_rtt', 'status', 'last_check', 'current_outage'}
_lock = threading.Lock()
running = True

def sanitize_ip(ip):
    return re.sub(r'[^a-zA-Z0-9_\-\.]', '_', ip)

def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS meta_stats (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        # Drop Events table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS outages (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                target_ip TEXT,
                target_name TEXT,
                start_time TEXT,
                end_time TEXT,
                duration_sec REAL,
                packets_lost INTEGER,
                max_loss_pct REAL,
                status TEXT,
                impact_summary TEXT,
                date_str TEXT,
                hour_str TEXT
            )
        """)
        # Migrations for existing DB if needed
        for col in ["target_ip", "target_name", "date_str", "hour_str"]:
            try:
                conn.execute(f"ALTER TABLE outages ADD COLUMN {col} TEXT")
            except Exception:
                pass

        conn.execute("CREATE INDEX IF NOT EXISTS idx_outages_ip ON outages (target_ip)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_outages_time ON outages (start_time)")
        conn.commit()

def load_targets():
    global _targets, _target_states
    default_targets = [{"ip": "192.168.2.13", "name": "SANS PBX", "enabled": True}]
    if not os.path.exists(TARGETS_FILE):
        with open(TARGETS_FILE, 'w') as f:
            json.dump(default_targets, f, indent=2)
        loaded = default_targets
    else:
        try:
            with open(TARGETS_FILE, 'r') as f:
                loaded = json.load(f)
        except Exception:
            loaded = default_targets

    with _lock:
        _targets = loaded
        with get_db() as conn:
            for t in _targets:
                ip = t['ip']
                if ip not in _target_states:
                    row_probes = conn.execute("SELECT value FROM meta_stats WHERE key=?", (f"probes_{ip}",)).fetchone()
                    row_sent = conn.execute("SELECT value FROM meta_stats WHERE key=?", (f"sent_{ip}",)).fetchone()
                    row_lost = conn.execute("SELECT value FROM meta_stats WHERE key=?", (f"lost_{ip}",)).fetchone()

                    _target_states[ip] = {
                        'ip': ip,
                        'name': t.get('name', ip),
                        'probes': int(row_probes[0]) if row_probes else 0,
                        'sent': int(row_sent[0]) if row_sent else 0,
                        'lost': int(row_lost[0]) if row_lost else 0,
                        'last_rtt': 0.0,
                        'status': 'OK',
                        'last_check': None,
                        'current_outage': None
                    }

def save_targets():
    with _lock:
        with open(TARGETS_FILE, 'w') as f:
            json.dump(_targets, f, indent=2)

def save_meta_stats_for_ip(ip):
    try:
        with _lock:
            st = _target_states.get(ip)
            if not st:
                return
            p, s, l = st['probes'], st['sent'], st['lost']

        with get_db() as conn:
            conn.execute("INSERT OR REPLACE INTO meta_stats (key, value) VALUES (?, ?)", (f"probes_{ip}", str(p)))
            conn.execute("INSERT OR REPLACE INTO meta_stats (key, value) VALUES (?, ?)", (f"sent_{ip}", str(s)))
            conn.execute("INSERT OR REPLACE INTO meta_stats (key, value) VALUES (?, ?)", (f"lost_{ip}", str(l)))
            conn.commit()
    except Exception as e:
        print(f"[!] Error saving stats for {ip}: {e}", file=sys.stderr)

def ping_single_ip(target):
    ip = target['ip']
    name = target.get('name', ip)
    cmd = ["ping", "-c", str(PACKETS_PER_PROBE), "-W", "1", "-i", "0.3", ip]

    try:
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True, timeout=5)
        out = p.stdout
    except Exception:
        return {
            'ip': ip,
            'name': name,
            'sent': PACKETS_PER_PROBE,
            'received': 0,
            'loss_pct': 100.0,
            'rtt_avg': 0.0,
            'status': 'DOWN'
        }

    sent = PACKETS_PER_PROBE
    received = 0
    loss_pct = 100.0
    rtt_avg = 0.0

    m_pkt = re.search(r'(\d+)\s+packets transmitted,\s+(\d+)\s+(?:packets\s+)?received,\s+([\d\.]+)%\s+packet loss', out)
    if m_pkt:
        sent = int(m_pkt.group(1))
        received = int(m_pkt.group(2))
        loss_pct = float(m_pkt.group(3))

    m_rtt = re.search(r'rtt min/avg/max/mdev\s*=\s*[\d\.]+/([\d\.]+)/[\d\.]+', out)
    if m_rtt:
        rtt_avg = float(m_rtt.group(1))

    if received == 0:
        status = 'DOWN'
    elif loss_pct > 0:
        status = 'PACKET_DROP'
    else:
        status = 'OK'

    return {
        'ip': ip,
        'name': name,
        'sent': sent,
        'received': received,
        'loss_pct': loss_pct,
        'rtt_avg': rtt_avg,
        'status': status
    }

def record_single_result(res):
    now = datetime.now()
    ts_str = now.strftime('%Y-%m-%d %H:%M:%S')
    date_str = now.strftime('%Y-%m-%d')
    hour_str = now.strftime('%H:00')
    ip = res['ip']
    name = res['name']
    lost = res['sent'] - res['received']
    is_drop = (res['status'] in ('DOWN', 'PACKET_DROP'))

    with _lock:
        if ip not in _target_states:
            _target_states[ip] = {
                'ip': ip,
                'name': name,
                'probes': 0,
                'sent': 0,
                'lost': 0,
                'last_rtt': 0.0,
                'status': 'OK',
                'last_check': None,
                'current_outage': None
            }
        st = _target_states[ip]
        st['probes'] += 1
        st['sent'] += res['sent']
        st['lost'] += lost
        st['last_rtt'] = res['rtt_avg']
        st['last_check'] = ts_str
        st['status'] = res['status']
        st['name'] = name

        outage = st['current_outage']

    # ONLY RECORD IF DROP / OUTAGE
    with get_db() as conn:
        if is_drop:
            if outage is None:
                new_outage = {
                    'start_time': ts_str,
                    'start_dt': now,
                    'packets_lost': lost,
                    'max_loss_pct': res['loss_pct'],
                    'status': res['status'],
                    'date_str': date_str,
                    'hour_str': hour_str
                }
                with _lock:
                    _target_states[ip]['current_outage'] = new_outage
            else:
                outage['packets_lost'] += lost
                if res['loss_pct'] > outage['max_loss_pct']:
                    outage['max_loss_pct'] = res['loss_pct']
                if res['status'] == 'DOWN':
                    outage['status'] = 'DOWN'
        else:
            # Recovery from outage
            if outage is not None:
                dur = (now - outage['start_dt']).total_seconds()
                if dur < 1.0:
                    dur = PING_INTERVAL
                summary = f"Downtime {dur:.1f}s. Severity: {outage['status']} ({outage['packets_lost']} pkts lost, Max Loss {outage['max_loss_pct']:.0f}%)"
                conn.execute("""
                    INSERT INTO outages (target_ip, target_name, start_time, end_time, duration_sec, packets_lost, max_loss_pct, status, impact_summary, date_str, hour_str)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    ip,
                    name,
                    outage['start_time'],
                    ts_str,
                    dur,
                    outage['packets_lost'],
                    outage['max_loss_pct'],
                    outage['status'],
                    summary,
                    outage['date_str'],
                    outage['hour_str']
                ))
                conn.commit()
                with _lock:
                    _target_states[ip]['current_outage'] = None

# ==============================================================================
# Excel Generator (Individual and Master)
# ==============================================================================
def get_excel_styles():
    f_title = xlwt.Font()
    f_title.name = 'Arial'
    f_title.bold = True
    f_title.height = 260
    f_title.colour_index = 4

    f_bold = xlwt.Font()
    f_bold.name = 'Arial'
    f_bold.bold = True

    f_norm = xlwt.Font()
    f_norm.name = 'Arial'

    f_white_bold = xlwt.Font()
    f_white_bold.name = 'Arial'
    f_white_bold.bold = True
    f_white_bold.colour_index = 1

    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    al_center = xlwt.Alignment()
    al_center.horz = xlwt.Alignment.HORZ_CENTER
    al_center.vert = xlwt.Alignment.VERT_CENTER

    al_left = xlwt.Alignment()
    al_left.horz = xlwt.Alignment.HORZ_LEFT
    al_left.vert = xlwt.Alignment.VERT_CENTER

    al_right = xlwt.Alignment()
    al_right.horz = xlwt.Alignment.HORZ_RIGHT
    al_right.vert = xlwt.Alignment.VERT_CENTER

    p_header = xlwt.Pattern()
    p_header.pattern = xlwt.Pattern.SOLID_PATTERN
    p_header.pattern_fore_colour = 4  # Dark Blue

    p_subhead = xlwt.Pattern()
    p_subhead.pattern = xlwt.Pattern.SOLID_PATTERN
    p_subhead.pattern_fore_colour = 22  # Grey/Blue

    p_yellow = xlwt.Pattern()
    p_yellow.pattern = xlwt.Pattern.SOLID_PATTERN
    p_yellow.pattern_fore_colour = 43  # Pale Yellow

    p_red = xlwt.Pattern()
    p_red.pattern = xlwt.Pattern.SOLID_PATTERN
    p_red.pattern_fore_colour = 29  # Light Red

    style_header = xlwt.XFStyle()
    style_header.font = f_white_bold
    style_header.pattern = p_header
    style_header.alignment = al_center
    style_header.borders = borders

    style_subhead = xlwt.XFStyle()
    style_subhead.font = f_bold
    style_subhead.pattern = p_subhead
    style_subhead.alignment = al_left
    style_subhead.borders = borders

    style_cell_center = xlwt.XFStyle()
    style_cell_center.font = f_norm
    style_cell_center.alignment = al_center
    style_cell_center.borders = borders

    style_cell_left = xlwt.XFStyle()
    style_cell_left.font = f_norm
    style_cell_left.alignment = al_left
    style_cell_left.borders = borders

    style_cell_right = xlwt.XFStyle()
    style_cell_right.font = f_norm
    style_cell_right.alignment = al_right
    style_cell_right.borders = borders

    style_kpi_label = xlwt.XFStyle()
    style_kpi_label.font = f_bold
    style_kpi_label.alignment = al_left
    style_kpi_label.borders = borders

    style_kpi_val = xlwt.XFStyle()
    style_kpi_val.font = f_bold
    style_kpi_val.alignment = al_right
    style_kpi_val.borders = borders

    style_drop = xlwt.XFStyle()
    style_drop.font = f_bold
    style_drop.pattern = p_yellow
    style_drop.alignment = al_center
    style_drop.borders = borders

    style_down = xlwt.XFStyle()
    style_down.font = f_bold
    style_down.pattern = p_red
    style_down.alignment = al_center
    style_down.borders = borders

    style_title = xlwt.XFStyle()
    style_title.font = f_title

    return {
        'title': style_title,
        'header': style_header,
        'subhead': style_subhead,
        'cell_center': style_cell_center,
        'cell_left': style_cell_left,
        'cell_right': style_cell_right,
        'kpi_label': style_kpi_label,
        'kpi_val': style_kpi_val,
        'drop': style_drop,
        'down': style_down
    }

def generate_individual_excel(target_ip, target_name):
    styles = get_excel_styles()
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet("Drop Events Log", cell_overwrite_ok=True)
    ws.show_grid = 1

    with _lock:
        st = _target_states.get(target_ip, {})
        probes = st.get('probes', 0)
        sent = st.get('sent', 0)
        lost = st.get('lost', 0)
        last_check = st.get('last_check', '--')
        last_rtt = st.get('last_rtt', 0.0)

    drop_rate = (lost / sent * 100.0) if sent > 0 else 0.0
    uptime = 100.0 - drop_rate

    with get_db() as conn:
        outages = conn.execute("SELECT * FROM outages WHERE target_ip=? ORDER BY id DESC", (target_ip,)).fetchall()
        tot_sec = conn.execute("SELECT SUM(duration_sec) FROM outages WHERE target_ip=?", (target_ip,)).fetchone()[0] or 0.0
        tot_hours = tot_sec / 3600.0

    ws.write(0, 0, f"DROP IMPACT REPORT: {target_name} ({target_ip})", styles['title'])
    ws.write(2, 0, "Target IP:", styles['kpi_label'])
    ws.write(2, 1, target_ip, styles['kpi_val'])
    ws.write(3, 0, "Target Name:", styles['kpi_label'])
    ws.write(3, 1, target_name, styles['kpi_val'])
    ws.write(4, 0, "Total Probes Evaluated:", styles['kpi_label'])
    ws.write(4, 1, f"{probes:,}", styles['kpi_val'])
    ws.write(5, 0, "Uptime Rate:", styles['kpi_label'])
    ws.write(5, 1, f"{uptime:.2f}%", styles['kpi_val'])

    ws.write(2, 3, "Total Packets Sent:", styles['kpi_label'])
    ws.write(2, 4, f"{sent:,}", styles['kpi_val'])
    ws.write(3, 3, "Total Packets Lost:", styles['kpi_label'])
    ws.write(3, 4, f"{lost:,} ({drop_rate:.2f}%)", styles['kpi_val'])
    ws.write(4, 3, "Drop Incidents Count:", styles['kpi_label'])
    ws.write(4, 4, str(len(outages)), styles['kpi_val'])
    ws.write(5, 3, "Total Drop Impact Hours:", styles['kpi_label'])
    ws.write(5, 4, f"{tot_sec:.1f}s ({tot_hours:.4f}h)", styles['kpi_val'])

    # Table
    ws.write(8, 0, "RECORDED DROP INCIDENTS LOG (REACHABLE PINGS EXCLUDED)", styles['subhead'])
    headers = ["#", "Date", "Start Time", "End Time", "Duration (Sec)", "Duration (Mins)", "Impact (Hours)", "Loss %", "Packets Lost", "Severity", "Impact Summary"]
    for idx, h in enumerate(headers):
        ws.write(9, idx, h, styles['header'])

    r = 10
    if not outages:
        ws.write(r, 0, "Zero drops recorded for this IP. 100% reachable.", styles['cell_left'])
    else:
        for o in outages:
            ws.write(r, 0, str(o['id']), styles['cell_center'])
            ws.write(r, 1, str(o['date_str']), styles['cell_center'])
            ws.write(r, 2, str(o['start_time']), styles['cell_center'])
            ws.write(r, 3, str(o['end_time'] or 'ONGOING'), styles['cell_center'])
            ws.write(r, 4, f"{o['duration_sec']:.1f}", styles['cell_right'])
            ws.write(r, 5, f"{(o['duration_sec']/60.0):.2f}", styles['cell_right'])
            ws.write(r, 6, f"{(o['duration_sec']/3600.0):.4f}", styles['cell_right'])
            ws.write(r, 7, f"{o['max_loss_pct']:.0f}%", styles['cell_center'])
            ws.write(r, 8, str(o['packets_lost']), styles['cell_right'])
            st_style = styles['down'] if o['status'] == 'DOWN' else styles['drop']
            ws.write(r, 9, str(o['status']), st_style)
            ws.write(r, 10, str(o['impact_summary']), styles['cell_left'])
            r += 1

    ws.col(0).width = 256 * 10
    ws.col(1).width = 256 * 14
    ws.col(2).width = 256 * 22
    ws.col(3).width = 256 * 22
    ws.col(4).width = 256 * 15
    ws.col(5).width = 256 * 15
    ws.col(6).width = 256 * 15
    ws.col(7).width = 256 * 12
    ws.col(8).width = 256 * 14
    ws.col(9).width = 256 * 16
    ws.col(10).width = 256 * 45

    file_name = f"report_{sanitize_ip(target_ip)}.xls"
    out_path = os.path.join(SCRIPT_DIR, file_name)
    tmp_path = out_path + ".tmp"
    wb.save(tmp_path)
    os.replace(tmp_path, out_path)
    return out_path

def generate_master_excel():
    styles = get_excel_styles()
    wb = xlwt.Workbook(encoding='utf-8')

    # Sheet 1: Master Summary (All Targets)
    ws_main = wb.add_sheet("All Targets Summary", cell_overwrite_ok=True)
    ws_main.show_grid = 1
    ws_main.write(0, 0, "MULTI-TARGET 24x7 PING DROP MONITOR - EXECUTIVE SUMMARY", styles['title'])

    headers_main = ["Target Name", "IP Address", "Status", "Latency (ms)", "Total Probes", "Packets Sent", "Packets Lost", "Drop Rate %", "Drop Incidents", "Drop Impact (Hours)"]
    for idx, h in enumerate(headers_main):
        ws_main.write(3, idx, h, styles['header'])

    r_main = 4
    with _lock:
        target_list = list(_targets)

    with get_db() as conn:
        for t in target_list:
            ip = t['ip']
            name = t.get('name', ip)
            with _lock:
                st = _target_states.get(ip, {})
                probes = st.get('probes', 0)
                sent = st.get('sent', 0)
                lost = st.get('lost', 0)
                last_rtt = st.get('last_rtt', 0.0)
                status = st.get('status', 'OK')

            dr = (lost / sent * 100.0) if sent > 0 else 0.0
            incidents = conn.execute("SELECT COUNT(*), SUM(duration_sec) FROM outages WHERE target_ip=?", (ip,)).fetchone()
            inc_cnt = incidents[0] or 0
            tot_sec = incidents[1] or 0.0
            impact_hrs = tot_sec / 3600.0

            ws_main.write(r_main, 0, name, styles['cell_left'])
            ws_main.write(r_main, 1, ip, styles['cell_center'])
            st_style = styles['down'] if status == 'DOWN' else (styles['drop'] if status == 'PACKET_DROP' else styles['cell_center'])
            ws_main.write(r_main, 2, status, st_style)
            ws_main.write(r_main, 3, f"{last_rtt:.1f}", styles['cell_right'])
            ws_main.write(r_main, 4, str(probes), styles['cell_right'])
            ws_main.write(r_main, 5, str(sent), styles['cell_right'])
            ws_main.write(r_main, 6, str(lost), styles['cell_right'])
            ws_main.write(r_main, 7, f"{dr:.2f}%", styles['cell_right'])
            ws_main.write(r_main, 8, str(inc_cnt), styles['cell_right'])
            ws_main.write(r_main, 9, f"{impact_hrs:.4f}h", styles['cell_right'])
            r_main += 1

    for c in range(10):
        ws_main.col(c).width = 256 * 18

    # Subsequent Sheets: One Sheet per Target with their drop logs
    with get_db() as conn:
        for t in target_list:
            ip = t['ip']
            name = t.get('name', ip)
            sheet_title = (name[:20] + f" ({ip[-6:]})") if len(name) > 20 else f"{name} ({ip})"
            # Sheet names in Excel cannot exceed 31 chars or contain invalid characters
            sheet_title = re.sub(r'[\:\\\/\?\*\[\]]', '_', sheet_title)[:30]
            ws = wb.add_sheet(sheet_title, cell_overwrite_ok=True)
            ws.show_grid = 1

            ws.write(0, 0, f"Target: {name} ({ip}) - Drop Incidents Log", styles['title'])
            headers = ["#", "Date", "Start Time", "End Time", "Duration (Sec)", "Impact (Hours)", "Max Loss %", "Packets Lost", "Status", "Summary"]
            for idx, h in enumerate(headers):
                ws.write(2, idx, h, styles['header'])

            outages = conn.execute("SELECT * FROM outages WHERE target_ip=? ORDER BY id DESC", (ip,)).fetchall()
            r = 3
            if not outages:
                ws.write(r, 0, "No drop events recorded. Connection is stable.", styles['cell_left'])
            else:
                for o in outages:
                    ws.write(r, 0, str(o['id']), styles['cell_center'])
                    ws.write(r, 1, str(o['date_str']), styles['cell_center'])
                    ws.write(r, 2, str(o['start_time']), styles['cell_center'])
                    ws.write(r, 3, str(o['end_time'] or 'ONGOING'), styles['cell_center'])
                    ws.write(r, 4, f"{o['duration_sec']:.1f}", styles['cell_right'])
                    ws.write(r, 5, f"{(o['duration_sec']/3600.0):.4f}", styles['cell_right'])
                    ws.write(r, 6, f"{o['max_loss_pct']:.0f}%", styles['cell_center'])
                    ws.write(r, 7, str(o['packets_lost']), styles['cell_right'])
                    st_style = styles['down'] if o['status'] == 'DOWN' else styles['drop']
                    ws.write(r, 8, str(o['status']), st_style)
                    ws.write(r, 9, str(o['impact_summary']), styles['cell_left'])
                    r += 1

            for c in range(10):
                ws.col(c).width = 256 * 16
            ws.col(9).width = 256 * 40

    tmp = EXCEL_MASTER_PATH + ".tmp"
    wb.save(tmp)
    os.replace(tmp, EXCEL_MASTER_PATH)

def generate_all_reports():
    with _lock:
        targets_copy = list(_targets)

    for t in targets_copy:
        try:
            generate_individual_excel(t['ip'], t.get('name', t['ip']))
        except Exception as e:
            print(f"[!] Error generating report for {t['ip']}: {e}", file=sys.stderr)

    try:
        generate_master_excel()
    except Exception as e:
        print(f"[!] Error generating master excel: {e}", file=sys.stderr)

# ==============================================================================
# Multi-IP Parallel Worker Loop
# ==============================================================================
def ping_worker_loop():
    print("[*] Multi-IP Ping Worker Thread running ...")
    cycle = 0
    while running:
        t0 = time.time()
        with _lock:
            active_targets = [t for t in _targets if t.get('enabled', True)]

        if active_targets:
            # Parallel execution over single network interface
            with ThreadPoolExecutor(max_workers=min(12, len(active_targets))) as executor:
                results = list(executor.map(ping_single_ip, active_targets))

            for res in results:
                record_single_result(res)
                save_meta_stats_for_ip(res['ip'])

        cycle += 1
        # Refresh Excel files periodically (every 30s) or immediately if drops detected
        if cycle % 3 == 0 or any(res['status'] in ('DOWN', 'PACKET_DROP') for res in results if active_targets):
            try:
                generate_all_reports()
            except Exception as e:
                print(f"[!] Excel refresh error: {e}", file=sys.stderr)

        elapsed = time.time() - t0
        time.sleep(max(0.5, PING_INTERVAL - elapsed))

# ==============================================================================
# Flask API & Routes (Port 8084)
# ==============================================================================
@app.route('/')
def route_index():
    return render_template('index.html')

@app.route('/api/stats')
def api_stats():
    targets_data = []
    with _lock:
        for t in _targets:
            ip = t['ip']
            name = t.get('name', ip)
            st = _target_states.get(ip, {})
            probes = st.get('probes', 0)
            sent = st.get('sent', 0)
            lost = st.get('lost', 0)
            last_rtt = st.get('last_rtt', 0.0)
            status = st.get('status', 'OK')

            with get_db() as conn:
                tot_drop_sec = conn.execute("SELECT SUM(duration_sec) FROM outages WHERE target_ip=?", (ip,)).fetchone()[0] or 0.0
                drop_cnt = conn.execute("SELECT COUNT(*) FROM outages WHERE target_ip=?", (ip,)).fetchone()[0] or 0

            dr = (lost / sent * 100.0) if sent > 0 else 0.0
            targets_data.append({
                'ip': ip,
                'name': name,
                'status': status,
                'last_rtt': last_rtt,
                'total_probes': probes,
                'total_packets_sent': sent,
                'total_packets_lost': lost,
                'drop_rate_pct': dr,
                'total_drop_events': drop_cnt,
                'total_drop_seconds': tot_drop_sec
            })

    with get_db() as conn:
        drops_raw = conn.execute("""
            SELECT id, target_ip, target_name, start_time, end_time, duration_sec, packets_lost, max_loss_pct, status, impact_summary
            FROM outages
            ORDER BY id DESC
            LIMIT 100
        """).fetchall()

    drops = []
    for d in drops_raw:
        drops.append({
            'id': d['id'],
            'target_ip': d['target_ip'],
            'target_name': d['target_name'],
            'start_time': d['start_time'],
            'end_time': d['end_time'],
            'duration_sec': d['duration_sec'],
            'packets_lost': d['packets_lost'],
            'max_loss_pct': d['max_loss_pct'],
            'status': d['status'],
            'impact_summary': d['impact_summary']
        })

    return jsonify({
        'targets': targets_data,
        'drops': drops
    })

@app.route('/api/targets/add', methods=['POST'])
def api_add_target():
    data = request.get_json() or {}
    ip = str(data.get('ip', '')).strip()
    name = str(data.get('name', '')).strip() or ip

    if not ip:
        return jsonify({'success': False, 'error': 'IP address is required'}), 400

    with _lock:
        for t in _targets:
            if t['ip'] == ip:
                return jsonify({'success': False, 'error': f'Target IP {ip} is already being monitored'}), 400

        _targets.append({'ip': ip, 'name': name, 'enabled': True})

    save_targets()
    # Immediately generate report for the new IP
    try:
        generate_individual_excel(ip, name)
        generate_master_excel()
    except Exception as e:
        print(f"[!] Initial excel generation error for {ip}: {e}")

    return jsonify({'success': True, 'ip': ip, 'name': name})

@app.route('/api/targets/delete', methods=['POST'])
def api_delete_target():
    data = request.get_json() or {}
    ip = str(data.get('ip', '')).strip()

    with _lock:
        orig_len = len(_targets)
        _targets[:] = [t for t in _targets if t['ip'] != ip]
        if len(_targets) == orig_len:
            return jsonify({'success': False, 'error': 'Target not found'}), 404

    save_targets()
    try:
        generate_master_excel()
    except Exception:
        pass
    return jsonify({'success': True})

@app.route('/download')
def route_download_master():
    try:
        generate_master_excel()
    except Exception:
        pass
    if os.path.exists(EXCEL_MASTER_PATH):
        return send_file(
            EXCEL_MASTER_PATH,
            as_attachment=True,
            download_name="all_targets_drop_report.xls",
            mimetype="application/vnd.ms-excel"
        )
    return "Report not ready", 404

@app.route('/download/<ip>')
def route_download_single(ip):
    target_ip = ip.strip()
    target_name = target_ip
    with _lock:
        for t in _targets:
            if t['ip'] == target_ip:
                target_name = t.get('name', target_ip)
                break

    try:
        path = generate_individual_excel(target_ip, target_name)
    except Exception as e:
        return f"Error: {e}", 500

    if os.path.exists(path):
        return send_file(
            path,
            as_attachment=True,
            download_name=f"drop_report_{sanitize_ip(target_ip)}.xls",
            mimetype="application/vnd.ms-excel"
        )
    return "Report not ready", 404

def main():
    print("[*] Starting Multi-IP Ping & Drop Rate Monitor ...")
    init_db()
    load_targets()

    try:
        generate_all_reports()
    except Exception as e:
        print(f"[!] Initial reports generation warning: {e}")

    # Launch background multi-threaded parallel ping worker
    t = threading.Thread(target=ping_worker_loop, daemon=True)
    t.start()

    print(f"[*] Multi-IP Web GUI listening on http://0.0.0.0:{WEB_PORT} ...")
    app.run(host='0.0.0.0', port=WEB_PORT, debug=False)

if __name__ == '__main__':
    main()
