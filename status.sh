#!/usr/bin/env bash
# ==============================================================================
# Quick CLI Status Checker for Ping Drop Monitor
# ==============================================================================
echo "=========================================================="
echo "  24x7 MULTI-IP PING DROP MONITOR (PORT 8084)"
echo "=========================================================="
systemctl is-active ping-monitor.service >/dev/null 2>&1 && echo " Service Status : ACTIVE (RUNNING)" || echo " Service Status : STOPPED"

echo ""
echo "--- CONFIGURED MONITORED TARGETS (/ping/targets.json) ---"
if [ -f /ping/targets.json ]; then
    cat /ping/targets.json | grep -E '"ip"|"name"' | tr -d '"' | tr -d ','
fi

echo ""
echo "--- DROP INCIDENTS SUMMARY FROM DATABASE ---"
if [ -f /ping/ping_history.db ]; then
    sqlite3 /ping/ping_history.db << 'SQL'
.mode column
.headers on
SELECT 
    target_ip as 'Target IP',
    target_name as 'Name',
    COUNT(*) as 'Total Drop Events',
    ROUND(IFNULL(SUM(duration_sec), 0), 1) || 's' as 'Total Drop Time',
    ROUND(IFNULL(SUM(duration_sec)/3600.0, 0), 4) || ' Hrs' as 'Drop Impact'
FROM outages
GROUP BY target_ip;
SQL
fi

echo ""
echo "--- GENERATED EXCEL REPORT FILES IN /ping/ ---"
ls -lh /ping/report*.xls 2>/dev/null | awk '{print " " $9 " (" $5 ", Updated: " $7 " " $8 ")"}'
echo "=========================================================="
