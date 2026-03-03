"""Generate an empty but styled Agent E2E Excel template.

Usage:
  python tests/generate_agent_e2e_template_excel.py
"""

from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path

import run_agent_e2e_suite as e2e


def main() -> int:
    suite_path = Path(__file__).resolve().parent / 'agent_e2e_suite.json'
    if not suite_path.exists():
        raise SystemExit(f'Suite file not found: {suite_path}')

    suite = json.loads(suite_path.read_text(encoding='utf-8'))
    report = {
        '测试集': suite.get('测试集') or {},
        '评测维度': suite.get('评测维度') or {},
        '验收门槛定义': suite.get('验收门槛') or {},
        '运行信息': {
            'run_id': 'TEMPLATE',
            'timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'base_url': 'https://api.coze.cn',
            'bot_id': (suite.get('测试集') or {}).get('bot_id') or '',
            'use_stream': (suite.get('默认配置') or {}).get('是否使用流式首响', True),
            'timeout_s': 120,
            'max_polls': 20,
            'poll_interval_s': 0.8,
            'start_index': 1,
            'max_cases': 0,
            'stop_on_exception': False,
            'default_user_id': (suite.get('默认配置') or {}).get('测试user_id') or '',
        },
        '结果': [],
        '汇总': {},
    }
    report['汇总'] = e2e._calc_summary(report['结果'], report['验收门槛定义'])

    out_dir = Path(__file__).resolve().parent / 'reports' / 'agent_e2e'
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / 'agent_e2e_template.xlsx'
    e2e._make_excel(report, out_path)

    print(f'TEMPLATE_XLSX: {out_path}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
