"""UI formatting helpers."""


def format_item_label(item: dict) -> str:
    org = item.get('org', '')
    name = item.get('name', '')
    reason = item.get('failure_reason', '')
    label = f'[{org}]  {name}' if org else f'{name}  (소속없음)'
    return f'{label}  — 실패: {reason}' if reason else label
