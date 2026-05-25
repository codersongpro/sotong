"""Small automation status helpers."""

FAIL_NO_USER = '사용자 없음'
FAIL_DUPLICATE = '중복'
FAIL_COORDINATE = '좌표 오류'
FAIL_MANUAL_STOP = '수동 중지'
FAIL_AUTOMATION = '자동화 오류'


def failure_reason_from_error(exc: Exception) -> str:
    return FAIL_COORDINATE if '좌표' in str(exc) else FAIL_AUTOMATION
