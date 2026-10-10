"""이전 import 경로 호환용. 실제 구현은 단일 실행 파일에 있습니다."""
from hometax_vat_downloader import (
    SERVICES, REQUIRED, clean, choices, biz_number, Period, Client, Job,
    load_clients, build_jobs, validate_artifact, Journal,
)
