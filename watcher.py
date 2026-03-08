#!/usr/bin/env python3
"""
폴더 감시 자동화 스크립트
input/ 폴더에 PDF를 넣으면 자동으로 소포수령증을 처리합니다.

사용법:
  python watcher.py          # 폴더 감시 시작 (Ctrl+C로 종료)
  python watcher.py --once   # 현재 input/ 폴더 즉시 처리 후 종료

설치 필요:
  pip install watchdog
"""

import sys
import time
import argparse
import logging
from pathlib import Path
from datetime import datetime

BASE_DIR  = Path(__file__).parent
INPUT_DIR = BASE_DIR / 'input'
LOGS_DIR  = BASE_DIR / 'logs'

# 로그 설정
LOGS_DIR.mkdir(exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(LOGS_DIR / 'watcher.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout),
    ]
)
log = logging.getLogger(__name__)

# ── watchdog 가져오기 ────────────────────────────────────────────
try:
    from watchdog.observers import Observer
    from watchdog.events    import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False
    log.warning('watchdog 미설치 — pip install watchdog 실행 후 재시도하세요.')


# ── 이벤트 핸들러 ────────────────────────────────────────────────

class PDFHandler(FileSystemEventHandler):
    """input/ 폴더에 PDF 파일이 추가되면 처리 트리거"""

    def __init__(self, wait_seconds=5):
        super().__init__()
        self.wait_seconds = wait_seconds   # 파일 안정화 대기 시간
        self._pending_files = set()
        self._last_modified  = 0

    def on_created(self, event):
        if event.is_directory:
            return
        path = Path(event.src_path)
        if path.suffix.lower() == '.pdf' and 'processed' not in str(path):
            log.info(f'새 PDF 감지: {path.name}')
            self._pending_files.add(str(path))
            self._last_modified = time.time()

    def on_moved(self, event):
        """파일 이동(복사 완료 시 발생하는 경우도 있음)"""
        self.on_created(type('E', (), {'is_directory': False, 'src_path': event.dest_path})())

    def check_and_process(self):
        """대기 시간 후 파일이 안정화되면 처리 실행"""
        if self._pending_files and (time.time() - self._last_modified) >= self.wait_seconds:
            files_to_process = list(self._pending_files)
            self._pending_files.clear()

            # 아직 파일이 있는 것만 처리
            existing = [p for p in files_to_process if Path(p).exists()]
            if not existing:
                return

            log.info(f'▶ 처리 시작: {len(existing)}개 PDF')
            try:
                from process import process
                process(pdf_paths=[Path(p) for p in existing])
            except Exception as e:
                log.error(f'처리 중 오류 발생: {e}', exc_info=True)


# ── 폴링 방식 폴더 감시 (watchdog 없을 때 대체) ─────────────────

def poll_watch(input_dir: Path, interval: int = 10):
    """watchdog 없을 때 주기적 폴링으로 감시"""
    log.info(f'📁 폴링 감시 시작 ({interval}초 간격): {input_dir}')
    log.info('   Ctrl+C 로 종료')

    seen_files = set(p.name for p in input_dir.glob('*.pdf'))

    while True:
        time.sleep(interval)
        try:
            current = set(p.name for p in input_dir.glob('*.pdf'))
            new_files = current - seen_files
            if new_files:
                log.info(f'새 PDF 감지: {new_files}')
                time.sleep(3)  # 파일 안정화 대기
                pdf_paths = [input_dir / f for f in new_files if (input_dir / f).exists()]
                if pdf_paths:
                    from process import process
                    process(pdf_paths=pdf_paths)
            seen_files = current
        except KeyboardInterrupt:
            raise
        except Exception as e:
            log.error(f'오류: {e}', exc_info=True)


# ── watchdog 기반 감시 ───────────────────────────────────────────

def watchdog_watch(input_dir: Path):
    log.info(f'📁 실시간 감시 시작: {input_dir}')
    log.info('   Ctrl+C 로 종료')

    handler  = PDFHandler(wait_seconds=5)
    observer = Observer()
    observer.schedule(handler, str(input_dir), recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(2)
            handler.check_and_process()
    except KeyboardInterrupt:
        log.info('감시 종료 중...')
    finally:
        observer.stop()
        observer.join()
    log.info('감시 종료.')


# ── 진입점 ──────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='소포수령증 폴더 감시 자동화')
    parser.add_argument('--once',     action='store_true', help='현재 PDF 즉시 처리 후 종료')
    parser.add_argument('--interval', type=int, default=10, help='폴링 간격(초), 기본값: 10')
    args = parser.parse_args()

    INPUT_DIR.mkdir(exist_ok=True)

    print(f"""
╔══════════════════════════════════════════════════════╗
║     소포수령증 자동화 시스템 v1.0                    ║
║     감시 폴더: input/                               ║
║     출력 폴더: output/                              ║
╚══════════════════════════════════════════════════════╝
""")

    if args.once:
        # 즉시 처리
        from process import process
        process()
        return

    if WATCHDOG_AVAILABLE:
        watchdog_watch(INPUT_DIR)
    else:
        log.warning('watchdog 없음 → 폴링 방식으로 전환 (pip install watchdog 권장)')
        poll_watch(INPUT_DIR, interval=args.interval)


if __name__ == '__main__':
    main()
