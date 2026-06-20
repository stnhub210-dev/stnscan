# -*- coding: utf-8 -*-
"""밀린 스캔본 일괄 처리 (무한 감시 없이 1회 실행)
   JPG→PDF 변환 → 미처리 PDF AI분석/이름변경/분류 → HTML 대장 갱신"""
import os, time, auto_manager as am

def run():
    if not am.load_keys():
        print("api_key.txt 확인 필요"); return
    am.log.info("=== 일괄처리 시작 ===")
    done = am.load_done()
    # 1) JPG 스캔본 → PDF 병합 (밀린 것은 대기 없이 즉시)
    am.convert_scan_jpgs(stable_sec=0)
    # 2) 미처리 PDF 처리
    pdfs = [f for f in os.listdir(am.SCAN_FOLDER) if f.lower().endswith(".pdf")]
    todo = [f for f in pdfs if f not in done]
    am.log.info(f"미처리 PDF {len(todo)}개 처리")
    for i, fname in enumerate(todo, 1):
        am.log.info(f"[{i}/{len(todo)}]")
        am.process_pdf(os.path.join(am.SCAN_FOLDER, fname), done)
        time.sleep(0.3)
    # 3) 드라이브ID 갱신 + HTML 대장 생성
    ids = am.load_drive_ids(force=True)
    am.generate_html(am.load_lawsuit_dict(), ids)
    am.log.info("=== 일괄처리 완료 ===")

if __name__ == "__main__":
    run()
