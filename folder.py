#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# 백업 및 복구 프로그램
# 사용법:
#     백업: python folder.py --backup "E:\backup"
#     전체 복구: python folder.py --restore "E:\backup\20251206"
#     선택 복구: python folder.py --restore "E:\backup\20251206\NPKI"
#     경로 추가: python folder.py --add-path "C:\경로\폴더"
#     경로 제거: python folder.py --remove-path "C:\경로\폴더"
#     경로 목록: python folder.py --list-paths
#

import argparse
import json
import os
import shutil
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path


CONFIG_FILE = Path(__file__).parent / "folder_config.json"

# 기본 제외 파일 목록 (시스템 파일)
DEFAULT_EXCLUDE_FILES = [
    "desktop.ini",
    "Thumbs.db",
    ".DS_Store",
]


# 환경변수 확장 (%APPDATA%, %LOCALAPPDATA% 등)
def expand_path(path):
    return os.path.normpath(os.path.expandvars(os.path.expanduser(path)))


# 경로 항목 정규화 (문자열 또는 딕셔너리 모두 지원)
def normalize_path_item(item):
    if isinstance(item, str):
        return {"path": item, "service": None, "exclude": [], "destination": None}
    return {
        "path": item.get("path", ""),
        "service": item.get("service"),
        "exclude": item.get("exclude", []),
        "destination": item.get("destination")
    }


# 서비스 상태 확인
def get_service_status(service_name: str) -> str | None:
    try:
        result = subprocess.run(
            ["sc", "query", service_name],
            capture_output=True, text=True, encoding="cp949"
        )
        if "RUNNING" in result.stdout:
            return "running"
        elif "STOPPED" in result.stdout:
            return "stopped"
        return "unknown"
    except Exception:
        return None


# 서비스 중지 (완전히 중지될 때까지 대기)
def stop_service(service_name: str, timeout: int = 30) -> bool:
    print(f"[서비스] {service_name} 중지 중...")
    
    # 이미 중지되어 있는지 확인
    status = get_service_status(service_name)
    if status == "stopped":
        print(f"[서비스] {service_name} 이미 중지됨")
        return True
    
    if status is None:
        print(f"[서비스] {service_name} 서비스를 찾을 수 없음")
        return False
    
    # 서비스 중지 명령
    try:
        subprocess.run(["sc", "stop", service_name], capture_output=True, encoding="cp949")
    except Exception as e:
        print(f"[서비스] 중지 명령 실패: {e}")
        return False
    
    # 완전히 중지될 때까지 대기
    for i in range(timeout):
        time.sleep(1)
        status = get_service_status(service_name)
        if status == "stopped":
            print(f"[서비스] {service_name} 중지 완료")
            return True
        print(f"[서비스] 대기 중... ({i+1}/{timeout}초)")
    
    print(f"[서비스] {service_name} 중지 타임아웃")
    return False


# 서비스 시작
def start_service(service_name: str, timeout: int = 30) -> bool:
    print(f"[서비스] {service_name} 시작 중...")
    
    # 이미 실행 중인지 확인
    status = get_service_status(service_name)
    if status == "running":
        print(f"[서비스] {service_name} 이미 실행 중")
        return True
    
    if status is None:
        print(f"[서비스] {service_name} 서비스를 찾을 수 없음")
        return False
    
    # 서비스 시작 명령
    try:
        subprocess.run(["sc", "start", service_name], capture_output=True, encoding="cp949")
    except Exception as e:
        print(f"[서비스] 시작 명령 실패: {e}")
        return False
    
    # 실행될 때까지 대기
    for i in range(timeout):
        time.sleep(1)
        status = get_service_status(service_name)
        if status == "running":
            print(f"[서비스] {service_name} 시작 완료")
            return True
        print(f"[서비스] 대기 중... ({i+1}/{timeout}초)")
    
    print(f"[서비스] {service_name} 시작 타임아웃")
    return False


# 설정 파일 로드
def load_config():
    if CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"backup_paths": [], "description": "백업할 폴더 경로 목록입니다."}


# 설정 파일 저장
def save_config(config):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


# 백업 경로 추가
def add_backup_path(path):
    config = load_config()
    # 환경변수는 그대로 저장 (이식성 유지)
    normalized_path = os.path.normpath(path)
    
    # 중복 체크는 확장된 경로로
    expanded_path = expand_path(path)
    existing_expanded = [expand_path(p) for p in config["backup_paths"]]
    
    if expanded_path in existing_expanded:
        print(f"이미 등록된 경로입니다: {normalized_path}")
        print(f"  -> {expanded_path}")
        return False
    
    if not os.path.exists(expanded_path):
        print(f"경고: 경로가 존재하지 않습니다: {expanded_path}")
        response = input("그래도 추가하시겠습니까? (y/n): ")
        if response.lower() != 'y':
            return False
    
    config["backup_paths"].append(normalized_path)
    save_config(config)
    print(f"경로가 추가되었습니다: {normalized_path}")
    print(f"  -> {expanded_path}")
    return True


# 백업 경로 제거
def remove_backup_path(path):
    config = load_config()
    normalized_path = os.path.normpath(path)
    
    if normalized_path not in config["backup_paths"]:
        print(f"등록되지 않은 경로입니다: {normalized_path}")
        return False
    
    config["backup_paths"].remove(normalized_path)
    save_config(config)
    print(f"경로가 제거되었습니다: {normalized_path}")
    return True


# 백업 경로 목록 출력
def list_backup_paths():
    config = load_config()
    paths = config.get("backup_paths", [])
    
    if not paths:
        print("등록된 백업 경로가 없습니다.")
        return
    
    print("\n=== 백업 경로 목록 ===\n")
    for i, item in enumerate(paths, 1):
        # dict 형태와 문자열 형태 모두 지원
        if isinstance(item, dict):
            path = item.get("path", "")
            service = item.get("service")
            exclude = item.get("exclude", [])
            destination = item.get("destination")
        else:
            path = item
            service = None
            exclude = []
            destination = None
        
        expanded = expand_path(path)
        exists = "✓" if os.path.exists(expanded) else "✗"
        
        # 복사 가능한 경로 출력
        print(f"  {i}. [{exists}] {expanded}")
        
        # 목적지 폴더명 표시
        if destination:
            print(f"       목적지: {destination}")
        # 서비스 정보 표시
        if service:
            print(f"       서비스: {service}")
        # 제외 폴더 표시
        if exclude:
            print(f"       제외: {', '.join(exclude)}")
    print()


# 제외 파일/폴더를 고려한 복사 (shutil.copytree의 ignore 함수)
def make_ignore_func(exclude_list=None):
    # 기본 제외 파일 + 사용자 지정 제외 목록 합치기
    all_excludes = set(DEFAULT_EXCLUDE_FILES)
    if exclude_list:
        all_excludes.update(exclude_list)
    
    def ignore_func(directory, files):
        ignored = []
        for f in files:
            if f in all_excludes:
                ignored.append(f)
        return ignored
    return ignore_func


# 스마트 복사 함수: 파일명, 크기, 수정시간이 동일하면 건너뛰기
def smart_copy2(src, dst):
    """파일명, 크기, 수정시간이 동일하면 복사 건너뛰기"""
    if os.path.exists(dst):
        src_stat = os.stat(src)
        dst_stat = os.stat(dst)
        # 크기와 수정시간이 같으면 건너뛰기
        if src_stat.st_size == dst_stat.st_size and int(src_stat.st_mtime) == int(dst_stat.st_mtime):
            return dst
    # 다르거나 파일이 없으면 복사
    return shutil.copy2(src, dst)


# 백업 수행
def backup(destination):
    config = load_config()
    paths = config.get("backup_paths", [])
    
    if not paths:
        print("백업할 경로가 등록되어 있지 않습니다.")
        print("'python folder.py --add-path \"경로\"' 명령으로 경로를 추가하세요.")
        return False
    
    # 백업 폴더 생성 (날짜 폴더 없이 바로 목적지에 백업)
    backup_folder = os.path.normpath(destination)
    
    try:
        os.makedirs(backup_folder, exist_ok=True)
    except Exception as e:
        print(f"백업 폴더 생성 실패: {e}")
        return False
    
    # 메타데이터 저장
    metadata = {
        "backup_date": datetime.now().isoformat(),
        "paths": []
    }
    
    # 기존 메타데이터 로드 (증분 백업용)
    meta_file = os.path.join(backup_folder, "backup_metadata.json")
    existing_backup_map = {}  # source_expanded -> backup folder name
    if os.path.exists(meta_file):
        try:
            with open(meta_file, "r", encoding="utf-8") as f:
                existing_meta = json.load(f)
                for item in existing_meta.get("paths", []):
                    existing_backup_map[item["source_expanded"]] = item["backup"]
        except:
            pass
    
    print(f"\n=== 백업 시작 ===")
    print(f"백업 위치: {backup_folder}\n")
    
    success_count = 0
    fail_count = 0
    services_to_restart = []  # 백업 후 재시작할 서비스 목록
    
    for item in paths:
        # 경로 항목 정규화 (문자열 또는 딕셔너리 지원)
        path_info = normalize_path_item(item)
        source_path = path_info["path"]
        service_name = path_info["service"]
        exclude_list = path_info["exclude"]
        custom_destination = path_info["destination"]
        
        # 환경변수 확장
        expanded_path = expand_path(source_path)
        
        # 서비스 중지 필요 시
        if service_name:
            original_status = get_service_status(service_name)
            if original_status == "running":
                if not stop_service(service_name):
                    print(f"[실패] 서비스 중지 실패로 백업 건너뜀: {expanded_path}")
                    fail_count += 1
                    continue
                services_to_restart.append(service_name)
        
        if not os.path.exists(expanded_path):
            print(f"[건너뜀] 경로가 존재하지 않음: {source_path}")
            if source_path != expanded_path:
                print(f"         -> {expanded_path}")
            fail_count += 1
            continue
        
        # 백업 폴더명 결정: custom_destination > 기존 메타데이터 > 자동 생성
        if custom_destination:
            # 사용자가 지정한 목적지 폴더명 사용
            folder_name = custom_destination
            dest_path = os.path.join(backup_folder, folder_name)
        elif expanded_path in existing_backup_map:
            # 기존 백업이 있으면 같은 폴더명 사용
            folder_name = existing_backup_map[expanded_path]
            dest_path = os.path.join(backup_folder, folder_name)
        else:
            # 새 백업: 항상 "상위폴더_폴더명" 형식 사용 (충돌 방지)
            base_name = os.path.basename(expanded_path)
            parent_name = os.path.basename(os.path.dirname(expanded_path))
            folder_name = f"{parent_name}_{base_name}"
            dest_path = os.path.join(backup_folder, folder_name)
        
        try:
            # 기존 백업이 있으면 증분 백업 (기존 파일 유지, 새 파일 추가/덮어쓰기)
            is_incremental = os.path.exists(dest_path)
            
            if exclude_list:
                print(f"[{'증분' if is_incremental else '백업'}] {expanded_path} -> {folder_name} (제외: {exclude_list})")
            else:
                print(f"[{'증분' if is_incremental else '백업'}] {expanded_path} -> {folder_name}")
            
            if os.path.isfile(expanded_path):
                smart_copy2(expanded_path, dest_path)
            else:
                # 항상 ignore 함수 사용 (기본 제외 파일 + 사용자 지정 제외)
                # dirs_exist_ok=True: 기존 백업 유지하면서 새 파일 추가
                # copy_function=smart_copy2: 동일 파일 건너뛰기
                shutil.copytree(expanded_path, dest_path, ignore=make_ignore_func(exclude_list), dirs_exist_ok=True, copy_function=smart_copy2)
            
            # 메타데이터에 원본 정보 저장
            meta_item = {
                "source": source_path,
                "source_expanded": expanded_path,
                "backup": folder_name,
                "type": "file" if os.path.isfile(expanded_path) else "directory"
            }
            if service_name:
                meta_item["service"] = service_name
            if exclude_list:
                meta_item["exclude"] = exclude_list
            
            metadata["paths"].append(meta_item)
            
            print(f"[완료] {expanded_path}")
            success_count += 1
            
        except Exception as e:
            print(f"[실패] {expanded_path}: {e}")
            fail_count += 1

    # 중지했던 서비스 재시작
    for service_name in services_to_restart:
        start_service(service_name)
    
    # 메타데이터 파일 저장
    metadata_file = os.path.join(backup_folder, "backup_metadata.json")
    with open(metadata_file, "w", encoding="utf-8") as f:
        json.dump(metadata, f, ensure_ascii=False, indent=4)
    
    print(f"\n=== 백업 완료 ===")
    print(f"성공: {success_count}개, 실패: {fail_count}개")
    print(f"백업 위치: {backup_folder}")
    
    return True


# 백업 경로에서 메타데이터 파일 위치와 타겟 폴더 자동 감지
def find_backup_root(path):
    path = os.path.normpath(path)
    
    # 현재 경로에 메타데이터가 있으면 해당 경로가 백업 루트
    metadata_file = os.path.join(path, "backup_metadata.json")
    if os.path.exists(metadata_file):
        return path, None
    
    # 상위 경로에서 메타데이터 찾기 (폴더명 추출)
    parent = os.path.dirname(path)
    folder_name = os.path.basename(path)
    
    parent_metadata = os.path.join(parent, "backup_metadata.json")
    if os.path.exists(parent_metadata):
        return parent, folder_name
    
    return path, None


# 복구 수행 (target: 특정 폴더명만 복구, None이면 전체 복구)
def restore(input_path, target=None):
    # 경로에서 백업 루트와 타겟 자동 감지
    backup_path, auto_target = find_backup_root(input_path)
    
    # 자동 감지된 타겟이 있으면 사용
    if auto_target and not target:
        target = auto_target
        print(f"[자동 감지] '{target}' 폴더만 복구합니다.")
    
    if not os.path.exists(backup_path):
        print(f"백업 폴더가 존재하지 않습니다: {backup_path}")
        return False
    
    # 메타데이터 파일 확인
    metadata_file = os.path.join(backup_path, "backup_metadata.json")
    
    if not os.path.exists(metadata_file):
        print(f"백업 메타데이터 파일이 없습니다: {metadata_file}")
        return False
    
    with open(metadata_file, "r", encoding="utf-8") as f:
        metadata = json.load(f)
    
    # 복구 대상 필터링
    all_items = metadata["paths"]
    if target:
        restore_items = [item for item in all_items if item["backup"] == target]
        if not restore_items:
            print(f"'{target}' 항목을 찾을 수 없습니다.")
            print("\n사용 가능한 항목:")
            for item in all_items:
                print(f"  - {item['backup']}")
            return False
    else:
        restore_items = all_items
    
    print(f"\n=== 복구 시작 ===")
    print(f"백업 날짜: {metadata.get('backup_date', '알 수 없음')}")
    print(f"복구할 항목 수: {len(restore_items)}개\n")
    
    # 확인 메시지
    print("경고: 기존 파일이 덮어쓰기될 수 있습니다!")
    response = input("계속하시겠습니까? (y/n): ")
    if response.lower() != 'y':
        print("복구가 취소되었습니다.")
        return False
    
    # 서비스별 항목 그룹화 (서비스 제어 최소화를 위해)
    service_groups: dict[str | None, list] = {}
    for item in restore_items:
        service_name = item.get("service")
        if service_name not in service_groups:
            service_groups[service_name] = []
        service_groups[service_name].append(item)
    
    success_count = 0
    fail_count = 0
    
    # 서비스 그룹별로 처리
    for service_name, items in service_groups.items():
        service_was_running = False
        
        # 서비스가 있으면 중지
        if service_name:
            print(f"\n[서비스] {service_name} 상태 확인 중...")
            status = get_service_status(service_name)
            if status == "RUNNING":
                service_was_running = True
                if not stop_service(service_name):
                    print(f"[경고] 서비스 {service_name} 중지 실패, 해당 항목들을 건너뜁니다.")
                    fail_count += len(items)
                    continue
        
        # 항목 복구
        for item in items:
            source = item["source"]
            # 환경변수 확장 (복구 시에도 현재 시스템의 환경변수 사용)
            source_expanded = expand_path(source)
            backup_relative = item["backup"]
            item_type = item.get("type", "directory")
            
            backup_item_path = os.path.join(backup_path, backup_relative)
            
            if not os.path.exists(backup_item_path):
                print(f"[건너뜀] 백업 항목이 없음: {backup_item_path}")
                fail_count += 1
                continue
            
            try:
                print(f"[복구중] {source_expanded}")
                
                # 기존 항목 제거
                if os.path.exists(source_expanded):
                    if os.path.isfile(source_expanded):
                        os.remove(source_expanded)
                    else:
                        shutil.rmtree(source_expanded)
                
                # 복구
                if item_type == "file":
                    os.makedirs(os.path.dirname(source_expanded), exist_ok=True)
                    shutil.copy2(backup_item_path, source_expanded)
                else:
                    shutil.copytree(backup_item_path, source_expanded)
                
                print(f"[완료] {source_expanded}")
                success_count += 1
                
            except Exception as e:
                print(f"[실패] {source_expanded}: {e}")
                fail_count += 1
        
        # 서비스가 실행 중이었으면 다시 시작
        if service_name and service_was_running:
            start_service(service_name)
    
    print(f"\n=== 복구 완료 ===")
    print(f"성공: {success_count}개, 실패: {fail_count}개")
    
    return True


def main():
    parser = argparse.ArgumentParser(
        description="파일 백업 및 복구 프로그램",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python folder.py --list                         경로 목록
  python folder.py --backup "E:\\backup"          백업 실행 (증분 백업)
  python folder.py --restore "E:\\backup"         전체 복구
  python folder.py --restore "E:\\backup\\NPKI"   선택 복구
  python folder.py --add "C:\\경로"               경로 추가
  python folder.py --remove "C:\\경로"            경로 제거
            """
    )
    
    parser.add_argument(
        "--backup", "-b",
        metavar="DEST",
        help="지정된 경로에 백업 수행"
    )
    
    parser.add_argument(
        "--restore", "-r",
        metavar="PATH",
        help="백업 폴더에서 복구 수행 (폴더명 포함 시 해당 폴더만 복구)"
    )
    
    parser.add_argument(
        "--add", "-a",
        metavar="PATH",
        help="백업할 경로 추가"
    )
    
    parser.add_argument(
        "--remove", "-d",
        metavar="PATH",
        help="백업 경로 제거"
    )
    
    parser.add_argument(
        "--list", "-l",
        action="store_true",
        help="등록된 백업 경로 목록 표시"
    )
    
    args = parser.parse_args()
    
    # 아무 옵션도 없으면 도움말 표시
    if len(sys.argv) == 1:
        parser.print_help()
        return
    
    if args.add:
        add_backup_path(args.add)
    elif args.remove:
        remove_backup_path(args.remove)
    elif args.list:
        list_backup_paths()
    elif args.backup:
        backup(args.backup)
    elif args.restore:
        restore(args.restore)


if __name__ == "__main__":
    main()
