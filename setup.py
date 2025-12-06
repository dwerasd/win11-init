#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Windows 11 시스템 설정 자동 적용 스크립트
# 사용법:
#     실행: python setup.py
#     시뮬레이션: python setup.py --dry-run
#     목록 보기: python setup.py --list
#

import json
import re
import socket
import subprocess
import sys
import winreg
from pathlib import Path
from typing import Any


# 설정 파일 경로
SCRIPT_DIR = Path(__file__).parent
REGISTRY_CONFIG = SCRIPT_DIR / "registry_config.json"
COMMANDS_CONFIG = SCRIPT_DIR / "commands_config.json"


# ===== 레지스트리 관련 =====

HKEY_MAP: dict[str, int] = {
    "HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
    "HKLM": winreg.HKEY_LOCAL_MACHINE,
    "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
    "HKCU": winreg.HKEY_CURRENT_USER,
    "HKEY_CLASSES_ROOT": winreg.HKEY_CLASSES_ROOT,
    "HKCR": winreg.HKEY_CLASSES_ROOT,
    "HKEY_USERS": winreg.HKEY_USERS,
    "HKU": winreg.HKEY_USERS,
    "HKEY_CURRENT_CONFIG": winreg.HKEY_CURRENT_CONFIG,
    "HKCC": winreg.HKEY_CURRENT_CONFIG,
}

TYPE_MAP = {
    "REG_SZ": winreg.REG_SZ,
    "REG_EXPAND_SZ": winreg.REG_EXPAND_SZ,
    "REG_BINARY": winreg.REG_BINARY,
    "REG_DWORD": winreg.REG_DWORD,
    "REG_QWORD": winreg.REG_QWORD,
    "REG_MULTI_SZ": winreg.REG_MULTI_SZ,
}


def parse_registry_path(full_path: str) -> tuple[int, str] | tuple[None, None]:
    parts = full_path.split("\\", 1)
    if len(parts) < 2:
        return None, None
    
    root_name = parts[0].upper()
    sub_key = parts[1]
    
    if root_name not in HKEY_MAP:
        return None, None
    
    return HKEY_MAP[root_name], sub_key


def write_registry_value(full_path: str, value_name: str, value: Any, reg_type: int) -> bool:
    root_key, sub_key = parse_registry_path(full_path)
    if root_key is None or sub_key is None:
        return False
    
    try:
        with winreg.CreateKey(root_key, sub_key) as key:
            winreg.SetValueEx(key, value_name, 0, reg_type, value)
        return True
    except PermissionError:
        print(f"    [권한 오류] 관리자 권한 필요: {full_path}")
        return False
    except Exception as e:
        print(f"    [오류] {e}")
        return False


def deserialize_value(value, reg_type):
    if reg_type == winreg.REG_BINARY:
        if isinstance(value, str):
            # 콤마로 구분된 hex 문자열 처리 (예: "00,a0,00,00")
            hex_str = value.replace(",", "").replace(" ", "")
            return bytes.fromhex(hex_str)
        return value if value else b""
    elif reg_type == winreg.REG_MULTI_SZ:
        return value if value else []
    elif reg_type in (winreg.REG_DWORD, winreg.REG_QWORD):
        return int(value) if isinstance(value, str) else value
    else:
        return value


# ===== 인트라넷 영역 자동 등록 =====

def get_local_ip() -> str | None:
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return None


def get_ip_range(ip: str) -> str:
    parts = ip.split(".")
    if len(parts) == 4:
        return f"{parts[0]}.{parts[1]}.{parts[2]}.*"
    return ip


def get_existing_intranet_ranges() -> dict[str, str]:
    ranges = {}
    base_path = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, base_path, 0, winreg.KEY_READ)
        i = 0
        while True:
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name, 0, winreg.KEY_READ)
                try:
                    range_value, _ = winreg.QueryValueEx(subkey, ":Range")
                    ranges[subkey_name] = range_value
                except FileNotFoundError:
                    pass
                winreg.CloseKey(subkey)
                i += 1
            except OSError:
                break
        winreg.CloseKey(key)
    except FileNotFoundError:
        pass
    
    return ranges


def setup_intranet_zone(dry_run: bool = False) -> bool:
    ip = get_local_ip()
    if not ip:
        print("    [인트라넷] IP 주소를 감지할 수 없습니다.")
        return False
    
    ip_range = get_ip_range(ip)
    
    # 이미 등록되어 있는지 확인
    existing = get_existing_intranet_ranges()
    for name, value in existing.items():
        if value == ip_range:
            print(f"    [인트라넷] 이미 등록됨: {ip_range}")
            return True
    
    if dry_run:
        print(f"    [인트라넷] 등록 예정: {ip_range}")
        return True
    
    # 새 Range 번호 결정
    max_num = 0
    for name in existing.keys():
        if name.startswith("Range"):
            try:
                num = int(name[5:])
                max_num = max(max_num, num)
            except ValueError:
                pass
    range_name = f"Range{max_num + 1}"
    base_path = r"Software\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Ranges"
    range_path = f"{base_path}\\{range_name}"
    
    try:
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, range_path)
        winreg.SetValueEx(key, ":Range", 0, winreg.REG_SZ, ip_range)
        winreg.SetValueEx(key, "*", 0, winreg.REG_DWORD, 1)
        winreg.CloseKey(key)
        print(f"    [인트라넷] 등록 완료: {ip_range}")
        return True
    except Exception as e:
        print(f"    [인트라넷] 등록 실패: {e}")
        return False


# ===== 명령어 실행 관련 =====

def run_powershell(command: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", command],
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace"
        )
        output = result.stdout + result.stderr
        return result.returncode == 0, output.strip()
    except Exception as e:
        return False, str(e)


def run_cmd(command: str) -> tuple[bool, str]:
    try:
        result = subprocess.run(
            ["cmd", "/c", command],
            capture_output=True,
            text=True,
            encoding="cp949",
            errors="replace"
        )
        output = result.stdout + result.stderr
        return result.returncode == 0, output.strip()
    except Exception as e:
        return False, str(e)


# ===== 설정 적용 함수들 =====

def apply_registry(dry_run: bool = False) -> tuple[int, int]:
    """레지스트리 설정 적용"""
    if not REGISTRY_CONFIG.exists():
        print("[레지스트리] 설정 파일이 없습니다.")
        return 0, 0
    
    with open(REGISTRY_CONFIG, "r", encoding="utf-8") as f:
        config = json.load(f)
    
    items = config.get("registry_items", [])
    if not items:
        print("[레지스트리] 적용할 항목이 없습니다.")
        return 0, 0
    
    print(f"\n{'='*50}")
    print(f"[1/3] 레지스트리 설정 적용 ({len(items)}개)")
    print(f"{'='*50}")
    
    success_count = 0
    fail_count = 0
    
    for item in items:
        path = item["path"]
        name = item["name"]
        type_name = item["type"]
        value = item["value"]
        desc = item.get("description", "")
        
        reg_type = TYPE_MAP.get(type_name, winreg.REG_SZ)
        final_value = deserialize_value(value, reg_type)
        
        # 출력
        short_path = path.split("\\")[-1] if "\\" in path else path
        print(f"  • {desc or name}")
        
        if dry_run:
            print(f"    [시뮬레이션] {short_path}\\{name} = {value}")
            success_count += 1
            continue
        
        if write_registry_value(path, name, final_value, reg_type):
            success_count += 1
        else:
            fail_count += 1
    
    print(f"\n  결과: 성공 {success_count}개, 실패 {fail_count}개")
    return success_count, fail_count


def apply_intranet(dry_run: bool = False) -> tuple[int, int]:
    """인트라넷 영역 설정"""
    print(f"\n{'='*50}")
    print("[2/3] 로컬 인트라넷 영역 설정")
    print(f"{'='*50}")
    
    ip = get_local_ip()
    if ip:
        print(f"  • 현재 IP: {ip}, 대역: {get_ip_range(ip)}")
    
    if setup_intranet_zone(dry_run):
        return 1, 0
    return 0, 1


def apply_commands(dry_run: bool = False) -> tuple[int, int]:
    """시스템 명령어 실행"""
    if not COMMANDS_CONFIG.exists():
        print("[명령어] 설정 파일이 없습니다.")
        return 0, 0
    
    with open(COMMANDS_CONFIG, "r", encoding="utf-8") as f:
        config = json.load(f)
    
    commands = config.get("commands", [])
    enabled_commands = [c for c in commands if c.get("enabled", True)]
    
    if not enabled_commands:
        print("[명령어] 실행할 명령어가 없습니다.")
        return 0, 0
    
    print(f"\n{'='*50}")
    print(f"[3/3] 시스템 명령어 실행 ({len(enabled_commands)}개)")
    print(f"{'='*50}")
    
    success_count = 0
    fail_count = 0
    
    for i, item in enumerate(enabled_commands, 1):
        command = item["command"]
        cmd_type = item["type"]
        desc = item.get("description", "")
        
        # 출력
        cmd_short = command[:60] + "..." if len(command) > 60 else command
        print(f"  [{i}/{len(enabled_commands)}] {desc or cmd_short}")
        
        if dry_run:
            print(f"    [시뮬레이션] {cmd_type.upper()}: {cmd_short}")
            success_count += 1
            continue
        
        # 실행
        if cmd_type == "powershell":
            success, output = run_powershell(command)
        else:
            success, output = run_cmd(command)
        
        if success:
            print(f"    [완료]")
            success_count += 1
        else:
            print(f"    [실패] {output[:100] if output else '알 수 없는 오류'}")
            fail_count += 1
    
    print(f"\n  결과: 성공 {success_count}개, 실패 {fail_count}개")
    return success_count, fail_count


def list_all() -> None:
    """모든 설정 항목 목록 출력"""
    print("\n" + "="*60)
    print(" Windows 11 시스템 설정 목록")
    print("="*60)
    
    # 레지스트리 항목
    if REGISTRY_CONFIG.exists():
        with open(REGISTRY_CONFIG, "r", encoding="utf-8") as f:
            config = json.load(f)
        items = config.get("registry_items", [])
        print(f"\n[레지스트리] {len(items)}개 항목")
        print("-"*40)
        for i, item in enumerate(items, 1):
            desc = item.get("description", item["name"])
            print(f"  {i:2}. {desc}")
    
    # 명령어 항목
    if COMMANDS_CONFIG.exists():
        with open(COMMANDS_CONFIG, "r", encoding="utf-8") as f:
            config = json.load(f)
        commands = config.get("commands", [])
        enabled = [c for c in commands if c.get("enabled", True)]
        print(f"\n[명령어] {len(enabled)}개 항목 (전체 {len(commands)}개)")
        print("-"*40)
        for i, item in enumerate(enabled, 1):
            desc = item.get("description", "")
            cmd_type = item["type"].upper()
            if desc:
                print(f"  {i:2}. [{cmd_type:4}] {desc}")
            else:
                cmd_short = item["command"][:50] + "..." if len(item["command"]) > 50 else item["command"]
                print(f"  {i:2}. [{cmd_type:4}] {cmd_short}")
    
    # 인트라넷
    print(f"\n[인트라넷 영역]")
    print("-"*40)
    ip = get_local_ip()
    if ip:
        print(f"  • 현재 IP 대역 ({get_ip_range(ip)}) 자동 등록")
    
    print()


def main():
    import argparse
    
    parser = argparse.ArgumentParser(
        description="Windows 11 시스템 설정 자동 적용",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예시:
  python setup.py              # 모든 설정 적용
  python setup.py --dry-run    # 시뮬레이션 (실제 적용 안함)
  python setup.py --list       # 설정 항목 목록 보기
        """
    )
    
    parser.add_argument(
        "--dry-run", "-n",
        action="store_true",
        help="시뮬레이션 모드 (실제 적용하지 않음)"
    )
    
    parser.add_argument(
        "--list", "-l",
        action="store_true",
        help="설정 항목 목록 보기"
    )
    
    args = parser.parse_args()
    
    # 목록 보기
    if args.list:
        list_all()
        return
    
    # 실행
    print("\n" + "="*60)
    print(" Windows 11 시스템 설정 자동 적용")
    print("="*60)
    
    if args.dry_run:
        print("\n[시뮬레이션 모드] 실제 변경은 적용되지 않습니다.\n")
    
    # 1. 레지스트리 적용
    reg_success, reg_fail = apply_registry(args.dry_run)
    
    # 2. 인트라넷 설정
    inet_success, inet_fail = apply_intranet(args.dry_run)
    
    # 3. 명령어 실행
    cmd_success, cmd_fail = apply_commands(args.dry_run)
    
    # 결과 요약
    total_success = reg_success + inet_success + cmd_success
    total_fail = reg_fail + inet_fail + cmd_fail
    
    print(f"\n{'='*60}")
    print(" 완료!")
    print(f"{'='*60}")
    print(f"  총 성공: {total_success}개")
    print(f"  총 실패: {total_fail}개")
    
    if not args.dry_run and total_fail == 0:
        print("\n  ※ 일부 설정은 재부팅 후 적용됩니다.")
    
    print()


if __name__ == "__main__":
    main()
