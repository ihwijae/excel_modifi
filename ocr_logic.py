from openpyxl import load_workbook
from config import RELATIVE_OFFSETS, COLUMN_MAP, RATIO_THRESHOLDS
from openpyxl.styles import PatternFill, Color, Font
from datetime import datetime
import re


def find_company_data(excel_path, biz_no_to_find):
    """
    엑셀 파일에서 업체를 찾아, 값과 함께 셀 배경색 정보도 반환합니다.
    """
    try:
        workbook = load_workbook(filename=excel_path, data_only=False)
    except Exception as e:
        return None, f"엑셀 파일 열기 오류: {e}"

    print(f"\n--- [진단 시작] 엑셀에서 데이터 조회 ---")
    print(f"  - 엑셀 파일 경로: {excel_path}")
    print(f"  - 찾으려는 사업자번호: '{biz_no_to_find}'")

    target_row, target_col, target_sheet_name = -1, -1, None;
    found = False
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        print(f"\n- 시트 '{sheet_name}' 검색 중...")
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None: continue
                if str(cell.value).strip().replace('-', '') == biz_no_to_find.strip().replace('-', ''):
                    target_row, target_col, target_sheet_name = cell.row, cell.column, sheet_name
                    found = True;
                    break
            if found: break
        if found: break

    if not found:
        print("  - [진단 결과] 엑셀의 모든 시트에서 해당 사업자번호를 찾지 못했습니다.")
        print("--- [진단 종료] ---\n")
        return None, None

    print(f"  - [진단] 업체 위치 찾기 성공!: '{target_sheet_name}' 시트, {target_row}행, {target_col}열")
    print(f"  - [진단] 이제 해당 위치를 기준으로 상대 오프셋 데이터를 읽습니다...")

    sheet = workbook[target_sheet_name]
    found_data = {}
    for key, excel_label in COLUMN_MAP.items():
        if excel_label in RELATIVE_OFFSETS:
            row_offset = RELATIVE_OFFSETS[excel_label]
            read_row = target_row + row_offset
            if 1 <= read_row <= sheet.max_row and 1 <= target_col <= sheet.max_column:
                cell = sheet.cell(row=read_row, column=target_col)
                value = cell.value

                color_hex = "#FFFFFF"
                if cell.fill and cell.fill.fgColor:
                    color_info = cell.fill.fgColor
                    if color_info.type == 'theme':
                        if color_info.theme == 6:
                            color_hex = "#E2EFDA"
                        elif color_info.theme == 3:
                            color_hex = "#DDEBF7"
                    elif color_info.type == 'rgb' and isinstance(color_info.rgb, str):
                        hex_val = color_info.rgb
                        color_hex = f"#{hex_val[2:]}" if len(hex_val) == 8 and hex_val.startswith(
                            "FF") else f"#{hex_val}"

                # [핵심 수정] '투명한 검은색'으로 인식되는 경우를 일반 '흰색'으로 처리
                if color_hex == '#00000000':
                    color_hex = '#FFFFFF'

                found_data[key] = {'value': value, 'color': color_hex}
                print(f"    - '{excel_label}' (행:{read_row}, 열:{target_col}) -> 값: {value} | 색상: {color_hex}")
            else:
                found_data[key] = {'value': 'N/A', 'color': '#FFFFFF'}

    print(f"\n  - [진단 결과] 최종적으로 반환하는 데이터: {found_data}")
    print("--- [진단 종료] ---\n")
    return found_data, None


def update_company_data(excel_path, biz_no_to_find, update_data, db_type):
    """
    엑셀에서 업체를 찾아 데이터를 업데이트하고, 조건에 따라 서식을 변경합니다.
    """
    try:
        workbook = load_workbook(filename=excel_path)
    except Exception as e:
        return None, f"엑셀 파일 열기 오류: {e}"

    # --- 서식 정의 ---
    THEME_GREEN_COLOR = Color(type='theme', theme=6, tint=0.7999816888943144)
    GREEN_FILL = PatternFill(fgColor=THEME_GREEN_COLOR, fill_type="solid")

    # [수정] 폰트 크기를 11로 고정
    DEFAULT_FONT = Font(color="000000", bold=False, size=9)
    HIGHLIGHT_FONT = Font(color="FF0000", bold=True, size=9)

    # --- 업체 위치 찾기 (기존과 동일) ---
    target_row, target_col, target_sheet_name = -1, -1, None;
    found = False
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None: continue
                if str(cell.value).strip().replace('-', '') == biz_no_to_find.strip().replace('-', ''):
                    target_row, target_col, target_sheet_name = cell.row, cell.column, sheet_name
                    found = True;
                    break
            if found: break
        if found: break
    if not found: return None, f"엑셀 파일에서 사업자번호 '{biz_no_to_find}'를 찾을 수 없습니다."

    # --- 데이터 업데이트 및 서식 적용 ---
    sheet = workbook[target_sheet_name]
    updated_log = []

    for key, excel_label in COLUMN_MAP.items():
        if excel_label in RELATIVE_OFFSETS:
            row_offset = RELATIVE_OFFSETS[excel_label]
            update_row = target_row + row_offset
            if 1 <= update_row <= sheet.max_row and 1 <= target_col <= sheet.max_column:
                cell = sheet.cell(row=update_row, column=target_col)
                if key not in ['상호', '신용평가']: cell.fill = GREEN_FILL

                if key in update_data and update_data[key]:
                    cell.font = DEFAULT_FONT  # 우선 기본 폰트로 초기화
                    value_str = str(update_data[key]).replace(",", "").replace("%", "")

                    try:
                        numeric_value = 0
                        if '비율' in key:
                            numeric_value = float(value_str)
                            cell.value = numeric_value / 100.0
                            cell.number_format = '0.00%'
                        elif key in ['시평액', '3년실적', '5년실적']:
                            cell.value = int(float(value_str)) * 1000
                        else:
                            cell.value = update_data[key]
                        updated_log.append(excel_label)

                        if db_type and key in ['부채비율', '유동비율']:
                            thresholds = RATIO_THRESHOLDS.get(db_type, {}).get(key, {})
                            if 'max' in thresholds and numeric_value > thresholds['max']:
                                cell.font = HIGHLIGHT_FONT
                            elif 'min' in thresholds and numeric_value < thresholds['min']:
                                cell.font = HIGHLIGHT_FONT

                    except (ValueError, TypeError):
                        pass

    try:
        workbook.save(excel_path)
        return updated_log, None
    except Exception as e:
        return None, f"엑셀 파일 저장 오류: {e}"


def batch_update_colors(excel_path):
    """
    엑셀 파일의 모든 데이터 셀을 순회하며, 상태 색상을 갱신합니다.
    - '신용평가' 행은 이 로직에서 제외됩니다.
    - 데이터가 있는 셀: 초록 -> 파랑, 파랑 -> 흰색
    - 데이터가 없는 셀: 모두 흰색(색 없음)으로 정리
    """
    try:
        workbook = load_workbook(filename=excel_path)
    except Exception as e:
        return f"엑셀 파일 열기 오류: {e}"

    GREEN_COLOR = Color(type='theme', theme=6, tint=0.7999816888943144)
    BLUE_COLOR = Color(type='theme', theme=3, tint=0.7999816888943144)
    NEW_BLUE_FILL = PatternFill(fgColor=BLUE_COLOR, fill_type="solid")
    NO_FILL = PatternFill(fill_type=None)

    update_count = 0
    for sheet in workbook.worksheets:
        # 데이터가 있는 모든 행을 순회 (min_row=2는 헤더 제외)
        for row in sheet.iter_rows(min_row=2):
            # [핵심] 행의 첫 번째 셀(A열)의 값을 확인하여 '신용평가'인지 검사
            label_cell = row[0]
            if label_cell.value and '신용평가' in str(label_cell.value):
                continue  # '신용평가'가 포함된 행은 건너뜀

            # '신용평가' 행이 아닐 경우에만, 기존 색상 변경 로직 실행
            # (A열을 제외한 B열부터 순회)
            for cell in row[1:]:
                current_color = cell.fill.fgColor if cell.fill else None

                if cell.value is None or str(cell.value).strip() == "":
                    if cell.fill and cell.fill.fill_type is not None:
                        cell.fill = NO_FILL
                        update_count += 1
                else:
                    if current_color == GREEN_COLOR:
                        cell.fill = NEW_BLUE_FILL
                        update_count += 1
                    elif current_color == BLUE_COLOR:
                        cell.fill = NO_FILL
                        update_count += 1
    try:
        workbook.save(excel_path)
        return f"총 {update_count}개 셀의 서식을 성공적으로 업데이트했습니다."
    except Exception as e:
        return f"엑셀 파일 저장 오류: {e}"

def update_credit_rating_only(excel_path, biz_no_to_find, new_credit_rating):
    """
    엑셀 파일에서 사업자번호로 업체를 찾아 '신용평가' 항목만 업데이트하고,
    해당 셀을 초록색으로 칠합니다.
    """
    try:
        workbook = load_workbook(filename=excel_path)
    except Exception as e:
        return None, f"엑셀 파일 열기 오류: {e}"

    # [수정] 신용평가 업데이트 시 사용할 색상을 초록색으로 변경
    THEME_GREEN_COLOR = Color(type='theme', theme=6, tint=0.7999816888943144)
    GREEN_FILL = PatternFill(fgColor=THEME_GREEN_COLOR, fill_type="solid")

    # --- 업체 위치 찾기 ---
    target_row, target_col, target_sheet_name = -1, -1, None
    found = False
    for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is None: continue
                if str(cell.value).strip().replace('-', '') == biz_no_to_find.strip().replace('-', ''):
                    target_row, target_col, target_sheet_name = cell.row, cell.column, sheet_name
                    found = True; break
            if found: break
        if found: break
    
    if not found:
        return f"해당 업체를 찾을 수 없습니다.", None

    # --- '신용평가' 셀 찾아서 업데이트 ---
    sheet = workbook[target_sheet_name]
    
    # config.py에서 '신용평가'의 상대 위치 가져오기
    credit_rating_offset = RELATIVE_OFFSETS.get('신용평가')
    if credit_rating_offset is None:
        return None, "'config.py'의 RELATIVE_OFFSETS에 '신용평가'가 정의되지 않았습니다."

    update_row = target_row + credit_rating_offset
    if not (1 <= update_row <= sheet.max_row and 1 <= target_col <= sheet.max_column):
        return None, f"'신용평가' 셀의 위치({update_row}행)가 유효하지 않습니다."
        
    cell_to_update = sheet.cell(row=update_row, column=target_col)
    cell_to_update.value = new_credit_rating
    
    # [수정] 셀 채우기를 초록색으로 적용
    cell_to_update.fill = GREEN_FILL

    try:
        workbook.save(excel_path)
        return "업데이트 완료!", None
    except Exception as e:
        return None, f"엑셀 파일 저장 오류: {e}"


# [ocr_logic.py 파일에서 batch_update_credit_rating_colors 함수를 이걸로 교체]

def batch_update_credit_rating_colors(excel_path):
    """
    (수정) '신용평가' 행을 찾아 유효기간에 따라 색상을 갱신합니다.
    - 빈 셀은 '색 없음'으로 처리합니다.
    """
    try:
        workbook = load_workbook(filename=excel_path)
    except Exception as e:
        return f"엑셀 파일 열기 오류: {e}"

    GREEN_FILL = PatternFill(fgColor=Color(type='theme', theme=6, tint=0.7999816888943144), fill_type="solid")
    BLUE_FILL = PatternFill(fgColor=Color(type='theme', theme=3, tint=0.7999816888943144), fill_type="solid")
    NO_FILL = PatternFill(fill_type=None)  # 색 없음을 위한 스타일
    today = datetime.now().date()
    update_count = 0

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=2):
            label_cell = row[0]
            if label_cell.value and '신용평가' in str(label_cell.value):
                for cell in row[1:]:
                    # [핵심 수정] 셀이 비어있는 경우, 색을 없애고 다음 셀로 넘어감
                    if cell.value is None or str(cell.value).strip() == "":
                        cell.fill = NO_FILL
                        update_count += 1
                        continue

                    match = re.search(r'~(\d{2,4}\.\d{2}\.\d{2})', str(cell.value))
                    if not match:
                        continue

                    end_date_str = match.group(1)
                    try:
                        expiry_date = datetime.strptime(end_date_str, '%y.%m.%d').date()
                        if expiry_date < today:
                            cell.fill = BLUE_FILL
                        else:
                            cell.fill = GREEN_FILL
                        update_count += 1
                    except ValueError:
                        continue
    try:
        workbook.save(excel_path)
        return f"총 {update_count}개의 신용평가 셀 색상을 갱신했습니다."
    except Exception as e:
        return f"엑셀 파일 저장 오류: {e}"