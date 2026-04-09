import argparse
import json
import os
import re
import time
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


def parse_mmdd(value: str) -> tuple[str, str]:
    normalized = ''.join(ch for ch in value if ch.isdigit())
    if len(normalized) != 4:
        raise ValueError('날짜는 MMDD 형식으로 입력해 주세요. 예: 0309')

    month_number = int(normalized[:2])
    day_number = int(normalized[2:])
    if not (1 <= month_number <= 12 and 1 <= day_number <= 31):
        raise ValueError('유효한 날짜를 입력해 주세요. 예: 0309, 0315')

    return normalized[:2], normalized[2:]


def click_picker_value(frame, mm_or_dd: str, unit: str, occurrence: int) -> bool:
    numeric = int(mm_or_dd)
    patterns = [
        re.compile(rf'^\s*0?{numeric}\s*{unit}\s*$'),
        re.compile(rf'^\s*{mm_or_dd}\s*{unit}\s*$'),
    ]

    for pattern in patterns:
        locator = frame.get_by_text(pattern)
        if locator.count() > occurrence:
            locator.nth(occurrence).click()
            return True

    return False


def normalize_digits(text: str) -> str:
    return ''.join(ch for ch in text if ch.isdigit())


def extract_series_name(product_name: str) -> str:
    text = str(product_name or '').strip()
    # "시디즈" 다음 첫 번째 단어를 시리즈명으로 사용
    match = re.search(r'시디즈\s+([^\s,]+)', text)
    if not match:
        return text
    return match.group(1).strip()


def replace_product_names_with_series(excel_path: Path) -> tuple[int, int]:
    workbook = load_workbook(str(excel_path))
    worksheet = workbook.active

    changed = 0
    total_rows = 0
    for row in range(2, worksheet.max_row + 1):
        cell = worksheet.cell(row=row, column=5)  # E열: 상품명
        value = cell.value
        if value is None:
            continue

        total_rows += 1
        original = str(value).strip()
        series = extract_series_name(original)

        if original != series:
            cell.value = series
            changed += 1

    workbook.save(str(excel_path))
    workbook.close()
    return changed, total_rows


def to_number(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    cleaned = text.replace(',', '')
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def fill_actual_quantity_column(excel_path: Path) -> tuple[int, int]:
    workbook = load_workbook(str(excel_path))
    worksheet = workbook.active

    # U열(21): 실제수량 = H열(결제상품수량) - S열(환불수량)
    worksheet.cell(row=1, column=21).value = '실제수량'

    changed = 0
    total_rows = 0
    for row in range(2, worksheet.max_row + 1):
        payment_qty = to_number(worksheet.cell(row=row, column=8).value)   # H열
        refund_qty = to_number(worksheet.cell(row=row, column=19).value)   # S열
        actual_qty = payment_qty - refund_qty

        total_rows += 1
        if actual_qty.is_integer():
            worksheet.cell(row=row, column=21).value = int(actual_qty)
        else:
            worksheet.cell(row=row, column=21).value = actual_qty
        changed += 1

    workbook.save(str(excel_path))
    workbook.close()
    return changed, total_rows


def fill_actual_amount_column(excel_path: Path) -> tuple[int, int]:
    workbook = load_workbook(str(excel_path))
    worksheet = workbook.active

    # V열(22): 실제금액 = J열(결제금액) - Q열(환불금액)
    worksheet.cell(row=1, column=22).value = '실제금액'

    changed = 0
    total_rows = 0
    for row in range(2, worksheet.max_row + 1):
        payment_amount = to_number(worksheet.cell(row=row, column=10).value)  # J열
        refund_amount = to_number(worksheet.cell(row=row, column=17).value)   # Q열
        actual_amount = payment_amount - refund_amount

        total_rows += 1
        if actual_amount.is_integer():
            worksheet.cell(row=row, column=22).value = int(actual_amount)
        else:
            worksheet.cell(row=row, column=22).value = actual_amount
        changed += 1

    workbook.save(str(excel_path))
    workbook.close()
    return changed, total_rows


def create_series_actual_pivot(excel_path: Path) -> tuple[bool, str]:
    try:
        import win32com.client as win32  # type: ignore
    except Exception:
        return False, 'pywin32를 찾지 못해 피벗테이블 자동 생성을 건너뜀'

    excel = None
    workbook = None

    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        excel.Visible = False
        excel.DisplayAlerts = False

        workbook = excel.Workbooks.Open(str(excel_path.resolve()))
        worksheet = workbook.Worksheets(1)

        xl_up = -4162
        xl_to_left = -4159
        xl_database = 1
        xl_row_field = 1
        xl_sum = -4157
        xl_descending = 2

        last_row = worksheet.Cells(worksheet.Rows.Count, 1).End(xl_up).Row
        if last_row < 2:
            return False, '데이터 행이 없어 피벗테이블을 생성하지 못함'

        last_col = worksheet.Cells(1, worksheet.Columns.Count).End(xl_to_left).Column
        last_col = max(last_col, 22)

        pivot_sheet_name = '피벗_시리즈실제수량'
        for sheet in workbook.Worksheets:
            if sheet.Name == pivot_sheet_name:
                sheet.Delete()
                break

        pivot_sheet = workbook.Worksheets.Add()
        pivot_sheet.Name = pivot_sheet_name

        source_range = worksheet.Range(
            worksheet.Cells(1, 1),
            worksheet.Cells(last_row, last_col),
        )
        row_field_name = str(worksheet.Cells(1, 5).Value or '').strip() or '상품명'
        qty_field_name = str(worksheet.Cells(1, 21).Value or '').strip() or '실제수량'
        amount_field_name = str(worksheet.Cells(1, 22).Value or '').strip() or '실제금액'

        try:
            pivot_cache = workbook.PivotCaches().Create(
                SourceType=xl_database,
                SourceData=source_range,
            )
        except Exception:
            source_data = f"'{worksheet.Name}'!{source_range.Address(ReferenceStyle=1)}"
            pivot_cache = workbook.PivotCaches().Create(
                SourceType=xl_database,
                SourceData=source_data,
            )
        pivot_table = pivot_cache.CreatePivotTable(
            TableDestination=f"'{pivot_sheet_name}'!R3C1",
            TableName='SeriesActualQtyPivot',
        )

        row_field = pivot_table.PivotFields(row_field_name)
        row_field.Orientation = xl_row_field
        row_field.Position = 1

        data_field = pivot_table.AddDataField(
            pivot_table.PivotFields(qty_field_name),
            '실제수량 합계',
            xl_sum,
        )
        data_field.NumberFormat = '#,##0'

        amount_field = pivot_table.AddDataField(
            pivot_table.PivotFields(amount_field_name),
            '실제금액 합계',
            xl_sum,
        )
        amount_field.NumberFormat = '#,##0'

        pivot_table.RefreshTable()

        try:
            row_field.AutoSort(xl_descending, amount_field.Name)
        except Exception:
            # 일부 Excel 버전/언어 환경에서 자동 정렬 호출이 실패할 수 있음
            pass

        pivot_sheet.Range('A1').Value = '시리즈별 실제판매수량'
        pivot_sheet.Range('D3').Value = '수량비중'

        # D열 수량비중 = 각 시리즈 실제수량 / 총 실제수량
        pivot_last_row = pivot_sheet.Cells(pivot_sheet.Rows.Count, 1).End(xl_up).Row
        grand_total_row = None
        for r in range(4, pivot_last_row + 1):
            label = str(pivot_sheet.Cells(r, 1).Value or '').strip()
            if '총합계' in label or 'Grand Total' in label:
                grand_total_row = r
                break

        if grand_total_row is not None:
            for r in range(4, grand_total_row):
                # B열: 실제수량 합계, D열: 수량비중
                pivot_sheet.Cells(r, 4).FormulaR1C1 = f"=IFERROR(RC[-2]/R{grand_total_row}C[-2],0)"
            ratio_range = pivot_sheet.Range(
                pivot_sheet.Cells(4, 4),
                pivot_sheet.Cells(max(4, grand_total_row - 1), 4),
            )
            ratio_range.NumberFormat = '0.00%'
            try:
                ratio_range.NumberFormatLocal = '0.00%'
            except Exception:
                pass

        pivot_sheet.Columns('A:C').AutoFit()
        pivot_sheet.Columns('D:D').AutoFit()

        workbook.Save()
        return True, pivot_sheet_name
    except Exception as exc:
        return False, f'피벗테이블 자동 생성 실패: {exc}'
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass


def wait_for_calendar_panel(target, timeout_ms: int = 8000) -> bool:
    deadline = time.time() + (timeout_ms / 1000)
    while time.time() < deadline:
        opened = target.locator('body').evaluate(
            r"""
            () => {
              const isVisible = (el) => {
                const rect = el.getBoundingClientRect();
                if (!rect || rect.width === 0 || rect.height === 0) return false;
                const style = window.getComputedStyle(el);
                return style.visibility !== 'hidden' && style.display !== 'none';
              };

                            const visibleComplete = Array.from(document.querySelectorAll('[data-test-id="DateCommonPickInfo_click_complate"], button, a'))
                                .some((el) => isVisible(el));

                            const visibleMonth = Array.from(document.querySelectorAll('button,li,div,span,a'))
                                .filter((el) => isVisible(el))
                                .map((el) => (el.textContent || '').trim())
                                .filter((text) => /^\d{1,2}\s*월$/.test(text)).length;

                            const visibleDay = Array.from(document.querySelectorAll('button,li,div,span,a'))
                                .filter((el) => isVisible(el))
                                .map((el) => (el.textContent || '').trim())
                                .filter((text) => /^\d{1,2}\s*일$/.test(text)).length;

                            const visibleInputs = Array.from(document.querySelectorAll('input'))
                                .filter((el) => isVisible(el))
                                .filter((el) => {
                                    const p = (el.getAttribute('placeholder') || '').toLowerCase();
                                    const n = (el.getAttribute('name') || '').toLowerCase();
                                    return /(date|날짜|시작|종료|from|to|period)/.test(p + ' ' + n);
                                }).length;

                            const visibleNumeric = Array.from(document.querySelectorAll('button,li,div,span,a'))
                                .filter((el) => isVisible(el))
                                .map((el) => (el.textContent || '').trim())
                                .filter((text) => /^\d{1,2}$/.test(text)).length;

                            const hasApplyText = Array.from(document.querySelectorAll('button,a,div,span'))
                                .filter((el) => isVisible(el))
                                .some((el) => /(적용|완료|확인)/.test((el.textContent || '').trim()));

                            return hasApplyText || visibleComplete || (visibleMonth >= 2 && visibleDay >= 2) || visibleInputs >= 2 || visibleNumeric >= 20;
            }
            """
        )
        if opened:
            return True
        time.sleep(0.2)
    return False


def open_calendar_panel(page, frame, timeout_ms: int = 8000) -> bool:
    if page.is_closed():
        return False

    def safe_count(locator, cap: int) -> int:
        try:
            return min(locator.count(), cap)
        except Exception:
            return 0

    def safe_click(locator, idx: int, timeout: int = 1200) -> bool:
        try:
            locator.nth(idx).click(timeout=timeout, force=True)
            return True
        except Exception:
            return False

    contexts = [frame, page]

    selectors = [
        '[data-test-id="DateRangeFixedArea_click_toggle"]',
        '[data-test-id*="DateRange"]',
        '[data-test-id*="date"]',
        'button:has-text("지난")',
        'a:has-text("지난")',
    ]
    text_patterns = [
        re.compile(r'\d{4}\.\s*\d{2}\.\s*\d{2}\.?\s*~\s*\d{4}\.\s*\d{2}\.\s*\d{2}\.?'),
        re.compile(r'지난\s*\d+일'),
    ]

    for ctx in contexts:
        if page.is_closed():
            return False
        for selector in selectors:
            locator = ctx.locator(selector)
            total = safe_count(locator, 5)
            for idx in range(total):
                if not safe_click(locator, idx):
                    continue
                if wait_for_calendar_panel(ctx, timeout_ms=2000):
                    return True

        for pattern in text_patterns:
            text_loc = ctx.get_by_text(pattern)
            total = safe_count(text_loc, 3)
            for idx in range(total):
                if not safe_click(text_loc, idx):
                    continue
                if wait_for_calendar_panel(ctx, timeout_ms=2000):
                    return True

        # DOM 레벨 강제 클릭 (일반 클릭이 막히는 경우 대응)
        try:
            forced = ctx.locator('body').evaluate(
                r"""
                () => {
                    const isVisible = (el) => {
                        const rect = el.getBoundingClientRect();
                        if (!rect || rect.width === 0 || rect.height === 0) return false;
                        const style = window.getComputedStyle(el);
                        return style.visibility !== 'hidden' && style.display !== 'none';
                    };

                    const fire = (node, type) => {
                        const evt = new MouseEvent(type, { bubbles: true, cancelable: true, view: window });
                        node.dispatchEvent(evt);
                    };

                    const firePointer = (node, type) => {
                        try {
                            const evt = new PointerEvent(type, { bubbles: true, cancelable: true, pointerType: 'mouse', isPrimary: true });
                            node.dispatchEvent(evt);
                        } catch {
                            // ignore when PointerEvent unsupported
                        }
                    };

                    const clickWithBubble = (node) => {
                        if (!node) return false;
                        node.focus?.();
                        firePointer(node, 'pointerdown');
                        fire(node, 'mousedown');
                        fire(node, 'mouseup');
                        firePointer(node, 'pointerup');
                        fire(node, 'click');
                        return true;
                    };

                    const candidates = [];
                    const byDataTest = Array.from(document.querySelectorAll('[data-test-id*="DateRange"], [data-test-id*="date"], [data-test-id*="Date"]'));
                    for (const el of byDataTest) {
                        if (isVisible(el)) candidates.push(el);
                    }

                    const byText = Array.from(document.querySelectorAll('a,button,div,span'))
                        .filter((el) => isVisible(el))
                        .filter((el) => /(지난\s*\d+일|\d{4}\.\s*\d{2}\.\s*\d{2}\.?\s*~\s*\d{4}\.\s*\d{2}\.\s*\d{2}\.?)/.test((el.textContent || '').replace(/\s+/g, ' ')));
                    for (const el of byText) candidates.push(el);

                    for (const el of candidates) {
                        let node = el;
                        for (let i = 0; i < 7 && node; i += 1) {
                            clickWithBubble(node);
                            node = node.parentElement;
                        }
                    }

                    return candidates.length > 0;
                }
                """
            )
            if forced and wait_for_calendar_panel(ctx, timeout_ms=2500):
                return True
        except Exception:
            continue

    return wait_for_calendar_panel(frame, timeout_ms=timeout_ms) or wait_for_calendar_panel(page, timeout_ms=timeout_ms)


def click_calendar_apply(frame) -> bool:
    candidates = [
        '[data-test-id="DateCommonPickInfo_click_complate"]:visible',
        'button:has-text("적용")',
        'a:has-text("적용")',
        'button:has-text("완료")',
        'a:has-text("완료")',
        'button:has-text("확인")',
        'a:has-text("확인")',
        '[data-test-id="DateCommonPickInfo_click_complate"]',
    ]

    for selector in candidates:
        locator = frame.locator(selector).first
        try:
            if locator.count() > 0:
                locator.click(timeout=1200)
                return True
        except Exception:
            continue
    return False


def is_month_visible(frame, month_value: int) -> bool:
    month = int(month_value)
    year = datetime.now().year
    labels = [
        f'{month:02d}.',
        f'{month}.',
        f'{year}. {month:02d}.',
        f'{year}. {month}.',
        f'{year - 1}. {month:02d}.',
        f'{year - 1}. {month}.',
        f'{year + 1}. {month:02d}.',
        f'{year + 1}. {month}.',
    ]

    for label in labels:
        try:
            locator = frame.get_by_text(label, exact=True)
            if locator.count() > 0:
                return True
        except Exception:
            continue
    return False


def click_month_nav(frame, direction: str) -> bool:
    selector = (
        '[data-test-id="DateCommonNavBarM_click_prev-month"]'
        if direction == 'prev'
        else '[data-test-id="DateCommonNavBarM_click_next-month"]'
    )
    try:
        locator = frame.locator(selector).first
        if locator.count() == 0:
            return False
        locator.click(timeout=1200)
        time.sleep(0.2)
        return True
    except Exception:
        return False


def ensure_months_visible(frame, start_month: int, end_month: int, max_steps: int = 12) -> bool:
    def visible() -> bool:
        return is_month_visible(frame, start_month) and is_month_visible(frame, end_month)

    if visible():
        return True

    current_month = datetime.now().month
    prev_steps = (current_month - start_month) % 12
    next_steps = (start_month - current_month) % 12
    first_dir, second_dir = ('prev', 'next') if prev_steps <= next_steps else ('next', 'prev')

    for _ in range(max_steps):
        if visible():
            return True
        if not click_month_nav(frame, first_dir):
            break

    if visible():
        return True

    for _ in range(max_steps * 2):
        if visible():
            return True
        if not click_month_nav(frame, second_dir):
            break

    return visible()


def click_day_in_month_panel(frame, month_value: int, day_value: int, prefer_right: bool) -> bool:
    try:
        return frame.locator('body').evaluate(
            r"""
            ({ monthValue, dayValue, preferRight }) => {
                const isVisible = (el) => {
                    const rect = el.getBoundingClientRect();
                    if (!rect || rect.width === 0 || rect.height === 0) return false;
                    const style = window.getComputedStyle(el);
                    return style.visibility !== 'hidden' && style.display !== 'none';
                };

                const parseMonth = (text) => {
                    const raw = (text || '').trim();
                    const y = raw.match(/^(\d{4})\.\s*(\d{1,2})\.$/);
                    if (y) return Number(y[2]);
                    const m = raw.match(/^(\d{1,2})\.$/);
                    return m ? Number(m[1]) : null;
                };

                const parseDay = (text) => {
                    const m = (text || '').trim().match(/^(\d{1,2})$/);
                    return m ? Number(m[1]) : null;
                };

                const monthLabels = Array.from(document.querySelectorAll('button,li,a,span,div'))
                    .filter((el) => isVisible(el))
                    .map((el) => {
                        const month = parseMonth(el.textContent || '');
                        if (month === null) return null;
                        const rect = el.getBoundingClientRect();
                        return { el, month, x: rect.left + rect.width / 2 };
                    })
                    .filter(Boolean)
                    .filter((item) => item.month === Number(monthValue))
                    .sort((a, b) => a.x - b.x);

                if (!monthLabels.length) return false;
                const monthTarget = preferRight ? monthLabels[monthLabels.length - 1] : monthLabels[0];
                monthTarget.el.click();

                const dayItems = Array.from(document.querySelectorAll('button,li,a,span,div'))
                    .filter((el) => isVisible(el))
                    .map((el) => {
                        const day = parseDay(el.textContent || '');
                        if (day === null) return null;
                        const rect = el.getBoundingClientRect();
                        return {
                            el,
                            day,
                            x: rect.left + rect.width / 2,
                            y: rect.top + rect.height / 2,
                        };
                    })
                    .filter(Boolean)
                    .filter((item) => item.day === Number(dayValue));

                if (!dayItems.length) return false;

                dayItems.sort((a, b) => Math.abs(a.x - monthTarget.x) - Math.abs(b.x - monthTarget.x));
                const xNearest = dayItems.slice(0, Math.min(4, dayItems.length));
                xNearest.sort((a, b) => a.y - b.y);
                const target = xNearest[Math.floor(xNearest.length / 2)] || xNearest[0];
                target.el.click();
                return true;
            }
            """,
            {
                'monthValue': month_value,
                'dayValue': day_value,
                'preferRight': prefer_right,
            },
        )
    except Exception:
        return False


def apply_cross_month_prev_pattern(frame, start_month: int, end_month: int, start_day: int, end_day: int) -> bool:
    if start_month == end_month:
        return False

    current_month = datetime.now().month
    prev_clicks = (current_month - ((start_month % 12) + 1)) % 12

    try:
        for _ in range(prev_clicks):
            if not click_month_nav(frame, 'prev'):
                return False

        start_locator = frame.get_by_text(str(start_day), exact=True)
        end_locator = frame.get_by_text(str(end_day), exact=True)
        if start_locator.count() <= 2 or end_locator.count() <= 4:
            return False

        start_locator.nth(2).click(timeout=1000)
        end_locator.nth(4).click(timeout=1000)
        return True
    except Exception:
        return False


def force_prev_month_clicks(frame, count: int) -> bool:
    if count <= 0:
        return True
    for _ in range(count):
        if not click_month_nav(frame, 'prev'):
            return False
    return True


def apply_date_range_via_inputs(frame, start_raw: str, end_raw: str) -> bool:
        year = datetime.now().year
        start_iso = f'{year}-{start_raw[:2]}-{start_raw[2:]}'
        end_iso = f'{year}-{end_raw[:2]}-{end_raw[2:]}'
        start_dot = f'{year}.{start_raw[:2]}.{start_raw[2:]}'
        end_dot = f'{year}.{end_raw[:2]}.{end_raw[2:]}'

        try:
            return frame.locator('body').evaluate(
                r"""
                ({ startIso, endIso, startDot, endDot }) => {
                    const isVisible = (el) => {
                        const rect = el.getBoundingClientRect();
                        if (!rect || rect.width === 0 || rect.height === 0) return false;
                        const style = window.getComputedStyle(el);
                        return style.visibility !== 'hidden' && style.display !== 'none';
                    };

                    const inputs = Array.from(document.querySelectorAll('input'))
                        .filter((el) => isVisible(el));
                    if (inputs.length < 2) return false;

                    const likely = inputs
                        .map((el) => {
                            const hint = `${(el.getAttribute('placeholder') || '')} ${(el.getAttribute('name') || '')} ${(el.getAttribute('aria-label') || '')}`;
                            const score = /(date|날짜|시작|종료|from|to|period)/i.test(hint) ? 10 : 0;
                            return { el, score, x: el.getBoundingClientRect().left };
                        })
                        .sort((a, b) => b.score - a.score || a.x - b.x)
                        .slice(0, 2)
                        .sort((a, b) => a.x - b.x)
                        .map((v) => v.el);

                    if (likely.length < 2) return false;

                    const setValue = (input, value) => {
                        input.focus();
                        input.value = '';
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.value = value;
                        input.dispatchEvent(new Event('input', { bubbles: true }));
                        input.dispatchEvent(new Event('change', { bubbles: true }));
                        input.blur();
                    };

                    const startInput = likely[0];
                    const endInput = likely[1];

                    if (startInput.type === 'date') {
                        setValue(startInput, startIso);
                        setValue(endInput, endIso);
                    } else {
                        setValue(startInput, startDot);
                        setValue(endInput, endDot);
                    }

                    return true;
                }
                """,
                {
                        'startIso': start_iso,
                        'endIso': end_iso,
                        'startDot': start_dot,
                        'endDot': end_dot,
                },
            )
        except Exception:
            return False


def capture_download_with_retries(page, frame):
    click_candidates = [
        lambda: frame.get_by_role('link', name='다운로드').click(),
        lambda: frame.get_by_role('button', name='다운로드').click(),
        lambda: frame.locator('a:has-text("다운로드")').first.click(),
        lambda: frame.locator('button:has-text("다운로드")').first.click(),
        lambda: page.get_by_role('link', name='다운로드').click(),
        lambda: page.get_by_role('button', name='다운로드').click(),
    ]

    for click_action in click_candidates:
        try:
            with page.expect_download(timeout=15000) as download_info:
                click_action()
            return download_info.value
        except Exception:
            continue

    try:
        with page.expect_download(timeout=120000) as download_info:
            input('다운로드 버튼을 직접 클릭한 뒤 Enter를 누르세요: ')
        return download_info.value
    except PlaywrightTimeoutError as exc:
        raise RuntimeError('다운로드를 감지하지 못했습니다. 화면에서 다운로드 버튼/권한 팝업 상태를 확인해 주세요.') from exc


def try_select_range_label(page, frame, start_raw: str, end_raw: str) -> bool:
    year = datetime.now().year
    start_dot = f'{year}.{start_raw[:2]}.{start_raw[2:]}'
    end_dot = f'{year}.{end_raw[:2]}.{end_raw[2:]}'

    contexts = [frame, page]
    patterns = [
        re.compile(rf'{re.escape(start_dot)}\.?\s*~\s*{re.escape(end_dot)}\.?'),
        re.compile(rf'{re.escape(start_dot)}\s*~\s*{re.escape(end_dot)}'),
    ]

    for ctx in contexts:
        for pattern in patterns:
            locator = ctx.get_by_text(pattern)
            count = min(locator.count(), 5)
            for idx in range(count):
                try:
                    locator.nth(idx).click(timeout=1500, force=True)
                    return True
                except Exception:
                    continue
    return False


def date_range_matches(frame, start_raw: str, end_raw: str) -> bool:
    selected_text = frame.locator('[data-test-id="DateRangeFixedArea_click_toggle"]').inner_text()
    date_pairs = re.findall(r'(?:\d{2,4}[./-])?\s*(\d{1,2})\s*[./-]\s*(\d{1,2})', selected_text)
    if len(date_pairs) < 2:
        return False

    start_month = int(start_raw[:2])
    start_day = int(start_raw[2:])
    end_month = int(end_raw[:2])
    end_day = int(end_raw[2:])

    first_month, first_day = (int(date_pairs[0][0]), int(date_pairs[0][1]))
    second_month, second_day = (int(date_pairs[1][0]), int(date_pairs[1][1]))
    return (
        first_month == start_month
        and first_day == start_day
        and second_month == end_month
        and second_day == end_day
    )


def any_date_range_matches(page, frame, start_raw: str, end_raw: str) -> bool:
    for ctx in (frame, page):
        try:
            if date_range_matches(ctx, start_raw, end_raw):
                return True
        except Exception:
            continue
    return False


def wait_for_date_range_match(page, frame, start_raw: str, end_raw: str, timeout_ms: int = 3500) -> bool:
    deadline = time.time() + (timeout_ms / 1000)
    while time.time() < deadline:
        if any_date_range_matches(page, frame, start_raw, end_raw):
            return True
        time.sleep(0.2)
    return any_date_range_matches(page, frame, start_raw, end_raw)


def get_selected_range_text(frame) -> str:
    try:
        return frame.locator('[data-test-id="DateRangeFixedArea_click_toggle"]').inner_text().strip()
    except Exception:
        return ''


def get_selected_range_text_any(page, frame) -> str:
    for ctx in (frame, page):
        text = get_selected_range_text(ctx)
        if text:
            return text
    return ''


def dump_calendar_debug(frame, output_path: Path) -> None:
        payload = frame.locator('body').evaluate(
        r"""
                () => {
                    const isVisible = (el) => {
                        const rect = el.getBoundingClientRect();
                        if (!rect || rect.width === 0 || rect.height === 0) return false;
                        const style = window.getComputedStyle(el);
                        return style.visibility !== 'hidden' && style.display !== 'none';
                    };

                    const textCandidates = Array.from(document.querySelectorAll('button,li,a,span,div'))
                        .filter((el) => isVisible(el))
                        .map((el) => (el.textContent || '').trim())
                        .filter((t) => /^(\d{1,2})(?:\s*[월일])?$/.test(t));

                    return {
                        time: new Date().toISOString(),
                        completeButtons: Array.from(document.querySelectorAll('[data-test-id="DateCommonPickInfo_click_complate"]')).map((el) => ({
                            text: (el.textContent || '').trim(),
                            visible: isVisible(el),
                        })),
                        numericItemsSample: textCandidates.slice(0, 80),
                        bodyTextSample: (document.body.innerText || '').slice(0, 2000),
                    };
                }
                """
        )
        output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')


def auto_pick_date_range(frame, start_month: str, start_day: str, end_month: str, end_day: str) -> bool:
    try:
        return frame.locator('body').evaluate(
            r"""
                ({ startMonth, startDay, endMonth, endDay }) => {
                    const isVisible = (el) => {
                        const rect = el.getBoundingClientRect();
                        if (!rect || rect.width === 0 || rect.height === 0) return false;
                        const style = window.getComputedStyle(el);
                        return style.visibility !== 'hidden' && style.display !== 'none';
                    };

                    const complete = Array.from(document.querySelectorAll('[data-test-id="DateCommonPickInfo_click_complate"], button, a'))
                        .find((el) => {
                            const byTestId = (el.getAttribute('data-test-id') || '').includes('DateCommonPickInfo_click_complate');
                            const byText = /(완료|적용|확인)/.test((el.textContent || '').trim());
                            return isVisible(el) && (byTestId || byText);
                        });

                    if (!complete) return false;

                    let root = complete.closest('[role="dialog"]') || complete.closest('section') || complete.closest('div');
                    if (!root || root === document.body) {
                        const candidates = Array.from(document.querySelectorAll('[role="dialog"], section, div')).filter(isVisible);
                        let best = null;
                        let bestCount = -1;
                        for (const c of candidates) {
                            const count = Array.from(c.querySelectorAll('button,li,a,span,div'))
                                .map((el) => (el.textContent || '').trim())
                                .filter((t) => /^(\d{1,2})(?:\s*[월일])?$/.test(t)).length;
                            if (count > bestCount) {
                                bestCount = count;
                                best = c;
                            }
                        }
                        if (best) root = best;
                    }
                    if (!root) root = document.body;

                    const parseItem = (text) => {
                        const raw = (text || '').trim();
                        const m = raw.match(/^(\d{1,2})(?:\s*([월일]))?$/);
                        if (!m) return null;
                        return { raw, value: Number(m[1]), unit: m[2] || '' };
                    };

                    const isLeaf = (el) => {
                        if (el.children.length === 0) return true;
                        return !Array.from(el.children).some((child) => /\d/.test((child.textContent || '').trim()));
                    };

                    const items = Array.from(root.querySelectorAll('button,li,a,span,div'))
                        .filter((el) => isVisible(el) && isLeaf(el))
                        .map((el) => {
                            const parsed = parseItem(el.textContent);
                            if (!parsed) return null;
                            const rect = el.getBoundingClientRect();
                            return {
                                el,
                                raw: parsed.raw,
                                value: parsed.value,
                                unit: parsed.unit,
                                x: rect.left + rect.width / 2,
                                y: rect.top + rect.height / 2,
                            };
                        })
                        .filter(Boolean);

                    if (items.length < 10) return false;

                    const bucket = (x) => Math.round(x / 18) * 18;
                    const colMap = new Map();
                    for (const item of items) {
                        const key = bucket(item.x);
                        if (!colMap.has(key)) colMap.set(key, []);
                        colMap.get(key).push(item);
                    }

                    const cols = Array.from(colMap.entries()).map(([x, colItems]) => {
                        const values = Array.from(new Set(colItems.map((v) => v.value))).sort((a, b) => a - b);
                        const monthUnit = colItems.filter((v) => v.unit === '월').length;
                        const dayUnit = colItems.filter((v) => v.unit === '일').length;
                        const max = values[values.length - 1] ?? 0;
                        return { x, colItems, values, monthUnit, dayUnit, max };
                    }).filter((c) => c.values.length >= 6).sort((a, b) => a.x - b.x);

                    if (cols.length < 4) return false;

                    const monthCols = cols
                        .filter((c) => c.monthUnit > 0 || c.max <= 12)
                        .sort((a, b) => a.x - b.x)
                        .slice(0, 2);

                    const dayCols = cols
                        .filter((c) => c.dayUnit > 0 || c.max >= 28)
                        .sort((a, b) => a.x - b.x)
                        .slice(0, 2);

                    if (monthCols.length < 2 || dayCols.length < 2) return false;

                    const clickValue = (col, value, unitHint) => {
                        const exactWithUnit = new RegExp(`^\\s*0?${value}\\s*${unitHint}\\s*$`);
                        const exactNoUnit = new RegExp(`^\\s*0?${value}\\s*$`);

                        let candidates = col.colItems.filter((item) => exactWithUnit.test(item.raw));
                        if (!candidates.length) candidates = col.colItems.filter((item) => exactNoUnit.test(item.raw));
                        if (!candidates.length) return false;

                        const ys = col.colItems.map((i) => i.y).sort((a, b) => a - b);
                        const midY = ys[Math.floor(ys.length / 2)] || (window.innerHeight / 2);
                        candidates.sort((a, b) => Math.abs(a.y - midY) - Math.abs(b.y - midY));
                        candidates[0].el.click();
                        return true;
                    };

                    return (
                        clickValue(monthCols[0], Number(startMonth), '월') &&
                        clickValue(dayCols[0], Number(startDay), '일') &&
                        clickValue(monthCols[1], Number(endMonth), '월') &&
                        clickValue(dayCols[1], Number(endDay), '일')
                    );
                }
                """,
                {
                        'startMonth': start_month,
                        'startDay': start_day,
                        'endMonth': end_month,
                        'endDay': end_day,
                },
            )
    except Exception:
        return False


def apply_recorded_sequence(frame, start_raw: str, end_raw: str) -> bool:
    try:
        start_month, start_day = parse_mmdd(start_raw)
        end_month, end_day = parse_mmdd(end_raw)
    except Exception:
        return False

    try:
        frame.locator('[data-test-id="DateRangeFixedArea_click_toggle"]').click()
    except Exception:
        pass

    start_month_num = int(start_month)
    end_month_num = int(end_month)

    if start_raw == '1201' and end_raw == '0101':
        force_prev_month_clicks(frame, 1)
        try:
            frame.get_by_text('12.').click(timeout=1000)
            frame.get_by_text('1').nth(4).click(timeout=1000)
            frame.get_by_text('2026. 01.', exact=True).click(timeout=1000)
            frame.get_by_text('1', exact=True).nth(2).click(timeout=1000)
            return True
        except Exception:
            pass

    ensure_months_visible(frame, start_month_num, end_month_num)

    if apply_cross_month_prev_pattern(
        frame,
        start_month_num,
        end_month_num,
        int(start_day),
        int(end_day),
    ):
        return True

    if start_raw == '0210' and end_raw == '0215':
        try:
            frame.get_by_text('02.').click()
            frame.get_by_text('10').first.click()
            frame.get_by_text('15').first.click()
            return True
        except Exception:
            try:
                year = str(datetime.now().year)
                frame.get_by_text(f'{year}. 02.', exact=True).click()
                frame.get_by_text('10').first.click()
                frame.get_by_text('15').first.click()
                return True
            except Exception:
                pass

    recorded_cases = {
        ('0309', '0315'): {
            'month': '03',
            'start_text': '9',
            'start_nth': 3,
            'end_text': '15',
            'end_nth': 1,
        },
        ('0201', '0215'): {
            'month': '02',
            'start_text': '1',
            'start_nth': 2,
            'end_text': '15',
            'end_nth': 4,
        },
    }

    case = recorded_cases.get((start_raw, end_raw))
    if case:
        try:
            frame.get_by_text(f'2026. {case["month"]}.', exact=True).click()
            frame.get_by_text(case['start_text']).nth(case['start_nth']).click()
            frame.get_by_text(case['end_text']).nth(case['end_nth']).click()
            return True
        except Exception:
            try:
                year = str(datetime.now().year)
                frame.get_by_text(f'{year}. {case["month"]}.', exact=True).click()
                frame.get_by_text(case['start_text']).nth(case['start_nth']).click()
                frame.get_by_text(case['end_text']).nth(case['end_nth']).click()
                return True
            except Exception:
                pass

    if start_month == end_month:
        month_num = int(start_month)
        start_num = str(int(start_day))
        end_num = str(int(end_day))
        month_candidates = [
            f'{month_num:02d}.',
            f'{month_num}.',
            f'{datetime.now().year}. {month_num:02d}.',
            f'{datetime.now().year}. {month_num}.',
        ]

        for month_text in month_candidates:
            try:
                month_locator = frame.get_by_text(month_text, exact=True)
                if month_locator.count() > 0:
                    month_locator.first.click(timeout=1000)
            except Exception:
                continue

            start_locator = frame.get_by_text(start_num, exact=True)
            end_locator = frame.get_by_text(end_num, exact=True)
            start_count = min(start_locator.count(), 6)
            end_count = min(end_locator.count(), 6)
            if start_count == 0 or end_count == 0:
                continue

            start_indexes = [0, 1, 2, 3, 4, 5]
            end_indexes = [0, 1, 2, 3, 4, 5]

            for s_idx in start_indexes:
                if s_idx >= start_count:
                    continue
                for e_idx in end_indexes:
                    if e_idx >= end_count:
                        continue
                    try:
                        start_locator.nth(s_idx).click(timeout=900)
                        end_locator.nth(e_idx).click(timeout=900)
                        return True
                    except Exception:
                        continue

    if start_month != end_month:
        start_month_num = int(start_month)
        end_month_num = int(end_month)
        start_num = str(int(start_day))
        end_num = str(int(end_day))
        current_year = datetime.now().year

        if click_day_in_month_panel(frame, start_month_num, int(start_num), prefer_right=False):
            if click_day_in_month_panel(frame, end_month_num, int(end_num), prefer_right=True):
                return True

        start_month_candidates = [
            f'{start_month_num:02d}.',
            f'{start_month_num}.',
            f'{current_year}. {start_month_num:02d}.',
            f'{current_year}. {start_month_num}.',
            f'{current_year - 1}. {start_month_num:02d}.',
            f'{current_year - 1}. {start_month_num}.',
        ]
        end_month_candidates = [
            f'{end_month_num:02d}.',
            f'{end_month_num}.',
            f'{current_year}. {end_month_num:02d}.',
            f'{current_year}. {end_month_num}.',
            f'{current_year - 1}. {end_month_num:02d}.',
            f'{current_year - 1}. {end_month_num}.',
        ]

        start_day_locator = frame.get_by_text(start_num, exact=True)
        end_day_locator = frame.get_by_text(end_num, exact=True)
        start_count = min(start_day_locator.count(), 8)
        end_count = min(end_day_locator.count(), 8)

        def click_month_label(candidates: list[str]) -> bool:
            for label in candidates:
                try:
                    locator = frame.get_by_text(label, exact=True)
                    if locator.count() > 0:
                        locator.first.click(timeout=1000)
                        return True
                except Exception:
                    continue
            return False

        if start_count > 0 and end_count > 0:
            for s_idx in range(start_count):
                for e_idx in range(end_count):
                    try:
                        if not click_month_label(start_month_candidates):
                            continue
                        start_day_locator.nth(s_idx).click(timeout=900)

                        if not click_month_label(end_month_candidates):
                            continue
                        end_day_locator.nth(e_idx).click(timeout=900)
                        return True
                    except Exception:
                        continue

    year = datetime.now().year

    try:
        return frame.locator('body').evaluate(
            r"""
            ({ year, startMonth, startDay, endMonth, endDay }) => {
                const isVisible = (el) => {
                    const rect = el.getBoundingClientRect();
                    if (!rect || rect.width === 0 || rect.height === 0) return false;
                    const style = window.getComputedStyle(el);
                    return style.visibility !== 'hidden' && style.display !== 'none';
                };

                const parseMonthLabel = (text) => {
                    const raw = (text || '').trim();
                    const withYear = raw.match(/^(\d{4})\.\s*(\d{1,2})\.$/);
                    if (withYear) {
                        return { year: Number(withYear[1]), month: Number(withYear[2]) };
                    }
                    const monthOnly = raw.match(/^(\d{1,2})\.$/);
                    if (monthOnly) {
                        return { year: null, month: Number(monthOnly[1]) };
                    }
                    return null;
                };

                const parseDay = (text) => {
                    const m = (text || '').trim().match(/^(\d{1,2})$/);
                    return m ? Number(m[1]) : null;
                };

                const complete = Array.from(document.querySelectorAll('[data-test-id="DateCommonPickInfo_click_complate"], button, a'))
                    .find((el) => {
                        const byTestId = (el.getAttribute('data-test-id') || '').includes('DateCommonPickInfo_click_complate');
                        const byText = /(완료|적용|확인)/.test((el.textContent || '').trim());
                        return isVisible(el) && (byTestId || byText);
                    });

                let root = complete?.closest('[role="dialog"]') || complete?.closest('section') || complete?.closest('div');
                if (!root || root === document.body) root = document.body;

                const pickMonth = (monthValue, panelIndex) => {
                    const monthItems = Array.from(root.querySelectorAll('button,li,a,span,div'))
                        .filter((el) => isVisible(el))
                        .map((el) => {
                            const parsed = parseMonthLabel(el.textContent || '');
                            if (!parsed) return null;
                            const rect = el.getBoundingClientRect();
                            return {
                                el,
                                year: parsed.year,
                                month: parsed.month,
                                x: rect.left + rect.width / 2,
                            };
                        })
                        .filter(Boolean)
                        .filter((item) => item.month === Number(monthValue) && (item.year === null || item.year === Number(year)))
                        .sort((a, b) => a.x - b.x);

                    if (!monthItems.length) return false;
                    const idx = Math.min(panelIndex, monthItems.length - 1);
                    monthItems[idx].el.click();
                    return true;
                };

                const pickDay = (dayValue, panelIndex) => {
                    const isLeaf = (el) => {
                        if (el.children.length === 0) return true;
                        return !Array.from(el.children).some((child) => /\d/.test((child.textContent || '').trim()));
                    };

                    const items = Array.from(root.querySelectorAll('button,li,a,span,div'))
                        .filter((el) => isVisible(el) && isLeaf(el))
                        .map((el) => {
                            const day = parseDay(el.textContent || '');
                            if (day === null) return null;
                            const rect = el.getBoundingClientRect();
                            return {
                                el,
                                day,
                                x: rect.left + rect.width / 2,
                                y: rect.top + rect.height / 2,
                            };
                        })
                        .filter(Boolean);

                    if (!items.length) return false;

                    const bucket = (x) => Math.round(x / 18) * 18;
                    const colMap = new Map();
                    for (const item of items) {
                        const key = bucket(item.x);
                        if (!colMap.has(key)) colMap.set(key, []);
                        colMap.get(key).push(item);
                    }

                    const cols = Array.from(colMap.entries())
                        .map(([x, colItems]) => {
                            const max = Math.max(...colItems.map((v) => v.day));
                            return { x, colItems, max };
                        })
                        .filter((c) => c.max >= 28)
                        .sort((a, b) => a.x - b.x);

                    if (!cols.length) return false;
                    const idx = Math.min(panelIndex, cols.length - 1);
                    const col = cols[idx];
                    let candidates = col.colItems.filter((item) => item.day === Number(dayValue));
                    if (!candidates.length) {
                        candidates = items.filter((item) => item.day === Number(dayValue));
                    }
                    if (!candidates.length) return false;

                    const ys = col.colItems.map((i) => i.y).sort((a, b) => a - b);
                    const midY = ys[Math.floor(ys.length / 2)] || (window.innerHeight / 2);
                    candidates.sort((a, b) => Math.abs(a.y - midY) - Math.abs(b.y - midY));
                    candidates[0].el.click();
                    return true;
                };

                const startMonthNum = Number(startMonth);
                const endMonthNum = Number(endMonth);
                const startDayNum = Number(startDay);
                const endDayNum = Number(endDay);

                const isSameMonth = startMonthNum === endMonthNum;
                const monthPicked = pickMonth(startMonthNum, 0) && (isSameMonth || pickMonth(endMonthNum, 1));
                const dayPicked = isSameMonth
                    ? (pickDay(startDayNum, 0) && pickDay(endDayNum, 0))
                    : (pickDay(startDayNum, 0) && pickDay(endDayNum, 1));

                return monthPicked && dayPicked;
            }
            """,
            {
                'year': year,
                'startMonth': start_month,
                'startDay': start_day,
                'endMonth': end_month,
                'endDay': end_day,
            },
        )
    except Exception:
        return False


def required_value(value: str | None, prompt: str) -> str:
    if value and value.strip():
        return value.strip()
    return input(prompt).strip()


def load_saved_credentials(path: Path) -> tuple[str, str]:
    if not path.exists():
        return '', ''

    try:
        payload = json.loads(path.read_text(encoding='utf-8'))
    except Exception:
        return '', ''

    saved_id = str(payload.get('login_id', '')).strip()
    saved_pw = str(payload.get('login_pw', '')).strip()
    return saved_id, saved_pw


def save_credentials(path: Path, login_id: str, login_pw: str) -> None:
    payload = {
        'login_id': login_id,
        'login_pw': login_pw,
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')


def main() -> None:
    parser = argparse.ArgumentParser(description='스마트스토어 판매데이터 자동 다운로드')
    parser.add_argument('--headless', action='store_true', help='헤드리스 실행')
    parser.add_argument('--start', help='시작일(MMDD), 예: 0309')
    parser.add_argument('--end', help='종료일(MMDD), 예: 0315')
    parser.add_argument(
        '--reset-credentials',
        action='store_true',
        help='저장된 아이디/비밀번호를 삭제하고 다시 입력',
    )
    parser.add_argument(
        '--manual-calendar',
        action='store_true',
        help='캘린더 날짜를 브라우저에서 직접 선택',
    )
    parser.add_argument(
        '--until-before-calendar',
        action='store_true',
        help='캘린더를 누르기 직전까지 진행하고 종료',
    )
    args = parser.parse_args()

    credentials_path = Path.cwd() / '.smartstore_credentials.json'
    if args.reset_credentials and credentials_path.exists():
        credentials_path.unlink()

    env_id = (os.getenv('SMARTSTORE_ID') or '').strip()
    env_pw = (os.getenv('SMARTSTORE_PW') or '').strip()
    saved_id, saved_pw = load_saved_credentials(credentials_path)

    login_id = env_id or saved_id
    login_pw = env_pw or saved_pw

    if not login_id:
        login_id = required_value(None, 'SMARTSTORE_ID 입력(최초 1회 저장): ')
    if not login_pw:
        login_pw = required_value(None, 'SMARTSTORE_PW 입력(최초 1회 저장): ')

    if not env_id and not env_pw and (not saved_id or not saved_pw):
        save_credentials(credentials_path, login_id, login_pw)

    start_raw = (args.start or '').strip()
    end_raw = (args.end or '').strip()
    start_month = ''
    start_day = ''
    end_month = ''
    end_day = ''

    if not args.manual_calendar:
        if not start_raw:
            start_raw = input('시작일(MMDD): ').strip()
        if not end_raw:
            end_raw = input('종료일(MMDD): ').strip()

        if not start_raw:
            raise ValueError('시작일은 비워둘 수 없습니다. 실행 전에 반드시 입력해 주세요.')
        if not end_raw:
            raise ValueError('종료일은 비워둘 수 없습니다. 실행 전에 반드시 입력해 주세요.')

        start_month, start_day = parse_mmdd(start_raw)
        end_month, end_day = parse_mmdd(end_raw)

    downloads_dir = Path.cwd() / 'downloads'
    downloads_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=args.headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        page.goto('https://sell.smartstore.naver.com/#/home/about')
        page.get_by_role('button', name='로그인하기').click()
        page.get_by_role('textbox', name='아이디 또는 이메일 주소').fill(login_id)
        page.get_by_role('textbox', name='비밀번호').fill(login_pw)
        page.get_by_role('button', name='로그인', exact=True).click()

        captcha_box = page.get_by_role('textbox', name='자동입력 방지 문자')
        if captcha_box.is_visible(timeout=8000):
            captcha = input('캡차: ').strip()
            if not captcha:
                raise ValueError('캡차가 표시된 경우에는 캡차를 입력해 주세요.')
            captcha_box.fill(captcha)
            page.get_by_role('button', name='로그인', exact=True).click()

        page.get_by_role('menuitem', name='데이터분석').click()
        page.get_by_role('link', name='판매분석').click()

        frame = page.frame_locator('#__delegate')
        if args.until_before_calendar:
            input('캘린더 클릭 직전입니다. 화면 확인 후 Enter를 누르면 종료합니다.')
            context.close()
            browser.close()
            return

        panel_opened = open_calendar_panel(page, frame, timeout_ms=8000)
        if not panel_opened:
            if not try_select_range_label(page, frame, start_raw, end_raw):
                raise RuntimeError('캘린더 패널을 열지 못했습니다. 페이지 상태를 확인해 주세요.')
        if args.manual_calendar:
            input('브라우저에서 캘린더 날짜를 직접 선택한 뒤 Enter를 누르세요.')
        else:
            applied = False
            if panel_opened:
                for _ in range(2):
                    # 녹화된 고정 시퀀스 우선 (시작일 -> 종료일)
                    if apply_recorded_sequence(frame, start_raw, end_raw):
                        already_matched = wait_for_date_range_match(page, frame, start_raw, end_raw, timeout_ms=900)
                        apply_clicked = click_calendar_apply(frame)
                        if (already_matched or apply_clicked) and wait_for_date_range_match(page, frame, start_raw, end_raw):
                            applied = True
                            break

                    for ctx in (frame, page):
                        if auto_pick_date_range(ctx, start_month, start_day, end_month, end_day):
                            apply_clicked = click_calendar_apply(ctx)
                            if apply_clicked and wait_for_date_range_match(page, frame, start_raw, end_raw):
                                applied = True
                                break
                    if applied:
                        break

                    for ctx in (frame, page):
                        input_set = apply_date_range_via_inputs(ctx, start_raw, end_raw)
                        if input_set:
                            apply_clicked = click_calendar_apply(ctx)
                            if apply_clicked and wait_for_date_range_match(page, frame, start_raw, end_raw):
                                applied = True
                                break
                    if applied:
                        break

                    if not open_calendar_panel(page, frame, timeout_ms=8000):
                        break

            if not applied and try_select_range_label(page, frame, start_raw, end_raw):
                applied = wait_for_date_range_match(page, frame, start_raw, end_raw)

            if not applied:
                dump_calendar_debug(frame, Path.cwd() / 'calendar-debug.json')
                current_range = get_selected_range_text_any(page, frame)
                raise RuntimeError(f'시작일/종료일 자동 선택 실패. 현재 선택값: {current_range}')

            current_range = get_selected_range_text_any(page, frame)
            print(f'선택된 기간: {current_range}')

        page.get_by_role('link', name='상품성과').click()

        download = capture_download_with_retries(page, frame)

        suggested = download.suggested_filename
        ext = Path(suggested).suffix or '.xlsx'
        if args.manual_calendar:
            timestamp = datetime.now().strftime('%Y%m%d-%H%M%S')
            output_path = downloads_dir / f'sales-manual-{timestamp}{ext}'
        else:
            output_path = downloads_dir / f'sales-{start_raw}-{end_raw}{ext}'
        download.save_as(str(output_path))

        changed_count, total_count = replace_product_names_with_series(output_path)
        actual_changed_count, actual_total_count = fill_actual_quantity_column(output_path)
        amount_changed_count, amount_total_count = fill_actual_amount_column(output_path)
        pivot_ok, pivot_message = create_series_actual_pivot(output_path)

        print(f'저장 완료: {output_path}')
        print(f'E열 시리즈명 변환 완료: {changed_count}/{total_count}행')
        print(f'U열 실제수량 계산 완료: {actual_changed_count}/{actual_total_count}행')
        print(f'V열 실제금액 계산 완료: {amount_changed_count}/{amount_total_count}행')
        if pivot_ok:
            print(f'피벗테이블 생성 완료: {pivot_message}')
        else:
            print(f'피벗테이블 생성 건너뜀/실패: {pivot_message}')
        context.close()
        browser.close()


if __name__ == '__main__':
    main()
