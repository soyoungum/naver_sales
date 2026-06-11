"""스마트스토어 판매·마케팅 통합 대시보드 (v1)"""
from __future__ import annotations

import json
import re
from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from streamlit_autorefresh import st_autorefresh

BASE = Path(__file__).parent
DOWNLOADS = BASE / 'downloads'
DOWNLOADS_SUPER = BASE / 'downloads_super'   # 슈퍼적립(월별) 전용 입력 폴더
CACHE_DIR = BASE / '.dashboard_cache'    # 진실의 원천 — downloads는 입력 채널만
CACHE_DIR.mkdir(exist_ok=True)
DOWNLOADS_SUPER.mkdir(exist_ok=True)

WEEK_RE = re.compile(r'(sales|marketing)-(\d{6})-(\d{6})\.xlsx$', re.IGNORECASE)
# 슈퍼적립 파일명: 두 가지 형식 모두 허용
#   - super-{YYMM}.xlsx                          (간결형)
#   - 상품성과_YYYY-MM-DD_YYYY-MM-DD.xlsx        (스마트스토어 기본 다운로드명)
SUPER_RE = re.compile(r'super-(\d{4})\.xlsx$', re.IGNORECASE)
SUPER_RAW_RE = re.compile(
    r'^상품성과_(\d{4})-(\d{2})-\d{2}_\d{4}-\d{2}-\d{2}\.xlsx$'
)
EXCLUDE_SERIES_PATTERN = r'토트넘|&'   # 조회에서 제외할 시리즈 키워드

# 다크 테마용 파스텔 팔레트
PASTEL_PREV = '#B8B8C4'      # 이전 주차 (중성 회보라)
PASTEL_BLUE = '#9BC4E2'      # 매출 계열
PASTEL_GREEN = '#A8D5B5'     # 수량 계열
PASTEL_CORAL = '#F5B7B1'     # 광고비 계열
PASTEL_PEACH = '#FFCB99'     # 강조 라인
PASTEL_QUAL = ['#9BC4E2', '#A8D5B5', '#FFCB99', '#F5B7B1',
               '#C9B4E8', '#FFE5A0', '#B5E0E0', '#F0B8D9',
               '#D4C5A9', '#A8C8B0']


# ── 데이터 로딩 ───────────────────────────────────────────────────────────────
# 정책: downloads/의 파일은 캐시(.dashboard_cache/)로 한 번 적재되면 영구 보존.
# downloads/에서 파일이 삭제돼도 대시보드는 캐시를 계속 보여줌.
# 캐시에서 빼려면 사이드바 '데이터 관리'에서 명시적으로 삭제.

def scan_downloads() -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {'sales': {}, 'marketing': {}}
    for p in sorted(DOWNLOADS.glob('*.xlsx')):
        m = WEEK_RE.match(p.name)
        if not m:
            continue
        kind, start, end = m.group(1).lower(), m.group(2), m.group(3)
        out[kind][f'{start}-{end}'] = str(p)
    return out


def sync_cache() -> None:
    """downloads/의 새/변경된 파일을 캐시에 반영. 캐시 파일은 절대 자동삭제 안 함."""
    files = scan_downloads()
    for kind, sheet in (('sales', '상품성과'), ('marketing', '전체채널')):
        for week, src in files[kind].items():
            cache_file = CACHE_DIR / f'{kind}-{week}.pkl'
            try:
                src_mtime = Path(src).stat().st_mtime
                cache_mtime = cache_file.stat().st_mtime if cache_file.exists() else 0
                if src_mtime > cache_mtime:
                    df = pd.read_excel(src, sheet_name=sheet)
                    if kind == 'sales' and '상품ID' in df.columns:
                        pid_col = df['상품ID'].astype(str).str.strip().str.removesuffix('.0')
                        mapped = pid_col.map(PRODUCT_CODE_MAP)
                        df['상품명'] = mapped.where(mapped.notna(), df.get('상품명'))
                    df['주차'] = week
                    df.to_pickle(cache_file)
            except Exception as e:
                st.warning(f'캐시 갱신 실패: {Path(src).name} — {e}')

    # 페이지 푸터 합계 JSON도 캐시로 복사 (sales만)
    for src in DOWNLOADS.glob('sales-*.totals.json'):
        m = re.match(r'sales-(\d{6}-\d{6})\.totals\.json$', src.name)
        if not m:
            continue
        dst = CACHE_DIR / src.name
        try:
            if not dst.exists() or src.stat().st_mtime > dst.stat().st_mtime:
                dst.write_bytes(src.read_bytes())
        except Exception as e:
            st.warning(f'totals JSON 캐시 갱신 실패: {src.name} — {e}')

    # 목표 거래액 JSON도 캐시로 복사
    for src in DOWNLOADS.glob('sales-*.target.json'):
        m = re.match(r'sales-(\d{6}-\d{6})\.target\.json$', src.name)
        if not m:
            continue
        dst = CACHE_DIR / src.name
        try:
            if not dst.exists() or src.stat().st_mtime > dst.stat().st_mtime:
                dst.write_bytes(src.read_bytes())
        except Exception as e:
            st.warning(f'target JSON 캐시 갱신 실패: {src.name} — {e}')


# ── 품목코드 → 시리즈명 매핑 ────────────────────────────────────────────────
# 상품ID(개별상품코드) 기준으로 시리즈명을 결정한다.
# sales xlsx 및 슈퍼적립 xlsx 모두 이 맵을 우선 적용하며,
# 미등록 코드는 기존 이름 추출 로직으로 fallback.
PRODUCT_CODE_MAP: dict[str, str] = {
    '11702025536': 'T50AIR', '11702025535': 'T50AIR', '467324523': 'T50AIR',
    '11702025538': 'T50AIR', '11702025537': 'T50AIR', '467324898': 'T50AIR',
    '11838839717': 'T50AIR', '6290832927': 'T50AIR',
    '12010460369': 'T50HDA',
    '11589281484': 'T50HLDA', '11589281483': 'T50HLDA', '134655631': 'T50HLDA',
    '11589281482': 'T50HLDA', '100127381': 'T50HLDA',
    '11589326677': 'T50HLDA', '11589326676': 'T50HLDA',
    '100126838': 'T50HF', '11701977457': 'T50HF', '11701977456': 'T50HF',
    '11700645975': 'T50HA', '6290833234': 'T50HA', '11700645976': 'T50HA',
    '12344899406': 'T60', '12775032955': 'T60', '12775032957': 'T60',
    '12416841336': 'T60 AIR', '12775356713': 'T60 AIR', '12775356715': 'T60 AIR',
    '12042037926': 'T20', '4974167370': 'T20', '11834696436': 'T20',
    '13370914201': 'T20', '11700517388': 'T20', '11700517389': 'T20',
    '11700517390': 'T20', '111993473': 'T20', '11700517387': 'T20',
    '11838823303': '아이블', '6290832721': '아이블',
    '11589555610': '아이블', '11589555609': '아이블', '5361883475': '아이블',
    '11700385594': '아이블높조', '9108487774': '아이블높조', '11700385593': '아이블높조',
    '437340742': 'T80', '11700473841': 'T80', '376965970': 'T80',
    '11700452519': 'T80', '11700415055': 'T80', '697587221': 'T80',
    '11609082711': '링고', '325093734': '링고',
    '11589206826': 'T90', '10045333478': 'T90', '11589188293': 'T90',
    '10045333560': 'T90', '12410629837': 'T90', '12379957016': 'T90',
    '11676176031': 'GX', '11676176030': 'GX', '10828007369': 'GX', '11676176029': 'GX',
    '8130702020': 'GC', '11676082590': 'GC', '7567795565': 'GC',
    '11589574601': 'GC', '10770156959': 'GC', '11589548593': 'GC',
    '11700339996': '리니에', '366841266': '리니에', '11700339994': '리니에',
    '11700339995': '리니에', '11700354887': '리니에', '366840759': '리니에',
    '11700354885': '리니에', '11700354886': '리니에',
    '9108487810': '탭스퀘어', '11589449301': '탭스퀘어', '11589449302': '탭스퀘어',
    '13204280111': '플릿',
    '13008808893': '뮤브',
    '12410300778': '에가 데일리', '12410300777': '에가 데일리', '12348541519': '에가 데일리',
    '11809993791': '에가', '101439394': '에가', '11809993788': '에가',
    '11809993790': '에가', '11809993789': '에가', '11810037576': '에가',
    '10765100902': '에가', '11810037573': '에가', '11810037575': '에가',
    '11810037574': '에가', '100126932': '에가', '11589612417': '에가',
    '11589612418': '에가', '11589612420': '에가', '11589612419': '에가',
    '10778511559': '에가', '11676216728': '에가', '11676216727': '에가',
    '11676216726': '에가', '11676216725': '에가',
    '11202343206': '트레보', '11676148382': '트레보', '11676148383': '트레보',
    '11589443707': '스테포', '11589443709': '스테포', '4836145633': '스테포',
    '11589475042': '버튼', '10790537178': '버튼', '11589507801': '버튼',
    '10778436023': '버튼', '11676462647': '버튼', '561843454': '버튼',
    '11676244218': '버튼', '7570461447': '버튼', '10778452389': '버튼',
    '11589576161': '버튼', '11589576159': '버튼', '10717434724': '버튼',
    '11589595003': '버튼', '11589595005': '버튼', '264588612': '버튼',
    '11589549995': '버튼', '11589549994': '버튼', '7570468820': '버튼',
    '11700401854': '버튼', '11700401853': '버튼',
    '11744333882': '펑거스', '11744333881': '펑거스', '10790502446': '펑거스',
    '11744333885': '펑거스', '11744333884': '펑거스', '11744333883': '펑거스',
    '11744442124': '펑거스 무브', '11744442123': '펑거스 무브', '3772569415': '펑거스 무브',
    '11744442127': '펑거스 무브', '11744442126': '펑거스 무브', '11744442125': '펑거스 무브',
    '11589573178': '위', '5974215913': '위', '11589587242': '위',
    '11589556238': '위', '11589526985': '위', '5974218721': '위',
    '11589526987': '위', '11589526986': '위', '11589549047': '위',
    '11589463041': '위', '11589549050': '위', '11589549048': '위',
    '11589448885': '필로', '369891819': '필로', '318550155': '필로',
    '11589432086': '필로', '11589432087': '필로', '11589432084': '필로',
    '140240040': '마네', '100128960': '마네', '5928957303': '마네',
    '11676056879': '마네', '11676056877': '마네', '11676056875': '마네', '12015489271': '마네',
    '4969694626': '몰티', '4963358016': '몰티', '4974829829': '몰티',
    '2152975441': '아띠', '11676022674': '아띠', '11676022673': '아띠',
    '2152978774': '아띠', '2152985723': '아띠', '11838794884': '아띠', '11838794882': '아띠',
    '11942778437': '토트넘 GC PRO',
    '11942762236': '토트넘 올리',
    '11942761924': '토트넘 펑거스',
    '3525409790': '액세서리', '11829643913': '액세서리', '11829643911': '액세서리',
    '11202342906': '액세서리', '3525503073': '액세서리', '3523982189': '액세서리',
    '4969694543': '액세서리',
}

# ── 슈퍼적립(월별) 처리 ───────────────────────────────────────────────────────
# downloads_super/의 raw 상품성과 xlsx에는 시리즈명 분류와 실제수량/금액이 없음.
# smartstore_recording.process_excel과 동일한 규칙을 여기서 다시 적용해 캐시한다.
# (smartstore_recording은 playwright 의존성이 있어 dashboard에서 직접 import 불가)
_SUPER_ACCESSORY_KEYWORDS = ('커버', '옷걸이', '레버', '받침대', '등판', '스프레이')
_SUPER_NO_GEN1_SUFFIX_SERIES = ('링고',)


def _super_extract_series(product_name: str) -> str:
    text = str(product_name or '').strip()
    if any(k in text for k in _SUPER_ACCESSORY_KEYWORDS):
        return '악세사리'
    m = re.search(r'시디즈\s+([^\s,\[]+)', text)
    if not m:
        # 시디즈 패턴 매칭 실패: 토트넘 콜라보는 EXCLUDE 필터(load 시점)에서
        # 잡히도록 raw 유지, 그 외 분류 불가 항목은 모두 악세사리로 묶기
        if '토트넘' in text:
            return text
        return '악세사리'
    series = m.group(1).strip()
    while series.endswith(' 1세대'):
        series = series[: -len(' 1세대')].strip()
    if series in _SUPER_NO_GEN1_SUFFIX_SERIES:
        return series
    if '1세대' in text:
        series = f'{series} 1세대'
    return series


def scan_downloads_super() -> dict[str, str]:
    """downloads_super/의 슈퍼적립 xlsx 스캔. {YYMM: path}.

    `super-{YYMM}.xlsx` 또는 `상품성과_YYYY-MM-DD_..._..._...xlsx` 모두 허용.
    같은 월에 두 형식이 모두 있으면 최근 수정된 파일을 사용.
    """
    out: dict[str, str] = {}
    for p in sorted(DOWNLOADS_SUPER.glob('*.xlsx')):
        ym: str | None = None
        m = SUPER_RE.match(p.name)
        if m:
            ym = m.group(1)
        else:
            m2 = SUPER_RAW_RE.match(p.name)
            if m2:
                yyyy, mm = m2.group(1), m2.group(2)
                ym = f'{yyyy[2:]}{mm}'   # 2026, 01 → 2601
        if ym is None:
            continue
        prev = out.get(ym)
        if prev is None or Path(p).stat().st_mtime > Path(prev).stat().st_mtime:
            out[ym] = str(p)
    return out


def sync_super_cache() -> None:
    """downloads_super/의 새/변경된 슈퍼적립 파일을 캐시(super-{YYMM}.pkl)에 반영.

    raw 상품성과 xlsx → 시리즈명 분류 + 실제수량/실제금액 계산 → pkl 저장.
    원본 xlsx는 건드리지 않음.
    """
    for ym, src in scan_downloads_super().items():
        cache_file = CACHE_DIR / f'super-{ym}.pkl'
        try:
            src_mtime = Path(src).stat().st_mtime
            cache_mtime = cache_file.stat().st_mtime if cache_file.exists() else 0
            if src_mtime <= cache_mtime:
                continue

            df = pd.read_excel(src, sheet_name='상품성과')
            # 위치 기반 컬럼 접근 (smartstore_recording.process_excel과 동일):
            # E(idx 4) 상품명, F(5) 상품ID, H(7) 결제상품수량, J(9) 결제금액,
            # Q(16) 환불금액, S(18) 환불수량
            if df.shape[1] < 19:
                st.warning(f'슈퍼적립 파일 컬럼 수 부족: {Path(src).name}')
                continue

            names = df.iloc[:, 4].astype(str)
            pids = (df.iloc[:, 5].astype(str).str.strip()
                    .str.removesuffix('.0'))
            series = [
                PRODUCT_CODE_MAP.get(pid, _super_extract_series(name))
                for name, pid in zip(names, pids)
            ]
            df['원본상품명'] = names           # 시리즈 분류 전 원본 상품명 보존
            df['상품명'] = series

            qty = pd.to_numeric(df.iloc[:, 7], errors='coerce').fillna(0)
            qty_ref = pd.to_numeric(df.iloc[:, 18], errors='coerce').fillna(0)
            amt = pd.to_numeric(df.iloc[:, 9], errors='coerce').fillna(0)
            amt_ref = pd.to_numeric(df.iloc[:, 16], errors='coerce').fillna(0)
            df['실제수량'] = qty - qty_ref
            df['실제금액'] = amt - amt_ref

            df['월'] = ym
            df.to_pickle(cache_file)
        except Exception as e:
            st.warning(f'슈퍼적립 캐시 갱신 실패: {Path(src).name} — {e}')


def load_totals(week: str) -> dict | None:
    """주차의 페이지 푸터 합계(JSON)를 캐시에서 로드. 없으면 None."""
    p = CACHE_DIR / f'sales-{week}.totals.json'
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return None


def load_target(week: str) -> float | None:
    """주차의 목표 거래액(억)을 캐시에서 로드. 없으면 None."""
    p = CACHE_DIR / f'sales-{week}.target.json'
    if not p.exists():
        return None
    try:
        v = json.loads(p.read_text(encoding='utf-8')).get('target_eok')
        return float(v) if v is not None else None
    except Exception:
        return None


def save_target(week: str, target_eok: float) -> None:
    """대시보드 사이드바에서 사용자가 목표 변경 시 캐시에 즉시 반영."""
    p = CACHE_DIR / f'sales-{week}.target.json'
    try:
        p.write_text(
            json.dumps({'target_eok': target_eok}, ensure_ascii=False, indent=2),
            encoding='utf-8',
        )
    except Exception as e:
        st.warning(f'목표 저장 실패: {e}')


# 진행 행사 구좌 옵션 (사용자 직접 선택) + 색상 ■ 매핑
EVENT_COLORS = {
    # 강조 (앞쪽 6개) — 사각형 색상
    '오늘의팝업':    '🟪',
    '브랜드데이':    '🟥',
    '네쇼페':       '🟦',
    '신상위크':     '🟫',
    '오늘끝딜':     '🟧',
    '강세일':       '🟩',
    # 일반 (3개) — 색상 중복 없음
    '리빙러빙데이':  '🟨',  # 노랑
    '퍼페페':       '🩷',  # 분홍
    '설+세일':      '🩵',  # 라이트블루
    # 회색 그룹 (중간 회색) — 묶어서 동일 색
    '브랜드위크':    '🩶',
    '슈세윜':       '🩶',
    '도착보장':     '🩶',
    '왕중왕':       '🩶',
    '상반기결산':    '🩶',
    '선물대첩':     '🩶',
    '가정의달':     '🩶',
    '빅브랜드페스타': '🩶',
    '더가구라이브':   '🩶',
}
EVENT_OPTIONS = list(EVENT_COLORS.keys())


# 주차별 특이사항 — 하드코딩된 narrative (캐시 삭제 후에도 유지)
WEEK_NOTES: dict[str, str] = {
    '260223-260301': (
        '**브랜드데이 1Day 매출 5.1억 원** (목표 4.5억 대비 **110% 달성**)\n\n'
        '**퍼패페** 기존 연 2회 행사에서 3회로 확대 운영. '
        "'신학기 전 마지막 찬스'라는 시즌 이슈가 가구 브랜드 일정과 부합하며 시너지를 냄.\n\n"
        '**AI 라이브 시너지** — 금요일 AI 쇼호스트 라이브에 5만 명 이상이 유입(퀴즈 이벤트)되면서, '
        '시디즈·데스커·슬로우베드 등 주요 브랜드의 행사 인지도 제고 및 목표 달성을 견인.\n\n'
        '**쿠폰 전략** — 자정에 발급된 선착순 쿠폰이 오전 11시(시디즈 라이브 전)에 전량 소진되며 '
        '초반 흥행 성공.\n\n'
        '**시디즈 T60** — 직전 행사 대비 큰 폭으로 성장.\n\n'
        '**링고** — 총 2,422개 판매 (링고데이 기간 1,500개 판매 및 종료 후에도 꾸준한 수요 유지).\n\n'
        "**T90** — 메쉬 구매자의 50% 이상이 라이브 당일(2/28)에 집중. '슈퍼적립 마지막 날' 소구가 주효.\n\n"
        '**뮤브** — 목표 수준 달성 (바퀴 모델 35개 / 글라이드 모델 8개).\n\n'
        '**유입 효율** — 네이버 채널 기준, 광고비를 20% 축소했음에도 유입 감소는 10%에 그침. '
        '1월 대비 브랜드 검색 유입은 오히려 크게 증가.\n\n'
        '**톡톡 메시지** — 라이브 10분 전 발송한 긴급 메시지(선착순/시간 안내)가 읽음률 45%를 기록하며 '
        '유입에 결정적 역할 (평균 20% 초반 대비 2배 이상 효율).'
    ),
}


def load_events(week: str) -> list[str]:
    """주차의 선택된 행사 리스트를 캐시에서 로드. 없으면 빈 리스트."""
    p = CACHE_DIR / f'sales-{week}.events.json'
    if not p.exists():
        return []
    try:
        v = json.loads(p.read_text(encoding='utf-8')).get('events', [])
        return [e for e in v if e in EVENT_OPTIONS]
    except Exception:
        return []


def save_events(week: str, events: list[str]) -> None:
    """행사 선택 변경 시 캐시에 즉시 반영."""
    p = CACHE_DIR / f'sales-{week}.events.json'
    try:
        p.write_text(
            json.dumps({'events': events}, ensure_ascii=False, indent=2),
            encoding='utf-8',
        )
    except Exception as e:
        st.warning(f'행사 저장 실패: {e}')


def _cache_signature() -> tuple[tuple[str, int], ...]:
    sigs = []
    for p in sorted(CACHE_DIR.glob('*.pkl')):
        try:
            sigs.append((p.name, p.stat().st_mtime_ns))
        except OSError:
            pass
    # totals.json 변경 시에도 캐시 무효화 (보정 계수가 달라지므로)
    for p in sorted(CACHE_DIR.glob('*.totals.json')):
        try:
            sigs.append((p.name, p.stat().st_mtime_ns))
        except OSError:
            pass
    return tuple(sigs)


@st.cache_data(show_spinner='데이터 로딩 중...')
def load_all(signature: tuple) -> tuple[pd.DataFrame, pd.DataFrame]:
    sales_frames, mkt_frames = [], []
    for p in sorted(CACHE_DIR.glob('sales-*.pkl')):
        try:
            sales_frames.append(pd.read_pickle(p))
        except Exception:
            pass
    for p in sorted(CACHE_DIR.glob('marketing-*.pkl')):
        try:
            mkt_frames.append(pd.read_pickle(p))
        except Exception:
            pass
    sales = pd.concat(sales_frames, ignore_index=True) if sales_frames else pd.DataFrame()
    mkt = pd.concat(mkt_frames, ignore_index=True) if mkt_frames else pd.DataFrame()
    if not sales.empty:
        # 상품ID + 주차 기준 중복 행 제거 (동일 카테고리 내 중복 행 처리)
        if '상품ID' in sales.columns:
            sales = sales.drop_duplicates(subset=['상품ID', '주차'], keep='first')

        # Cross-listing 보정: Excel 결제금액 합계 vs totals.json 실제 결제금액 비율로
        # 실제수량·실제금액을 스케일 다운 → 시리즈별 수량/금액이 실제값에 근접
        if '결제금액' in sales.columns:
            corrected = []
            for week, w_df in sales.groupby('주차'):
                w_df = w_df.copy()
                totals = load_totals(str(week))
                excel_pay = float(w_df['결제금액'].sum())
                if totals and excel_pay > 0:
                    true_pay = float(totals['결제금액'])
                    factor = true_pay / excel_pay
                    if 0.2 < factor < 1.0:  # 합리적 보정 범위만 적용
                        for col in ('실제수량', '실제금액'):
                            if col in w_df.columns:
                                w_df[col] = (w_df[col] * factor).round()
                corrected.append(w_df)
            sales = pd.concat(corrected, ignore_index=True)

        mask = sales['상품명'].astype(str).str.contains(
            EXCLUDE_SERIES_PATTERN, na=False, regex=True)
        sales = sales[~mask].reset_index(drop=True)
    return sales, mkt


def cached_weeks() -> dict[str, list[str]]:
    out: dict[str, list[str]] = {'sales': [], 'marketing': []}
    for p in sorted(CACHE_DIR.glob('*.pkl')):
        m = re.match(r'(sales|marketing)-(\d{6}-\d{6})\.pkl$', p.name)
        if m:
            out[m.group(1)].append(m.group(2))
    return out


# ── KPI 계산 ─────────────────────────────────────────────────────────────────
def kpis_for_week(sales: pd.DataFrame, mkt: pd.DataFrame, week: str | None) -> dict:
    if not week:
        return {k: 0 for k in ['실제매출', '실제수량', '결제수', '객단가', '유입수', '광고비', 'ROAS']}
    s = sales[sales['주차'] == week]
    m = mkt[mkt['주차'] == week]
    revenue = float(s['실제금액'].sum())
    qty = float(s['실제수량'].sum())
    pays = float(s['결제수'].sum())
    visits = float(m['유입수'].sum())
    ad_cost = float(m['광고비'].sum())

    # 페이지 푸터값(JSON)이 있으면 실제매출은 그 값으로 override
    # — 엑셀 행은 카테고리 cross-listing으로 부풀려진 합이라 페이지 푸터가 진실의 원천
    totals = load_totals(week)
    if totals and '결제금액' in totals and '환불금액' in totals:
        revenue = float(totals['결제금액']) - float(totals['환불금액'])

    return {
        '실제매출': revenue,
        '실제수량': qty,
        '결제수': pays,
        '객단가': revenue / pays if pays else 0,
        '유입수': visits,
        '광고비': ad_cost,
        'ROAS': revenue / ad_cost if ad_cost else 0,
    }


def fmt_won(v: float) -> str:
    return f'₩{v:,.0f}'


def fmt_eok(v: float) -> str:
    """원 → '0.0억' 형태."""
    return f'{v / 1e8:.1f}억'


def fmt_int(v: float) -> str:
    return f'{v:,.0f}'


def delta_str(curr: float, prev: float, *, money: bool = False) -> str | None:
    if prev == 0:
        return None
    diff = curr - prev
    pct = diff / prev * 100
    sign = '+' if diff >= 0 else ''
    return f'{sign}{fmt_won(diff) if money else fmt_int(diff)} ({sign}{pct:.1f}%)'


def delta_str_eok(curr: float, prev: float) -> str | None:
    """매출 변화 → '+0.5억 (+5.0%)' 형태 (₩ 없음)."""
    if prev == 0:
        return None
    diff = curr - prev
    pct = diff / prev * 100
    sign = '+' if diff >= 0 else ''
    return f'{sign}{(diff / 1e8):.2f}억 ({sign}{pct:.1f}%)'


def _pct_change(curr: float, prev: float) -> float | None:
    if not prev:
        return None
    return (curr - prev) / prev * 100


def _week_days(week_code: str) -> int:
    """주차 코드 'YYMMDD-YYMMDD'에서 포함 일수(끝-시작+1) 계산.

    스마트스토어는 7/14/21일 등 가변 길이로 다운로드되므로,
    비교 시 같은 일수 기준으로 정규화하기 위해 사용.
    """
    m = re.match(r'^(\d{2})(\d{2})(\d{2})-(\d{2})(\d{2})(\d{2})$', week_code)
    if not m:
        return 7
    y1, m1, d1, y2, m2, d2 = (int(g) for g in m.groups())
    try:
        start = date(2000 + y1, m1, d1)
        end = date(2000 + y2, m2, d2)
        return max(1, (end - start).days + 1)
    except ValueError:
        return 7


def _pct_change_normalized(curr: float, prev: float,
                           curr_days: int, prev_days: int) -> float | None:
    """일수 정규화 후 변화율. 절대 누적값(매출/유입/광고비)에 사용.
    비율값(객단가·ROAS)에는 _pct_change를 그대로 쓰면 됨.
    """
    if not prev or not prev_days or not curr_days:
        return None
    curr_avg = curr / curr_days
    prev_avg = prev / prev_days
    if not prev_avg:
        return None
    return (curr_avg - prev_avg) / prev_avg * 100


def _arrow(diff: float, *, neutral_thresh: float = 0.5) -> str:
    if abs(diff) < neutral_thresh:
        return '→'
    return '▲' if diff > 0 else '▼'


def _week_to_md(w: str) -> str:
    """'260420-260426' → '4/20~4/26' 형태로 변환."""
    try:
        s, e = w.split('-')
        return f'{int(s[2:4])}/{int(s[4:6])}~{int(e[2:4])}/{int(e[4:6])}'
    except Exception:
        return w


def _signed_phrase(p: float, *, unit: str = '%') -> str:
    """변화율 → '12.3% 증가' / '12.3% 감소' / '변동 없음'."""
    if abs(p) < 1:
        return f'변동 없음 ({p:+.1f}{unit})'
    verb = '증가' if p > 0 else '감소'
    return f'{abs(p):.1f}{unit} {verb}'


def _signed(p: float) -> str:
    """+12.3% / -8.5% 형태 (간결)."""
    return f'{p:+.1f}%'


# 색상 정의 — 증가는 파랑, 감소는 빨강
_COLOR_UP = '#1E88E5'
_COLOR_DN = '#E53935'
_COLOR_FLAT = '#888'


def _paren_dir(p: float, *, unit: str = '%') -> str:
    """괄호+화살표+색상 HTML — '(▲ 12.3% 증가)' / '(▼ 8.5% 감소)'."""
    if abs(p) < 1:
        return f'<span style="color:{_COLOR_FLAT};">(변동 없음 {p:+.1f}{unit})</span>'
    if p > 0:
        return f'<span style="color:{_COLOR_UP};">(▲ {p:.1f}{unit} 증가)</span>'
    return f'<span style="color:{_COLOR_DN};">(▼ {abs(p):.1f}{unit} 감소)</span>'


def _paren_pp(curr: float, prev: float, *, decimals: int = 0,
              prev_decimals: int = 0) -> str:
    """이전→이번 %p 변화 괄호 HTML — '(이전 X% → ▲ Y%p)'."""
    diff = curr - prev
    if abs(diff) < (10 ** -decimals) / 2:
        return (f'<span style="color:{_COLOR_FLAT};">'
                f'(이전 {prev:,.{prev_decimals}f}% → 변동 없음)</span>')
    arrow = '▲' if diff > 0 else '▼'
    color = _COLOR_UP if diff > 0 else _COLOR_DN
    return (f'<span style="color:{color};">'
            f'(이전 {prev:,.{prev_decimals}f}% → '
            f'{arrow} {abs(diff):,.{decimals}f}%p)</span>')


def _achievement_line(revenue: float, target_eok: float | None) -> str | None:
    """목표 달성률 라인 — 달성률에 따라 색·아이콘. None이면 라인 안 그림."""
    if not target_eok or target_eok <= 0:
        return None
    target_won = target_eok * 1e8
    pct = revenue / target_won * 100
    rev_eok = revenue / 1e8
    if pct >= 100:
        color, icon = _COLOR_UP, '✓'
    elif pct >= 95:
        color, icon = '#FFA000', '◎'
    else:
        color, icon = _COLOR_DN, '✗'
    return (f'- 목표 대비 <span style="color:{color};font-weight:600">'
            f'{icon} {pct:.1f}% 달성</span> '
            f'({rev_eok:.2f}억 / {target_eok:.1f}억)')


def _render_summary_no_compare(sales_df: pd.DataFrame, mkt_df: pd.DataFrame,
                               base_kpi: dict, base_week: str,
                               target_eok: float | None = None) -> dict:
    """비교 주차 없이 — 기준 주차만 단독 요약."""
    title = ("<div style='font-size:1rem;color:#ccc;font-weight:600;"
             "margin:6px 0 10px 0'>주간 실적 요약</div>")

    revenue = base_kpi['실제매출']
    visits = base_kpi['유입수']
    cost = base_kpi['광고비']
    aov = base_kpi['객단가']
    pays = base_kpi['결제수']
    roas = base_kpi['ROAS'] * 100
    conv = (pays / visits * 100) if visits else 0
    rev_eok = revenue / 1e8

    # 1. 매출 (단독 — 비교 없음)
    parts = ['### 1 . 매출', '']
    parts.append(f'금주 매출 **{rev_eok:.2f}억**.')
    parts.append('')
    parts.append('> 비교 주차 미선택 — 이번 주 절대값만 표시. 비교 주차 선택 시 변화율과 효율 분석이 추가됨.')
    parts.append('')
    achv = _achievement_line(revenue, target_eok)
    if achv:
        parts.append(achv)
    parts.append(f'- 유입수 **{visits:,.0f}명**')
    parts.append(f'- 광고비 **{fmt_won(cost)}**')
    parts.append(f'- ROAS **{roas:,.0f}%**')
    parts.append(f'- 객단가 **{fmt_won(aov)}**')
    parts.append(f'- 결제전환율 **{conv:.2f}%**')
    overview_md = '\n'.join(parts)

    # 2. 유입 (채널그룹별 분포 상위)
    channel_md = None
    base_mkt = mkt_df[mkt_df['주차'] == base_week]
    if not base_mkt.empty:
        parts = ['### 2 . 유입', '']
        ch = (base_mkt.groupby('채널그룹', as_index=False)
              .agg(유입=('유입수', 'sum'),
                   광고=('광고비', 'sum'),
                   결제=('결제금액(마지막클릭)', 'sum')))
        ch_top = ch.sort_values('유입', ascending=False).head(6)
        ch_strs = [f'**{r["채널그룹"]}** {r["유입"]:,.0f}'
                   for _, r in ch_top.iterrows()]
        parts.append('채널그룹별 유입 분포: ' + ', '.join(ch_strs) + '.')

        # ROAS 상위 채널그룹
        ch['ROAS'] = ch.apply(
            lambda r: (r['결제'] / r['광고'] * 100) if r['광고'] else None, axis=1)
        roas_top = (ch.dropna(subset=['ROAS'])
                    .pipe(lambda d: d[d['광고'] >= 100_000])
                    .sort_values('ROAS', ascending=False)
                    .head(2))
        if not roas_top.empty:
            roas_strs = [f'**{r["채널그룹"]}**({r["ROAS"]:,.0f}%)'
                         for _, r in roas_top.iterrows()]
            parts.append('')
            parts.append(f'광고 효율 상위: {", ".join(roas_strs)}.')
        channel_md = '\n'.join(parts)

    # 3. 시리즈별 (단독 — top 5 by 매출)
    series_md = None
    base_s = sales_df[sales_df['주차'] == base_week]
    b_se = (base_s.groupby('상품명', as_index=False)
            .agg(매출=('실제금액', 'sum'),
                 수량=('실제수량', 'sum'),
                 결제=('결제금액', 'sum'),
                 환불=('환불금액', 'sum')))
    s_top = b_se[b_se['매출'] > 0].sort_values('매출', ascending=False)
    top5 = s_top.head(5)
    if len(top5) >= 1:
        parts = ['### 3 . 상품별 이슈', '*매출 상위 시리즈*', '']
        rank_str = ' > '.join(top5['상품명'].astype(str).tolist())
        parts.append(f'매출 순위: **{rank_str}**')
        parts.append('')
        for _, r in top5.iterrows():
            name = r['상품명']
            parts.append(
                f'**{name}** — 매출 **{fmt_eok(r["매출"])}**, 수량 {int(r["수량"]):,}건.'
            )
            parts.append('')
        series_md = '\n'.join(parts).rstrip()

    # 4. 특이사항 — 주차별 하드코딩 노트 우선, 없으면 자동 환불율 알림
    alerts_md = None
    custom_note = WEEK_NOTES.get(base_week)
    if custom_note:
        alerts_md = '### 4 . 특이사항\n\n' + custom_note
    else:
        b_se['환불율'] = b_se.apply(
            lambda r: (r['환불'] / r['결제'] * 100) if r['결제'] else 0, axis=1)
        alerts = b_se[
            (b_se['결제'] >= 1_000_000) & (b_se['환불율'] >= 30)
        ].sort_values('환불율', ascending=False).head(5)
        if len(alerts) > 0:
            parts = ['### 4 . 특이사항', '',
                     '환불율 30% 이상 시리즈 — 원인 확인 필요:', '']
            for _, r in alerts.iterrows():
                parts.append(
                    f'- **{r["상품명"]}** — 환불율 **{r["환불율"]:.1f}%** '
                    f'(매출 {fmt_won(r["매출"])})'
                )
            alerts_md = '\n'.join(parts)

    return {
        'title': title, 'overview': overview_md,
        'channel': channel_md, 'series': series_md, 'alerts': alerts_md,
    }


def render_integrated_summary(sales_df: pd.DataFrame, mkt_df: pd.DataFrame,
                              base_kpi: dict, prev_kpi: dict,
                              base_week: str, compare_week: str | None,
                              target_eok: float | None = None) -> dict:
    """모든 탭(요약·시리즈·마케팅) 내용을 종합한 주간 실적 요약 — narrative 형식.

    반환: {'title': str, 'overview': str, 'channel': str | None,
           'series': str | None, 'alerts': str | None}
    """
    if not compare_week:
        return _render_summary_no_compare(
            sales_df, mkt_df, base_kpi, base_week, target_eok=target_eok)

    prev_md = _week_to_md(compare_week)

    # 일수 — 7/14/21일 가변. 절대 누적값은 일평균으로 정규화해 변화율 계산.
    base_days = _week_days(base_week)
    prev_days = _week_days(compare_week)
    days_mismatch = (base_days != prev_days)

    # ── 요약 탭 KPI ──
    revenue = base_kpi['실제매출']
    visits = base_kpi['유입수']
    cost = base_kpi['광고비']
    aov = base_kpi['객단가']
    pays = base_kpi['결제수']
    roas = base_kpi['ROAS'] * 100
    conv = (pays / visits * 100) if visits else 0

    # 절대 누적값(매출·유입·광고비) → 일평균 정규화 변화율
    rev_p = _pct_change_normalized(revenue, prev_kpi['실제매출'], base_days, prev_days)
    vis_p = _pct_change_normalized(visits, prev_kpi['유입수'], base_days, prev_days)
    cost_p = _pct_change_normalized(cost, prev_kpi['광고비'], base_days, prev_days)
    # 비율값(객단가·ROAS·전환율)은 일수 무관 — 단순 비교
    aov_p = _pct_change(aov, prev_kpi['객단가'])
    roas_prev = prev_kpi['ROAS'] * 100
    conv_prev = (prev_kpi['결제수'] / prev_kpi['유입수'] * 100) if prev_kpi['유입수'] else 0

    prev_revenue = prev_kpi['실제매출']
    rev_eok = revenue / 1e8
    prev_rev_eok = prev_revenue / 1e8

    title = ("<div style='font-size:1rem;color:#ccc;font-weight:600;"
             "margin:6px 0 10px 0'>주간 실적 요약</div>")

    # ════════════════════════════════════════════════════════════════
    # 1. 총괄
    # ════════════════════════════════════════════════════════════════
    parts: list[str] = []
    parts.append('### 1 . 매출')
    parts.append('')

    # 매출 본문
    if rev_p is not None:
        verb = '감소' if rev_p < 0 else '증가'
        basis = ' (일평균 기준)' if days_mismatch else ''
        parts.append(
            f'금주 매출은 **{rev_eok:.2f}억**({base_days}일)으로 이전 {prev_md}'
            f'({prev_rev_eok:.2f}억, {prev_days}일) 대비 '
            f'**{abs(rev_p):.1f}% {verb}**{basis}.'
        )
    else:
        parts.append(f'금주 매출 **{rev_eok:.2f}억** ({base_days}일).')
    parts.append('')

    # 인사이트 (blockquote)
    if vis_p is not None and rev_p is not None:
        eff_diff = rev_p - vis_p
        if abs(eff_diff) < 3:
            insight = (
                f'> 유입은 {_signed_phrase(vis_p)}, 매출도 비슷한 폭으로 움직여 '
                f'효율 차이는 {eff_diff:+.1f}%p로 미미. 유입 변화 그대로 따라간 형태.'
            )
        elif eff_diff > 0:
            if vis_p < 0:
                insight = (
                    f'> 같은 기간 유입이 **{abs(vis_p):.1f}% 빠진 점**을 감안하면, '
                    f'매출 감소 폭은 유입 대비 **+{eff_diff:.1f}%p 적게 빠진 편**. '
                    f'유입 감소 환경 속에서 매출 효율은 오히려 개선된 상태.'
                )
            else:
                insight = (
                    f'> 유입은 {_signed_phrase(vis_p)} 했지만 매출은 그보다 '
                    f'**+{eff_diff:.1f}%p 더 잘 나온 편** — 유입 증가 효과를 제대로 흡수.'
                )
        else:
            if vis_p < 0:
                insight = (
                    f'> 유입이 **{abs(vis_p):.1f}% 빠진 와중에** 매출은 그보다 '
                    f'**{eff_diff:.1f}%p 더 빠진 상태**. 유입 감소 폭을 그대로 흡수하지 '
                    f'못해 유입 대비 효율 악화. 원인 확인 필요.'
                )
            else:
                insight = (
                    f'> 유입은 {_signed_phrase(vis_p)} 했음에도 매출은 그만큼 따라가지 '
                    f'못한 상태(효율 {eff_diff:+.1f}%p). 유입 증가가 매출로 충분히 '
                    f'전환되지 못함.'
                )
        parts.append(insight)
        parts.append('')

    # 목표 달성률
    achv = _achievement_line(revenue, target_eok)
    if achv:
        parts.append(achv)

    # 숫자 메트릭 — 줄바꿈으로 한 줄씩, 괄호 안 변화 색·화살표
    if vis_p is not None:
        parts.append(f'- 유입수 **{visits:,.0f}명**  {_paren_dir(vis_p)}')
    cost_line = f'- 광고비 **{fmt_won(cost)}**'
    if cost_p is not None:
        cost_line += f'  {_paren_dir(cost_p)}'
    parts.append(cost_line)
    roas_line = f'- ROAS **{roas:,.0f}%**'
    if abs(roas - roas_prev) >= 0.5:
        roas_line += f'  {_paren_pp(roas, roas_prev, decimals=0, prev_decimals=0)}'
    parts.append(roas_line)
    aov_line = f'- 객단가 **{fmt_won(aov)}**'
    if aov_p is not None:
        aov_line += f'  {_paren_dir(aov_p)}'
    parts.append(aov_line)
    parts.append(
        f'- 결제전환율 **{conv:.2f}%**  '
        f'{_paren_pp(conv, conv_prev, decimals=2, prev_decimals=2)}'
    )
    overview_md = '\n'.join(parts)

    # ════════════════════════════════════════════════════════════════
    # 2. 마케팅 채널
    # ════════════════════════════════════════════════════════════════
    channel_md: str | None = None
    base_mkt = mkt_df[mkt_df['주차'] == base_week]
    prev_mkt = mkt_df[mkt_df['주차'] == compare_week]
    if not base_mkt.empty and not prev_mkt.empty:
        parts = []
        parts.append('### 2 . 유입')
        parts.append('')

        # 전체 유입·광고비·ROAS 비교 (1. 매출 톤과 통일, 일평균 정규화)
        b_visits = float(base_mkt['유입수'].sum())
        p_visits = float(prev_mkt['유입수'].sum())
        b_cost = float(base_mkt['광고비'].sum())
        p_cost = float(prev_mkt['광고비'].sum())
        b_settle = float(base_mkt['결제금액(마지막클릭)'].sum())
        p_settle = float(prev_mkt['결제금액(마지막클릭)'].sum())
        b_roas_m = (b_settle / b_cost * 100) if b_cost else 0
        p_roas_m = (p_settle / p_cost * 100) if p_cost else 0
        v_pct = _pct_change_normalized(b_visits, p_visits, base_days, prev_days)
        c_pct = _pct_change_normalized(b_cost, p_cost, base_days, prev_days)

        if v_pct is not None:
            verb_v = '감소' if v_pct < 0 else '증가'
            parts.append(
                f'금주 유입은 **{b_visits:,.0f}명**으로 이전 {prev_md}'
                f'({p_visits:,.0f}명) 대비 **{abs(v_pct):.1f}% {verb_v}**.'
            )
            parts.append('')
        if c_pct is not None:
            verb_c = '감소' if c_pct < 0 else '증가'
            roas_change = ''
            if abs(b_roas_m - p_roas_m) >= 0.5:
                roas_change = f' (이전 {p_roas_m:,.0f}%, {_paren_pp(b_roas_m, p_roas_m)})'
            parts.append(
                f'광고비는 **{fmt_won(b_cost)}**로 이전 {fmt_won(p_cost)} 대비 '
                f'**{abs(c_pct):.1f}% {verb_c}**, 마케팅 ROAS는 '
                f'**{b_roas_m:,.0f}%**{roas_change}.'
            )
            parts.append('')

        b_ch = base_mkt.groupby('채널그룹', as_index=False).agg(
            이번=('유입수', 'sum'),
            이번광고=('광고비', 'sum'),
            이번결제=('결제금액(마지막클릭)', 'sum'))
        p_ch = prev_mkt.groupby('채널그룹', as_index=False).agg(이전=('유입수', 'sum'))
        ch_m = b_ch.merge(p_ch, on='채널그룹', how='outer').fillna(0)
        ch_m['max유입'] = ch_m[['이번', '이전']].max(axis=1)
        ch_m = ch_m[ch_m['max유입'] >= 100]
        ch_m['변화'] = ch_m.apply(
            lambda r: _pct_change_normalized(
                r['이번'], r['이전'], base_days, prev_days), axis=1)

        # 변화 narrative
        ch_changed = ch_m.dropna(subset=['변화']).sort_values('max유입', ascending=False).head(6)
        if not ch_changed.empty:
            ch_strs = [f'**{r["채널그룹"]}** {r["변화"]:+.0f}%' for _, r in ch_changed.iterrows()]
            parts.append(
                f'채널그룹별 유입 변화는 {", ".join(ch_strs)} 순으로 나타남.'
            )
            parts.append('')

        # ROAS / 빠진 채널 narrative
        ch_m['ROAS'] = ch_m.apply(
            lambda r: (r['이번결제'] / r['이번광고'] * 100) if r['이번광고'] else None, axis=1)
        roas_rank = ch_m.dropna(subset=['ROAS'])
        roas_rank = (roas_rank[roas_rank['이번광고'] >= 100_000]
                     .sort_values('ROAS', ascending=False))
        biggest_drop = ch_m.dropna(subset=['변화']).sort_values('변화').head(1)

        narrative_bits = []
        if not roas_rank.empty:
            top_roas = roas_rank.head(2)
            roas_strs = [f'**{r["채널그룹"]}**({r["ROAS"]:,.0f}%)' for _, r in top_roas.iterrows()]
            narrative_bits.append(
                f'광고 효율은 {", ".join(roas_strs)}이 ROAS 상위로, '
                f'예산 집행 대비 매출 기여가 가장 좋은 채널그룹.'
            )
        if not biggest_drop.empty:
            r = biggest_drop.iloc[0]
            if r['변화'] <= -10:
                narrative_bits.append(
                    f'반면 유입이 가장 크게 빠진 채널그룹은 **{r["채널그룹"]}**'
                    f'({r["변화"]:+.0f}%)으로, 노출 위치·소재 점검 필요.'
                )
        if narrative_bits:
            parts.append(' '.join(narrative_bits))
        channel_md = '\n'.join(parts)

    # ════════════════════════════════════════════════════════════════
    # 3. 시리즈 (유입 변화 대비 효율)
    # ════════════════════════════════════════════════════════════════
    series_md: str | None = None
    base_s = sales_df[sales_df['주차'] == base_week]
    prev_s = sales_df[sales_df['주차'] == compare_week]
    b_se = (base_s.groupby('상품명', as_index=False)
            .agg(이번매출=('실제금액', 'sum'),
                 이번수량=('실제수량', 'sum'),
                 이번결제=('결제금액', 'sum'),
                 이번환불=('환불금액', 'sum')))
    p_se = (prev_s.groupby('상품명', as_index=False)
            .agg(이전매출=('실제금액', 'sum'),
                 이전수량=('실제수량', 'sum'),
                 이전결제=('결제금액', 'sum'),
                 이전환불=('환불금액', 'sum')))
    s_m = b_se.merge(p_se, on='상품명', how='outer').fillna(0)
    s_top = s_m[s_m['이번매출'] > 0].sort_values('이번매출', ascending=False)

    top5 = s_top.head(5)
    if len(top5) >= 1:
        parts = []
        parts.append('### 3 . 상품별 이슈')
        parts.append('*유입 변화 대비 효율*')
        parts.append('')
        rank_str = ' > '.join(top5['상품명'].astype(str).tolist())
        parts.append(f'매출 순위는 **{rank_str}** 순.')
        parts.append('')
        if vis_p is not None:
            parts.append(
                f'> 이전 {prev_md} 대비 전체 유입이 {_signed_phrase(vis_p)} 한 상황이라, '
                f'시리즈별 매출·수량 절대 변화율은 유입 변동에 휩쓸려 의미가 약함. '
                f'**유입 변화(baseline {_signed(vis_p)}) 대비 시리즈별 효율(±%p)**을 보면 다음과 같음.'
            )
            parts.append('')

        for _, r in top5.iterrows():
            name = r['상품명']
            rev_pp = _pct_change_normalized(
                r['이번매출'], r['이전매출'], base_days, prev_days)

            if rev_pp is None:
                parts.append(f'**{name}** — 신규 시리즈로 이전 비교 데이터 없음.')
                parts.append('')
                continue

            sent1 = f'**{name}** — 매출 **{fmt_eok(r["이번매출"])}**.'

            sent2 = ''
            if vis_p is not None:
                rev_diff = rev_pp - vis_p
                if rev_diff > 15:
                    verdict = '유입 대비 효율 매우 우수'
                elif rev_diff > 5:
                    verdict = '유입 대비 효율 양호'
                elif rev_diff < -15:
                    verdict = '유입 대비 효율 매우 저조 — 원인 확인 필요'
                elif rev_diff < -5:
                    verdict = '유입 변화 폭보다 다소 더 빠진 상태'
                else:
                    verdict = '유입 변화 폭과 비슷한 수준'
                sent2 = f'유입({_signed(vis_p)}) 대비 **{rev_diff:+.1f}%p** → {verdict}.'

            if sent2:
                parts.append(sent1 + ' ' + sent2)
            else:
                parts.append(sent1)
            parts.append('')

        series_md = '\n'.join(parts).rstrip()

    # ════════════════════════════════════════════════════════════════
    # 4. 주의 — 환불율 급증
    # ════════════════════════════════════════════════════════════════
    s_m['이번환불율'] = s_m.apply(
        lambda r: (r['이번환불'] / r['이번결제'] * 100) if r['이번결제'] else 0, axis=1)
    s_m['이전환불율'] = s_m.apply(
        lambda r: (r['이전환불'] / r['이전결제'] * 100) if r['이전결제'] else 0, axis=1)
    s_m['환불율차이'] = s_m['이번환불율'] - s_m['이전환불율']
    alerts = s_m[
        (s_m['이번결제'] >= 1_000_000)
        & (s_m['이번환불율'] >= 30)
        & (s_m['환불율차이'] >= 10)
    ].sort_values('이번환불율', ascending=False).head(5)

    alerts_md: str | None = None
    custom_note = WEEK_NOTES.get(base_week)
    if custom_note:
        alerts_md = '### 4 . 특이사항\n\n' + custom_note
    else:
        notice_blocks: list[str] = []

        # 4-1. 채널그룹 유입 큰 변화 (±30% 이상, 유입 규모 500명 이상)
        if not base_mkt.empty and not prev_mkt.empty:
            b_ch2 = base_mkt.groupby('채널그룹', as_index=False).agg(이번=('유입수', 'sum'))
            p_ch2 = prev_mkt.groupby('채널그룹', as_index=False).agg(이전=('유입수', 'sum'))
            ch_n = b_ch2.merge(p_ch2, on='채널그룹', how='outer').fillna(0)
            ch_n['최대유입'] = ch_n[['이번', '이전']].max(axis=1)
            ch_n = ch_n[ch_n['최대유입'] >= 500].copy()
            ch_n['변화'] = ch_n.apply(
                lambda r: _pct_change(r['이번'], r['이전']), axis=1)
            ch_n = ch_n.dropna(subset=['변화'])

            big_up = ch_n[ch_n['변화'] >= 30].sort_values('변화', ascending=False).head(3)
            big_down = ch_n[ch_n['변화'] <= -30].sort_values('변화').head(3)
            if not big_up.empty or not big_down.empty:
                ch_lines = ['**채널그룹 유입 큰 변화**', '']
                for _, r in big_up.iterrows():
                    ch_lines.append(
                        f'- **{r["채널그룹"]}** — 유입 **{r["변화"]:+.0f}%** '
                        f'(이전 {r["이전"]:,.0f}명 → 이번 {r["이번"]:,.0f}명)'
                    )
                for _, r in big_down.iterrows():
                    ch_lines.append(
                        f'- **{r["채널그룹"]}** — 유입 **{r["변화"]:+.0f}%** '
                        f'(이전 {r["이전"]:,.0f}명 → 이번 {r["이번"]:,.0f}명)'
                    )
                notice_blocks.append('\n'.join(ch_lines))

        # 4-2. 상품별 특이사항 — TOP5 외 신규 진입 또는 유입 대비 크게 성장한 시리즈
        top5_names = top5['상품명'].astype(str).tolist() if not top5.empty else []
        out_top = s_m[~s_m['상품명'].astype(str).isin(top5_names)].copy()
        # 신규 시리즈 — 이전 매출 0 + 이번 매출 1천만원 이상
        new_series = out_top[
            (out_top['이전매출'] == 0) & (out_top['이번매출'] >= 10_000_000)
        ].sort_values('이번매출', ascending=False).head(3)
        # 큰 성장 — 매출 1천만원 이상 + 유입 대비 +30%p 이상
        big_growers = pd.DataFrame()
        if vis_p is not None:
            out_with_prev = out_top[out_top['이전매출'] > 0].copy()
            out_with_prev['rev_pp'] = out_with_prev.apply(
                lambda r: _pct_change_normalized(
                    r['이번매출'], r['이전매출'], base_days, prev_days), axis=1)
            out_with_prev = out_with_prev.dropna(subset=['rev_pp']).copy()
            out_with_prev['rev_diff'] = out_with_prev['rev_pp'] - vis_p
            big_growers = out_with_prev[
                (out_with_prev['이번매출'] >= 10_000_000)
                & (out_with_prev['rev_diff'] >= 30)
            ].sort_values('rev_diff', ascending=False).head(3)

        if not new_series.empty or not big_growers.empty:
            s_lines = ['**상품별 특이사항**', '']
            for _, r in new_series.iterrows():
                s_lines.append(
                    f'- **{r["상품명"]}** — 신규 진입, 매출 **{fmt_eok(r["이번매출"])}** 발생.'
                )
            for _, r in big_growers.iterrows():
                s_lines.append(
                    f'- **{r["상품명"]}** — 매출 **{fmt_eok(r["이번매출"])}**, '
                    f'유입({_signed(vis_p)}) 대비 **+{r["rev_diff"]:.1f}%p** 크게 성장.'
                )
            notice_blocks.append('\n'.join(s_lines))

        # 4-3. 환불율 급증 (기존)
        if len(alerts) > 0:
            r_lines = ['**환불율 급증**', '']
            r_lines.append(
                '환불율이 이전 대비 +10%p 이상 상승, 절대 수치 30% 초과 — 원인 확인 시급:'
            )
            for _, r in alerts.iterrows():
                r_lines.append(
                    f'- **{r["상품명"]}** — 환불율 **{r["이번환불율"]:.1f}%** '
                    f'(이전 {r["이전환불율"]:.1f}% → **+{r["환불율차이"]:.1f}%p**)'
                )
            notice_blocks.append('\n'.join(r_lines))

        if notice_blocks:
            alerts_md = '### 4 . 특이사항\n\n' + '\n\n'.join(notice_blocks)

    return {
        'title': title,
        'overview': overview_md,
        'channel': channel_md,
        'series': series_md,
        'alerts': alerts_md,
    }


# ── 팀장님 보고용 메시지 ─────────────────────────────────────────────────────
SLACK_CONFIG_FILE = BASE / '.slack_config.json'


def load_slack_config() -> dict:
    """팀장님 Slack 설정 로드. {dm_url, team_id, manager_user_id, manager_name}."""
    try:
        if SLACK_CONFIG_FILE.exists():
            return json.loads(SLACK_CONFIG_FILE.read_text(encoding='utf-8'))
    except Exception:
        pass
    return {}


def send_to_slack_rpa(message: str, slack_url: str,
                      auto_send: bool = False,
                      wait_seconds: float = 3.0) -> tuple[bool, str]:
    """RPA — 클립보드 복사 → 슬랙 앱 열기 → Ctrl+V → (옵션) Enter.

    Streamlit 서버가 도는 PC의 키보드를 조작합니다(=대시보드를 실행한 본인 PC).
    슬랙 데스크탑 앱이 설치돼 있어야 하고, 자동화 중에는 마우스·키보드를
    건드리지 말아야 합니다(pyautogui 보안: 마우스 모서리 = 중단).
    """
    try:
        import time
        import webbrowser
        import pyperclip
        import pyautogui
    except ImportError as e:
        return False, f'필수 패키지 누락: {e}. `pip install pyperclip pyautogui` 실행 필요.'

    if not message.strip():
        return False, '빈 메시지는 보낼 수 없습니다.'
    if not slack_url:
        return False, '슬랙 URL이 설정되지 않았습니다.'

    try:
        pyautogui.FAILSAFE = True   # 마우스 좌상단으로 = 즉시 중단(안전장치)

        pyperclip.copy(message)
        webbrowser.open(slack_url)
        time.sleep(wait_seconds)     # 슬랙 앱 로딩/포커스 대기

        pyautogui.hotkey('ctrl', 'v')
        time.sleep(0.4)

        if auto_send:
            pyautogui.press('enter')

        return True, '슬랙 전송 시퀀스 완료.' if auto_send else '붙여넣기 완료. 슬랙에서 Enter로 전송하세요.'
    except pyautogui.FailSafeException:
        return False, '자동화 중단됨 — 마우스가 화면 모서리로 이동했습니다.'
    except Exception as e:
        return False, f'RPA 실행 실패: {e}'


def build_report_message(
    sales_df: pd.DataFrame, mkt_df: pd.DataFrame,
    weeks: list[str], event_name: str,
    target_eok: float | None = None,
    forced_compare_weeks: list[str] | None = None,
) -> str:
    """팀장님 보고용 줄글 메시지. 선택한 주차들을 합산해 단일 보고서 생성.

    fact(수치)만 자동 채우고, 운영 컨텍스트는 `[ ... 추가 ]` 자리표시자로 표기.
    슬랙에 붙여넣기 전에 사용자가 컨텍스트를 직접 채워 넣는 흐름.
    """
    if not weeks:
        return '데이터 없음 — 주차를 선택해 주세요.'

    starts, ends = [], []
    for w in sorted(weeks):
        m = re.match(r'(\d{6})-(\d{6})', w)
        if m:
            starts.append(m.group(1))
            ends.append(m.group(2))
    period = f'{min(starts)[2:]}~{max(ends)[2:]}' if starts else '기간 미확정'

    base_s = sales_df[sales_df['주차'].isin(weeks)]
    base_m = mkt_df[mkt_df['주차'].isin(weeks)]

    # 주차별 totals JSON 합산 — 없는 주차는 엑셀 합으로 대체
    revenue_total = 0.0
    for w in weeks:
        t = load_totals(w)
        if t and '결제금액' in t and '환불금액' in t:
            revenue_total += float(t['결제금액']) - float(t['환불금액'])
        else:
            revenue_total += float(sales_df[sales_df['주차'] == w]['실제금액'].sum())

    pays = float(base_s['결제수'].sum()) if not base_s.empty else 0
    visits = float(base_m['유입수'].sum()) if not base_m.empty else 0
    cost = float(base_m['광고비'].sum()) if not base_m.empty else 0
    settle = (float(base_m['결제금액(마지막클릭)'].sum())
              if not base_m.empty and '결제금액(마지막클릭)' in base_m.columns else 0)

    aov = (revenue_total / pays) if pays else 0
    roas = (settle / cost * 100) if cost else 0
    rev_eok = revenue_total / 1e8

    # 보고 주차 총 일수 (7/14/21일 등 가변)
    report_days = sum(_week_days(w) for w in weeks)

    # 비교 주차: 사용자가 명시한 경우 그대로 사용, 아니면 직전 주차 자동 산출
    if forced_compare_weeks is not None:
        compare_weeks: list[str] = forced_compare_weeks
        compare_days = sum(_week_days(w) for w in compare_weeks)
    else:
        all_weeks_sorted = sorted(sales_df['주차'].unique().tolist())
        sorted_report = sorted(weeks)
        compare_weeks = []
        compare_days = 0
        try:
            earliest_idx = all_weeks_sorted.index(sorted_report[0])
        except ValueError:
            earliest_idx = 0
        i = earliest_idx - 1
        while i >= 0 and compare_days < report_days:
            cand = all_weeks_sorted[i]
            compare_weeks.insert(0, cand)
            compare_days += _week_days(cand)
            i -= 1

    # 비교 주차 합산 KPI
    prev_revenue = 0.0
    prev_visits = prev_cost = prev_settle = prev_pays = 0.0
    prev_period = ''
    if compare_weeks:
        prev_s_df = sales_df[sales_df['주차'].isin(compare_weeks)]
        prev_m_df = mkt_df[mkt_df['주차'].isin(compare_weeks)]
        for w in compare_weeks:
            t = load_totals(w)
            if t and '결제금액' in t and '환불금액' in t:
                prev_revenue += float(t['결제금액']) - float(t['환불금액'])
            else:
                prev_revenue += float(sales_df[sales_df['주차'] == w]['실제금액'].sum())
        prev_visits = float(prev_m_df['유입수'].sum()) if not prev_m_df.empty else 0
        prev_cost = float(prev_m_df['광고비'].sum()) if not prev_m_df.empty else 0
        prev_settle = (float(prev_m_df['결제금액(마지막클릭)'].sum())
                       if not prev_m_df.empty and '결제금액(마지막클릭)' in prev_m_df.columns else 0)
        prev_pays = float(prev_s_df['결제수'].sum()) if not prev_s_df.empty else 0
        cs, ce = [], []
        for w in sorted(compare_weeks):
            m = re.match(r'(\d{6})-(\d{6})', w)
            if m:
                cs.append(m.group(1))
                ce.append(m.group(2))
        prev_period = f'{min(cs)[2:]}~{max(ce)[2:]}' if cs else ''

    prev_roas = (prev_settle / prev_cost * 100) if prev_cost else 0
    prev_aov_val = (prev_revenue / prev_pays) if prev_pays else 0

    # 변화율 — 절대 누적값은 일평균 정규화 (일수 다를 때 공정 비교)
    rev_pct = _pct_change_normalized(
        revenue_total, prev_revenue, report_days, compare_days)
    vis_pct = _pct_change_normalized(
        visits, prev_visits, report_days, compare_days)
    days_mismatch = (report_days != compare_days) and compare_days > 0

    # 유입 대비 매출 효율 인사이트 — 보고 메시지의 핵심 평가
    insight_line = ''
    if rev_pct is not None and vis_pct is not None:
        eff_diff = rev_pct - vis_pct
        if abs(eff_diff) < 3:
            insight_line = (
                f'→ 같은 기간 유입은 {vis_pct:+.1f}% 변화, 매출도 비슷한 폭'
                f'({rev_pct:+.1f}%)으로 움직여 유입 변화를 그대로 따라간 형태'
                f' (효율 차이 {eff_diff:+.1f}%p).'
            )
        elif eff_diff > 0:
            if vis_pct < 0:
                insight_line = (
                    f'→ 같은 기간 유입이 {abs(vis_pct):.1f}% 빠진 점을 감안하면, '
                    f'매출 감소 폭은 유입 대비 +{eff_diff:.1f}%p 적게 빠진 편. '
                    f'유입 감소 환경 속에서 매출 효율은 오히려 개선된 상태.'
                )
            else:
                insight_line = (
                    f'→ 같은 기간 유입은 {vis_pct:+.1f}% 변화, 매출은 그보다 '
                    f'+{eff_diff:.1f}%p 더 잘 나옴 — 유입 증가 효과를 제대로 흡수.'
                )
        else:
            if vis_pct < 0:
                insight_line = (
                    f'→ 같은 기간 유입이 {abs(vis_pct):.1f}% 빠진 와중에 매출은 '
                    f'그보다 {eff_diff:.1f}%p 더 빠진 상태. 유입 감소 폭을 그대로 '
                    f'흡수하지 못해 효율 악화 — 원인 확인 필요.'
                )
            else:
                insight_line = (
                    f'→ 같은 기간 유입은 {vis_pct:+.1f}% 증가했음에도 매출은 그만큼 '
                    f'따라가지 못한 상태 (효율 {eff_diff:+.1f}%p). 유입 증가가 매출로 '
                    f'충분히 전환되지 못함.'
                )

    def _delta_pct(cur: float, prev: float, normalize: bool = False) -> str:
        """변화율 (%). normalize=True면 일평균 환산 후 비교."""
        if not prev:
            return ''
        if normalize and compare_days and report_days:
            cur_n = cur / report_days
            prev_n = prev / compare_days
            if not prev_n:
                return ''
            diff = (cur_n - prev_n) / prev_n * 100
        else:
            diff = (cur - prev) / prev * 100
        sign = '+' if diff >= 0 else ''
        return f' ({sign}{diff:.1f}%)'

    def _delta_pp(cur: float, prev: float) -> str:
        if not prev:
            return ''
        diff = cur - prev
        sign = '+' if diff >= 0 else ''
        return f' ({sign}{diff:.1f}%p)'

    out: list[str] = []
    ev = event_name.strip() if event_name else '행사명'
    out.append(f'[{period} 네이버 <{ev}> 실적 요약]')
    out.append('')

    # 매출 — 절대값 + 목표 + 직전 대비
    # 기간이 다를 때: 총액 증감(직관적)을 주 지표로, 일평균 증감을 참고 지표로 병기
    rev_change_text = ''
    if prev_revenue and prev_period:
        raw_pct = (revenue_total - prev_revenue) / prev_revenue * 100
        raw_verb = '감소' if raw_pct < 0 else '증가'
        if days_mismatch:
            day_verb = '감소' if (rev_pct or 0) < 0 else '증가'
            rev_change_text = (
                f', 직전 {prev_period}({prev_revenue / 1e8:.1f}억, {compare_days}일) 대비 '
                f'총액 {abs(raw_pct):.1f}% {raw_verb}'
                f' / 일평균 {abs(rev_pct):.1f}% {day_verb}'
            )
        else:
            rev_change_text = (
                f', 직전 {prev_period}({prev_revenue / 1e8:.1f}억) 대비 '
                f'{abs(raw_pct):.1f}% {raw_verb}'
            )
    period_days_text = f' ({report_days}일)'
    if target_eok and target_eok > 0:
        pct = revenue_total / (target_eok * 1e8) * 100
        out.append(
            f'매출: {rev_eok:.1f}억 원 기록{period_days_text} (목표 {target_eok}억 대비 '
            f'{pct:.1f}% 달성){rev_change_text}.'
        )
    else:
        out.append(f'매출: {rev_eok:.1f}억 원 기록{period_days_text}{rev_change_text}.')
    if insight_line:
        out.append(insight_line)
    out.append('')

    # 효율 — 절대 누적값은 일평균 정규화로 변화율 계산 (객단가·ROAS는 ratio라 그대로)
    eff_parts = [
        f'기간 누적 유입 {visits:,.0f}명{_delta_pct(visits, prev_visits, normalize=True)}',
        f'광고비 {fmt_won(cost)}{_delta_pct(cost, prev_cost, normalize=True)}',
    ]
    if roas:
        eff_parts.append(f'마케팅 ROAS {roas:,.0f}%{_delta_pp(roas, prev_roas)}')
    if aov:
        eff_parts.append(f'객단가 {fmt_won(aov)}{_delta_pct(aov, prev_aov_val)}')

    if compare_weeks and prev_period:
        basis = ', 일평균 기준' if days_mismatch else ''
        out.append(
            f'효율 (직전 {prev_period} 대비{basis}): ' + ', '.join(eff_parts) + '.'
        )
    else:
        out.append('효율: ' + ', '.join(eff_parts) + '.')
    out.append('')

    top5 = (base_s.groupby('상품명', as_index=False)
            .agg(매출=('실제금액', 'sum'))
            .sort_values('매출', ascending=False).head(5))
    if not top5.empty:
        rank = ' > '.join(top5['상품명'].astype(str).tolist())
        out.append(f'품목별 매출 TOP 5: {rank}')
        out.append('')

    # ── 특이사항 (가장 중요) — 통합 분석 4. 특이사항을 요약해 가져옴 ──
    alert_lines: list[str] = []
    if compare_weeks:
        prev_s_df = sales_df[sales_df['주차'].isin(compare_weeks)]
        prev_m_df = mkt_df[mkt_df['주차'].isin(compare_weeks)]

        # 1) 채널그룹 유입 큰 변화 (±30% 이상, 유입 500명+) — 일평균 정규화
        if not base_m.empty and not prev_m_df.empty:
            b_ch = base_m.groupby('채널그룹', as_index=False).agg(이번=('유입수', 'sum'))
            p_ch = prev_m_df.groupby('채널그룹', as_index=False).agg(이전=('유입수', 'sum'))
            ch_n = b_ch.merge(p_ch, on='채널그룹', how='outer').fillna(0)
            ch_n['최대유입'] = ch_n[['이번', '이전']].max(axis=1)
            ch_n = ch_n[ch_n['최대유입'] >= 500].copy()
            ch_n['변화'] = ch_n.apply(
                lambda r: _pct_change_normalized(
                    r['이번'], r['이전'], report_days, compare_days), axis=1)
            ch_n = ch_n.dropna(subset=['변화'])
            big_chg = ch_n[(ch_n['변화'] >= 30) | (ch_n['변화'] <= -30)].copy()
            big_chg['abs변화'] = big_chg['변화'].abs()
            big_chg = big_chg.sort_values('abs변화', ascending=False).head(4)
            if not big_chg.empty:
                alert_lines.append('채널 유입 큰 변화:')
                for _, r in big_chg.iterrows():
                    sign = '+' if r['변화'] >= 0 else ''
                    alert_lines.append(
                        f'  - {r["채널그룹"]} — 유입 {sign}{r["변화"]:.0f}%'
                        f' (이전 {int(r["이전"]):,}명 → 이번 {int(r["이번"]):,}명)'
                    )

        # 2) 상품별 특이사항 (TOP5 외) — 유입 대비 크게 성장 + 신규 진입
        cur_series = base_s.groupby('상품명', as_index=False).agg(
            이번매출=('실제금액', 'sum'))
        prev_series = prev_s_df.groupby('상품명', as_index=False).agg(
            이전매출=('실제금액', 'sum'))
        s_m_local = cur_series.merge(prev_series, on='상품명', how='outer').fillna(0)
        top5_names = top5['상품명'].astype(str).tolist() if not top5.empty else []
        out_top = s_m_local[~s_m_local['상품명'].astype(str).isin(top5_names)].copy()

        if vis_pct is not None:
            out_with_prev = out_top[
                (out_top['이전매출'] > 0) & (out_top['이번매출'] >= 10_000_000)].copy()
            # 시리즈 매출도 일평균 정규화 — 일수 다른 기간 비교의 공정성 확보
            out_with_prev['rev_pp'] = out_with_prev.apply(
                lambda r: _pct_change_normalized(
                    r['이번매출'], r['이전매출'], report_days, compare_days), axis=1)
            out_with_prev = out_with_prev.dropna(subset=['rev_pp']).copy()
            out_with_prev['rev_diff'] = out_with_prev['rev_pp'] - vis_pct
            growers = out_with_prev[out_with_prev['rev_diff'] >= 30].sort_values(
                'rev_diff', ascending=False).head(3)
            if not growers.empty:
                bits = [f'{r["상품명"]} (+{r["rev_diff"]:.0f}%p)'
                        for _, r in growers.iterrows()]
                alert_lines.append(f'- 유입 대비 큰 성장 시리즈 — {", ".join(bits)}')

        new_series = out_top[
            (out_top['이전매출'] == 0) & (out_top['이번매출'] >= 10_000_000)
        ].sort_values('이번매출', ascending=False).head(3)
        if not new_series.empty:
            bits = [f'{r["상품명"]} ({fmt_eok(r["이번매출"])})'
                    for _, r in new_series.iterrows()]
            alert_lines.append(f'- 신규 진입 — {", ".join(bits)}')

        # 3) 환불율 급증 (이번 결제 100만+, 환불율 30%+, 차이 +10%p+)
        cur_ref = base_s.groupby('상품명', as_index=False).agg(
            이번결제=('결제금액', 'sum'),
            이번환불=('환불금액', 'sum'),
        )
        prev_ref = prev_s_df.groupby('상품명', as_index=False).agg(
            이전결제=('결제금액', 'sum'),
            이전환불=('환불금액', 'sum'),
        )
        s_ref = cur_ref.merge(prev_ref, on='상품명', how='outer').fillna(0)
        s_ref['이번환불율'] = s_ref.apply(
            lambda r: (r['이번환불'] / r['이번결제'] * 100) if r['이번결제'] else 0, axis=1)
        s_ref['이전환불율'] = s_ref.apply(
            lambda r: (r['이전환불'] / r['이전결제'] * 100) if r['이전결제'] else 0, axis=1)
        s_ref['환불율차이'] = s_ref['이번환불율'] - s_ref['이전환불율']
        refund_alerts = s_ref[
            (s_ref['이번결제'] >= 1_000_000)
            & (s_ref['이번환불율'] >= 30)
            & (s_ref['환불율차이'] >= 10)
        ].sort_values('이번환불율', ascending=False).head(3)
        if not refund_alerts.empty:
            bits = [f'{r["상품명"]} ({r["이번환불율"]:.0f}%, +{r["환불율차이"]:.0f}%p)'
                    for _, r in refund_alerts.iterrows()]
            alert_lines.append(f'- 환불율 급증 — {", ".join(bits)}')

    if alert_lines:
        out.append('특이사항:')
        out.extend(alert_lines)

    return '\n'.join(out).rstrip()


# ── 페이지 설정 ──────────────────────────────────────────────────────────────
st.set_page_config(page_title='네이버 프로모션 대시보드', page_icon='📊', layout='wide')

# 상단 여백 축소 + 사이드바 위로 + 탭 4~6 우정렬
st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem !important; }
    /* 사이드바 — 모든 상단 여백 최소화 */
    section[data-testid="stSidebar"] > div,
    section[data-testid="stSidebar"] > div > div,
    [data-testid="stSidebarContent"],
    [data-testid="stSidebarUserContent"] {
        padding-top: 0 !important;
        margin-top: 0 !important;
    }
    [data-testid="stSidebarHeader"] {
        padding: 0 !important;
        min-height: 0 !important;
    }
    section[data-testid="stSidebar"] h2 {
        margin-top: 0.5rem !important;
        padding-top: 0 !important;
    }
    /* 사이드바 구분선 여백 */
    section[data-testid="stSidebar"] hr {
        margin-top: 1rem !important;
        margin-bottom: 1rem !important;
    }
    /* 도움말(?) 아이콘 — 배경 없이 테두리(stroke)만 어둡게 */
    [data-testid="stTooltipIcon"],
    [data-testid="stTooltipIcon"] svg {
        color: #555 !important;
    }
    /* 4번째 탭부터 우측으로 밀기 */
    [data-baseweb="tab-list"] [data-baseweb="tab"]:nth-child(4) {
        margin-left: auto !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
# 컴팩트한 페이지 제목
st.markdown(
    "<h2 style='font-size:1.4rem;font-weight:700;margin:0 0 0.6rem 0;'>"
    "📊 네이버 프로모션 대시보드</h2>",
    unsafe_allow_html=True,
)

# 10초마다 자동 새로고침 → downloads/ 새 파일 즉시 캐시 적재
st_autorefresh(interval=10_000, key='auto_refresh_data')

sync_cache()
sync_super_cache()
sales_df, mkt_df = load_all(_cache_signature())

if sales_df.empty and mkt_df.empty:
    st.error('downloads/ 폴더에 데이터가 없습니다.')
    st.stop()

# ── 사이드바 ─────────────────────────────────────────────────────────────────
sales_weeks = sorted(sales_df['주차'].unique().tolist()) if not sales_df.empty else []
mkt_weeks = sorted(mkt_df['주차'].unique().tolist()) if not mkt_df.empty else []
all_weeks = sorted(set(sales_weeks) | set(mkt_weeks))

with st.sidebar:
    st.header('필터')
    base_week = st.selectbox('기준 주차', all_weeks, index=len(all_weeks) - 1)

    compare_options = ['(비교 안함)'] + [w for w in all_weeks if w != base_week]
    default_compare_idx = 1 if len(compare_options) > 1 else 0
    compare_choice = st.selectbox('비교 주차', compare_options, index=default_compare_idx)
    compare_week = None if compare_choice == '(비교 안함)' else compare_choice

    # 매출목표 (억) — 다운로드 시 입력받은 값 / 사이드바에서 수정 가능
    saved_target = load_target(base_week) or 0.0
    target_input = st.number_input(
        '매출목표 (억)',
        min_value=0.0,
        value=float(saved_target),
        step=0.1,
        format='%.1f',
        help='0이면 목표 미설정으로 처리',
    )
    target_eok_active = target_input if target_input > 0 else None
    # 사용자가 변경했으면 즉시 캐시에 반영
    if target_eok_active is not None and abs(target_eok_active - saved_target) > 1e-9:
        save_target(base_week, target_eok_active)

    st.divider()
    # 진행 행사 — 색상 ■ + 둥근 chip(pills), 다중 선택, 항상 노출
    saved_events = load_events(base_week)
    selected_events = st.pills(
        '진행 행사',
        EVENT_OPTIONS,
        selection_mode='multi',
        default=saved_events,
        format_func=lambda x: f'{EVENT_COLORS.get(x, "⚪")} {x}',
        help='이 주차에 진행한 행사 구좌 (다중 선택). 분석에는 아직 반영되지 않음.',
        key=f'sidebar_events_{base_week}',
    )
    if set(selected_events) != set(saved_events):
        save_events(base_week, selected_events)

    st.divider()
    series_options = sorted(sales_df['상품명'].dropna().unique().tolist()) if not sales_df.empty else []
    with st.expander('시리즈 필터', expanded=False):
        series_filter = [o for o in series_options
                         if st.checkbox(o, value=True, key=f'series_chk__{o}')]

    group_options = sorted(mkt_df['채널그룹'].dropna().unique().tolist()) if not mkt_df.empty else []
    with st.expander('채널그룹 필터', expanded=False):
        group_filter = [o for o in group_options
                        if st.checkbox(o, value=True, key=f'group_chk__{o}')]

    st.divider()
    # 캐시관리 popover 버튼 — 작은 회색 정사각형 + V 아이콘을 ▼ 로 교체
    st.markdown("""
        <style>
        [data-testid="stPopover"] button,
        [data-testid="stPopoverButton"] {
            background-color: #2a2a2a !important;
            color: #999 !important;
            border: 1px solid #3a3a3a !important;
            position: relative !important;
            padding: 6px 10px !important;
            line-height: 1.3 !important;
            white-space: nowrap !important;
            text-align: center !important;
        }
        [data-testid="stPopover"] button:hover,
        [data-testid="stPopoverButton"]:hover {
            background-color: #353535 !important;
            color: #bbb !important;
            border-color: #555 !important;
        }
        /* 원래 V 아이콘 (svg) 숨기기 */
        [data-testid="stPopover"] svg,
        [data-testid="stPopoverButton"] svg {
            display: none !important;
        }
        /* 채워진 ▼ 삼각형으로 교체 */
        [data-testid="stPopover"] button::after,
        [data-testid="stPopoverButton"]::after {
            content: '▼' !important;
            display: inline-block !important;
            color: #bbb !important;
            font-size: 0.7em !important;
            margin-left: 6px !important;
        }
        /* popover 펼쳤을 때 컨텐츠 폭을 사이드바 정도로 좁히기 + 내용 줄바꿈 허용 */
        [data-baseweb="popover"] > div,
        [data-testid="stPopoverBody"],
        [role="dialog"][data-baseweb="popover"] {
            max-width: 280px !important;
            width: 280px !important;
            box-sizing: border-box !important;
        }
        [data-baseweb="popover"] p,
        [data-baseweb="popover"] span,
        [data-baseweb="popover"] div,
        [data-baseweb="popover"] label,
        [data-testid="stPopoverBody"] p,
        [data-testid="stPopoverBody"] span,
        [data-testid="stPopoverBody"] div {
            word-wrap: break-word !important;
            overflow-wrap: break-word !important;
            white-space: normal !important;
            max-width: 100% !important;
        }
        </style>
    """, unsafe_allow_html=True)
    _, cm_col = st.columns([3, 2])
    with cm_col:
        with st.popover('캐시관리', use_container_width=True):
            st.caption('downloads/에서 파일을 지워도 대시보드에 남습니다. '
                       '대시보드에서 빼려면 여기서 명시적으로 삭제하세요.')
            cw = cached_weeks()
            all_cached = sorted(set(cw['sales']) | set(cw['marketing']))
            if not all_cached:
                st.info('주차 캐시 비어있음')
            else:
                to_remove = st.multiselect('캐시에서 삭제할 주차', all_cached, default=[])
                if st.button('선택 주차 삭제', disabled=not to_remove):
                    removed = []
                    for w in to_remove:
                        for kind in ('sales', 'marketing'):
                            f = CACHE_DIR / f'{kind}-{w}.pkl'
                            if f.exists():
                                f.unlink()
                                removed.append(f.name)
                        for suffix in ('totals.json', 'target.json', 'events.json'):
                            sf = CACHE_DIR / f'sales-{w}.{suffix}'
                            if sf.exists():
                                sf.unlink()
                                removed.append(sf.name)
                    st.cache_data.clear()
                    st.success(f'{len(removed)}개 캐시 파일 삭제됨')
                    st.rerun()

# 필터 적용
if series_filter:
    sales_df = sales_df[sales_df['상품명'].isin(series_filter)]
if group_filter:
    mkt_df = mkt_df[mkt_df['채널그룹'].isin(group_filter)]

base_kpi = kpis_for_week(sales_df, mkt_df, base_week)
prev_kpi = kpis_for_week(sales_df, mkt_df, compare_week)

tab1, tab2, tab3, tab4 = st.tabs([
    '통합 분석', '시리즈 매출 성과', '마케팅 채널 효율', '슬랙 보고📩',
])

# ── 탭 1: 통합 분석 ──────────────────────────────────────────────────────────
with tab1:
    weekly = []
    for w in all_weeks:
        k = kpis_for_week(sales_df, mkt_df, w)
        weekly.append({'주차': w, '실제매출': k['실제매출'], '유입수': k['유입수'],
                       '광고비': k['광고비'], 'ROAS': k['ROAS']})
    weekly_df = pd.DataFrame(weekly)

    summary = render_integrated_summary(
        sales_df, mkt_df, base_kpi, prev_kpi, base_week, compare_week,
        target_eok=target_eok_active)
    if summary.get('title'):
        st.markdown(summary['title'], unsafe_allow_html=True)

    kpi_col, rank_col, chart_col = st.columns([1, 1, 2], gap='medium')

    def _metric_html(label, value, delta_text, *, big: bool = False,
                     achievement_pct: float | None = None):
        if not delta_text:
            color, arrow = '#888', ''
        elif delta_text.startswith('-'):
            color, arrow = '#E53935', '▼ '
        else:
            color, arrow = '#1E88E5', '▲ '
        if big:
            label_size, value_size, delta_size = '0.85rem', '1.6rem', '0.8rem'
            margin = '14px'
        else:
            label_size, value_size, delta_size = '0.72rem', '1.05rem', '0.7rem'
            margin = '10px'

        # 값 옆에 붙는 달성률 배지 (있을 때)
        ach_html = ''
        if achievement_pct is not None:
            if achievement_pct >= 100:
                ach_color, ach_bg, ach_icon = '#1E88E5', 'rgba(30,136,229,0.15)', '✓'
            else:
                ach_color, ach_bg, ach_icon = '#E53935', 'rgba(229,57,53,0.15)', '✗'
            ach_html = (
                f"<span style='display:inline-block;margin-left:10px;"
                f"padding:2px 9px;border-radius:10px;"
                f"background:{ach_bg};color:{ach_color};"
                f"font-size:0.72rem;font-weight:600;vertical-align:middle'>"
                f"{ach_icon} {achievement_pct:.1f}% 달성"
                f"</span>"
            )

        return (
            f"<div style='line-height:1.25;margin-top:{margin}'>"
            f"<div style='font-size:{label_size};color:#888'>{label}</div>"
            f"<div style='font-size:{value_size};font-weight:600'>"
            f"{value}{ach_html}</div>"
            f"<div style='font-size:{delta_size};color:{color}'>"
            f"{arrow}{delta_text or '—'}</div>"
            f"</div>"
        )

    with kpi_col:
        # 매출 목표 달성률
        revenue_achievement = None
        if target_eok_active and target_eok_active > 0:
            revenue_achievement = base_kpi['실제매출'] / (target_eok_active * 1e8) * 100

        # 기간 일수 계산 — 누적값(매출·유입·광고비) delta 정규화
        _bd = _week_days(base_week)
        _pd = _week_days(compare_week) if compare_week else _bd
        _days_diff = compare_week is not None and _bd != _pd

        def _delta_cumul_eok(curr: float, prev: float) -> str | None:
            if not prev:
                return None
            if _days_diff:
                p = _pct_change_normalized(curr, prev, _bd, _pd)
                if p is None:
                    return None
                sign = '+' if p >= 0 else ''
                return f'{sign}{p:.1f}% (일평균)'
            return delta_str_eok(curr, prev)

        def _delta_cumul(curr: float, prev: float) -> str | None:
            if not prev:
                return None
            if _days_diff:
                p = _pct_change_normalized(curr, prev, _bd, _pd)
                if p is None:
                    return None
                sign = '+' if p >= 0 else ''
                return f'{sign}{p:.1f}% (일평균)'
            return delta_str(curr, prev)

        # 기간 일수 불일치 안내
        if _days_diff:
            st.markdown(
                f"<div style='font-size:0.7rem;color:#FFA000;margin-bottom:4px'>"
                f"⚠ 기준 {_bd}일 / 비교 {_pd}일 — 매출·유입·광고비 변화율은 일평균 기준</div>",
                unsafe_allow_html=True,
            )

        # 모든 메트릭 세로 나열 — 회색 라벨, 큰 2개 + 구분선 + 작은 3개
        kpi_html = ''.join([
            _metric_html('총 실제 매출', fmt_eok(base_kpi['실제매출']),
                         _delta_cumul_eok(base_kpi['실제매출'], prev_kpi['실제매출']),
                         big=True, achievement_pct=revenue_achievement),
            _metric_html('총 유입수', fmt_int(base_kpi['유입수']),
                         _delta_cumul(base_kpi['유입수'], prev_kpi['유입수']),
                         big=True),
            "<hr style='margin:18px 0 4px 0;border:none;"
            "border-top:1px solid #333' />",
            _metric_html('평균 객단가', fmt_int(base_kpi['객단가']),
                         delta_str(base_kpi['객단가'], prev_kpi['객단가'])),
            _metric_html('총 광고비', fmt_int(base_kpi['광고비']),
                         _delta_cumul(base_kpi['광고비'], prev_kpi['광고비'])),
            _metric_html('통합 ROAS', f'{base_kpi["ROAS"]*100:,.1f}%',
                         delta_str(base_kpi['ROAS'] * 100, prev_kpi['ROAS'] * 100)),
        ])
        st.markdown(kpi_html, unsafe_allow_html=True)

    with rank_col:
        # 매출 순위 TOP 10 — 세로 나열, 매출 + 수량 비중%
        rank_base = sales_df[sales_df['주차'] == base_week]
        total_qty = float(rank_base['실제수량'].sum() or 0)
        rank_df = (rank_base.groupby('상품명', as_index=False)
                   .agg(매출=('실제금액', 'sum'),
                        수량=('실제수량', 'sum'))
                   .sort_values('매출', ascending=False)
                   .head(10))
        st.markdown(
            "<div style='font-size:1rem;color:#ccc;font-weight:600;"
            "margin:6px 0 10px 0'>매출 순위 TOP 10</div>",
            unsafe_allow_html=True,
        )
        if not rank_df.empty:
            rows = []
            # 헤더 행 — 컬럼 라벨
            rows.append(
                "<div style='display:flex;padding:4px 0;font-size:0.7rem;"
                "color:#888;border-bottom:1px solid #444'>"
                "<span style='flex:2'></span>"
                "<span style='flex:1;text-align:right'>매출</span>"
                "<span style='flex:1;text-align:right'>수량</span>"
                "<span style='flex:1;text-align:right'>비중</span>"
                "</div>"
            )
            sum_qty_pct = 0.0
            sum_qty = 0
            # 상위 3위 색 강조 (금/은/동), 4~10위 회색
            rank_colors = {1: '#FFC857', 2: '#C7CDD3', 3: '#D89F76'}
            for i, (_, r) in enumerate(rank_df.iterrows(), start=1):
                qty = int(r['수량'])
                qty_pct = (r['수량'] / total_qty * 100) if total_qty else 0
                sum_qty_pct += qty_pct
                sum_qty += qty
                rank_color = rank_colors.get(i, '#888')
                rows.append(
                    f"<div style='display:flex;padding:5px 0;font-size:0.82rem;"
                    f"border-bottom:1px solid #2a2a2a;align-items:center'>"
                    f"<span style='flex:2;display:flex;align-items:center'>"
                    f"<b style='color:{rank_color};min-width:32px'>{i}위</b>"
                    f"<span style='color:#444;margin:0 8px'>│</span>"
                    f"<span>{r['상품명']}</span>"
                    f"</span>"
                    f"<span style='flex:1;text-align:right;color:#bbb'>"
                    f"{fmt_eok(r['매출'])}</span>"
                    f"<span style='flex:1;text-align:right;color:#bbb'>"
                    f"{qty:,}ea</span>"
                    f"<span style='flex:1;text-align:right;color:#666;font-size:0.75rem'>"
                    f"{qty_pct:.1f}%</span>"
                    f"</div>"
                )
            # 합계 행
            rows.append(
                f"<div style='display:flex;justify-content:flex-end;gap:16px;"
                f"padding:8px 0 4px 0;font-size:0.85rem;font-weight:600;"
                f"border-top:2px solid #555;margin-top:6px'>"
                f"<span style='color:#888'>TOP 10 합계</span>"
                f"<span style='color:#bbb'>{sum_qty:,}ea</span>"
                f"<span style='color:#1E88E5'>{sum_qty_pct:.1f}%</span>"
                f"</div>"
            )
            st.markdown(''.join(rows), unsafe_allow_html=True)
        else:
            st.caption('데이터 없음')

    with chart_col:
        st.markdown(
            "<div style='font-size:1rem;color:#ccc;font-weight:600;"
            "margin:6px 0 10px 0'>주차별 매출 · 유입수 추이</div>",
            unsafe_allow_html=True,
        )
        if len(weekly_df) >= 1:
            fig = go.Figure()
            sales_eok = weekly_df['실제매출'] / 1e8
            fig.add_bar(x=weekly_df['주차'], y=sales_eok, name='실제매출(억)',
                        yaxis='y1', marker_color=PASTEL_BLUE,
                        customdata=weekly_df['실제매출'],
                        hovertemplate='%{x}<br>실제매출: %{customdata:,.0f}원'
                                      ' (%{y:.2f}억)<extra></extra>')
            fig.add_scatter(x=weekly_df['주차'], y=weekly_df['유입수'], name='유입수',
                            yaxis='y2', mode='lines+markers',
                            line=dict(color=PASTEL_PEACH, width=3),
                            marker=dict(size=9),
                            hovertemplate='%{x}<br>유입수: %{y:,.0f}<extra></extra>')
            fig.update_layout(
                yaxis=dict(title='실제매출(억)', tickformat=',.1f'),
                yaxis2=dict(title='유입수', overlaying='y', side='right', tickformat=','),
                xaxis=dict(title=None),
                height=420, margin=dict(t=40, b=40, l=50, r=50),
                legend=dict(orientation='h', y=1.08, x=0.5, xanchor='center'),
            )
            st.plotly_chart(fig, use_container_width=True)

    st.divider()
    st.markdown('## 프로모션 분석')

    s1, s2, s3, s4 = st.columns(4, gap='small')
    with s1:
        with st.container(border=True):
            st.markdown(summary.get('overview') or '_데이터 없음_',
                        unsafe_allow_html=True)
    with s2:
        with st.container(border=True):
            st.markdown(summary.get('channel') or '### 2 . 유입\n\n_마케팅 데이터 없음_',
                        unsafe_allow_html=True)
    with s3:
        with st.container(border=True):
            st.markdown(summary.get('series') or '### 3 . 상품별 이슈\n\n*유입 변화 대비 효율*\n\n_데이터 없음_',
                        unsafe_allow_html=True)
    with s4:
        with st.container(border=True):
            st.markdown(summary.get('alerts') or '### 4 . 특이사항\n\n_특이사항 없음_',
                        unsafe_allow_html=True)
    st.divider()

    matrix_rows = []
    for w in all_weeks:
        k = kpis_for_week(sales_df, mkt_df, w)
        conv = (k['결제수'] / k['유입수'] * 100) if k['유입수'] else 0
        days = _week_days(w)
        matrix_rows.append({
            '주차': w,
            '일수': days,
            '실제매출': k['실제매출'],
            '일평균매출': k['실제매출'] / days if days else 0,
            '광고비': k['광고비'],
            'ROAS(%)': k['ROAS'] * 100,
            '유입수': k['유입수'],
            '결제전환율(%)': conv,
            '객단가': k['객단가'],
        })
    matrix_df = pd.DataFrame(matrix_rows)
    st.markdown('##### 주차별 매트릭스')
    mtx_col, _ = st.columns([2, 1])
    with mtx_col:
        st.dataframe(
            matrix_df,
            use_container_width=True,
            column_config={
                '일수': st.column_config.NumberColumn(format='%d일'),
                '실제매출': st.column_config.NumberColumn(format='localized'),
                '일평균매출': st.column_config.NumberColumn(format='localized'),
                '광고비': st.column_config.NumberColumn(format='localized'),
                'ROAS(%)': st.column_config.NumberColumn(format='%.0f%%'),
                '유입수': st.column_config.NumberColumn(format='localized'),
                '결제전환율(%)': st.column_config.NumberColumn(format='%.2f%%'),
                '객단가': st.column_config.NumberColumn(format='localized'),
            },
            hide_index=True,
        )

# ── 탭 2: 시리즈 매출 성과 ───────────────────────────────────────────────────
with tab2:
    head_col, opt_amt_col, opt_qty_col = st.columns([4, 1, 1])
    head_col.subheader('시리즈별 매출 성과')
    show_amt_label = opt_amt_col.checkbox('매출 막대 숫자', value=True, key='show_amt_label')
    show_qty_label = opt_qty_col.checkbox('수량 막대 숫자', value=True, key='show_qty_label')

    s_base = sales_df[sales_df['주차'] == base_week]
    base_agg = (s_base.groupby('상품명', as_index=False)
                .agg(이번주매출=('실제금액', 'sum'),
                     이번주수량=('실제수량', 'sum'))
                .sort_values('이번주매출', ascending=False))

    if base_agg.empty:
        st.info('선택한 기준 주차/필터에 데이터가 없습니다.')
    elif not compare_week:
        # 비교 주차 없이 — 기준 주차만 표시
        st.caption(f'비교 주차 미선택 — {base_week} 단독 표시 ({_week_days(base_week)}일)')
        merged = base_agg.copy().sort_values('이번주매출', ascending=False)

        # 매출 단일 막대
        st.markdown(f'##### 시리즈별 실제매출 — {base_week}')
        fig_amt = px.bar(
            merged, x='상품명', y='이번주매출',
            text_auto=(',d' if show_amt_label else False),
            color_discrete_sequence=[PASTEL_BLUE],
            labels={'이번주매출': '실제매출(₩)'},
        )
        fig_amt.update_layout(
            height=420, xaxis_title=None, showlegend=False,
            yaxis=dict(title='실제매출(₩)', tickformat=','),
            margin=dict(t=40, b=80, l=60, r=20),
        )
        if show_amt_label:
            fig_amt.update_traces(
                textfont_size=9,
                textposition='outside',
                textangle=-30,
                cliponaxis=False,
            )
        st.plotly_chart(fig_amt, use_container_width=True)

        # 수량 단일 막대
        st.markdown(f'##### 시리즈별 실제수량 — {base_week}')
        fig_qty = px.bar(
            merged, x='상품명', y='이번주수량',
            text_auto=(',d' if show_qty_label else False),
            color_discrete_sequence=[PASTEL_GREEN],
            labels={'이번주수량': '실제수량'},
        )
        fig_qty.update_layout(
            height=380, xaxis_title=None, showlegend=False,
            yaxis=dict(title='실제수량', tickformat=','),
            margin=dict(t=40, b=80, l=60, r=20),
        )
        if show_qty_label:
            fig_qty.update_traces(
                textfont_size=9,
                textposition='outside',
                textangle=-30,
                cliponaxis=False,
            )
        st.plotly_chart(fig_qty, use_container_width=True)

        # 표 — 단일 주차 (변동 없이)
        st.markdown(f'##### {base_week} 시리즈별 매출·수량')
        st.dataframe(
            merged[['상품명', '이번주매출', '이번주수량']],
            use_container_width=True,
            column_config={
                '이번주매출': st.column_config.NumberColumn(format='localized'),
                '이번주수량': st.column_config.NumberColumn(format='localized'),
            },
            hide_index=True,
        )
    else:
        s_prev = sales_df[sales_df['주차'] == compare_week]
        prev_agg = (s_prev.groupby('상품명', as_index=False)
                    .agg(이전주매출=('실제금액', 'sum'),
                         이전주수량=('실제수량', 'sum')))
        merged = base_agg.merge(prev_agg, on='상품명', how='outer').fillna(0)
        merged = merged.sort_values('이번주매출', ascending=False)

        # 기간 일수 불일치 안내 — 막대값은 기간 합계이므로 직접 비교 시 주의
        _t2_bd = _week_days(base_week)
        _t2_pd = _week_days(compare_week)
        if _t2_bd != _t2_pd:
            st.caption(
                f'⚠ 기준 {base_week} ({_t2_bd}일) vs 비교 {compare_week} ({_t2_pd}일) — '
                f'막대는 기간 합계입니다. 일평균 기준 비교는 통합분석 탭의 요약을 참고하세요.'
            )

        # 비교 막대차트 (매출)
        long_amt = pd.melt(
            merged, id_vars='상품명',
            value_vars=['이전주매출', '이번주매출'],
            var_name='주차구분', value_name='실제매출')
        long_amt['주차'] = long_amt['주차구분'].map({
            '이전주매출': f'{compare_week} (이전)',
            '이번주매출': f'{base_week} (이번)',
        })
        fig_amt = px.bar(
            long_amt, x='상품명', y='실제매출', color='주차', barmode='group',
            text_auto=(',d' if show_amt_label else False),
            color_discrete_sequence=[PASTEL_PREV, PASTEL_BLUE],
        )
        fig_amt.update_layout(
            height=460, xaxis_title=None,
            yaxis=dict(title='실제매출(₩)', tickformat=','),
            margin=dict(t=40, b=80, l=60, r=20),
            legend=dict(orientation='h', y=1.10, x=0.5, xanchor='center'),
        )
        if show_amt_label:
            fig_amt.update_traces(
                textfont_size=9,
                textposition='outside',
                textangle=-30,
                cliponaxis=False,
            )
        st.plotly_chart(fig_amt, use_container_width=True)

        # 비교 막대차트 (수량)
        long_qty = pd.melt(
            merged, id_vars='상품명',
            value_vars=['이전주수량', '이번주수량'],
            var_name='주차구분', value_name='실제수량')
        long_qty['주차'] = long_qty['주차구분'].map({
            '이전주수량': f'{compare_week} (이전)',
            '이번주수량': f'{base_week} (이번)',
        })
        st.markdown(f'##### 시리즈별 실제수량 비교 — {compare_week} vs {base_week}')
        fig_qty = px.bar(
            long_qty, x='상품명', y='실제수량', color='주차', barmode='group',
            text_auto=(',d' if show_qty_label else False),
            color_discrete_sequence=[PASTEL_PREV, PASTEL_GREEN],
        )
        fig_qty.update_layout(
            height=420, xaxis_title=None,
            yaxis=dict(title='실제수량', tickformat=','),
            margin=dict(t=40, b=80, l=60, r=20),
            legend=dict(orientation='h', y=1.10, x=0.5, xanchor='center'),
        )
        if show_qty_label:
            fig_qty.update_traces(
                textfont_size=9,
                textposition='outside',
                textangle=-30,
                cliponaxis=False,
            )
        st.plotly_chart(fig_qty, use_container_width=True)

        # 표 — 매출·수량 변동
        st.markdown(f'##### WoW 변동 ({compare_week} → {base_week})')
        merged['Δ금액'] = merged['이번주매출'] - merged['이전주매출']
        merged['Δ금액%'] = merged.apply(
            lambda r: (r['Δ금액'] / r['이전주매출'] * 100) if r['이전주매출'] else None, axis=1)
        merged['Δ수량'] = merged['이번주수량'] - merged['이전주수량']
        merged['Δ수량%'] = merged.apply(
            lambda r: (r['Δ수량'] / r['이전주수량'] * 100) if r['이전주수량'] else None, axis=1)
        table = merged[['상품명', '이번주매출', '이전주매출', 'Δ금액', 'Δ금액%',
                        '이번주수량', '이전주수량', 'Δ수량', 'Δ수량%']]
        st.dataframe(
            table,
            use_container_width=True,
            column_config={
                '이번주매출': st.column_config.NumberColumn(format='localized'),
                '이전주매출': st.column_config.NumberColumn(format='localized'),
                'Δ금액': st.column_config.NumberColumn(format='localized'),
                'Δ금액%': st.column_config.NumberColumn(format='%.1f%%'),
                '이번주수량': st.column_config.NumberColumn(format='localized'),
                '이전주수량': st.column_config.NumberColumn(format='localized'),
                'Δ수량': st.column_config.NumberColumn(format='localized'),
                'Δ수량%': st.column_config.NumberColumn(format='%.1f%%'),
            },
            hide_index=True,
        )

# ── 탭 3: 마케팅 채널 효율 ──────────────────────────────────────────────────
with tab3:
    st.subheader('마케팅 채널 효율')

    m_base = mkt_df[mkt_df['주차'] == base_week]
    if m_base.empty:
        st.info('선택한 기준 주차에 마케팅 데이터가 없습니다.')
    else:
        col1, col2 = st.columns([1, 1])
        with col1:
            st.markdown('##### 채널별 유입수 (상위 15)')
            ch_traffic = (m_base.groupby(['채널명', '채널그룹'], as_index=False)
                          .agg(유입수=('유입수', 'sum')))
            ch_traffic = (ch_traffic[ch_traffic['유입수'] > 0]
                          .sort_values('유입수', ascending=False)
                          .head(15))
            fig_traffic = px.bar(
                ch_traffic, x='채널명', y='유입수', color='채널그룹',
                text_auto=',.0f',
                color_discrete_sequence=PASTEL_QUAL,
            )
            fig_traffic.update_layout(
                height=420, xaxis_title=None,
                yaxis=dict(tickformat=','),
                xaxis=dict(tickangle=-30),
                margin=dict(t=40, b=110, l=50, r=20),
                legend=dict(orientation='h', y=1.10, x=0.5, xanchor='center', title=None),
            )
            st.plotly_chart(fig_traffic, use_container_width=True)

        with col2:
            st.markdown('##### 주차별 광고비 vs 매출 추이')
            if len(weekly_df) >= 1:
                fig_trend = go.Figure()
                fig_trend.add_bar(x=weekly_df['주차'], y=weekly_df['광고비'],
                                  name='광고비', marker_color=PASTEL_CORAL,
                                  hovertemplate='%{x}<br>광고비: %{y:,.0f}원<extra></extra>')
                fig_trend.add_scatter(x=weekly_df['주차'], y=weekly_df['실제매출'],
                                      name='실제매출', yaxis='y2',
                                      mode='lines+markers',
                                      line=dict(color=PASTEL_BLUE, width=3),
                                      marker=dict(size=9),
                                      hovertemplate='%{x}<br>실제매출: %{y:,.0f}원<extra></extra>')
                fig_trend.update_layout(
                    yaxis=dict(title='광고비(₩)', tickformat=','),
                    yaxis2=dict(title='실제매출(₩)', overlaying='y', side='right',
                                tickformat=','),
                    height=420, margin=dict(t=40, b=50, l=60, r=60),
                    legend=dict(orientation='h', y=1.10, x=0.5, xanchor='center'),
                )
                st.plotly_chart(fig_trend, use_container_width=True)

        # 채널별 유입수 WoW 비교
        if compare_week:
            m_prev = mkt_df[mkt_df['주차'] == compare_week]
            base_ch = (m_base.groupby('채널명', as_index=False)
                       .agg(이번주유입=('유입수', 'sum')))
            prev_ch = (m_prev.groupby('채널명', as_index=False)
                       .agg(이전주유입=('유입수', 'sum')))
            ch_cmp = base_ch.merge(prev_ch, on='채널명', how='outer').fillna(0)
            ch_cmp['최대유입'] = ch_cmp[['이번주유입', '이전주유입']].max(axis=1)
            ch_cmp['변화율'] = ch_cmp.apply(
                lambda r: ((r['이번주유입'] - r['이전주유입']) / r['이전주유입'] * 100)
                          if r['이전주유입'] else None,
                axis=1)
            ch_cmp = (ch_cmp[ch_cmp['최대유입'] > 0]
                      .sort_values('최대유입', ascending=False)
                      .head(15))

            long_cmp = pd.melt(
                ch_cmp, id_vars='채널명',
                value_vars=['이전주유입', '이번주유입'],
                var_name='주차구분', value_name='유입수')
            long_cmp['주차'] = long_cmp['주차구분'].map({
                '이전주유입': f'{compare_week} (이전)',
                '이번주유입': f'{base_week} (이번)',
            })
            st.markdown(
                f'##### 채널별 유입수 비교 — {compare_week} vs {base_week} (상위 15)')

            # 평균 / 전체 변화율 캡션 (차트 위쪽으로)
            avg_pct = ch_cmp['변화율'].mean(skipna=True)
            base_total = ch_cmp['이번주유입'].sum()
            prev_total = ch_cmp['이전주유입'].sum()
            total_pct = ((base_total - prev_total) / prev_total * 100) if prev_total else None
            cap_bits = []
            if pd.notna(avg_pct):
                cap_bits.append(f'표시 채널 평균 변화율 **{avg_pct:+.1f}%**')
            if total_pct is not None:
                cap_bits.append(f'전체 합계 변화율 **{total_pct:+.1f}%**')
            if cap_bits:
                st.caption('  ·  '.join(cap_bits))

            fig_cmp = px.bar(
                long_cmp, x='채널명', y='유입수', color='주차', barmode='group',
                text_auto=',d',
                color_discrete_sequence=[PASTEL_PREV, PASTEL_GREEN],
            )
            # y축 상한을 최대 유입의 1.18배로 늘려 변화율 라벨 공간 확보
            y_max = float(ch_cmp['최대유입'].max() or 0) * 1.18
            fig_cmp.update_layout(
                height=480, xaxis_title=None,
                yaxis=dict(title='유입수', tickformat=',', range=[0, y_max] if y_max else None),
                xaxis=dict(tickangle=-30),
                margin=dict(t=50, b=110, l=50, r=20),
                legend=dict(orientation='h', y=1.10, x=0.5, xanchor='center', title=None),
            )

            # 채널별 변화율 annotation (각 채널 막대 위에 +X% / -X% 표시)
            for _, r in ch_cmp.iterrows():
                if pd.isna(r['변화율']):
                    text, color = '신규', '#888'
                else:
                    sign = '+' if r['변화율'] >= 0 else ''
                    text = f'{sign}{r["변화율"]:.0f}%'
                    color = '#2ECC71' if r['변화율'] >= 0 else '#E74C3C'
                fig_cmp.add_annotation(
                    x=r['채널명'], y=r['최대유입'],
                    text=f'<b>{text}</b>',
                    showarrow=False, yshift=18,
                    font=dict(size=12, color=color),
                )

            st.plotly_chart(fig_cmp, use_container_width=True)
        else:
            st.info('비교 차트를 보려면 사이드바에서 비교 주차를 선택하세요.')

        st.markdown('##### 채널 상세 (상위 20)')
        ch = (m_base.groupby('채널명', as_index=False)
              .agg(유입수=('유입수', 'sum'),
                   광고비=('광고비', 'sum'),
                   결제수=('결제수(마지막클릭)', 'sum'),
                   결제금액=('결제금액(마지막클릭)', 'sum')))
        ch['ROAS(%)'] = ch.apply(
            lambda r: (r['결제금액'] / r['광고비'] * 100) if r['광고비'] else None, axis=1)
        ch['유입당결제율(%)'] = ch.apply(
            lambda r: (r['결제수'] / r['유입수'] * 100) if r['유입수'] else 0, axis=1)
        ch = ch.sort_values('유입수', ascending=False).head(20)
        st.dataframe(
            ch,
            use_container_width=True,
            column_config={
                '유입수': st.column_config.NumberColumn(format='localized'),
                '광고비': st.column_config.NumberColumn(format='localized'),
                '결제수': st.column_config.NumberColumn(format='localized'),
                '결제금액': st.column_config.NumberColumn(format='localized'),
                'ROAS(%)': st.column_config.NumberColumn(format='%.0f%%'),
                '유입당결제율(%)': st.column_config.NumberColumn(format='%.2f%%'),
            },
            hide_index=True,
        )




# ── 탭 4: 팀장님 보고 ─────────────────────────────────────────────────────
with tab4:
    st.markdown('### 📩 슬랙 메세지 전송')
    # 슬랙 전송 버튼 — 핑크 색상으로 구분
    st.markdown(
        """
        <style>
        div[data-testid="stButton"] button[kind="primary"] {
            background-color: #FF4FA3 !important;
            border-color: #FF4FA3 !important;
            color: white !important;
        }
        div[data-testid="stButton"] button[kind="primary"]:hover {
            background-color: #E6398C !important;
            border-color: #E6398C !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    # 보고 주차 = 사이드바 기준 주차 (통합 분석과 동일한 기준)
    # 진행 행사는 기준 주차에 등록된 것만 사용
    report_weeks = [base_week] if base_week else []
    report_event_name = ' / '.join(sorted(load_events(base_week))) if base_week else ''

    # 현재 옵션 기준 자동 생성된 메시지 (편집 전 자동 갱신용)
    forced_cw = [compare_week] if compare_week else None
    auto_msg = build_report_message(
        sales_df, mkt_df, report_weeks, report_event_name,
        target_eok=target_eok_active,
        forced_compare_weeks=forced_cw,
    )

    # 자동 동기화 — 사용자가 편집하기 전까지는 옵션 변경 시 자동 갱신
    if 'report_msg_content' not in st.session_state:
        st.session_state['report_msg_content'] = auto_msg
        st.session_state['_report_last_auto_msg'] = auto_msg

    _last_auto = st.session_state.get('_report_last_auto_msg', '')
    _current = st.session_state.get('report_msg_content', '')
    _is_user_edited = (_current != _last_auto)

    if not _is_user_edited and auto_msg != _last_auto:
        st.session_state['report_msg_content'] = auto_msg
        st.session_state['_report_last_auto_msg'] = auto_msg

    slack_cfg = load_slack_config()
    dm_url = slack_cfg.get('dm_url', '').strip()
    self_url = slack_cfg.get('self_url', '').strip()
    team_id = slack_cfg.get('team_id', '').strip()
    manager_id = slack_cfg.get('manager_user_id', '').strip()
    manager_name = slack_cfg.get('manager_name', '팀장님')

    slack_url = ''
    if dm_url:
        slack_url = dm_url
    elif team_id and manager_id:
        slack_url = f'slack://user?team={team_id}&id={manager_id}'

    # 메시지 박스 + 우측 액션 패널 (한 화면에 보이도록)
    msg_col, action_col = st.columns([3, 1])
    with msg_col:
        st.text_area(
            '실적 요약 (직접 수정 가능)',
            key='report_msg_content',
            height=420,
        )
    report_msg = st.session_state.get('report_msg_content', '')

    with action_col:
        # 왼쪽 text_area label 높이만큼 빈 공간 → 버튼이 메시지창 상단과 정렬
        st.markdown(
            '<div style="height:1.85rem;"></div>',
            unsafe_allow_html=True,
        )
        if slack_url:
            rpa_clicked = st.button(
                f'🚀 {manager_name}께 전송',
                type='primary',
                key='rpa_send_btn',
                use_container_width=True,
            )
            if rpa_clicked:
                with st.spinner('슬랙 앱을 열고 메시지를 붙여넣고 전송 중...'):
                    ok, info = send_to_slack_rpa(
                        report_msg, slack_url, auto_send=True)
                if ok:
                    st.success(f'✅ {info}')
                else:
                    st.error(f'❌ {info}')

            if self_url:
                self_clicked = st.button(
                    '📌 나에게 전송',
                    key='rpa_self_btn',
                    use_container_width=True,
                )
                if self_clicked:
                    with st.spinner('내 슬랙으로 전송 중...'):
                        ok, info = send_to_slack_rpa(
                            report_msg, self_url, auto_send=True)
                    if ok:
                        st.success(f'✅ {info}')
                    else:
                        st.error(f'❌ {info}')
        else:
            st.info(
                'ℹ️ 슬랙 자동 열기를 사용하려면 프로젝트 루트에 `.slack_config.json` '
                '파일 필요. 아래 형식으로 만들어주세요:\n\n'
                '```json\n'
                '{\n'
                '  "dm_url": "https://.../archives/D...",\n'
                '  "manager_name": "팀장님 성함"\n'
                '}\n'
                '```'
            )
