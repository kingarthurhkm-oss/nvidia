#!/usr/bin/env python3
"""
NVIDIA 재무제표 자동 업데이트 생성기
SEC EDGAR XBRL API (data.sec.gov) 에서 최신 10-Q 데이터를 가져와 Excel 파일을 생성합니다.

사용법:
    python nvidia_auto_update.py                  # 기본 실행 (캐시 사용)
    python nvidia_auto_update.py --no-cache       # 강제 새로 다운로드
    python nvidia_auto_update.py --quarters 6     # 표시할 분기 수 (기본: 5)
    python nvidia_auto_update.py --output 파일명.xlsx
"""

import argparse
import json
import logging
import time
import urllib.request
import urllib.error
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── 상수 ─────────────────────────────────────────────────────────────────────
CIK            = "0001045810"
COMPANY_NAME   = "NVIDIA Corporation"
FACTS_URL      = f"https://data.sec.gov/api/xbrl/companyfacts/CIK{CIK}.json"
SEC_USER_AGENT = "NVIDIA-FinancialUpdater researcher@example.com"
SCRIPT_DIR     = Path(__file__).parent
CACHE_FILE     = SCRIPT_DIR / ".nvda_xbrl_cache.json"
CACHE_TTL_H    = 12  # 캐시 유효시간 (시간)

# ── 엑셀 색상 ─────────────────────────────────────────────────────────────────
C_DARK  = "1F3864"
C_MID   = "2E75B6"
C_LIGHT = "BDD7EE"
C_TOTAL = "E2EFDA"
C_GREEN = "375623"
C_SUB   = "D6E4F0"
C_GRAY  = "F2F2F2"
C_WHITE = "FFFFFF"

# ── XBRL 개념 매핑 (우선순위 순) ─────────────────────────────────────────────
INCOME_MAP = {
    "revenue":         ["Revenues", "RevenueFromContractWithCustomerExcludingAssessedTax"],
    "cost_of_revenue": ["CostOfRevenue", "CostOfGoodsAndServicesSold"],
    "gross_profit":    ["GrossProfit"],
    "rd":              ["ResearchAndDevelopmentExpense"],
    "sga":             ["SellingGeneralAndAdministrativeExpense"],
    "op_income":       ["OperatingIncomeLoss"],
    "int_income":      ["InvestmentIncomeInterest", "InterestAndDividendIncomeOperating"],
    "int_expense":     ["InterestExpense"],
    "other_income":    ["NonoperatingIncomeExpense", "OtherNonoperatingIncomeExpense"],
    "pretax":          ["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"],
    "tax":             ["IncomeTaxExpenseBenefit"],
    "net_income":      ["NetIncomeLoss"],
    "eps_basic":       ["EarningsPerShareBasic"],
    "eps_diluted":     ["EarningsPerShareDiluted"],
}

BALANCE_MAP = {
    "cash":         ["CashAndCashEquivalentsAtCarryingValue"],
    "st_invest":    ["ShortTermInvestments", "MarketableSecuritiesCurrent"],
    "ar":           ["AccountsReceivableNetCurrent"],
    "inventories":  ["InventoryNet"],
    "prepaid":      ["PrepaidExpenseAndOtherAssetsCurrent"],
    "total_ca":     ["AssetsCurrent"],
    "ppe":          ["PropertyPlantAndEquipmentNet"],
    "goodwill":     ["Goodwill"],
    "intangibles":  ["IntangibleAssetsNetExcludingGoodwill", "FiniteLivedIntangibleAssetsNet"],
    "def_tax_a":    ["DeferredIncomeTaxAssetsNet", "DeferredTaxAssetsNetNoncurrent"],
    "other_nca":    ["OtherAssetsNoncurrent"],
    "total_assets": ["Assets"],
    "ap":           ["AccountsPayableCurrent"],
    "accrued":      ["AccruedLiabilitiesCurrent", "OtherLiabilitiesCurrent"],
    "st_debt":      ["ShortTermBorrowings", "LongTermDebtCurrent"],
    "total_cl":     ["LiabilitiesCurrent"],
    "lt_debt":      ["LongTermDebtNoncurrent", "LongTermDebt"],
    "lt_lease":     ["OperatingLeaseLiabilityNoncurrent"],
    "other_ncl":    ["OtherLiabilitiesNoncurrent"],
    "total_liab":   ["Liabilities"],
    "equity":       ["StockholdersEquity", "StockholdersEquityAttributableToParent"],
}

CASHFLOW_MAP = {
    "net_income":    ["NetIncomeLoss"],
    "sbc":           ["ShareBasedCompensation"],
    "dna":           ["DepreciationDepletionAndAmortization", "Depreciation"],
    "def_tax":       ["DeferredIncomeTaxExpenseBenefit"],
    "cfo":           ["NetCashProvidedByUsedInOperatingActivities"],
    "capex":         ["PaymentsToAcquirePropertyPlantAndEquipment"],
    "buy_invest":    ["PaymentsToAcquireAvailableForSaleSecuritiesDebt",
                      "PaymentsToAcquireMarketableSecurities"],
    "sell_invest":   ["ProceedsFromSaleAndMaturityOfAvailableForSaleSecurities",
                      "ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities"],
    "cfi":           ["NetCashProvidedByUsedInInvestingActivities"],
    "buybacks":      ["PaymentsForRepurchaseOfCommonStock"],
    "dividends":     ["PaymentsOfDividends"],
    "stock_proceeds":["ProceedsFromStockOptionsExercised", "ProceedsFromIssuanceOfCommonStock"],
    "cff":           ["NetCashProvidedByUsedInFinancingActivities"],
}

EPS_KEYS = {"eps_basic", "eps_diluted"}  # 스케일 없음

# ═════════════════════════════════════════════════════════════════════════════
# 1. 시드 데이터 (SEC EDGAR 접근 불가 시 폴백)
#    출처: NVIDIA Form 10-Q, nvda-20251026, SEC EDGAR CIK 0001045810
# ═════════════════════════════════════════════════════════════════════════════

def _fact(val_m, start, end, form, fy, fp, filed):
    """백만 달러 값 → XBRL 팩트 딕셔너리 (USD 전액 기준)"""
    return {"val": int(val_m * 1_000_000), "start": start, "end": end,
            "form": form, "fy": fy, "fp": fp, "filed": filed, "accn": f"seed-{end}-{fp}"}

def _fact_i(val_m, end, form, fy, fp, filed):
    """Instant (대차대조표) 팩트"""
    return {"val": int(val_m * 1_000_000), "end": end,
            "form": form, "fy": fy, "fp": fp, "filed": filed, "accn": f"seed-bs-{end}"}

def _eps(val, start, end, form, fy, fp, filed):
    return {"val": val, "start": start, "end": end,
            "form": form, "fy": fy, "fp": fp, "filed": filed, "accn": f"seed-eps-{end}-{fp}"}


def get_seed_data() -> dict:
    """
    SEC 10-Q 공시 기반 시드 데이터 (XBRL 형식 호환).
    출처: nvda-20251026 (Q3 FY2026), nvda-20250727 (Q2), nvda-20250427 (Q1)
    """
    # 기간 정의: (start, end, form, fy, fp, filed)  ← _fact 시그니처와 일치
    Q1    = ("2025-01-27", "2025-04-27", "10-Q", 2026, "Q1", "2025-05-28")
    Q2    = ("2025-04-28", "2025-07-27", "10-Q", 2026, "Q2", "2025-08-27")
    Q3    = ("2025-07-28", "2025-10-26", "10-Q", 2026, "Q3", "2025-11-19")
    YTD9  = ("2025-01-27", "2025-10-26", "10-Q", 2026, "Q3", "2025-11-19")
    Q3Y   = ("2024-07-29", "2024-10-27", "10-Q", 2025, "Q3", "2024-11-20")
    YTD9Y = ("2024-01-29", "2024-10-27", "10-Q", 2025, "Q3", "2024-11-20")

    def dur(vals, periods):
        return [_fact(*((v,) + p)) for v, p in zip(vals, periods)]

    def inst(vals, dates_forms):
        return [_fact_i(*((v,) + d)) for v, d in zip(vals, dates_forms)]

    # 대차대조표 날짜: (end, form, fy, fp, filed)
    BS_Q3 = ("2025-10-26", "10-Q", 2026, "Q3", "2025-11-19")
    BS_FY = ("2025-01-26", "10-K", 2025, "FY", "2025-02-26")

    g = {}  # us-gaap 팩트 딕셔너리

    # ── 손익계산서 ──────────────────────────────────────────────────────────
    rev = [44062, 46743, 57006, 147811, 35082, 91166]
    g["Revenues"] = {"units": {"USD": dur(rev, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    cor = [17394, 12890, 15157, 45441, 8926, 22031]
    g["CostOfRevenue"] = {"units": {"USD": dur(cor, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    gp = [26668, 33853, 41849, 102370, 26156, 69135]
    g["GrossProfit"] = {"units": {"USD": dur(gp, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    rd = [3989, 4291, 4705, 12985, 3390, 9200]
    g["ResearchAndDevelopmentExpense"] = {"units": {"USD": dur(rd, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    sga = [1041, 1122, 1134, 3297, 897, 2516]
    g["SellingGeneralAndAdministrativeExpense"] = {"units": {"USD": dur(sga, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    oi = [21638, 28440, 36010, 86088, 21869, 57419]
    g["OperatingIncomeLoss"] = {"units": {"USD": dur(oi, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    ii = [None, None, 624, 1732, 472, 1275]
    g["InvestmentIncomeInterest"] = {"units": {"USD": dur(
        [x for x in ii if x is not None],
        [p for x, p in zip(ii, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y]) if x is not None]
    )}}

    ie = [None, None, 61, 186, 61, 186]
    g["InterestExpense"] = {"units": {"USD": dur(
        [x for x in ie if x is not None],
        [p for x, p in zip(ie, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y]) if x is not None]
    )}}

    oth = [None, None, 1363, 3418, 36, 301]
    g["NonoperatingIncomeExpense"] = {"units": {"USD": dur(
        [x for x in oth if x is not None],
        [p for x, p in zip(oth, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y]) if x is not None]
    )}}

    pt = [None, None, 37936, 91052, 22316, 58809]
    g["IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest"] = {
        "units": {"USD": dur(
            [x for x in pt if x is not None],
            [p for x, p in zip(pt, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y]) if x is not None]
        )}
    }

    tax = [None, None, 6026, 13945, 3007, 8020]
    g["IncomeTaxExpenseBenefit"] = {"units": {"USD": dur(
        [x for x in tax if x is not None],
        [p for x, p in zip(tax, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y]) if x is not None]
    )}}

    ni = [18775, 26422, 31910, 77107, 19309, 50789]
    g["NetIncomeLoss"] = {"units": {"USD": dur(ni, [Q1, Q2, Q3, YTD9, Q3Y, YTD9Y])}}

    # Q3 tuple: (start, end, form, fy, fp, filed) — all 6 elements
    g["EarningsPerShareBasic"] = {"units": {"USD/shares": [
        _eps(1.31, *Q3), _eps(0.79, *Q3Y)
    ]}}
    g["EarningsPerShareDiluted"] = {"units": {"USD/shares": [
        _eps(1.30, *Q3), _eps(0.78, *Q3Y)
    ]}}

    # ── 대차대조표 ──────────────────────────────────────────────────────────
    g["CashAndCashEquivalentsAtCarryingValue"] = {"units": {"USD": inst(
        [8589, 8589], [BS_Q3, BS_FY])}}
    g["ShortTermInvestments"] = {"units": {"USD": inst([52019, 34621], [BS_Q3, BS_FY])}}
    g["AccountsReceivableNetCurrent"] = {"units": {"USD": inst([33391, 23065], [BS_Q3, BS_FY])}}
    g["InventoryNet"] = {"units": {"USD": inst([19784, 10080], [BS_Q3, BS_FY])}}
    g["PrepaidExpenseAndOtherAssetsCurrent"] = {"units": {"USD": inst([2709, 3771], [BS_Q3, BS_FY])}}
    g["AssetsCurrent"] = {"units": {"USD": inst([116492, 80126], [BS_Q3, BS_FY])}}
    g["PropertyPlantAndEquipmentNet"] = {"units": {"USD": inst([9780, 6283], [BS_Q3, BS_FY])}}
    g["Goodwill"] = {"units": {"USD": inst([6261, 5188], [BS_Q3, BS_FY])}}
    g["IntangibleAssetsNetExcludingGoodwill"] = {"units": {"USD": inst([936, 807], [BS_Q3, BS_FY])}}
    g["DeferredIncomeTaxAssetsNet"] = {"units": {"USD": inst([13674, 10979], [BS_Q3, BS_FY])}}
    g["OtherAssetsNoncurrent"] = {"units": {"USD": inst([11724, 6425], [BS_Q3, BS_FY])}}
    g["Assets"] = {"units": {"USD": inst([161148, 111601], [BS_Q3, BS_FY])}}
    g["AccountsPayableCurrent"] = {"units": {"USD": inst([8624, 6310], [BS_Q3, BS_FY])}}
    g["AccruedLiabilitiesCurrent"] = {"units": {"USD": inst([16452, 11737], [BS_Q3, BS_FY])}}
    g["LongTermDebtCurrent"] = {"units": {"USD": inst([999, 0], [BS_Q3, BS_FY])}}
    g["LiabilitiesCurrent"] = {"units": {"USD": inst([26075, 18047], [BS_Q3, BS_FY])}}
    g["LongTermDebtNoncurrent"] = {"units": {"USD": inst([7468, 8463], [BS_Q3, BS_FY])}}
    g["OperatingLeaseLiabilityNoncurrent"] = {"units": {"USD": inst([2014, 1519], [BS_Q3, BS_FY])}}
    g["OtherLiabilitiesNoncurrent"] = {"units": {"USD": inst([6694, 4245], [BS_Q3, BS_FY])}}
    g["Liabilities"] = {"units": {"USD": inst([42251, 32274], [BS_Q3, BS_FY])}}
    g["StockholdersEquity"] = {"units": {"USD": inst([118897, 79327], [BS_Q3, BS_FY])}}

    # ── 현금흐름표 (YTD) ────────────────────────────────────────────────────
    cf_pairs = [YTD9, YTD9Y]
    g["ShareBasedCompensation"] = {"units": {"USD": dur([4753, 3416], cf_pairs)}}
    g["DepreciationDepletionAndAmortization"] = {"units": {"USD": dur([2031, 1128], cf_pairs)}}
    g["DeferredIncomeTaxExpenseBenefit"] = {"units": {"USD": dur([-2670, -2461], cf_pairs)}}
    g["NetCashProvidedByUsedInOperatingActivities"] = {"units": {"USD": dur([66530, 50099], cf_pairs)}}
    g["PaymentsToAcquirePropertyPlantAndEquipment"] = {"units": {"USD": dur([4758, 1560], cf_pairs)}}
    g["PaymentsToAcquireAvailableForSaleSecuritiesDebt"] = {"units": {"USD": dur([121949, 55053], cf_pairs)}}
    g["ProceedsFromSaleAndMaturityOfAvailableForSaleSecurities"] = {"units": {"USD": dur([113526, 44018], cf_pairs)}}
    g["NetCashProvidedByUsedInInvestingActivities"] = {"units": {"USD": dur([-13223, -12600], cf_pairs)}}
    g["PaymentsForRepurchaseOfCommonStock"] = {"units": {"USD": dur([36271, 25162], cf_pairs)}}
    g["PaymentsOfDividends"] = {"units": {"USD": dur([732, 458], cf_pairs)}}
    g["ProceedsFromStockOptionsExercised"] = {"units": {"USD": dur([643, 521], cf_pairs)}}
    g["NetCashProvidedByUsedInFinancingActivities"] = {"units": {"USD": dur([-42266, -28862], cf_pairs)}}
    g["CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect"] = {
        "units": {"USD": dur([11041, 8637], cf_pairs)}
    }

    return {
        "cik": CIK.lstrip("0"),
        "entityName": COMPANY_NAME,
        "facts": {"us-gaap": g},
        "_seed": True,
        "_seed_source": "NVIDIA Form 10-Q (nvda-20251026), SEC EDGAR CIK 0001045810",
    }


# ═════════════════════════════════════════════════════════════════════════════
# 2. 데이터 수집
# ═════════════════════════════════════════════════════════════════════════════

def fetch_json(url: str, retries: int = 2) -> dict:
    req = urllib.request.Request(
        url,
        headers={"User-Agent": SEC_USER_AGENT, "Accept": "application/json"},
    )
    for attempt in range(retries):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return json.loads(resp.read().decode("utf-8"))
        except Exception as exc:
            wait = 2 ** attempt
            logger.warning(f"요청 오류 (시도 {attempt+1}/{retries}): {exc}")
            if attempt < retries - 1:
                time.sleep(wait)
    raise RuntimeError(f"API 요청 실패: {url}")


def load_facts(no_cache: bool = False) -> dict:
    # 1) 유효한 캐시가 있으면 사용
    if not no_cache and CACHE_FILE.exists():
        age_h = (time.time() - CACHE_FILE.stat().st_mtime) / 3600
        if age_h < CACHE_TTL_H:
            logger.info(f"캐시 사용 (최종 갱신: {age_h:.1f}시간 전)")
            return json.loads(CACHE_FILE.read_text(encoding="utf-8"))

    # 2) SEC EDGAR API 시도
    logger.info("SEC EDGAR XBRL 데이터 다운로드 시도 중...")
    try:
        data = fetch_json(FACTS_URL)
        CACHE_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
        logger.info(f"SEC EDGAR 다운로드 성공 → 캐시 저장: {CACHE_FILE.name}")
        return data
    except RuntimeError as e:
        logger.warning(f"SEC EDGAR 접근 실패: {e}")
        logger.warning("→ 내장 시드 데이터(FY2026 Q1-Q3)를 사용합니다.")
        logger.warning("  최신 데이터를 받으려면 네트워크 연결 후 --no-cache 옵션으로 재실행하세요.")
        return get_seed_data()


# ═════════════════════════════════════════════════════════════════════════════
# 2. XBRL 파싱
# ═════════════════════════════════════════════════════════════════════════════

def _get_usd_values(us_gaap: dict, candidates: list) -> list:
    """후보 개념 중 첫 번째 USD 단위 값 목록을 반환"""
    for c in candidates:
        if c not in us_gaap:
            continue
        units = us_gaap[c].get("units", {})
        if "USD" in units:
            return units["USD"]
        for v in units.values():  # USD/shares 등
            return v
    return []


def _duration_map(values: list, min_d: int, max_d: int, scale: float) -> dict:
    """
    10-Q 팩트 중 특정 기간(min_d~max_d일)의 값을 {end_date: value} 로 반환.
    같은 end_date에 여러 팩트가 있으면 가장 최근 filed 우선.
    """
    best: dict[str, tuple] = {}  # {end: (val, filed)}
    for v in values:
        if v.get("form") != "10-Q":
            continue
        s, e = v.get("start"), v.get("end")
        if not s or not e:
            continue
        days = (datetime.strptime(e, "%Y-%m-%d") - datetime.strptime(s, "%Y-%m-%d")).days
        if not (min_d <= days <= max_d):
            continue
        filed = v.get("filed", "")
        if e not in best or filed > best[e][1]:
            best[e] = (v["val"], filed)
    return {dt: round(val / scale, 2) for dt, (val, _) in best.items()}


def _instant_map(values: list, scale: float) -> dict:
    """10-Q / 10-K instant 팩트를 {end_date: value} 로 반환"""
    best: dict[str, tuple] = {}
    for v in values:
        if v.get("form") not in ("10-Q", "10-K"):
            continue
        if v.get("start"):  # instant는 start 없음
            continue
        e = v.get("end", "")
        filed = v.get("filed", "")
        if e not in best or filed > best[e][1]:
            best[e] = (v["val"], filed)
    return {dt: round(val / scale, 2) for dt, (val, _) in best.items()}


def extract(facts: dict, num_quarters: int) -> dict:
    gaap = facts.get("facts", {}).get("us-gaap", {})
    S = 1_000_000

    # 손익계산서: 3개월 단기 (75~105일)
    inc3 = {}
    for k, cands in INCOME_MAP.items():
        vals = _get_usd_values(gaap, cands)
        inc3[k] = _duration_map(vals, 75, 105, 1.0 if k in EPS_KEYS else S)

    # 손익계산서: YTD (75~295일, Q1~Q3 누적 포함)
    inc_ytd = {}
    for k, cands in INCOME_MAP.items():
        vals = _get_usd_values(gaap, cands)
        inc_ytd[k] = _duration_map(vals, 75, 295, 1.0 if k in EPS_KEYS else S)

    # 대차대조표: instant
    bal = {}
    for k, cands in BALANCE_MAP.items():
        vals = _get_usd_values(gaap, cands)
        bal[k] = _instant_map(vals, S)

    # 현금흐름표: YTD
    cf = {}
    for k, cands in CASHFLOW_MAP.items():
        vals = _get_usd_values(gaap, cands)
        cf[k] = _duration_map(vals, 75, 295, S)

    # 최근 N분기 날짜 (revenue 기준, 최신 순)
    qtrs = sorted(inc3.get("revenue", {}).keys(), reverse=True)[:num_quarters]

    return {
        "qtrs":    qtrs,
        "inc3":    inc3,
        "inc_ytd": inc_ytd,
        "bal":     bal,
        "cf":      cf,
        "meta": {
            "company": facts.get("entityName", COMPANY_NAME),
            "cik":     facts.get("cik", CIK),
            "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "url":     FACTS_URL,
        },
    }


# ═════════════════════════════════════════════════════════════════════════════
# 3. 레이블 유틸
# ═════════════════════════════════════════════════════════════════════════════

def fq_label(date_str: str) -> str:
    """'2025-10-26' → 'Q3 FY2026'  (NVIDIA 회계연도 기준)"""
    d = datetime.strptime(date_str, "%Y-%m-%d")
    m, y = d.month, d.year
    if m == 1:
        return f"Q4 FY{y}"
    elif m <= 4:
        return f"Q1 FY{y+1}"
    elif m <= 7:
        return f"Q2 FY{y+1}"
    elif m <= 10:
        return f"Q3 FY{y+1}"
    else:
        return f"Q4 FY{y+1}"


def ytd_label(date_str: str, start_str: str) -> str:
    """YTD 기간 라벨: '9M FY2026\n(2025.10.26)' 형식"""
    d_end = datetime.strptime(date_str, "%Y-%m-%d")
    d_st  = datetime.strptime(start_str, "%Y-%m-%d")
    months = round((d_end - d_st).days / 30.4)
    fq = fq_label(date_str)
    fy = fq.split()[-1]
    return f"{months}M {fy}\n({d_end.strftime('%Y.%m.%d')})"


def col_header(date_str: str) -> str:
    return f"{fq_label(date_str)}\n({datetime.strptime(date_str, '%Y-%m-%d').strftime('%Y.%m.%d')})"


def get_ytd_dates(cf: dict, key: str = "cfo") -> list:
    """현금흐름 YTD 날짜 목록 (최신 순)"""
    return sorted(cf.get(key, {}).keys(), reverse=True)


# ═════════════════════════════════════════════════════════════════════════════
# 4. 엑셀 스타일 헬퍼
# ═════════════════════════════════════════════════════════════════════════════

def _f(bold=False, italic=False, color="000000", size=10):
    return Font(name="Arial", bold=bold, italic=italic, color=color, size=size)

def _fill(c):
    return PatternFill("solid", fgColor=c)

def _al(h="left", v="center", wrap=False, ind=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=ind)


def hdr(cell, text, bg=C_DARK, fc="FFFFFF", size=10, wrap=True):
    cell.value = text
    cell.font  = _f(bold=True, color=fc, size=size)
    cell.fill  = _fill(bg)
    cell.alignment = _al(h="center", wrap=wrap)

def sec(cell, text, bg=C_MID):
    cell.value = text
    cell.font  = _f(bold=True, color="FFFFFF")
    cell.fill  = _fill(bg)
    cell.alignment = _al(h="left", ind=1)

def lbl(cell, text, ind=0, bold=False, italic=False, bg=None):
    cell.value = text
    cell.font  = _f(bold=bold, italic=italic)
    cell.alignment = _al(h="left", ind=ind)
    if bg:
        cell.fill = _fill(bg)

def tot_lbl(cell, text=None):
    if text is not None:
        cell.value = text
    cell.font  = _f(bold=True, color=C_GREEN)
    cell.fill  = _fill(C_TOTAL)
    cell.alignment = _al(h="left", ind=1)

def num(cell, val, bold=False, italic=False, bg=None, fmt="#,##0"):
    cell.value = val
    neg_color = "CC0000" if val is not None and val < 0 else "000000"
    cell.font  = _f(bold=bold, italic=italic, color=neg_color)
    cell.alignment = _al(h="right")
    cell.number_format = fmt
    if bg:
        cell.fill = _fill(bg)

def tot_num(cell, val=None, fmt="#,##0"):
    if val is not None:
        cell.value = val
    cell.font  = _f(bold=True, color=C_GREEN)
    cell.fill  = _fill(C_TOTAL)
    cell.alignment = _al(h="right")
    cell.number_format = fmt

def pct_cell(cell, val, bg=None):
    cell.value = (val / 100) if val is not None else None
    cell.font  = _f(italic=True, size=9, color="555555")
    cell.alignment = _al(h="right")
    cell.number_format = "0.0%"
    if bg:
        cell.fill = _fill(bg)

def na(cell, bg=None):
    cell.value = "N/A"
    cell.font  = _f(italic=True, color="999999", size=9)
    cell.alignment = _al(h="right")
    if bg:
        cell.fill = _fill(bg)


def _v(data_dict: dict, date_key: str, scale_back: float = 1.0):
    """data_dict에서 date_key의 값 반환. 없으면 None."""
    v = data_dict.get(date_key)
    return round(v * scale_back, 2) if v is not None else None


def write_title(ws, title, subtitle, ncols):
    ws.sheet_view.showGridLines = False
    ws.row_dimensions[1].height = 38
    ws.merge_cells(f"A1:{chr(64+ncols)}1")
    c = ws["A1"]
    c.value = title
    c.font  = _f(bold=True, color="FFFFFF", size=15)
    c.fill  = _fill(C_DARK)
    c.alignment = _al(h="center")

    ws.row_dimensions[2].height = 20
    ws.merge_cells(f"A2:{chr(64+ncols)}2")
    c = ws["A2"]
    c.value = subtitle
    c.font  = _f(italic=True, color="FFFFFF", size=9)
    c.fill  = _fill("2E4057")
    c.alignment = _al(h="center")

    ws.row_dimensions[3].height = 6


def write_note(ws, note, row, ncols):
    ws.row_dimensions[row].height = 28
    ws.merge_cells(f"A{row}:{chr(64+ncols)}{row}")
    c = ws.cell(row=row, column=1)
    c.value = note
    c.font  = _f(italic=True, color="666666", size=8)
    c.alignment = _al(h="left", wrap=True)


# ═════════════════════════════════════════════════════════════════════════════
# 5. 손익계산서 시트
# ═════════════════════════════════════════════════════════════════════════════

def sheet_income(wb, data: dict):
    ws = wb.create_sheet("손익계산서")
    qtrs = data["qtrs"]   # 최신 순 날짜 리스트
    inc  = data["inc3"]
    ncols = 1 + len(qtrs)

    ws.column_dimensions["A"].width = 40
    for i in range(len(qtrs)):
        ws.column_dimensions[chr(66+i)].width = 17

    write_title(ws,
        "NVIDIA Corporation - 손익계산서 (Condensed Consolidated Statements of Income)",
        f"출처: SEC EDGAR XBRL API | Form 10-Q | 단위: 백만 달러(USD Millions) | 미감사 | 자동 업데이트: {data['meta']['updated']}",
        ncols)

    # 헤더 행
    R = 4
    ws.row_dimensions[R].height = 42
    hdr(ws.cell(R, 1), "항목 (Line Item)", bg=C_DARK, size=10)
    for i, dt in enumerate(qtrs):
        hdr(ws.cell(R, 2+i), col_header(dt), bg=C_MID if i % 2 == 0 else C_DARK, size=9)

    # 데이터 정의: (label, key, indent, bold, is_total, is_pct, pct_num_key, pct_den_key)
    rows = [
        ("§ 매출",                              None,          0, True,  False, False, None, None),
        ("매출 (Revenue)",                       "revenue",     0, True,  True,  False, None, None),
        ("§ 원가·비용",                           None,          0, True,  False, False, None, None),
        ("  매출원가 (Cost of Revenue)",          "cost_of_revenue", 1, False, False, False, None, None),
        ("  매출총이익 (Gross Profit)",            "gross_profit", 0, True, True, False, None, None),
        ("  매출총이익률 (Gross Margin)",           None,          1, False, False, True, "gross_profit", "revenue"),
        ("§ 영업비용",                            None,          0, True,  False, False, None, None),
        ("  연구개발비 (R&D)",                    "rd",          1, False, False, False, None, None),
        ("  판매비와관리비 (SG&A)",               "sga",         1, False, False, False, None, None),
        ("§ 영업이익",                            None,          0, True,  False, False, None, None),
        ("영업이익 (Operating Income)",           "op_income",   0, True,  True,  False, None, None),
        ("영업이익률 (Operating Margin)",          None,          1, False, False, True, "op_income", "revenue"),
        ("§ 영업외손익",                          None,          0, True,  False, False, None, None),
        ("  이자수익 (Interest Income)",          "int_income",  1, False, False, False, None, None),
        ("  이자비용 (Interest Expense)",         "int_expense", 1, False, False, False, None, None),
        ("  기타수익(비용) (Other Income, Net)",  "other_income",1, False, False, False, None, None),
        ("§ 세전·순이익",                         None,          0, True,  False, False, None, None),
        ("세전이익 (Pretax Income)",              "pretax",      0, True,  True,  False, None, None),
        ("  법인세비용 (Income Tax Expense)",      "tax",         1, False, False, False, None, None),
        ("순이익 (Net Income)",                   "net_income",  0, True,  True,  False, None, None),
        ("순이익률 (Net Margin)",                 None,          1, False, False, True, "net_income", "revenue"),
        ("§ 주당순이익 (EPS)",                    None,          0, True,  False, False, None, None),
        ("  기본 EPS (Basic, USD)",               "eps_basic",   1, False, False, False, None, None),
        ("  희석 EPS (Diluted, USD)",             "eps_diluted", 1, False, False, False, None, None),
    ]

    for ri, row_def in enumerate(rows):
        label, key, indent, bold, is_total, is_pct, pn_key, pd_key = row_def
        r = R + 1 + ri
        ws.row_dimensions[r].height = 18
        is_sec = label.startswith("§")
        bg_row = C_GRAY if ri % 2 == 0 else C_WHITE

        if is_sec:
            sec(ws.cell(r, 1), label[2:].strip())
            for i in range(len(qtrs)):
                ws.cell(r, 2+i).fill = _fill(C_MID)
        elif is_pct:
            lbl(ws.cell(r, 1), label.strip(), ind=indent, italic=True, bg="EEF5FB")
            for i, dt in enumerate(qtrs):
                pn = inc.get(pn_key, {}).get(dt)
                pd_ = inc.get(pd_key, {}).get(dt)
                pct_v = (pn / pd_ * 100) if (pn and pd_) else None
                pct_cell(ws.cell(r, 2+i), pct_v, bg="EEF5FB")
        elif is_total:
            tot_lbl(ws.cell(r, 1))
            ws.cell(r, 1).value = label.strip()
            for i, dt in enumerate(qtrs):
                v = _v(inc.get(key, {}), dt)
                fmt = '$#,##0.00' if key in EPS_KEYS else "#,##0"
                if v is not None:
                    tot_num(ws.cell(r, 2+i), v, fmt=fmt)
                else:
                    na(ws.cell(r, 2+i), bg=C_TOTAL)
        else:
            lbl(ws.cell(r, 1), label.strip(), ind=indent, bold=bold, bg=bg_row)
            for i, dt in enumerate(qtrs):
                v = _v(inc.get(key, {}), dt) if key else None
                fmt = '$#,##0.00' if key in EPS_KEYS else "#,##0"
                if v is not None:
                    num(ws.cell(r, 2+i), v, bold=bold, bg=bg_row, fmt=fmt)
                else:
                    na(ws.cell(r, 2+i), bg=bg_row) if key else None

    note_r = R + 1 + len(rows) + 1
    write_note(ws, (
        "※ 데이터 출처: SEC EDGAR XBRL API (data.sec.gov) | Form 10-Q | "
        f"CIK {CIK} | 단위: 백만 달러 | EPS는 달러 기준 | "
        "SEC EDGAR: https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK=0001045810&type=10-Q"
    ), note_r, ncols)
    ws.freeze_panes = f"B{R+1}"


# ═════════════════════════════════════════════════════════════════════════════
# 6. 재무상태표 시트
# ═════════════════════════════════════════════════════════════════════════════

def sheet_balance(wb, data: dict):
    ws = wb.create_sheet("재무상태표")
    qtrs = data["qtrs"]
    bal  = data["bal"]
    ncols = 1 + len(qtrs)

    ws.column_dimensions["A"].width = 44
    for i in range(len(qtrs)):
        ws.column_dimensions[chr(66+i)].width = 17

    write_title(ws,
        "NVIDIA Corporation - 재무상태표 (Condensed Consolidated Balance Sheets)",
        f"출처: SEC EDGAR XBRL API | Form 10-Q | 단위: 백만 달러(USD Millions) | 미감사 | 자동 업데이트: {data['meta']['updated']}",
        ncols)

    R = 4
    ws.row_dimensions[R].height = 42
    hdr(ws.cell(R, 1), "항목 (Line Item)", bg=C_DARK)
    for i, dt in enumerate(qtrs):
        bg = C_MID if i % 2 == 0 else C_DARK
        hdr(ws.cell(R, 2+i), f"기말 기준\n({datetime.strptime(dt, '%Y-%m-%d').strftime('%Y.%m.%d')})\n{fq_label(dt)}", bg=bg, size=9)

    rows = [
        ("자산 (ASSETS)",                             None,         "section"),
        ("유동자산 (Current Assets)",                  None,         "subsec"),
        ("현금 및 현금성자산 (Cash)",                  "cash",       "item"),
        ("단기금융상품 (Short-term Investments)",       "st_invest",  "item"),
        ("매출채권 (Accounts Receivable, Net)",        "ar",         "item"),
        ("재고자산 (Inventories)",                     "inventories","item"),
        ("선급금 및 기타유동자산 (Prepaid & Other)",    "prepaid",    "item"),
        ("유동자산 합계 (Total Current Assets)",       "total_ca",   "total"),
        ("비유동자산 (Non-Current Assets)",            None,         "subsec"),
        ("유형자산 (Property & Equipment, Net)",       "ppe",        "item"),
        ("영업권 (Goodwill)",                          "goodwill",   "item"),
        ("무형자산 (Intangible Assets, Net)",          "intangibles","item"),
        ("이연법인세자산 (Deferred Tax Assets)",        "def_tax_a",  "item"),
        ("기타비유동자산 (Other Non-Current Assets)",  "other_nca",  "item"),
        ("자산 총계 (Total Assets)",                   "total_assets","grand"),
        ("부채 (LIABILITIES)",                         None,         "section"),
        ("유동부채 (Current Liabilities)",             None,         "subsec"),
        ("매입채무 (Accounts Payable)",                "ap",         "item"),
        ("미지급금 및 기타유동부채 (Accrued & Other)", "accrued",    "item"),
        ("단기차입금 (Short-term Debt)",               "st_debt",    "item"),
        ("유동부채 합계 (Total Current Liabilities)", "total_cl",   "total"),
        ("비유동부채 (Non-Current Liabilities)",       None,         "subsec"),
        ("장기차입금 (Long-term Debt)",                "lt_debt",    "item"),
        ("장기 리스부채 (LT Operating Lease)",         "lt_lease",   "item"),
        ("기타비유동부채 (Other Non-Current Liab.)",   "other_ncl",  "item"),
        ("부채 총계 (Total Liabilities)",              "total_liab", "grand"),
        ("자기자본 (SHAREHOLDERS' EQUITY)",             None,         "section"),
        ("자기자본 합계 (Total Shareholders' Equity)", "equity",     "grand"),
        ("부채 및 자기자본 합계",                       "total_assets","grand"),
    ]

    for ri, (label, key, rtype) in enumerate(rows):
        r = R + 1 + ri
        ws.row_dimensions[r].height = 18
        bg_row = C_GRAY if ri % 2 == 0 else C_WHITE

        if rtype == "section":
            sec(ws.cell(r, 1), label)
            for i in range(len(qtrs)):
                ws.cell(r, 2+i).fill = _fill(C_MID)
        elif rtype == "subsec":
            for col in range(1, 2+len(qtrs)):
                ws.cell(r, col).fill = _fill(C_LIGHT)
            lbl(ws.cell(r, 1), label, bold=True, bg=C_LIGHT)
            ws.cell(r, 1).font = _f(bold=True, color="1F3864")
        elif rtype in ("total", "grand"):
            tot_lbl(ws.cell(r, 1))
            ws.cell(r, 1).value = label
            for i, dt in enumerate(qtrs):
                v = _v(bal.get(key, {}), dt)
                if v is not None:
                    tot_num(ws.cell(r, 2+i))
                    ws.cell(r, 2+i).value = v
                else:
                    na(ws.cell(r, 2+i), bg=C_TOTAL)
        else:
            lbl(ws.cell(r, 1), label, ind=1, bg=bg_row)
            for i, dt in enumerate(qtrs):
                v = _v(bal.get(key, {}), dt)
                if v is not None:
                    num(ws.cell(r, 2+i), v, bg=bg_row)
                else:
                    na(ws.cell(r, 2+i), bg=bg_row)

    note_r = R + 1 + len(rows) + 1
    write_note(ws, (
        "※ 대차대조표는 각 분기말 기준 잔액. 비유동자산 일부 항목(리스자산 등)은 "
        "XBRL 태깅 방식에 따라 표시되지 않을 수 있음. "
        f"출처: SEC EDGAR XBRL API | CIK {CIK}"
    ), note_r, ncols)
    ws.freeze_panes = f"B{R+1}"


# ═════════════════════════════════════════════════════════════════════════════
# 7. 현금흐름표 시트
# ═════════════════════════════════════════════════════════════════════════════

def sheet_cashflow(wb, data: dict):
    ws = wb.create_sheet("현금흐름표")
    # 현금흐름은 YTD 기준 — 가장 최근 N개 (Q3=9M, Q2=6M, Q1=3M)
    cf   = data["cf"]
    qtrs = data["qtrs"]
    ncols = 1 + len(qtrs)

    ws.column_dimensions["A"].width = 50
    for i in range(len(qtrs)):
        ws.column_dimensions[chr(66+i)].width = 17

    write_title(ws,
        "NVIDIA Corporation - 현금흐름표 (Condensed Consolidated Statements of Cash Flows)",
        f"출처: SEC EDGAR XBRL API | Form 10-Q | YTD 누적 기준 | 단위: 백만 달러 | 자동 업데이트: {data['meta']['updated']}",
        ncols)

    R = 4
    ws.row_dimensions[R].height = 42
    hdr(ws.cell(R, 1), "항목 (Line Item)", bg=C_DARK)
    for i, dt in enumerate(qtrs):
        bg = C_MID if i % 2 == 0 else C_DARK
        # YTD 기간 표시
        d = datetime.strptime(dt, "%Y-%m-%d")
        m = d.month
        if m <= 4:
            ytd_m = "Q1 (3M)"
        elif m <= 7:
            ytd_m = "Q2 (6M)"
        else:
            ytd_m = "Q3 (9M)"
        hdr(ws.cell(R, 2+i), f"{fq_label(dt)} YTD\n{ytd_m}\n({d.strftime('%Y.%m.%d')})", bg=bg, size=9)

    rows = [
        ("영업활동 (Operating Activities)",                      None,         "section"),
        ("  당기순이익 (Net Income)",                            "net_income", "item"),
        ("  주식보상비용 (Stock-Based Compensation)",             "sbc",        "item"),
        ("  감가상각비 (Depreciation & Amortization)",           "dna",        "item"),
        ("  이연법인세 변동 (Deferred Income Tax)",              "def_tax",    "item"),
        ("영업활동 순현금 (Net Cash from Operations)",           "cfo",        "total"),
        ("투자활동 (Investing Activities)",                      None,         "section"),
        ("  유형자산 취득 (Capital Expenditures)",               "capex",      "item"),
        ("  투자유가증권 매입 (Purchases of Investments)",       "buy_invest", "item"),
        ("  투자유가증권 매각·만기 (Proceeds from Investments)", "sell_invest","item"),
        ("투자활동 순현금 (Net Cash from Investing)",            "cfi",        "total"),
        ("재무활동 (Financing Activities)",                      None,         "section"),
        ("  자사주 매입 (Share Repurchases)",                    "buybacks",   "item"),
        ("  배당금 지급 (Dividends Paid)",                       "dividends",  "item"),
        ("  주식 행사 수입 (Stock Plan Proceeds)",               "stock_proceeds","item"),
        ("재무활동 순현금 (Net Cash from Financing)",            "cff",        "total"),
        ("현금 순증감 (Net Change in Cash)",                     "net_change", "grand"),
    ]

    for ri, (label, key, rtype) in enumerate(rows):
        r = R + 1 + ri
        ws.row_dimensions[r].height = 18
        bg_row = C_GRAY if ri % 2 == 0 else C_WHITE

        if rtype == "section":
            sec(ws.cell(r, 1), label)
            for i in range(len(qtrs)):
                ws.cell(r, 2+i).fill = _fill(C_MID)
        elif rtype in ("total", "grand"):
            tot_lbl(ws.cell(r, 1))
            ws.cell(r, 1).value = label
            for i, dt in enumerate(qtrs):
                v = _v(cf.get(key, {}), dt)
                if v is not None:
                    tot_num(ws.cell(r, 2+i))
                    ws.cell(r, 2+i).value = v
                else:
                    na(ws.cell(r, 2+i), bg=C_TOTAL)
        else:
            lbl(ws.cell(r, 1), label.strip(), ind=1, bg=bg_row)
            for i, dt in enumerate(qtrs):
                v = _v(cf.get(key, {}), dt)
                if v is not None:
                    num(ws.cell(r, 2+i), v, bg=bg_row)
                else:
                    na(ws.cell(r, 2+i), bg=bg_row)

    note_r = R + 1 + len(rows) + 1
    write_note(ws, (
        "※ YTD 누적 기준 표시 (Q1=3개월, Q2=6개월, Q3=9개월). "
        "일부 투자활동 세부 항목은 XBRL 태깅 방식에 따라 달라질 수 있음. "
        f"출처: SEC EDGAR XBRL API | CIK {CIK}"
    ), note_r, ncols)
    ws.freeze_panes = f"B{R+1}"


# ═════════════════════════════════════════════════════════════════════════════
# 8. 업데이트 정보 시트
# ═════════════════════════════════════════════════════════════════════════════

def sheet_meta(wb, data: dict):
    ws = wb.create_sheet("업데이트 정보")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 60

    ws.row_dimensions[1].height = 38
    ws.merge_cells("A1:B1")
    c = ws["A1"]
    c.value = "업데이트 정보 / Data Source Information"
    c.font  = _f(bold=True, color="FFFFFF", size=14)
    c.fill  = _fill(C_DARK)
    c.alignment = _al(h="center")

    meta_rows = [
        ("회사명 (Company)",       data["meta"]["company"]),
        ("CIK",                    data["meta"]["cik"]),
        ("데이터 소스 (Source)",   "SEC EDGAR XBRL API (data.sec.gov)"),
        ("공시 유형 (Filing Type)","Form 10-Q (분기보고서, 미감사)"),
        ("데이터 갱신 시각",        data["meta"]["updated"]),
        ("표시 단위 (Unit)",        "USD Millions (백만 달러)"),
        ("회계연도 기준",           "NVIDIA 회계연도 (1월 말 결산)"),
        ("표시 분기 수",            f"{len(data['qtrs'])}개 분기"),
        ("최신 분기",               fq_label(data["qtrs"][0]) if data["qtrs"] else "N/A"),
        ("최구기 분기",             fq_label(data["qtrs"][-1]) if data["qtrs"] else "N/A"),
        ("캐시 파일",               str(CACHE_FILE)),
        ("캐시 유효시간",           f"{CACHE_TTL_H}시간"),
        ("재실행 방법",             "python nvidia_auto_update.py --no-cache"),
        ("EDGAR XBRL URL",         FACTS_URL),
        ("SEC 10-Q 검색",          f"https://www.sec.gov/cgi-bin/browse-edgar?action=getcompany&CIK={CIK}&type=10-Q"),
    ]

    for ri, (k, v) in enumerate(meta_rows):
        r = 2 + ri
        ws.row_dimensions[r].height = 20
        bg = C_GRAY if ri % 2 == 0 else C_WHITE
        c1 = ws.cell(r, 1)
        c1.value = k
        c1.font  = _f(bold=True)
        c1.fill  = _fill(bg)
        c1.alignment = _al(h="left", ind=1)
        c2 = ws.cell(r, 2)
        c2.value = v
        c2.font  = _f()
        c2.fill  = _fill(bg)
        c2.alignment = _al(h="left", ind=1)


# ═════════════════════════════════════════════════════════════════════════════
# 9. 메인
# ═════════════════════════════════════════════════════════════════════════════

def build_excel(data: dict, output_path: Path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # 기본 빈 시트 제거

    sheet_income(wb, data)
    sheet_balance(wb, data)
    sheet_cashflow(wb, data)
    sheet_meta(wb, data)

    wb.save(output_path)
    logger.info(f"Excel 저장 완료: {output_path}")


def main():
    parser = argparse.ArgumentParser(description="NVIDIA 10-Q 재무제표 자동 업데이트 생성기")
    parser.add_argument("--no-cache", action="store_true", help="캐시 무시하고 새로 다운로드")
    parser.add_argument("--quarters", type=int, default=5, help="표시할 분기 수 (기본: 5)")
    parser.add_argument("--output", type=str,
                        default=str(SCRIPT_DIR / "NVIDIA_Financial_Statements_10Q_AutoUpdate.xlsx"),
                        help="출력 파일 경로")
    args = parser.parse_args()

    logger.info("=" * 60)
    logger.info("NVIDIA 재무제표 자동 업데이트 시작")
    logger.info("=" * 60)

    # 1) 데이터 수집
    facts = load_facts(no_cache=args.no_cache)

    # 2) 파싱
    logger.info(f"XBRL 데이터 파싱 중 (최근 {args.quarters}개 분기)...")
    data = extract(facts, num_quarters=args.quarters)

    qtrs = data["qtrs"]
    if not qtrs:
        logger.error("분기 데이터를 찾을 수 없습니다. SEC EDGAR XBRL 데이터를 확인하세요.")
        return

    logger.info(f"분기 발견: {', '.join(fq_label(q) for q in qtrs)}")

    # 3) Excel 생성
    out = Path(args.output)
    build_excel(data, out)

    logger.info("=" * 60)
    logger.info(f"완료! 출력 파일: {out}")
    logger.info(f"최신 분기: {fq_label(qtrs[0])} ({qtrs[0]})")
    logger.info("=" * 60)


if __name__ == "__main__":
    main()
