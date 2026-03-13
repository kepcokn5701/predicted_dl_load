"""
=============================================================================
  배전선로 휴전 가능월 검토 프로그램 v9.0  —  AI 미래 예측 엔진
  Distribution Line Suspension Feasibility Review (AI Prediction Engine)

  ■ 입력
    ① 전환선로 매핑 파일 (변전소명, 대상선로, 전환선로)
    ② 과거 부하량 데이터 폴더 (연도별 하위 폴더 또는 단일 폴더)

  ■ 핵심 로직
    1) 과거 데이터로 XGBoost 학습 (대상선로 / 전환선로 각각)
    2) 미래 1년(Target Year) 365일 전체 부하 예측
    3) 예측 합산 최대 ≤ 기준값 → 가능(O) / 초과 → 불가(X)
    4) 월별 가능(O) 일수를 AI 예측 기반으로 표시

  ■ 색상 기준: 22일+ 유력(Green) | 13~21일 고려(Orange) | 12일- 불가(Red)
=============================================================================
"""

import os
import sys
import re
import calendar
import pickle
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime, timedelta
import traceback

import numpy as np
import pandas as pd
import customtkinter as ctk

import matplotlib
matplotlib.use("TkAgg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from xgboost import XGBRegressor

# matplotlib 한글 폰트 설정
matplotlib.rcParams["font.family"] = "Malgun Gothic" if sys.platform != "darwin" else "AppleGothic"
matplotlib.rcParams["axes.unicode_minus"] = False


# ═══════════════════════════════════════════════════════════════
#  전역 설정
# ═══════════════════════════════════════════════════════════════
FONT_FAMILY = "맑은 고딕"
if sys.platform == "darwin":
    FONT_FAMILY = "Apple SD Gothic Neo"

LEVEL_HIGH = 22   # 이상 → 유력
LEVEL_MID  = 13   # 이상 → 고려, 미만 → 불가
OVERLOAD_THRESHOLD = 10  # 합산 부하 초과 기준

# ── 라이트 테마 ──
LIGHT = {
    "bg":           "#f0f2f5",
    "card_bg":      "#ffffff",
    "title_bg":     "#1e3a5f",
    "title_fg":     "#ffffff",
    "subtitle_fg":  "#90afd4",
    "text":         "#2c3e50",
    "text_sub":     "#7f8c8d",
    "accent":       "#2980b9",
    "accent_hover": "#2471a3",
    "highlight":    "#e67e22",
    "highlight_bg": "#fef9e7",
    "entry_bg":     "#f8f9fa",
    "ok":           "#27ae60",
    "ok_light":     "#eafaf1",
    "ok_cell":      "#d5f5e3",
    "warn":         "#e67e22",
    "warn_light":   "#fef5e7",
    "warn_cell":    "#fdebd0",
    "ng":           "#e74c3c",
    "ng_light":     "#fdedec",
    "ng_cell":      "#fadbd8",
    "nodata":       "#bdc3c7",
    "nodata_light": "#f4f6f7",
    "nodata_cell":  "#f2f3f4",
    "tree_hdr_bg":  "#1e3a5f",
    "tree_hdr_fg":  "#ffffff",
    "tree_sel":     "#d5e8f0",
    "grid_hdr":     "#2c3e50",
    "grid_hdr_fg":  "#ffffff",
    "grid_line_bg": "#f8f9fa",
    "link":         "#2471a3",
    "link_hover":   "#1a5276",
    "popup_bg":     "#ffffff",
    "overload_bg":  "#fadbd8",
    "overload_fg":  "#c0392b",
    "graph_target": "#3498db",
    "graph_transfer":"#e67e22",
    "graph_total":  "#e74c3c",
    "graph_line":   "#bdc3c7",
    "weekend_bg":   "#e8f0fe",
    "weekend_fg":   "#1a5276",
    "weekend_chart":"#dce8f5",
}

# ── 다크 테마 ──
DARK = {
    "bg":           "#1a1a2e",
    "card_bg":      "#16213e",
    "title_bg":     "#0f3460",
    "title_fg":     "#e0e0e0",
    "subtitle_fg":  "#7f8fa6",
    "text":         "#e0e0e0",
    "text_sub":     "#a0a0a0",
    "accent":       "#4fc3f7",
    "accent_hover": "#0288d1",
    "highlight":    "#f39c12",
    "highlight_bg": "#3d2e0f",
    "entry_bg":     "#1e2d45",
    "ok":           "#2ecc71",
    "ok_light":     "#1b4332",
    "ok_cell":      "#1b4332",
    "warn":         "#f39c12",
    "warn_light":   "#3d2e0f",
    "warn_cell":    "#3d2e0f",
    "ng":           "#ff6b6b",
    "ng_light":     "#4a1a1a",
    "ng_cell":      "#4a1a1a",
    "nodata":       "#636e72",
    "nodata_light": "#2d3436",
    "nodata_cell":  "#2d3436",
    "tree_hdr_bg":  "#0f3460",
    "tree_hdr_fg":  "#e0e0e0",
    "tree_sel":     "#1e3a5f",
    "grid_hdr":     "#0f3460",
    "grid_hdr_fg":  "#e0e0e0",
    "grid_line_bg": "#1a2540",
    "link":         "#4fc3f7",
    "link_hover":   "#81d4fa",
    "popup_bg":     "#16213e",
    "overload_bg":  "#4a1a1a",
    "overload_fg":  "#ff6b6b",
    "graph_target": "#4fc3f7",
    "graph_transfer":"#f39c12",
    "graph_total":  "#ff6b6b",
    "graph_line":   "#636e72",
    "weekend_bg":   "#1e2d45",
    "weekend_fg":   "#81d4fa",
    "weekend_chart":"#1a2540",
}


# ═══════════════════════════════════════════════════════════════
#  데이터 매니저 (자동 계산 엔진)
# ═══════════════════════════════════════════════════════════════
class DataManager:
    """
    Raw Data를 직접 읽어와 프로그램 내부에서 합산/판정하는 자동 계산 엔진.

    ■ 매핑 파일: {변전소: [(대상선로, 전환선로), ...]}
    ■ 사용량 파일: Master DataFrame
        columns = [변전소명, 회선명, 일자(YYYYMMDD), 1시~24시]
        12개 월별 파일을 하나로 병합
    """

    # 요일 한글 매핑
    WEEKDAY_KR = ("월", "화", "수", "목", "금", "토", "일")

    # 정규식: 한글·영문·숫자만 남기고 나머지 모두 제거
    _RE_CLEAN = re.compile(r'[^가-힣a-zA-Z0-9]')

    def __init__(self):
        self.substations: dict = {}          # {변전소: [(대상선로, 전환선로), ...]}
        self.master_df: pd.DataFrame = None  # 병합된 시간대별 사용량
        self._year: int = datetime.now().year

    @classmethod
    def _clean_text(cls, val) -> str:
        """
        정규식으로 한글·영문·숫자 외 모든 문자를 완전 제거.

        마스킹 기호(*, X, ○, ●), 일반 공백, 탭(\\t), 줄바꿈(\\n),
        영폭 공백(\\u200b), 전각 공백(\\u3000) 등 모두 제거.

        예: '월*촌'  → '월촌'
            '구 산 ' → '구산'
            '월\\t촌' → '월촌'
            '서*울 D/L' → '서울DL'
        """
        return cls._RE_CLEAN.sub('', str(val))

    # ─────────────────────────────────
    #  매핑 파일 로드
    # ─────────────────────────────────
    def load_mapping(self, filepath: str) -> tuple[bool, str]:
        """전환선로 매핑 엑셀을 파싱한다."""
        try:
            df = pd.read_excel(filepath, header=None)
        except Exception as e:
            return False, f"매핑 파일 열기 실패: {e}"

        try:
            self.substations = {}
            current_sub = None

            for i in range(len(df)):
                c0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ""
                c1 = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ""
                c3 = str(df.iloc[i, 3]).strip() if df.shape[1] > 3 and pd.notna(df.iloc[i, 3]) else ""

                # 변전소명 행 감지
                if c0 == "변전소명":
                    current_sub = self._clean_text(c1) if c1 else c1
                    continue
                # 헤더 행 건너뛰기
                if c0 in ("대상선로", ""):
                    continue
                # 데이터 행
                if current_sub and c0:
                    self.substations.setdefault(
                        self._clean_text(current_sub), []
                    ).append((self._clean_text(c0), self._clean_text(c3)))

            n_s = len(self.substations)
            n_p = sum(len(v) for v in self.substations.values())
            return True, f"{n_s}개 변전소, {n_p}개 선로 매핑 완료"
        except Exception as e:
            return False, f"매핑 파싱 오류: {e}"

    # ─────────────────────────────────
    #  사용량 폴더 로드
    # ─────────────────────────────────
    def load_usage_folder(self, folder_path: str) -> tuple[bool, str]:
        """
        1월~12월 시간대별 사용량 엑셀 파일들을 읽어 Master DataFrame으로 병합.

        각 파일 구조:
          - Row 0~2: 빈 행 (병합 헤더 등)
          - Row 3: [변전소명, 회선명, 일자, 일일 사용량, ...]
          - Row 4: [, , , 1시, 2시, ..., 24시]
          - Row 5~: 데이터
        """
        xlsx_files = sorted([
            f for f in os.listdir(folder_path)
            if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')
        ])

        if not xlsx_files:
            return False, "폴더에 엑셀 파일이 없습니다."

        frames = []
        errors = []
        col_names = ["변전소명", "회선명", "일자"] + [f"{h}시" for h in range(1, 25)]

        for fname in xlsx_files:
            fpath = os.path.join(folder_path, fname)
            try:
                # engine='calamine' 시도, 실패 시 openpyxl 폴백
                try:
                    df = pd.read_excel(fpath, header=None, skiprows=4, engine="calamine")
                except Exception:
                    df = pd.read_excel(fpath, header=None, skiprows=4)

                # 컬럼 수 확인 및 보정
                if df.shape[1] < 27:
                    # 부족한 컬럼 패딩
                    for _ in range(27 - df.shape[1]):
                        df[df.shape[1]] = None
                elif df.shape[1] > 27:
                    df = df.iloc[:, :27]

                df.columns = col_names

                # 헤더 행 제거 (skiprows 후에도 남을 수 있는 부제목 행)
                df = df[df["회선명"].notna()].copy()
                df = df[df["회선명"].apply(lambda x: str(x).strip() not in ("", "회선명"))].copy()

                # 변전소명 전방 채움 (병합셀 → NaN 처리)
                df["변전소명"] = df["변전소명"].ffill()
                df["변전소명"] = df["변전소명"].astype(str).str.strip().apply(self._clean_text)
                df["회선명"] = df["회선명"].astype(str).str.strip().apply(self._clean_text)

                # 일자를 문자열로 통일
                df["일자"] = df["일자"].apply(self._normalize_date)

                # 1시~24시 숫자 변환 (결측치 → 0)
                for h in range(1, 25):
                    col = f"{h}시"
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

                frames.append(df)
            except Exception as e:
                errors.append(f"{fname}: {e}")

        if not frames:
            return False, f"유효한 데이터 없음. 오류: {'; '.join(errors)}"

        self.master_df = pd.concat(frames, ignore_index=True)

        # ── 시계열 전처리 구조화 (예측 모델 대비) ──
        # 1) datetime 컬럼: pd.to_datetime으로 연도/요일 자동 인식
        self.master_df["날짜"] = pd.to_datetime(
            self.master_df["일자"], format="%Y%m%d", errors="coerce"
        )
        # 2) 요일 컬럼 (0=월 ~ 6=일)
        self.master_df["요일번호"] = self.master_df["날짜"].dt.weekday          # 0~6
        self.master_df["요일"] = self.master_df["요일번호"].map(
            lambda x: self.WEEKDAY_KR[int(x)] if pd.notna(x) else ""
        )
        # 3) 주말 여부 (토=5, 일=6)
        self.master_df["주말"] = self.master_df["요일번호"].isin([5, 6])
        # 4) 연/월/일 분리 (예측 모델 피처용)
        self.master_df["연"] = self.master_df["날짜"].dt.year
        self.master_df["월"] = self.master_df["날짜"].dt.month
        self.master_df["일"] = self.master_df["날짜"].dt.day

        # 연도 자동 감지 (datetime 기반)
        try:
            valid_years = self.master_df["연"].dropna()
            if not valid_years.empty:
                self._year = int(valid_years.mode().iloc[0])
        except Exception:
            try:
                sample_date = str(self.master_df["일자"].dropna().iloc[0])
                if len(sample_date) >= 4:
                    self._year = int(sample_date[:4])
            except Exception:
                pass

        # 5) 결측치 정리: datetime 변환 실패한 행 경고 (삭제하지 않음)
        n_nat = self.master_df["날짜"].isna().sum()

        msg = f"{len(xlsx_files)}개 파일, {len(self.master_df):,}건 로드 완료 (연도: {self._year})"
        if n_nat > 0:
            msg += f" | 날짜 변환 실패: {n_nat}건"
        if errors:
            msg += f" (오류 {len(errors)}건)"
        return True, msg

    @staticmethod
    def _normalize_date(val) -> str:
        """일자를 'YYYYMMDD' 문자열로 정규화."""
        if pd.isna(val):
            return ""
        if isinstance(val, (int, float)):
            return str(int(val))
        s = str(val).strip().replace("-", "").replace("/", "").replace(".", "")
        return s[:8] if len(s) >= 8 else s

    # ─────────────────────────────────
    #  조회 헬퍼
    # ─────────────────────────────────
    def get_substation_list(self) -> list[str]:
        return sorted(self.substations.keys())

    def get_target_lines(self, sub: str) -> list[str]:
        return [t for t, _ in self.substations.get(sub, [])]

    def get_transfer_line(self, sub: str, target: str) -> str:
        for t, tr in self.substations.get(sub, []):
            if t == target:
                return tr
        return ""

    def has_data(self) -> bool:
        return self.master_df is not None and len(self.master_df) > 0

    # ─────────────────────────────────
    #  핵심 계산 엔진
    # ─────────────────────────────────
    def calc_monthly_possible_days(self, sub: str, target: str, threshold: float = OVERLOAD_THRESHOLD) -> list:
        """
        월별 절체 가능 일수를 계산하여 [1월, 2월, ..., 12월] 리스트로 반환.
        데이터가 없으면 None.
        """
        if not self.has_data():
            return [None] * 12

        transfer = self.get_transfer_line(sub, target)
        if not transfer:
            return [None] * 12

        hour_cols = [f"{h}시" for h in range(1, 25)]

        # 대상선로 / 전환선로 데이터 필터
        df_target = self.master_df[
            (self.master_df["변전소명"] == sub) &
            (self.master_df["회선명"] == target)
        ].copy()

        df_transfer = self.master_df[
            (self.master_df["회선명"] == transfer)
        ].copy()

        if df_target.empty and df_transfer.empty:
            return [None] * 12

        results = []
        for month in range(1, 13):
            # 해당 월 문자열 패턴
            month_str = f"{self._year}{month:02d}"

            dt = df_target[df_target["일자"].str.startswith(month_str)]
            dtr = df_transfer[df_transfer["일자"].str.startswith(month_str)]

            if dt.empty and dtr.empty:
                results.append(None)
                continue

            # 날짜 기준 병합
            merged = pd.merge(
                dt[["일자"] + hour_cols],
                dtr[["일자"] + hour_cols],
                on="일자", how="outer",
                suffixes=("_t", "_tr"),
            )

            possible_count = 0
            for _, row in merged.iterrows():
                max_sum = 0.0
                for h in range(1, 25):
                    t_val = row.get(f"{h}시_t", 0.0)
                    tr_val = row.get(f"{h}시_tr", 0.0)
                    t_val = float(t_val) if pd.notna(t_val) else 0.0
                    tr_val = float(tr_val) if pd.notna(tr_val) else 0.0
                    s = t_val + tr_val
                    if s > max_sum:
                        max_sum = s
                if max_sum <= threshold:
                    possible_count += 1

            results.append(possible_count)

        return results

    def get_daily_detail(self, sub: str, target: str, month: int, threshold: float = OVERLOAD_THRESHOLD) -> list[dict]:
        """
        특정 월의 일별 상세 데이터를 반환.

        Returns: [{
            "day": int, "date_str": str,
            "target_hours": [24 float], "transfer_hours": [24 float],
            "sum_hours": [24 float],
            "target_max": float, "transfer_max": float,
            "sum_max": float, "possible": bool
        }, ...]
        """
        if not self.has_data():
            return []

        transfer = self.get_transfer_line(sub, target)
        hour_cols = [f"{h}시" for h in range(1, 25)]
        month_str = f"{self._year}{month:02d}"

        df_target = self.master_df[
            (self.master_df["변전소명"] == sub) &
            (self.master_df["회선명"] == target) &
            (self.master_df["일자"].str.startswith(month_str))
        ].copy()

        df_transfer = pd.DataFrame()
        if transfer:
            df_transfer = self.master_df[
                (self.master_df["회선명"] == transfer) &
                (self.master_df["일자"].str.startswith(month_str))
            ].copy()

        # 모든 일자 수집
        all_dates = set()
        if not df_target.empty:
            all_dates.update(df_target["일자"].tolist())
        if not df_transfer.empty:
            all_dates.update(df_transfer["일자"].tolist())

        if not all_dates:
            return []

        result = []
        for date_str in sorted(all_dates):
            t_row = df_target[df_target["일자"] == date_str]
            tr_row = df_transfer[df_transfer["일자"] == date_str] if not df_transfer.empty else pd.DataFrame()

            target_hours = []
            transfer_hours = []
            sum_hours = []

            for h in range(1, 25):
                col = f"{h}시"
                t_val = float(t_row[col].iloc[0]) if not t_row.empty and pd.notna(t_row[col].iloc[0]) else 0.0
                tr_val = float(tr_row[col].iloc[0]) if not tr_row.empty and pd.notna(tr_row[col].iloc[0]) else 0.0
                target_hours.append(round(t_val, 2))
                transfer_hours.append(round(tr_val, 2))
                sum_hours.append(round(t_val + tr_val, 2))

            sum_max = max(sum_hours) if sum_hours else 0.0

            # 일자에서 day 추출
            try:
                day = int(date_str[6:8]) if len(date_str) >= 8 else 0
            except ValueError:
                day = 0

            # 요일 자동 인식 (pd.to_datetime 활용, 연도 무관)
            try:
                dt_obj = pd.to_datetime(date_str, format="%Y%m%d")
                weekday_num = dt_obj.weekday()  # 0=월 ~ 6=일
                weekday_kr = self.WEEKDAY_KR[weekday_num]
                is_weekend = weekday_num >= 5   # 토(5), 일(6)
            except Exception:
                weekday_kr = ""
                is_weekend = False

            result.append({
                "day": day,
                "date_str": date_str,
                "weekday": weekday_kr,
                "is_weekend": is_weekend,
                "target_hours": target_hours,
                "transfer_hours": transfer_hours,
                "sum_hours": sum_hours,
                "target_max": max(target_hours) if target_hours else 0.0,
                "transfer_max": max(transfer_hours) if transfer_hours else 0.0,
                "sum_max": round(sum_max, 2),
                "possible": sum_max <= threshold,
            })

        return result

    def get_all_lines_data(self, sub: str) -> list[tuple[str, str, list]]:
        """변전소 종합 조회: 모든 선로의 월별 가능 일수."""
        lines = self.substations.get(sub, [])
        result = []
        for target, transfer in lines:
            monthly = self.calc_monthly_possible_days(sub, target)
            result.append((target, transfer, monthly))
        return result

    def get_month_actual_days(self, month: int) -> int:
        """month(1-indexed)에 해당하는 월의 실제 일수."""
        try:
            return calendar.monthrange(self._year, month)[1]
        except Exception:
            return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month - 1]

    # ─────────────────────────────────
    #  예측 모델용 데이터프레임 (미래 ML 대비)
    # ─────────────────────────────────
    def get_ml_ready_df(self, sub: str = None, line: str = None) -> pd.DataFrame:
        """
        시계열 예측 모델이 바로 학습할 수 있는 깔끔한 DataFrame 반환.

        구조:
          - 인덱스: DatetimeIndex (날짜)
          - 컬럼: 변전소명, 회선명, 1시~24시, 일최대, 요일, 요일번호, 주말, 연, 월, 일
          - 결측치: 0으로 채움
          - 정렬: 날짜 오름차순

        Parameters:
            sub: 특정 변전소만 필터 (None이면 전체)
            line: 특정 회선만 필터 (None이면 전체)
        """
        if not self.has_data():
            return pd.DataFrame()

        df = self.master_df.copy()

        # 필터
        if sub:
            df = df[df["변전소명"] == sub]
        if line:
            df = df[df["회선명"] == line]

        if df.empty:
            return pd.DataFrame()

        # datetime 파싱 실패 행 제거
        df = df.dropna(subset=["날짜"])

        # 일최대 컬럼 추가
        hour_cols = [f"{h}시" for h in range(1, 25)]
        df["일최대"] = df[hour_cols].max(axis=1)

        # 인덱스 설정 + 정렬
        df = df.set_index("날짜").sort_index()

        # 필요 컬럼만 유지
        keep_cols = ["변전소명", "회선명"] + hour_cols + \
                    ["일최대", "요일", "요일번호", "주말", "연", "월", "일"]
        df = df[[c for c in keep_cols if c in df.columns]]

        return df

    # ─────────────────────────────────
    #  연도별 하위 폴더 순회 로드 (AI 학습용)
    # ─────────────────────────────────
    def load_usage_multi_year(self, root_folder: str) -> tuple[bool, str]:
        """
        root_folder 아래 '2023년', '2024년' 등 연도 폴더를 순회하며
        각 폴더 안의 1월~12월.xlsx 파일을 모두 병합.
        기존 load_usage_folder 로직을 재활용한다.
        """
        year_dirs = []
        try:
            for name in sorted(os.listdir(root_folder)):
                full = os.path.join(root_folder, name)
                if os.path.isdir(full):
                    # '2023년', '2024', '2025년' 등 숫자가 포함된 폴더
                    digits = "".join(c for c in name if c.isdigit())
                    if len(digits) == 4:
                        year_dirs.append(full)
        except Exception as e:
            return False, f"폴더 탐색 실패: {e}"

        if not year_dirs:
            # 연도 폴더가 없으면 root 자체를 단일 폴더로 시도
            return self.load_usage_folder(root_folder)

        all_frames = []
        errors = []
        col_names = ["변전소명", "회선명", "일자"] + [f"{h}시" for h in range(1, 25)]

        for ydir in year_dirs:
            xlsx_files = sorted([
                f for f in os.listdir(ydir)
                if f.endswith(('.xlsx', '.xls')) and not f.startswith('~$')
            ])
            for fname in xlsx_files:
                fpath = os.path.join(ydir, fname)
                try:
                    try:
                        df = pd.read_excel(fpath, header=None, skiprows=4, engine="calamine")
                    except Exception:
                        df = pd.read_excel(fpath, header=None, skiprows=4)

                    if df.shape[1] < 27:
                        for _ in range(27 - df.shape[1]):
                            df[df.shape[1]] = None
                    elif df.shape[1] > 27:
                        df = df.iloc[:, :27]

                    df.columns = col_names
                    df = df[df["회선명"].notna()].copy()
                    df = df[df["회선명"].apply(lambda x: str(x).strip() not in ("", "회선명"))].copy()
                    df["변전소명"] = df["변전소명"].ffill()
                    df["변전소명"] = df["변전소명"].astype(str).str.strip().apply(self._clean_text)
                    df["회선명"] = df["회선명"].astype(str).str.strip().apply(self._clean_text)
                    df["일자"] = df["일자"].apply(self._normalize_date)
                    for h in range(1, 25):
                        col = f"{h}시"
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

                    all_frames.append(df)
                except Exception as e:
                    errors.append(f"{fname}: {e}")

        if not all_frames:
            return False, f"유효한 데이터 없음. 오류: {'; '.join(errors)}"

        self.master_df = pd.concat(all_frames, ignore_index=True)

        # 시계열 전처리 (기존과 동일)
        self.master_df["날짜"] = pd.to_datetime(
            self.master_df["일자"], format="%Y%m%d", errors="coerce"
        )
        self.master_df["요일번호"] = self.master_df["날짜"].dt.weekday
        self.master_df["요일"] = self.master_df["요일번호"].map(
            lambda x: self.WEEKDAY_KR[int(x)] if pd.notna(x) else ""
        )
        self.master_df["주말"] = self.master_df["요일번호"].isin([5, 6])
        self.master_df["연"] = self.master_df["날짜"].dt.year
        self.master_df["월"] = self.master_df["날짜"].dt.month
        self.master_df["일"] = self.master_df["날짜"].dt.day

        try:
            valid_years = self.master_df["연"].dropna()
            if not valid_years.empty:
                self._year = int(valid_years.max())
        except Exception:
            pass

        n_files = len(all_frames)
        n_years = len(year_dirs)
        msg = f"{n_years}개 연도, {n_files}개 파일, {len(self.master_df):,}건 로드 완료"
        if errors:
            msg += f" (오류 {len(errors)}건)"
        return True, msg


# ═══════════════════════════════════════════════════════════════
#  AI 부하량 예측 엔진 (XGBoost)
# ═══════════════════════════════════════════════════════════════
class LoadPredictor:
    """
    XGBoost 기반 익일 전력 부하량(1~24시) 예측 모델.

    ■ 피처: 월, 일, 요일, 주말 여부, 전일 1~24시 부하(Lag-1), 전일 일최대
    ■ 타겟: 당일 1~24시 부하 (24개 개별 모델 또는 MultiOutput)
    ■ 학습: 특정 회선의 전체 일별 데이터로 학습
    """

    HOUR_COLS = [f"{h}시" for h in range(1, 25)]

    def __init__(self):
        self.models: dict = {}       # {hour_col: XGBRegressor}
        self.is_trained: bool = False
        self.train_line: str = ""
        self.train_sub: str = ""
        self.last_date: pd.Timestamp = None
        self.train_df: pd.DataFrame = None   # 학습에 사용된 원본
        self._train_msg: str = ""

    def train(self, dm: DataManager, sub: str, line: str) -> tuple[bool, str]:
        """특정 변전소/회선의 데이터로 24개 XGBoost 모델 학습."""
        self.is_trained = False
        self.models = {}

        df = dm.get_ml_ready_df(sub, line)
        if df.empty or len(df) < 14:
            # 디버깅용 상세 메시지: 왜 데이터가 부족한지 표시
            n_found = len(df)
            # master_df에서 해당 변전소의 회선명 목록 추출 (유사 이름 힌트)
            hint = ""
            if dm.has_data():
                if sub:
                    sub_df = dm.master_df[dm.master_df["변전소명"] == sub]
                    available = sorted(sub_df["회선명"].unique().tolist()) if not sub_df.empty else []
                    if available:
                        hint = f"\n\n[참고] '{sub}' 변전소에 존재하는 회선명 목록:\n  " + ", ".join(available[:20])
                        if len(available) > 20:
                            hint += f" ... 외 {len(available)-20}개"
                    else:
                        hint = f"\n\n[참고] '{sub}' 변전소와 일치하는 데이터가 없습니다."
                else:
                    available = sorted(dm.master_df["회선명"].unique().tolist())
                    if available:
                        hint = f"\n\n[참고] 전체 데이터에 존재하는 회선명 목록:\n  " + ", ".join(available[:20])
                        if len(available) > 20:
                            hint += f" ... 외 {len(available)-20}개"
            sub_label = f"'{sub}'" if sub else "(전체)"
            return False, (
                f"학습 데이터 부족!\n"
                f"- 변전소: {sub_label}\n"
                f"- 회선명: '{line}'\n"
                f"- 검색된 데이터: {n_found}건 (최소 14일 필요)"
                f"{hint}"
            )

        # 일별 1행으로 정리 (중복 날짜 시 평균)
        df = df.groupby(df.index).mean(numeric_only=True)
        df = df.sort_index()

        # Lag 피처 생성 (전일 부하)
        for hc in self.HOUR_COLS:
            df[f"lag1_{hc}"] = df[hc].shift(1)
        df["lag1_일최대"] = df["일최대"].shift(1)

        # 첫 행(lag 없음) 제거
        df = df.dropna(subset=[f"lag1_1시"])

        # 피처 컬럼
        feature_cols = ["월", "일", "요일번호"] + \
                       [f"lag1_{hc}" for hc in self.HOUR_COLS] + \
                       ["lag1_일최대"]
        # 주말 피처 (bool → int)
        if "주말" in df.columns:
            df["주말_int"] = df["주말"].astype(int)
            feature_cols.append("주말_int")

        X = df[feature_cols].values
        self.train_df = df
        self.last_date = df.index.max()
        self.train_line = line
        self.train_sub = sub

        # 24개 시간대별 모델 학습
        for hc in self.HOUR_COLS:
            y = df[hc].values
            model = XGBRegressor(
                n_estimators=200,
                max_depth=5,
                learning_rate=0.08,
                subsample=0.8,
                colsample_bytree=0.8,
                random_state=42,
                verbosity=0,
            )
            model.fit(X, y)
            self.models[hc] = model

        self.is_trained = True
        n_days = len(df)
        date_range = f"{df.index.min().strftime('%Y-%m-%d')} ~ {df.index.max().strftime('%Y-%m-%d')}"
        self._train_msg = f"{sub}/{line} | {n_days}일 학습 완료 ({date_range})"
        return True, self._train_msg

    def predict_next_day(self) -> dict | None:
        """학습된 모델로 마지막 날짜 다음 날 1~24시 부하 예측."""
        if not self.is_trained or self.train_df is None:
            return None

        df = self.train_df
        last_row = df.iloc[-1]
        next_date = self.last_date + timedelta(days=1)

        # 피처 구성 (다음 날 예측)
        features = {
            "월": next_date.month,
            "일": next_date.day,
            "요일번호": next_date.weekday(),
        }
        # lag = 오늘(마지막 날) 부하
        for hc in self.HOUR_COLS:
            features[f"lag1_{hc}"] = last_row[hc]
        features["lag1_일최대"] = last_row["일최대"]
        if "주말_int" in df.columns:
            features["주말_int"] = 1 if next_date.weekday() >= 5 else 0

        feature_cols = ["월", "일", "요일번호"] + \
                       [f"lag1_{hc}" for hc in self.HOUR_COLS] + \
                       ["lag1_일최대"]
        if "주말_int" in df.columns:
            feature_cols.append("주말_int")

        X_pred = np.array([[features[c] for c in feature_cols]])

        pred_hours = []
        for hc in self.HOUR_COLS:
            val = float(self.models[hc].predict(X_pred)[0])
            pred_hours.append(round(max(val, 0.0), 2))

        weekday_kr = DataManager.WEEKDAY_KR[next_date.weekday()]
        return {
            "date": next_date,
            "date_str": next_date.strftime("%Y%m%d"),
            "weekday": weekday_kr,
            "is_weekend": next_date.weekday() >= 5,
            "pred_hours": pred_hours,
            "pred_max": round(max(pred_hours), 2),
        }

    def predict_date(self, target_date: pd.Timestamp) -> dict | None:
        """임의의 미래 날짜에 대해 예측 (반복 예측으로 중간 lag 전파)."""
        if not self.is_trained or self.train_df is None:
            return None

        df = self.train_df
        current_last = self.last_date
        current_hours = {hc: df.iloc[-1][hc] for hc in self.HOUR_COLS}
        current_max = df.iloc[-1]["일최대"]

        feature_cols = ["월", "일", "요일번호"] + \
                       [f"lag1_{hc}" for hc in self.HOUR_COLS] + \
                       ["lag1_일최대"]
        if "주말_int" in df.columns:
            feature_cols.append("주말_int")

        # 하루씩 전파하며 예측
        cursor = current_last + timedelta(days=1)
        while cursor <= target_date:
            features = {
                "월": cursor.month,
                "일": cursor.day,
                "요일번호": cursor.weekday(),
            }
            for hc in self.HOUR_COLS:
                features[f"lag1_{hc}"] = current_hours[hc]
            features["lag1_일최대"] = current_max
            if "주말_int" in df.columns:
                features["주말_int"] = 1 if cursor.weekday() >= 5 else 0

            X_pred = np.array([[features[c] for c in feature_cols]])
            pred = {}
            for hc in self.HOUR_COLS:
                val = float(self.models[hc].predict(X_pred)[0])
                pred[hc] = round(max(val, 0.0), 2)

            current_hours = pred
            current_max = max(pred.values())
            cursor += timedelta(days=1)

        weekday_kr = DataManager.WEEKDAY_KR[target_date.weekday()]
        pred_hours = [current_hours[hc] for hc in self.HOUR_COLS]
        return {
            "date": target_date,
            "date_str": target_date.strftime("%Y%m%d"),
            "weekday": weekday_kr,
            "is_weekend": target_date.weekday() >= 5,
            "pred_hours": pred_hours,
            "pred_max": round(max(pred_hours), 2),
        }

    def predict_year(self, target_year: int) -> dict | None:
        """
        target_year 전체 365(또는 366)일을 한 번에 효율적으로 예측.

        마지막 학습일 → target_year 12/31까지 순차 roll-forward 하되,
        target_year에 해당하는 날짜만 결과 dict에 저장.

        Returns: {
            "YYYYMMDD": {
                "pred_hours": [24 floats],
                "pred_max": float,
                "weekday": str,
                "is_weekend": bool,
            }, ...
        }
        """
        if not self.is_trained or self.train_df is None:
            return None

        df = self.train_df
        current_hours = {hc: df.iloc[-1][hc] for hc in self.HOUR_COLS}
        current_max = df.iloc[-1]["일최대"]

        feature_cols = ["월", "일", "요일번호"] + \
                       [f"lag1_{hc}" for hc in self.HOUR_COLS] + \
                       ["lag1_일최대"]
        has_weekend = "주말_int" in df.columns
        if has_weekend:
            feature_cols.append("주말_int")

        # target_year의 시작~끝
        year_start = pd.Timestamp(target_year, 1, 1)
        year_end = pd.Timestamp(target_year, 12, 31)

        # 학습 마지막 날짜 다음날부터 roll-forward
        cursor = self.last_date + timedelta(days=1)
        results = {}

        while cursor <= year_end:
            features = {
                "월": cursor.month,
                "일": cursor.day,
                "요일번호": cursor.weekday(),
            }
            for hc in self.HOUR_COLS:
                features[f"lag1_{hc}"] = current_hours[hc]
            features["lag1_일최대"] = current_max
            if has_weekend:
                features["주말_int"] = 1 if cursor.weekday() >= 5 else 0

            X_pred = np.array([[features[c] for c in feature_cols]])
            pred = {}
            for hc in self.HOUR_COLS:
                val = float(self.models[hc].predict(X_pred)[0])
                pred[hc] = round(max(val, 0.0), 2)

            current_hours = pred
            current_max = max(pred.values())

            # target_year에 해당하는 날짜만 저장
            if cursor >= year_start:
                date_str = cursor.strftime("%Y%m%d")
                weekday_kr = DataManager.WEEKDAY_KR[cursor.weekday()]
                pred_hours = [pred[hc] for hc in self.HOUR_COLS]
                results[date_str] = {
                    "pred_hours": pred_hours,
                    "pred_max": round(max(pred_hours), 2),
                    "weekday": weekday_kr,
                    "is_weekend": cursor.weekday() >= 5,
                }

            cursor += timedelta(days=1)

        return results if results else None


# ═══════════════════════════════════════════════════════════════
#  색상 판정 유틸리티
# ═══════════════════════════════════════════════════════════════
def get_level_info(days, C: dict) -> dict:
    if days is None:
        return {"border": C["nodata"], "bg": C["nodata_light"], "cell": C["nodata_cell"],
                "color": C["nodata"], "badge": C["nodata"],
                "icon": "—", "level": "데이터 없음", "text": "—"}
    if days >= LEVEL_HIGH:
        return {"border": C["ok"], "bg": C["ok_light"], "cell": C["ok_cell"],
                "color": C["ok"], "badge": C["ok"],
                "icon": "O", "level": "유력", "text": f"{days}"}
    if days >= LEVEL_MID:
        return {"border": C["warn"], "bg": C["warn_light"], "cell": C["warn_cell"],
                "color": C["warn"], "badge": C["warn"],
                "icon": "△", "level": "고려", "text": f"{days}"}
    return {"border": C["ng"], "bg": C["ng_light"], "cell": C["ng_cell"],
            "color": C["ng"], "badge": C["ng"],
            "icon": "X", "level": "불가", "text": f"{days}"}


# ═══════════════════════════════════════════════════════════════
#  메인 GUI 애플리케이션
# ═══════════════════════════════════════════════════════════════
class App(ctk.CTk):

    def __init__(self):
        super().__init__()
        self.dm = DataManager()
        self.predictor = LoadPredictor()
        self.threshold = OVERLOAD_THRESHOLD  # 사용자 설정 가능 임계값
        self._pred_cache = {}  # {(sub, target): {year, target_preds, transfer_preds}}
        self._cache_file_path = ""  # pkl 캐시 파일 경로
        self._cached_target_year = 0  # 캐시된 예측 대상 연도
        self._data_folder = ""  # 데이터 폴더 경로
        self.is_dark = False
        self.C = LIGHT
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self._current_sub = ""
        self._current_target = ""
        self._current_transfer = ""
        self._setup_window()
        self._build_ui()

    def _setup_window(self):
        self.title("배전선로 휴전 가능월 검토 프로그램 v9 — AI 예측")
        self.geometry("1340x920")
        self.minsize(1140, 820)
        self.configure(fg_color=self.C["bg"])
        self.update_idletasks()
        x = (self.winfo_screenwidth() - 1340) // 2
        y = (self.winfo_screenheight() - 920) // 2
        self.geometry(f"+{x}+{y}")

    # ──────────────────────────────────
    #  UI 전체 구성
    # ──────────────────────────────────
    def _build_ui(self):
        C = self.C

        # ═══ 타이틀 바 ═══
        title_bar = ctk.CTkFrame(self, fg_color=C["title_bg"], corner_radius=0, height=50)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)

        ctk.CTkLabel(
            title_bar, text="  배전선로 휴전 가능월 검토",
            font=ctk.CTkFont(family=FONT_FAMILY, size=18, weight="bold"),
            text_color=C["title_fg"],
        ).pack(side="left", padx=14)

        ctk.CTkLabel(
            title_bar,
            text=f"유력({LEVEL_HIGH}일+)  |  고려({LEVEL_MID}~{LEVEL_HIGH-1}일)  |  불가({LEVEL_MID-1}일-)",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=C["subtitle_fg"],
        ).pack(side="left", padx=(14, 0))

        self.theme_btn = ctk.CTkButton(
            title_bar, text="🌙 다크모드", width=110, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            fg_color="#4a6a8a", hover_color="#5a7a9a",
            text_color=C["title_fg"], border_width=0,
            command=self._toggle_theme,
        )
        self.theme_btn.pack(side="right", padx=14)

        self.retrain_btn = ctk.CTkButton(
            title_bar, text="🔄 AI 재학습 (데이터 갱신)", width=200, height=28,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            fg_color="#c0392b", hover_color="#e74c3c",
            text_color="#ffffff", border_width=0,
            command=self._on_retrain_all, state="disabled",
        )
        self.retrain_btn.pack(side="right", padx=(0, 6))

        # ═══ 데이터 로드 카드 ═══
        load_card = ctk.CTkFrame(self, fg_color=C["card_bg"], corner_radius=10)
        load_card.pack(fill="x", padx=14, pady=(8, 4))

        ctk.CTkLabel(
            load_card, text="  데이터 로드",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=C["text"],
        ).grid(row=0, column=0, padx=14, pady=(8, 2), sticky="w", columnspan=3)

        # ① 전환선로 매핑 파일
        ctk.CTkLabel(
            load_card, text="전환선로 매핑 파일 :",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12), text_color=C["text"],
        ).grid(row=1, column=0, padx=(14, 6), pady=4, sticky="e")

        self.mapping_var = ctk.StringVar(value="파일을 선택하세요")
        ctk.CTkEntry(
            load_card, textvariable=self.mapping_var, width=620,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            state="readonly", fg_color=C["entry_bg"], text_color=C["text"],
        ).grid(row=1, column=1, padx=4, pady=4, sticky="w")

        ctk.CTkButton(
            load_card, text="📂 파일 선택", width=130, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            command=self._on_select_mapping,
        ).grid(row=1, column=2, padx=8, pady=4)

        # ② 과거 부하량 데이터 폴더
        ctk.CTkLabel(
            load_card, text="과거 부하량 데이터 폴더 :",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12), text_color=C["text"],
        ).grid(row=2, column=0, padx=(14, 6), pady=4, sticky="e")

        self.usage_var = ctk.StringVar(value="폴더를 선택하세요 (연도별 하위 폴더 또는 단일 폴더)")
        ctk.CTkEntry(
            load_card, textvariable=self.usage_var, width=620,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            state="readonly", fg_color=C["entry_bg"], text_color=C["text"],
        ).grid(row=2, column=1, padx=4, pady=4, sticky="w")

        ctk.CTkButton(
            load_card, text="📁 폴더 선택", width=130, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            command=self._on_select_usage_folder,
        ).grid(row=2, column=2, padx=8, pady=4)

        # ③ 판정 기준값 (임계값)
        ctk.CTkLabel(
            load_card, text="판정 기준값 (합산 부하) :",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12), text_color=C["text"],
        ).grid(row=3, column=0, padx=(14, 6), pady=4, sticky="e")

        threshold_fr = tk.Frame(load_card, bg=C["card_bg"])
        threshold_fr.grid(row=3, column=1, padx=4, pady=4, sticky="w")

        self.threshold_var = ctk.StringVar(value=str(self.threshold))
        self.threshold_entry = ctk.CTkEntry(
            threshold_fr, textvariable=self.threshold_var, width=100,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            fg_color=C["entry_bg"], text_color=C["text"],
        )
        self.threshold_entry.pack(side="left")

        ctk.CTkLabel(
            threshold_fr,
            text="  (합산 최대 부하 ≤ 기준값 → 가능(O), 초과 → 불가(X))",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=C["text_sub"],
        ).pack(side="left", padx=(8, 0))

        # 로드 상태
        self.status_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            load_card, textvariable=self.status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=C["ok"],
        ).grid(row=4, column=0, columnspan=3, padx=14, pady=(0, 6), sticky="w")

        load_card.columnconfigure(1, weight=1)

        # ═══ CTkTabview ═══
        self.tabview = ctk.CTkTabview(
            self, fg_color=C["card_bg"], corner_radius=10,
            segmented_button_fg_color=C["entry_bg"],
            segmented_button_selected_color=C["accent"],
            segmented_button_unselected_color=C["entry_bg"],
            segmented_button_selected_hover_color=C["accent_hover"],
        )
        self.tabview.pack(fill="both", expand=True, padx=14, pady=(4, 12))

        self.tab2 = self.tabview.add("변전소 종합 조회")
        self.tab1 = self.tabview.add("단일 선로 상세 조회")

        self._build_tab2(self.tab2)
        self._build_tab1(self.tab1)

    # ──────────────────────────────────
    #  탭 1: 단일 선로 상세 조회
    # ──────────────────────────────────
    def _build_tab1(self, parent):
        C = self.C

        filter_fr = ctk.CTkFrame(parent, fg_color=C["card_bg"], corner_radius=8)
        filter_fr.pack(fill="x", padx=6, pady=(6, 4))

        ctk.CTkLabel(
            filter_fr, text="  조건 선택",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=C["text"],
        ).grid(row=0, column=0, padx=10, pady=(8, 4), sticky="w", columnspan=8)

        ctk.CTkLabel(filter_fr, text="변전소",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                     text_color=C["text"],
        ).grid(row=1, column=0, padx=(10, 4), pady=8, sticky="e")

        self.t1_sub_combo = ctk.CTkComboBox(
            filter_fr, width=165, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            dropdown_font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            values=["데이터를 먼저 로드하세요"],
            command=self._t1_on_sub, state="readonly",
        )
        self.t1_sub_combo.grid(row=1, column=1, padx=4, pady=8, sticky="w")

        ctk.CTkLabel(filter_fr, text="휴전선로(대상)",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                     text_color=C["text"],
        ).grid(row=1, column=2, padx=(16, 4), pady=8, sticky="e")

        self.t1_target_combo = ctk.CTkComboBox(
            filter_fr, width=165, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            dropdown_font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            values=["변전소를 선택하세요"],
            command=self._t1_on_target, state="readonly",
        )
        self.t1_target_combo.grid(row=1, column=3, padx=4, pady=8, sticky="w")

        ctk.CTkLabel(filter_fr, text="전환선로",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                     text_color=C["text"],
        ).grid(row=1, column=4, padx=(16, 4), pady=8, sticky="e")

        self.t1_transfer_var = ctk.StringVar(value="—")
        ctk.CTkLabel(
            filter_fr, textvariable=self.t1_transfer_var, width=120,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=C["highlight"], fg_color=C["highlight_bg"], corner_radius=6,
        ).grid(row=1, column=5, padx=4, pady=8, sticky="w")

        ctk.CTkButton(
            filter_fr, text="  결과 조회", width=130, height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=C["accent"], hover_color=C["accent_hover"],
            command=self._t1_on_run,
        ).grid(row=1, column=6, padx=(20, 10), pady=8)

        filter_fr.columnconfigure(7, weight=1)

        # ── 결과 영역 ──
        result_fr = ctk.CTkFrame(parent, fg_color=C["card_bg"], corner_radius=8)
        result_fr.pack(fill="both", expand=True, padx=6, pady=(4, 6))

        leg_row = tk.Frame(result_fr, bg=C["card_bg"])
        leg_row.pack(fill="x", padx=12, pady=(8, 4))

        for color, label in [
            (C["ok"], "유력"), (C["warn"], "고려"), (C["ng"], "불가"), (C["nodata"], "없음"),
        ]:
            box = tk.Frame(leg_row, bg=color, width=12, height=12)
            box.pack(side="left", padx=(0, 3))
            box.pack_propagate(False)
            tk.Label(leg_row, text=label, font=(FONT_FAMILY, 10),
                     bg=C["card_bg"], fg=C["text_sub"]).pack(side="left", padx=(0, 14))

        self.t1_summary_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            leg_row, textvariable=self.t1_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=C["text_sub"],
        ).pack(side="right", padx=4)

        self.t1_dashboard = ctk.CTkFrame(result_fr, fg_color=C["card_bg"])
        self.t1_dashboard.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        ctk.CTkLabel(
            self.t1_dashboard,
            text="변전소와 휴전선로를 선택한 후  [ 결과 조회 ]  버튼을 눌러주세요.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13), text_color=C["text_sub"],
        ).place(relx=0.5, rely=0.5, anchor="center")

    # ──────────────────────────────────
    #  탭 2: 변전소 종합 조회
    # ──────────────────────────────────
    def _build_tab2(self, parent):
        C = self.C

        filter_fr = ctk.CTkFrame(parent, fg_color=C["card_bg"], corner_radius=8)
        filter_fr.pack(fill="x", padx=6, pady=(6, 4))

        ctk.CTkLabel(
            filter_fr, text="  변전소 종합 조회",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=C["text"],
        ).grid(row=0, column=0, padx=10, pady=(8, 4), sticky="w", columnspan=4)

        ctk.CTkLabel(filter_fr, text="변전소",
                     font=ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
                     text_color=C["text"],
        ).grid(row=1, column=0, padx=(10, 4), pady=8, sticky="e")

        self.t2_sub_combo = ctk.CTkComboBox(
            filter_fr, width=175, height=34,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            dropdown_font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            values=["데이터를 먼저 로드하세요"],
            state="readonly",
        )
        self.t2_sub_combo.grid(row=1, column=1, padx=4, pady=8, sticky="w")

        ctk.CTkButton(
            filter_fr, text="  전체 조회", width=130, height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=C["accent"], hover_color=C["accent_hover"],
            command=self._t2_on_run,
        ).grid(row=1, column=2, padx=(20, 10), pady=8)

        self.t2_summary_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            filter_fr, textvariable=self.t2_summary_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=C["text_sub"],
        ).grid(row=1, column=3, padx=10, pady=8, sticky="w")

        filter_fr.columnconfigure(3, weight=1)

        # ── 범례 ──
        leg = ctk.CTkFrame(parent, fg_color=C["card_bg"], corner_radius=0)
        leg.pack(fill="x", padx=12, pady=(4, 2))

        for color, label in [
            (C["ok"], f"유력({LEVEL_HIGH}일+)"),
            (C["warn"], f"고려({LEVEL_MID}~{LEVEL_HIGH-1}일)"),
            (C["ng"], f"불가({LEVEL_MID-1}일-)"),
            (C["nodata"], "없음"),
        ]:
            box = tk.Frame(leg, bg=color, width=12, height=12)
            box.pack(side="left", padx=(0, 3))
            box.pack_propagate(False)
            tk.Label(leg, text=label, font=(FONT_FAMILY, 10),
                     bg=C["card_bg"], fg=C["text_sub"]).pack(side="left", padx=(0, 14))

        tk.Label(leg, text="※ 대상선로를 클릭하면 상세 조회로 이동합니다",
                 font=(FONT_FAMILY, 10), bg=C["card_bg"], fg=C["accent"],
        ).pack(side="right", padx=8)

        # ── 고정 헤더 ──
        self.t2_header_fr = tk.Frame(parent, bg=C["card_bg"])
        self.t2_header_fr.pack(fill="x", padx=6, pady=(4, 0))

        # ── 스크롤 데이터 ──
        self.t2_scroll = ctk.CTkScrollableFrame(
            parent, fg_color=C["card_bg"], corner_radius=8,
            label_text="", label_fg_color=C["card_bg"],
        )
        self.t2_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        self.t2_placeholder = ctk.CTkLabel(
            self.t2_scroll,
            text="변전소를 선택한 후  [ 전체 조회 ]  버튼을 눌러주세요.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13), text_color=C["text_sub"],
        )
        self.t2_placeholder.pack(pady=40)

    # ══════════════════════════════════
    #  데이터 로드 이벤트
    # ══════════════════════════════════
    def _on_select_mapping(self):
        path = filedialog.askopenfilename(
            title="전환선로 매핑 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if not path:
            return
        self.mapping_var.set(path)
        ok, msg = self.dm.load_mapping(path)

        if ok:
            self.status_var.set(f"  매핑: {msg}")
            subs = self.dm.get_substation_list()
            self.t1_sub_combo.configure(values=subs)
            if subs:
                self.t1_sub_combo.set(subs[0])
                self._t1_on_sub(subs[0])
            self.t2_sub_combo.configure(values=subs)
            if subs:
                self.t2_sub_combo.set(subs[0])
        else:
            self.status_var.set(f"  오류: {msg}")
            messagebox.showerror("매핑 로드 오류", msg)

    def _on_select_usage_folder(self):
        folder = filedialog.askdirectory(title="과거 부하량 데이터 폴더 선택")
        if not folder:
            return
        self.usage_var.set(folder)
        self._data_folder = folder
        self._cache_file_path = os.path.join(folder, "ai_total_prediction_cache.pkl")
        self.retrain_btn.configure(state="normal")

        # pkl 캐시 존재 확인 → 있으면 즉시 로드 (엑셀/학습 전면 생략)
        if os.path.exists(self._cache_file_path) and self.dm.substations:
            self.status_var.set("  캐시 파일 로딩 중...")
            self.update_idletasks()
            if self._load_cache_file():
                n_keys = len(self._pred_cache)
                self.status_var.set(
                    f"  캐시 로드 완료 (초고속 모드) | {self._cached_target_year}년 예측 | {n_keys}개 선로")
                return

        # pkl 없음 → 기존 흐름: 엑셀 로드
        self.status_var.set("  데이터 로딩 중...")
        self.update_idletasks()
        ok, msg = self.dm.load_usage_multi_year(folder)
        if not ok:
            self.status_var.set(f"  오류: {msg}")
            messagebox.showerror("데이터 로드 오류", msg)
            return

        # 매핑 로드 완료 상태면 자동 일괄 학습
        if self.dm.substations:
            self.status_var.set(f"  데이터: {msg} | 전체 선로 일괄 학습을 시작합니다...")
            self.update_idletasks()
            self._batch_train_all()
        else:
            self.status_var.set(f"  데이터: {msg} (매핑 파일 로드 후 [AI 재학습] 버튼을 눌러주세요)")

    def _get_threshold(self) -> float:
        """임계값 Entry에서 현재 값을 읽어 반환."""
        try:
            val = float(self.threshold_var.get())
            if val <= 0:
                raise ValueError
            self.threshold = val
            return val
        except (ValueError, AttributeError):
            self.threshold = OVERLOAD_THRESHOLD
            return OVERLOAD_THRESHOLD

    # ──────────────────────────────────
    #  pkl 캐시 로드 / 저장 / 일괄 학습
    # ──────────────────────────────────
    def _load_cache_file(self) -> bool:
        """pkl 캐시 파일 로드. 성공하면 True."""
        if not self._cache_file_path or not os.path.exists(self._cache_file_path):
            return False
        try:
            with open(self._cache_file_path, "rb") as f:
                cache_data = pickle.load(f)
            self._pred_cache = cache_data.get("predictions", {})
            self._cached_target_year = cache_data.get("target_year", 0)
            return True
        except Exception as e:
            print(f"캐시 로드 실패: {e}")
            return False

    def _save_cache_file(self):
        """현재 _pred_cache를 pkl 파일로 저장."""
        if not self._cache_file_path:
            return
        cache_data = {
            "version": 1,
            "target_year": self._cached_target_year,
            "data_folder": self._data_folder,
            "predictions": self._pred_cache,
        }
        try:
            with open(self._cache_file_path, "wb") as f:
                pickle.dump(cache_data, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as e:
            messagebox.showerror("캐시 저장 오류", f"pkl 파일 저장 실패:\n{e}")

    def _batch_train_all(self):
        """전체 선로 일괄 XGBoost 학습 + 예측 → pkl 저장."""
        if not self.dm.substations:
            messagebox.showwarning("알림", "전환선로 매핑 파일을 먼저 로드해주세요.")
            return
        if not self.dm.has_data():
            messagebox.showwarning("알림", "과거 부하량 데이터를 먼저 로드해주세요.")
            return

        # 전체 선로 쌍 수집
        all_pairs = []
        for sub, lines in self.dm.substations.items():
            for target, transfer in lines:
                all_pairs.append((sub, target, transfer))
        total = len(all_pairs)
        if total == 0:
            return

        target_year = self.dm._year + 1
        self._cached_target_year = target_year
        self._pred_cache = {}

        # Progress Bar 팝업
        prog_win = ctk.CTkToplevel(self)
        prog_win.title("AI 전체 선로 일괄 학습")
        prog_win.geometry("480x170")
        prog_win.resizable(False, False)
        prog_win.attributes("-topmost", True)
        prog_win.grab_set()
        # 화면 중앙
        prog_win.update_idletasks()
        sw = prog_win.winfo_screenwidth()
        sh = prog_win.winfo_screenheight()
        x = (sw - 480) // 2
        y = (sh - 170) // 2
        prog_win.geometry(f"480x170+{x}+{y}")

        ctk.CTkLabel(
            prog_win, text="AI 전체 선로 일괄 학습 중...",
            font=ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
        ).pack(pady=(18, 6))

        prog_label = ctk.CTkLabel(
            prog_win, text=f"준비 중... (0/{total})",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
        )
        prog_label.pack(pady=(0, 8))

        prog_bar = ctk.CTkProgressBar(prog_win, width=400, height=18)
        prog_bar.pack(pady=(0, 6))
        prog_bar.set(0)

        fail_label = ctk.CTkLabel(
            prog_win, text="",
            font=ctk.CTkFont(family=FONT_FAMILY, size=10),
            text_color="#e74c3c",
        )
        fail_label.pack(pady=(0, 4))

        self.update_idletasks()

        success_count = 0
        fail_count = 0

        for idx, (sub, target, transfer) in enumerate(all_pairs):
            prog_label.configure(text=f"연산 중... {sub}/{target} ({idx+1}/{total})")
            prog_bar.set((idx) / total)
            self.update_idletasks()

            # _train_and_predict_year 내부 로직 직접 실행 (캐시 체크 포함)
            result = self._train_and_predict_year(sub, target)
            if result:
                success_count += 1
            else:
                fail_count += 1
                fail_label.configure(text=f"실패: {fail_count}건")

        prog_bar.set(1.0)
        prog_label.configure(text=f"완료! 성공 {success_count} / 실패 {fail_count} (총 {total})")
        self.update_idletasks()

        # pkl 저장
        self._save_cache_file()

        # 잠시 후 팝업 닫기
        prog_win.after(1200, prog_win.destroy)

        self.status_var.set(
            f"  AI 일괄 학습 완료 | {target_year}년 예측 | "
            f"성공 {success_count} / 실패 {fail_count} (총 {total}) | 캐시 저장됨")

    def _on_retrain_all(self):
        """기존 캐시 무시, 엑셀 재로드 + 전체 일괄 재학습."""
        if not self.dm.substations:
            messagebox.showwarning("알림", "전환선로 매핑 파일을 먼저 로드해주세요.")
            return
        if not self._data_folder:
            messagebox.showwarning("알림", "데이터 폴더를 먼저 선택해주세요.")
            return

        # 기존 캐시 삭제
        if self._cache_file_path and os.path.exists(self._cache_file_path):
            try:
                os.remove(self._cache_file_path)
            except Exception:
                pass
        self._pred_cache = {}

        # 엑셀 데이터 (재)로드
        self.status_var.set("  데이터 재로딩 중...")
        self.update_idletasks()
        ok, msg = self.dm.load_usage_multi_year(self._data_folder)
        if not ok:
            self.status_var.set(f"  오류: {msg}")
            messagebox.showerror("데이터 로드 오류", msg)
            return

        self.status_var.set(f"  데이터: {msg} | 전체 선로 일괄 재학습을 시작합니다...")
        self.update_idletasks()
        self._batch_train_all()

    def _train_and_predict_year(self, sub: str, target: str) -> dict | None:
        """대상선로 + 전환선로 각각 XGBoost 학습 → Target Year 전체 예측."""
        if not self.dm.has_data():
            messagebox.showwarning("알림", "과거 부하량 데이터를 먼저 로드해주세요.")
            return None

        transfer = self.dm.get_transfer_line(sub, target)
        if not transfer:
            messagebox.showwarning("알림",
                f"'{target}'의 전환선로를 찾을 수 없습니다.\n\n"
                f"매핑 파일에서 '{sub}' 변전소의 선로 매핑을 확인해 주세요.")
            return None

        last_year = self.dm._year
        target_year = last_year + 1

        # 캐시 확인
        cache_key = (sub, target)
        if cache_key in self._pred_cache and self._pred_cache[cache_key]["year"] == target_year:
            return self._pred_cache[cache_key]

        # ── 학습 전 데이터 존재 여부 사전 점검 ──
        target_df = self.dm.master_df[
            (self.dm.master_df["변전소명"] == sub) &
            (self.dm.master_df["회선명"] == target)
        ]
        transfer_df = self.dm.master_df[
            (self.dm.master_df["회선명"] == transfer)
        ]
        if target_df.empty or transfer_df.empty:
            detail = (
                f"데이터 누락 발생!\n\n"
                f"- 대상선로('{target}'): {len(target_df)}건\n"
                f"- 전환선로('{transfer}'): {len(transfer_df)}건\n\n"
                f"둘 중 하나라도 데이터가 0건이면 예측할 수 없습니다.\n"
                f"엑셀 파일의 회선명을 확인해 주세요."
            )
            # 유사 이름 힌트
            if target_df.empty:
                sub_lines = self.dm.master_df[self.dm.master_df["변전소명"] == sub]
                if not sub_lines.empty:
                    available = sorted(sub_lines["회선명"].unique().tolist())
                    detail += f"\n\n['{sub}' 변전소 회선명 목록]\n  " + ", ".join(available[:30])
            if transfer_df.empty:
                all_lines = sorted(self.dm.master_df["회선명"].unique().tolist())
                detail += f"\n\n[전체 데이터 회선명 목록]\n  " + ", ".join(all_lines[:30])
            self.status_var.set(f"  데이터 누락: {target} 또는 {transfer}")
            messagebox.showerror("데이터 누락", detail)
            return None

        self.status_var.set(f"  AI 학습 중... ({sub}/{target}: {len(target_df)}건)")
        self.update_idletasks()

        # 대상선로 학습 + 예측
        pred_target = LoadPredictor()
        ok, msg = pred_target.train(self.dm, sub, target)
        if not ok:
            self.status_var.set(f"  대상선로 학습 실패")
            messagebox.showerror("AI 학습 오류", f"대상선로({target}):\n{msg}")
            return None

        self.status_var.set(
            f"  대상선로 학습 완료. 전환선로 학습 중... ({transfer}: {len(transfer_df)}건)")
        self.update_idletasks()

        # 전환선로 학습 + 예측
        pred_transfer = LoadPredictor()
        ok, msg = pred_transfer.train(self.dm, None, transfer)
        if not ok:
            self.status_var.set(f"  전환선로 학습 실패")
            messagebox.showerror("AI 학습 오류", f"전환선로({transfer}):\n{msg}")
            return None

        self.status_var.set(f"  {target_year}년 예측 계산 중...")
        self.update_idletasks()

        # Target Year 전체 예측
        target_preds = pred_target.predict_year(target_year)
        transfer_preds = pred_transfer.predict_year(target_year)

        if not target_preds or not transfer_preds:
            self.status_var.set("  예측 실패")
            messagebox.showerror("예측 오류", "연간 예측 생성에 실패했습니다.")
            return None

        result = {
            "year": target_year,
            "target_preds": target_preds,
            "transfer_preds": transfer_preds,
        }
        self._pred_cache[cache_key] = result

        self.status_var.set(
            f"  AI 예측 완료 | {target_year}년 | {sub}/{target} → {transfer}")
        return result

    # ══════════════════════════════════
    #  탭 1 이벤트 핸들러
    # ══════════════════════════════════
    def _t1_on_sub(self, val):
        targets = self.dm.get_target_lines(val)
        self.t1_target_combo.configure(values=targets or ["(선로 없음)"])
        if targets:
            self.t1_target_combo.set(targets[0])
            self._t1_on_target(targets[0])
        else:
            self.t1_target_combo.set("(선로 없음)")
            self.t1_transfer_var.set("—")

    def _t1_on_target(self, val):
        sub = self.t1_sub_combo.get()
        tr = self.dm.get_transfer_line(sub, val)
        self.t1_transfer_var.set(tr if tr else "—")

    def _t1_on_run(self):
        sub = self.t1_sub_combo.get()
        target = self.t1_target_combo.get()
        transfer = self.t1_transfer_var.get()

        if not self.dm.substations:
            messagebox.showwarning("알림", "전환선로 매핑 파일을 먼저 로드해주세요.")
            return
        if not self._pred_cache:
            messagebox.showwarning("알림",
                "예측 데이터가 없습니다.\n"
                "데이터 폴더를 선택하거나 [AI 재학습] 버튼을 눌러주세요.")
            return
        if target in ("(선로 없음)", "변전소를 선택하세요"):
            messagebox.showwarning("알림", "휴전선로(대상선로)를 선택해주세요.")
            return

        self._current_sub = sub
        self._current_target = target
        self._current_transfer = transfer

        threshold = self._get_threshold()

        # 캐시에서 즉시 조회 (재학습 없음)
        cache_key = (sub, target)
        pred_result = self._pred_cache.get(cache_key)
        if not pred_result:
            messagebox.showwarning("알림",
                f"'{target}' 선로의 예측 데이터가 없습니다.\n"
                "[AI 재학습] 버튼을 눌러 전체 데이터를 갱신해 주세요.")
            return

        target_year = pred_result["year"]
        monthly = self._calc_predicted_monthly(
            target_year, pred_result["target_preds"],
            pred_result["transfer_preds"], threshold)
        self._t1_render(monthly, sub, target, transfer, target_year)

    def _calc_predicted_monthly(self, target_year: int,
                                 target_preds: dict, transfer_preds: dict,
                                 threshold: float) -> list:
        """예측 결과에서 월별 절체 가능 일수를 계산."""
        results = []
        for month in range(1, 13):
            n_days = calendar.monthrange(target_year, month)[1]
            possible_count = 0
            has_data = False
            for day in range(1, n_days + 1):
                date_str = f"{target_year}{month:02d}{day:02d}"
                t_pred = target_preds.get(date_str)
                tr_pred = transfer_preds.get(date_str)
                if t_pred and tr_pred:
                    has_data = True
                    sum_hours = [t + tr for t, tr in
                                 zip(t_pred["pred_hours"], tr_pred["pred_hours"])]
                    sum_max = max(sum_hours)
                    if sum_max <= threshold:
                        possible_count += 1
            results.append(possible_count if has_data else None)
        return results

    def _t1_render(self, monthly: list, sub: str, target: str, transfer: str,
                   target_year: int = None):
        C = self.C
        for w in self.t1_dashboard.winfo_children():
            w.destroy()

        year_tag = f"  ({target_year}년 AI 예측)" if target_year else ""
        ok_c = sum(1 for d in monthly if d is not None and d >= LEVEL_HIGH)
        wa_c = sum(1 for d in monthly if d is not None and LEVEL_MID <= d < LEVEL_HIGH)
        ng_c = sum(1 for d in monthly if d is not None and d < LEVEL_MID)
        nd_c = sum(1 for d in monthly if d is None)
        self.t1_summary_var.set(
            f"{sub}  |  {target} → {transfer}{year_tag}      "
            f"[ 유력 {ok_c}  /  고려 {wa_c}  /  불가 {ng_c}"
            + (f"  /  없음 {nd_c}" if nd_c else "") + " ]"
        )

        self._t1_render_table(monthly, sub, target, transfer, target_year)

    def _t1_render_table(self, monthly: list, sub: str, target: str, transfer: str,
                         target_year: int = None):
        C = self.C
        year_label = f"  {target_year}년 AI 예측 — 절체 가능 일수" if target_year else "  절체 가능 일수"

        title_bar = tk.Frame(self.t1_dashboard, bg=C["accent"])
        title_bar.pack(fill="x", padx=2, pady=(6, 0))
        tk.Label(title_bar, text=year_label, font=(FONT_FAMILY, 12, "bold"),
                 bg=C["accent"], fg="#ffffff",
        ).pack(side="left", padx=4, pady=3)
        tk.Label(title_bar, text="※ 🔍 버튼을 클릭하면 일별 AI 예측 부하를 상세 조회합니다  ",
                 font=(FONT_FAMILY, 10), bg=C["accent"], fg="#d0e8ff",
        ).pack(side="right", padx=4, pady=3)

        wrap = tk.Frame(self.t1_dashboard, bg=C["card_bg"])
        wrap.pack(fill="both", expand=True, padx=2, pady=(0, 2))

        grid = tk.Frame(wrap, bg=C["card_bg"])
        grid.pack(fill="both", expand=True, padx=4, pady=4)

        headers = [("월", 6), ("절체 가능 일수", 14), ("판정", 10), ("상세", 6)]
        for col, (txt, w) in enumerate(headers):
            tk.Label(grid, text=txt, font=(FONT_FAMILY, 11, "bold"),
                     bg=C["tree_hdr_bg"], fg=C["tree_hdr_fg"],
                     width=w, height=2, relief="flat",
            ).grid(row=0, column=col, padx=(0, 1), pady=(0, 1), sticky="nsew")

        for i in range(12):
            row_num = i + 1
            d = monthly[i]
            info = get_level_info(d, C)
            d_str = f"{d}일" if d is not None else "—"
            row_bg = info["bg"]

            tk.Label(grid, text=f"{i+1}월", font=(FONT_FAMILY, 11, "bold"),
                     bg=row_bg, fg=C["text"],
                     width=6, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")

            tk.Label(grid, text=d_str, font=(FONT_FAMILY, 13, "bold"),
                     bg=row_bg, fg=info["color"],
                     width=14, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

            tk.Label(grid, text=f'{info["icon"]}  {info["level"]}',
                     font=(FONT_FAMILY, 11, "bold"),
                     bg=row_bg, fg=info["color"],
                     width=10, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=2, padx=(0, 1), pady=(0, 1), sticky="nsew")

            btn = tk.Button(
                grid, text="🔍 상세", font=(FONT_FAMILY, 10),
                bg=C["accent"], fg="#ffffff", activebackground=C["accent_hover"],
                activeforeground="#ffffff", relief="flat", cursor="hand2",
                command=lambda m=i, s=sub, t=target, tr=transfer, ty=target_year:
                    self._open_daily_popup(s, t, tr, m + 1, ty),
            )
            btn.grid(row=row_num, column=3, padx=(2, 1), pady=(1, 1), sticky="nsew")

        grid.columnconfigure(0, weight=1, uniform="t1")
        grid.columnconfigure(1, weight=3, uniform="t1")
        grid.columnconfigure(2, weight=2, uniform="t1")
        grid.columnconfigure(3, weight=1, uniform="t1")
        for r in range(13):
            grid.rowconfigure(r, weight=1)

    # ══════════════════════════════════
    #  탭 2 이벤트 핸들러
    # ══════════════════════════════════
    def _t2_on_run(self):
        sub = self.t2_sub_combo.get()
        if not self.dm.substations:
            messagebox.showwarning("알림", "전환선로 매핑 파일을 먼저 로드해주세요.")
            return
        if not self._pred_cache:
            messagebox.showwarning("알림",
                "예측 데이터가 없습니다.\n"
                "데이터 폴더를 선택하거나 [AI 재학습] 버튼을 눌러주세요.")
            return
        if sub == "데이터를 먼저 로드하세요":
            messagebox.showwarning("알림", "변전소를 선택해주세요.")
            return

        threshold = self._get_threshold()

        lines = self.dm.substations.get(sub, [])
        target_year = self._cached_target_year
        all_data = []

        for target, transfer in lines:
            cache_key = (sub, target)
            pred_result = self._pred_cache.get(cache_key)
            if not pred_result:
                all_data.append((target, transfer, [None] * 12))
                continue

            monthly = self._calc_predicted_monthly(
                target_year, pred_result["target_preds"],
                pred_result["transfer_preds"], threshold)
            all_data.append((target, transfer, monthly))

        self.status_var.set(
            f"  조회 완료 | {target_year}년 AI 예측 | {sub} 변전소 {len(lines)}개 선로")
        self._t2_render(sub, all_data, target_year)

    def _t2_render(self, sub: str, all_data: list, target_year: int = None):
        C = self.C

        for w in self.t2_header_fr.winfo_children():
            w.destroy()
        for w in self.t2_scroll.winfo_children():
            w.destroy()

        if not all_data:
            ctk.CTkLabel(
                self.t2_scroll,
                text=f"'{sub}' 변전소에 해당하는 선로 데이터가 없습니다.",
                font=ctk.CTkFont(family=FONT_FAMILY, size=13), text_color=C["text_sub"],
            ).pack(pady=40)
            self.t2_summary_var.set("")
            return

        year_tag = f"  ({target_year}년 AI 예측)" if target_year else ""
        self.t2_summary_var.set(f"{sub} 변전소  —  총 {len(all_data)}개 선로{year_tag}")

        # ── 고정 헤더 ──
        hdr_fr = tk.Frame(self.t2_header_fr, bg=C["card_bg"])
        hdr_fr.pack(fill="x", padx=(0, 14), pady=(0, 0))

        tk.Label(hdr_fr, text="", font=(FONT_FAMILY, 10),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 height=1, relief="flat",
        ).grid(row=0, column=0, columnspan=2, padx=(0, 1), pady=(0, 0), sticky="nsew")

        hdr_title = f"{target_year}년 AI 예측 — 절체 가능 일수" if target_year else "절체 가능 일수"
        tk.Label(hdr_fr, text=hdr_title, font=(FONT_FAMILY, 12, "bold"),
                 bg=C["accent"], fg="#ffffff",
                 height=1, relief="flat",
        ).grid(row=0, column=2, columnspan=12, padx=(0, 1), pady=(0, 0), sticky="nsew")

        tk.Label(hdr_fr, text="대상선로", font=(FONT_FAMILY, 11, "bold"),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 width=10, height=2, relief="flat",
        ).grid(row=1, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")

        tk.Label(hdr_fr, text="전환선로", font=(FONT_FAMILY, 11, "bold"),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 width=10, height=2, relief="flat",
        ).grid(row=1, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

        for m in range(12):
            tk.Label(hdr_fr, text=f"{m+1}월", font=(FONT_FAMILY, 11, "bold"),
                     bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                     width=6, height=2, relief="flat",
            ).grid(row=1, column=m+2, padx=(0, 1), pady=(0, 1), sticky="nsew")

        hdr_fr.columnconfigure(0, weight=2, uniform="g")
        hdr_fr.columnconfigure(1, weight=2, uniform="g")
        for m in range(12):
            hdr_fr.columnconfigure(m+2, weight=1, uniform="g")

        # ── 데이터 행 ──
        grid_fr = tk.Frame(self.t2_scroll, bg=C["card_bg"])
        grid_fr.pack(fill="x", padx=0, pady=(0, 2))

        for r_idx, (target, transfer, monthly) in enumerate(all_data):
            stripe = C["card_bg"] if r_idx % 2 == 0 else C["grid_line_bg"]

            target_lbl = tk.Label(
                grid_fr, text=target, font=(FONT_FAMILY, 11, "bold"),
                bg=stripe, fg=C["link"],
                width=10, height=2, anchor="center", relief="flat",
                cursor="hand2",
            )
            target_lbl.grid(row=r_idx, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")
            target_lbl.bind(
                "<Button-1>",
                lambda e, s=sub, t=target: self._t2_navigate_to_t1(s, t),
            )
            target_lbl.bind(
                "<Enter>",
                lambda e, lbl=target_lbl: lbl.configure(fg=C["link_hover"], font=(FONT_FAMILY, 11, "bold underline")),
            )
            target_lbl.bind(
                "<Leave>",
                lambda e, lbl=target_lbl: lbl.configure(fg=C["link"], font=(FONT_FAMILY, 11, "bold")),
            )

            tk.Label(grid_fr, text=transfer, font=(FONT_FAMILY, 10),
                     bg=stripe, fg=C["highlight"],
                     width=10, height=2, anchor="center", relief="flat",
            ).grid(row=r_idx, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

            for m in range(12):
                days = monthly[m] if m < len(monthly) else None
                info = get_level_info(days, C)

                cell_text = info["text"] if days is not None else "—"

                lbl = tk.Label(
                    grid_fr, text=cell_text,
                    font=(FONT_FAMILY, 11, "bold"),
                    bg=info["cell"], fg=info["color"],
                    width=6, height=2, anchor="center", relief="flat",
                )
                lbl.grid(row=r_idx, column=m+2, padx=(0, 1), pady=(0, 1), sticky="nsew")

                tip = (f"{target} → {transfer}\n"
                       f"{m+1}월: {cell_text}일\n판정: {info['level']}")
                self._bind_tooltip(lbl, tip)

        grid_fr.columnconfigure(0, weight=2, uniform="g")
        grid_fr.columnconfigure(1, weight=2, uniform="g")
        for m in range(12):
            grid_fr.columnconfigure(m+2, weight=1, uniform="g")

    # ══════════════════════════════════
    #  탭 2 → 탭 1 자동 전환
    # ══════════════════════════════════
    def _t2_navigate_to_t1(self, sub: str, target: str):
        self.t1_sub_combo.set(sub)
        targets = self.dm.get_target_lines(sub)
        self.t1_target_combo.configure(values=targets or ["(선로 없음)"])
        self.t1_target_combo.set(target)

        transfer = self.dm.get_transfer_line(sub, target)
        self.t1_transfer_var.set(transfer if transfer else "—")

        self.tabview.set("단일 선로 상세 조회")
        self._t1_on_run()

    # ══════════════════════════════════
    #  ★★★ 일별 시간대 부하 상세 팝업 ★★★
    # ══════════════════════════════════
    def _open_daily_popup(self, sub: str, target: str, transfer: str, month: int,
                          target_year: int = None):
        """
        팝업 상단: 일자 리스트 (Treeview) — 일자, 대상 최대, 전환 최대, 합산 최대, 가능여부
        팝업 하단: 선택한 일자의 1~24시 시간대별 부하 Treeview (가로 스크롤)

        target_year가 지정되면 AI 예측 데이터를 사용.
        """
        C = self.C
        threshold = self._get_threshold()

        # 예측 데이터 기반 daily_data 생성
        if target_year:
            cache_key = (sub, target)
            cache = self._pred_cache.get(cache_key)
            if not cache:
                messagebox.showwarning("알림", "예측 데이터가 없습니다. 먼저 결과 조회를 실행하세요.")
                return

            target_preds = cache["target_preds"]
            transfer_preds = cache["transfer_preds"]
            n_days = calendar.monthrange(target_year, month)[1]
            daily_data = []
            for day in range(1, n_days + 1):
                date_str = f"{target_year}{month:02d}{day:02d}"
                t_pred = target_preds.get(date_str)
                tr_pred = transfer_preds.get(date_str)
                if t_pred and tr_pred:
                    sum_hours = [round(t + tr, 2) for t, tr in
                                 zip(t_pred["pred_hours"], tr_pred["pred_hours"])]
                    sum_max = round(max(sum_hours), 2)
                    daily_data.append({
                        "day": day,
                        "date_str": date_str,
                        "weekday": t_pred["weekday"],
                        "is_weekend": t_pred["is_weekend"],
                        "target_hours": t_pred["pred_hours"],
                        "transfer_hours": tr_pred["pred_hours"],
                        "sum_hours": sum_hours,
                        "target_max": t_pred["pred_max"],
                        "transfer_max": tr_pred["pred_max"],
                        "sum_max": sum_max,
                        "possible": sum_max <= threshold,
                    })
            month_label = f"{target_year}년 {month}월 (AI 예측)"
        else:
            month_label = f"{month}월"
            daily_data = self.dm.get_daily_detail(sub, target, month)

        popup = ctk.CTkToplevel(self)
        popup.title(f"일별 부하 상세 — {sub} | {target} → {transfer} | {month_label}")
        popup.geometry("1200x920")
        popup.resizable(True, True)
        popup.minsize(900, 700)
        popup.attributes("-topmost", True)
        popup.after(100, popup.focus_force)

        # ★ 최상단 노출: 열릴 때 맨 앞으로, 이후 사용자 자유 조작 허용
        popup.attributes("-topmost", True)
        popup.after(300, lambda: popup.attributes("-topmost", False))
        popup.focus_force()
        popup.lift()

        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 1200) // 2
        py = self.winfo_y() + (self.winfo_height() - 920) // 2
        popup.geometry(f"+{max(px, 0)}+{max(py, 0)}")

        # ── 헤더 ──
        header_fr = ctk.CTkFrame(popup, fg_color=C["title_bg"], corner_radius=0, height=50)
        header_fr.pack(fill="x")
        header_fr.pack_propagate(False)

        ctk.CTkLabel(
            header_fr,
            text=f"  {month_label} 일별 시간대 부하 상세",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=C["title_fg"],
        ).pack(side="left", padx=14)

        ctk.CTkLabel(
            header_fr,
            text=f"{sub}  |  대상: {target}  |  전환: {transfer}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=C["subtitle_fg"],
        ).pack(side="left", padx=14)

        # ── 범례 ──
        legend_fr = tk.Frame(popup, bg=C["card_bg"])
        legend_fr.pack(fill="x", padx=14, pady=(6, 2))

        pred_tag = "  [AI 예측]" if target_year else ""
        tk.Label(legend_fr,
                 text=f"※ 합산 최대 부하 > {threshold} → 불가(X) / ≤ {threshold} → 가능(O){pred_tag}",
                 font=(FONT_FAMILY, 10), bg=C["card_bg"], fg=C["ng"],
        ).pack(side="left")

        # 주말 범례
        wknd_box = tk.Frame(legend_fr, bg=C["weekend_bg"], width=14, height=14)
        wknd_box.pack(side="left", padx=(20, 3))
        wknd_box.pack_propagate(False)
        tk.Label(legend_fr, text="주말(토/일)", font=(FONT_FAMILY, 10),
                 bg=C["card_bg"], fg=C["weekend_fg"]).pack(side="left")

        if not daily_data:
            ctk.CTkLabel(
                popup, text="해당 월의 부하 데이터가 없습니다.",
                font=ctk.CTkFont(family=FONT_FAMILY, size=13), text_color=C["text_sub"],
            ).pack(expand=True)
            return

        # ══════════════════════════════════
        #  상단: 일자별 판정 리스트 (Treeview)
        # ══════════════════════════════════
        top_label = tk.Label(popup, text="  ▼ 일자별 판정 리스트 (클릭하면 하단에 시간대별 부하 표시)",
                             font=(FONT_FAMILY, 11, "bold"), bg=C["card_bg"], fg=C["text"],
                             anchor="w")
        top_label.pack(fill="x", padx=14, pady=(6, 2))

        tree_fr = tk.Frame(popup, bg=C["card_bg"])
        tree_fr.pack(fill="x", padx=14, pady=(0, 4))

        style = ttk.Style()
        style.theme_use("clam")
        sname = "DailyPopup.Treeview"
        style.configure(sname, font=(FONT_FAMILY, 11), rowheight=26,
                        background=C["card_bg"], fieldbackground=C["card_bg"],
                        foreground=C["text"], borderwidth=0)
        style.configure(f"{sname}.Heading", font=(FONT_FAMILY, 11, "bold"),
                        background=C["tree_hdr_bg"], foreground=C["tree_hdr_fg"],
                        borderwidth=0, relief="flat")
        style.map(sname, background=[("selected", C["tree_sel"])])

        cols = ("day", "weekday", "target_max", "transfer_max", "sum_max", "judge")
        tree = ttk.Treeview(tree_fr, columns=cols, show="headings",
                            style=sname, height=min(len(daily_data), 12))
        tree.heading("day", text="일자")
        tree.heading("weekday", text="요일")
        tree.heading("target_max", text=f"대상 일최대 ({target})")
        tree.heading("transfer_max", text=f"전환 일최대 ({transfer})")
        tree.heading("sum_max", text="합산 일최대")
        tree.heading("judge", text="가능여부")
        tree.column("day", width=65, anchor="center")
        tree.column("weekday", width=50, anchor="center")
        tree.column("target_max", width=185, anchor="center")
        tree.column("transfer_max", width=185, anchor="center")
        tree.column("sum_max", width=140, anchor="center")
        tree.column("judge", width=80, anchor="center")

        tree.tag_configure("normal", background=C["card_bg"])
        tree.tag_configure("overload", background=C["overload_bg"], foreground=C["overload_fg"])
        tree.tag_configure("stripe", background=C["grid_line_bg"])
        tree.tag_configure("weekend", background=C["weekend_bg"], foreground=C["weekend_fg"])
        tree.tag_configure("weekend_overload", background=C["overload_bg"], foreground=C["overload_fg"])

        possible_count = 0
        for idx, dd in enumerate(daily_data):
            is_wknd = dd.get("is_weekend", False)
            if not dd["possible"]:
                tag = "weekend_overload" if is_wknd else "overload"
            elif is_wknd:
                tag = "weekend"
            else:
                tag = "stripe" if idx % 2 == 1 else "normal"

            judge_str = "O (가능)" if dd["possible"] else "X (불가)"
            if dd["possible"]:
                possible_count += 1

            weekday_str = dd.get("weekday", "")
            # 토/일 표시 강조
            if weekday_str in ("토", "일"):
                weekday_str = f"({weekday_str})"

            tree.insert("", "end",
                        values=(f"{dd['day']}일",
                                weekday_str,
                                f"{dd['target_max']:.1f}",
                                f"{dd['transfer_max']:.1f}",
                                f"{dd['sum_max']:.1f}",
                                judge_str),
                        tags=(tag,))

        sb = ttk.Scrollbar(tree_fr, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # 요약 (주말 일수 표시 추가)
        weekend_count = sum(1 for dd in daily_data if dd.get("is_weekend"))
        sum_label = tk.Label(popup,
                             text=(f"  총 {len(daily_data)}일  |  가능(O): {possible_count}일  |  "
                                   f"불가(X): {len(daily_data) - possible_count}일  |  "
                                   f"주말: {weekend_count}일"),
                             font=(FONT_FAMILY, 11, "bold"), bg=C["card_bg"],
                             fg=C["ok"] if possible_count > len(daily_data) // 2 else C["ng"],
                             anchor="w")
        sum_label.pack(fill="x", padx=14, pady=(0, 4))

        # ══════════════════════════════════════════
        #  하단: 시간대별 부하 프로필 (Treeview + 가로/세로 스크롤)
        # ══════════════════════════════════════════
        profile_title_var = tk.StringVar(
            value="  ▼ 시간대별 부하 프로필 (상단 리스트에서 일자를 클릭하세요)")
        profile_title_lbl = tk.Label(popup, textvariable=profile_title_var,
                                     font=(FONT_FAMILY, 11, "bold"),
                                     bg=C["card_bg"], fg=C["text"], anchor="w")
        profile_title_lbl.pack(fill="x", padx=14, pady=(4, 2))

        # 프로필 Treeview 영역
        profile_outer = tk.Frame(popup, bg=C["card_bg"])
        profile_outer.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        # 스타일
        pname = "HourlyProfile.Treeview"
        style.configure(pname, font=(FONT_FAMILY, 11), rowheight=30,
                        background=C["card_bg"], fieldbackground=C["card_bg"],
                        foreground=C["text"], borderwidth=0)
        style.configure(f"{pname}.Heading", font=(FONT_FAMILY, 10, "bold"),
                        background=C["tree_hdr_bg"], foreground=C["tree_hdr_fg"],
                        borderwidth=0, relief="flat")
        style.map(pname, background=[("selected", C["tree_sel"])])

        # 컬럼 정의: 구분 + 1시~24시 + MAX = 26개
        pcols = ("label",) + tuple(f"h{h}" for h in range(1, 25)) + ("max_val",)
        profile_tree = ttk.Treeview(profile_outer, columns=pcols, show="headings",
                                     style=pname, height=4)

        profile_tree.heading("label", text="구분")
        profile_tree.column("label", width=130, minwidth=110, stretch=True, anchor="center")
        for h in range(1, 25):
            profile_tree.heading(f"h{h}", text=f"{h}시")
            profile_tree.column(f"h{h}", width=65, minwidth=55, stretch=True, anchor="center")
        profile_tree.heading("max_val", text="MAX")
        profile_tree.column("max_val", width=75, minwidth=60, stretch=True, anchor="center")

        # 행 태그 (전체 행 색상)
        profile_tree.tag_configure("target",
                                    foreground=C["graph_target"])
        profile_tree.tag_configure("transfer",
                                    background=C["grid_line_bg"], foreground=C["graph_transfer"])
        profile_tree.tag_configure("sum_ok",
                                    foreground=C["ok"], background=C["ok_light"])
        profile_tree.tag_configure("sum_ng",
                                    background=C["overload_bg"], foreground=C["overload_fg"])

        # 가로 + 세로 스크롤바
        profile_xsb = ttk.Scrollbar(profile_outer, orient="horizontal",
                                     command=profile_tree.xview)
        profile_ysb = ttk.Scrollbar(profile_outer, orient="vertical",
                                     command=profile_tree.yview)
        profile_tree.configure(xscrollcommand=profile_xsb.set,
                                yscrollcommand=profile_ysb.set)

        profile_tree.grid(row=0, column=0, sticky="nsew")
        profile_ysb.grid(row=0, column=1, sticky="ns")
        profile_xsb.grid(row=1, column=0, sticky="ew")
        profile_outer.grid_rowconfigure(0, weight=1)
        profile_outer.grid_columnconfigure(0, weight=1)

        # 임계값 안내
        pred_marker = "  [AI 예측 데이터]" if target_year else ""
        threshold_lbl = tk.Label(popup,
                                  text=f"  임계값: {threshold}{pred_marker}  |  ※ 창을 최대화하거나 넓히면 더 많은 시간대를 한눈에 볼 수 있습니다",
                                  font=(FONT_FAMILY, 9), bg=C["card_bg"], fg=C["ng"],
                                  anchor="w")
        threshold_lbl.pack(fill="x", padx=14, pady=(0, 2))

        # ══════════════════════════════════════════
        #  하단: 시간대별 부하 그래프 (matplotlib)
        # ══════════════════════════════════════════
        chart_title_var = tk.StringVar(
            value="  ▼ 시간대별 부하 그래프 (상단 리스트에서 일자를 클릭하세요)")
        chart_title_lbl = tk.Label(popup, textvariable=chart_title_var,
                                    font=(FONT_FAMILY, 11, "bold"),
                                    bg=C["card_bg"], fg=C["text"], anchor="w")
        chart_title_lbl.pack(fill="x", padx=14, pady=(4, 2))

        chart_fr = tk.Frame(popup, bg=C["card_bg"])
        chart_fr.pack(fill="both", expand=True, padx=14, pady=(0, 4))

        # matplotlib Figure 생성
        fig_bg = C["card_bg"]
        fig = Figure(figsize=(10, 3), dpi=96, facecolor=fig_bg)
        ax = fig.add_subplot(111)
        ax.set_facecolor(fig_bg)
        ax.set_xlim(0.5, 24.5)
        ax.set_xticks(range(1, 25))
        ax.set_xticklabels([f"{h}시" for h in range(1, 25)], fontsize=8)
        ax.set_xlabel("시간대", fontsize=10)
        ax.set_ylabel("부하", fontsize=10)
        ax.tick_params(colors=C["text"], labelsize=8)
        ax.xaxis.label.set_color(C["text"])
        ax.yaxis.label.set_color(C["text"])
        for spine in ax.spines.values():
            spine.set_color(C["grid_line_bg"])
        ax.text(0.5, 0.5, "일자를 선택하세요", transform=ax.transAxes,
                ha="center", va="center", fontsize=14, color=C["text_sub"], alpha=0.5)
        fig.tight_layout(pad=2)

        canvas_widget = FigureCanvasTkAgg(fig, master=chart_fr)
        canvas_widget.draw()
        canvas_widget.get_tk_widget().pack(fill="both", expand=True)

        # ── 일자 클릭 → 프로필 갱신 + 그래프 갱신 ──
        is_pred_mode = target_year is not None
        label_suffix = " (예측)" if is_pred_mode else ""

        def on_tree_select(event):
            sel = tree.selection()
            if not sel:
                return
            idx = tree.index(sel[0])
            if idx >= len(daily_data):
                return
            dd = daily_data[idx]

            # 타이틀 갱신
            judge = "가능(O)" if dd["possible"] else "불가(X)"
            wkday = dd.get("weekday", "")
            wknd_tag = "  [주말]" if dd.get("is_weekend") else ""
            profile_title_var.set(
                f"  ▼ {dd['day']}일({wkday}) 시간대별 부하{label_suffix}  |  합산 최대: {dd['sum_max']:.1f}  |  판정: {judge}{wknd_tag}")
            profile_title_lbl.configure(
                bg="#8e44ad" if is_pred_mode else C["accent"], fg="#ffffff")

            # 기존 행 삭제
            for item in profile_tree.get_children():
                profile_tree.delete(item)

            # 행 1: 대상선로
            vals = [f"대상({target}){label_suffix}"] + \
                   [f"{v:.1f}" for v in dd["target_hours"]] + \
                   [f"{dd['target_max']:.1f}"]
            profile_tree.insert("", "end", values=vals, tags=("target",))

            # 행 2: 전환선로
            vals = [f"전환({transfer}){label_suffix}"] + \
                   [f"{v:.1f}" for v in dd["transfer_hours"]] + \
                   [f"{dd['transfer_max']:.1f}"]
            profile_tree.insert("", "end", values=vals, tags=("transfer",))

            # 행 3: 합산 부하 — 초과 시간대는 ▲ 표시
            sum_tag = "sum_ok" if dd["possible"] else "sum_ng"
            sum_vals = []
            for v in dd["sum_hours"]:
                if v > threshold:
                    sum_vals.append(f"▲{v:.1f}")
                else:
                    sum_vals.append(f"{v:.1f}")
            vals = [f"합산 부하{label_suffix}"] + sum_vals + [f"{dd['sum_max']:.1f}"]
            profile_tree.insert("", "end", values=vals, tags=(sum_tag,))

            # ── 그래프 갱신 ──
            chart_title_var.set(
                f"  ▼ {dd['day']}일({wkday}) 시간대별 부하 그래프{label_suffix}  |  판정: {judge}{wknd_tag}")
            chart_title_lbl.configure(
                bg="#8e44ad" if is_pred_mode else C["accent"], fg="#ffffff")

            ax.clear()
            hours = list(range(1, 25))

            # 대상선로
            ax.plot(hours, dd["target_hours"], color=C["graph_target"],
                    linestyle="--", linewidth=1.5, marker="o", markersize=4,
                    label=f"대상({target}){label_suffix}")

            # 전환선로
            ax.plot(hours, dd["transfer_hours"], color=C["graph_transfer"],
                    linestyle="--", linewidth=1.5, marker="s", markersize=4,
                    label=f"전환({transfer}){label_suffix}")

            # 합산 부하 (실선, 굵게)
            ax.plot(hours, dd["sum_hours"], color=C["graph_total"],
                    linestyle="-", linewidth=2.5, marker="D", markersize=4,
                    label=f"합산 부하{label_suffix}", zorder=5)

            # 초과 구간 강조
            over_hours = [h for h, v in zip(hours, dd["sum_hours"]) if v > threshold]
            over_vals = [v for v in dd["sum_hours"] if v > threshold]
            if over_hours:
                ax.scatter(over_hours, over_vals, color=C["graph_total"],
                           s=80, zorder=6, edgecolors="white", linewidths=1.2)

            # 임계값 기준선
            ax.axhline(y=threshold, color=C["ng"], linestyle=":",
                       linewidth=1.5, alpha=0.8, label=f"임계값({threshold})")

            # 초과 영역 배경
            ax.fill_between(hours, threshold, dd["sum_hours"],
                            where=[v > threshold for v in dd["sum_hours"]],
                            color=C["graph_total"], alpha=0.1, interpolate=True)

            # 축 설정
            ax.set_xlim(0.5, 24.5)
            ax.set_xticks(range(1, 25))
            ax.set_xticklabels([f"{h}시" for h in range(1, 25)], fontsize=8)
            ax.set_xlabel("시간대", fontsize=10)
            ax.set_ylabel("부하" + (" (예측)" if is_pred_mode else ""), fontsize=10)
            ax.set_facecolor(fig_bg)
            ax.tick_params(colors=C["text"], labelsize=8)
            ax.xaxis.label.set_color(C["text"])
            ax.yaxis.label.set_color(C["text"])
            for spine in ax.spines.values():
                spine.set_color(C["grid_line_bg"])

            # 범례
            ax.legend(loc="upper right", fontsize=9, framealpha=0.8,
                      facecolor=fig_bg, edgecolor=C["grid_line_bg"],
                      labelcolor=C["text"])

            ax.grid(True, alpha=0.3, color=C["grid_line_bg"])
            fig.tight_layout(pad=2)
            canvas_widget.draw()

        tree.bind("<<TreeviewSelect>>", on_tree_select)

        # 버튼 영역
        btn_fr = tk.Frame(popup, bg=C["card_bg"])
        btn_fr.pack(fill="x", padx=14, pady=(0, 8))

        ctk.CTkButton(
            btn_fr, text="닫기", width=80, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            command=popup.destroy,
        ).pack(side="right")

    # ══════════════════════════════════
    #  다크/라이트 토글
    # ══════════════════════════════════
    def _toggle_theme(self):
        self.is_dark = not self.is_dark
        self.C = DARK if self.is_dark else LIGHT
        ctk.set_appearance_mode("dark" if self.is_dark else "light")
        for w in self.winfo_children():
            w.destroy()
        self.configure(fg_color=self.C["bg"])
        self._build_ui()
        self.theme_btn.configure(
            text="☀️ 라이트모드" if self.is_dark else "🌙 다크모드"
        )

    # ══════════════════════════════════
    #  툴팁
    # ══════════════════════════════════
    def _bind_tooltip(self, widget, text: str):
        tip_win = None

        def show(e):
            nonlocal tip_win
            if tip_win:
                return
            tip_win = tk.Toplevel(widget)
            tip_win.wm_overrideredirect(True)
            tip_win.wm_geometry(f"+{e.x_root + 14}+{e.y_root + 10}")
            tk.Label(
                tip_win, text=text, font=(FONT_FAMILY, 10),
                bg="#2c3e50", fg="white", justify="left", padx=10, pady=6,
            ).pack()

        def hide(e):
            nonlocal tip_win
            if tip_win:
                tip_win.destroy()
                tip_win = None

        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)


# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    app = App()
    app.mainloop()
