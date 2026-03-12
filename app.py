"""
=============================================================================
  배전선로 휴전 가능월 검토 프로그램 v5.0
  Distribution Line Suspension Monthly Feasibility Review

  - 종합 결과 엑셀 파일 1개에서 네 시트를 읽어 UI에 표출
    ① '전환선로' 시트 → 콤보박스 매핑
    ② '절체가능여부 판단결과' 시트 → 월별 절체 가능 일수
    ③ '일일 최대부하' 시트 → 대상선로 일별 부하
    ④ '전환선로 부하' 시트 → 전환선로 일별 부하
  - 탭 1: 단일 선로 상세 조회 (카드형 대시보드 + 월 클릭 시 일별 팝업)
  - 탭 2: 변전소 종합 조회 (전 선로 히트맵 Grid, 대상선로 클릭 → 탭 1 연동)
  - 색상 기준: 22일+ 유력(Green) | 13~21일 고려(Orange) | 12일- 불가(Red)
=============================================================================
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
import customtkinter as ctk


# ═══════════════════════════════════════════════════════════════
#  전역 설정
# ═══════════════════════════════════════════════════════════════
FONT_FAMILY = "맑은 고딕"
if sys.platform == "darwin":
    FONT_FAMILY = "Apple SD Gothic Neo"

LEVEL_HIGH = 22   # 이상 → 유력
LEVEL_MID  = 13   # 이상 → 고려, 미만 → 불가
OVERLOAD_THRESHOLD = 10  # 합산 부하 초과 기준 (빨간색 표시)

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
}


# ═══════════════════════════════════════════════════════════════
#  데이터 매니저
# ═══════════════════════════════════════════════════════════════
class DataManager:
    """종합 결과 엑셀에서 네 시트를 파싱"""

    SHEET_MAPPING        = "전환선로"
    SHEET_RESULT         = "절체가능여부 판단결과"
    SHEET_DAILY_TARGET   = "일일 최대부하"
    SHEET_DAILY_TRANSFER = "전환선로 부하"

    def __init__(self):
        self.substations: dict = {}   # {변전소: [(대상선로, 전환선로), ...]}
        self.results: dict     = {}   # {(변전소, 대상선로): [1월~12월 가능일수]}
        # ── 일별 부하 데이터 ──
        self.daily_target: dict   = {}  # {(변전소, 대상선로, month_idx): [day1..day31]}
        self.daily_transfer: dict = {}  # {(변전소, 대상선로, month_idx): [day1..day31]}
        self.month_offsets: list  = []  # [(col_start, num_days), ...] 12개월

    def load_excel(self, filepath: str) -> tuple[bool, str]:
        try:
            xls = pd.ExcelFile(filepath)
            available = xls.sheet_names
        except Exception as e:
            return False, f"파일 열기 실패: {e}"

        msgs = []
        ok1, m1 = self._parse_mapping(filepath, available)
        msgs.append(m1)
        ok2, m2 = self._parse_result(filepath, available)
        msgs.append(m2)
        ok3, m3 = self._parse_daily_load(filepath, available)
        msgs.append(m3)
        return ok1, "  |  ".join(msgs)

    # ─── 전환선로 매핑 파싱 ───
    def _parse_mapping(self, filepath, available) -> tuple[bool, str]:
        if self.SHEET_MAPPING not in available:
            return False, f"'{self.SHEET_MAPPING}' 시트 없음"
        try:
            df = pd.read_excel(filepath, sheet_name=self.SHEET_MAPPING)
            self.substations = {}
            current_sub = str(df.columns[1]).strip()
            for i in range(len(df)):
                c0 = df.iloc[i, 0]
                if pd.isna(c0):
                    continue
                c0 = str(c0).strip()
                if c0 == "변전소명":
                    current_sub = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else current_sub
                    continue
                if c0 == "대상선로":
                    continue
                transfer = str(df.iloc[i, 3]).strip() if pd.notna(df.iloc[i, 3]) else ""
                self.substations.setdefault(current_sub, []).append((c0, transfer))
            n_s = len(self.substations)
            n_p = sum(len(v) for v in self.substations.values())
            return True, f"매핑: {n_s}개 변전소, {n_p}개 선로"
        except Exception as e:
            return False, f"매핑 오류: {e}"

    # ─── 절체가능여부 판단결과 파싱 ───
    def _parse_result(self, filepath, available) -> tuple[bool, str]:
        if self.SHEET_RESULT not in available:
            return False, f"'{self.SHEET_RESULT}' 시트 없음"
        try:
            df = pd.read_excel(filepath, sheet_name=self.SHEET_RESULT)
            self.results = {}
            current_sub = None
            for i in range(3, len(df)):
                row = df.iloc[i]
                if pd.notna(row.iloc[1]):
                    current_sub = str(row.iloc[1]).strip()
                if pd.isna(row.iloc[2]) or current_sub is None:
                    continue
                target = str(row.iloc[2]).strip()
                months = []
                for j in range(3, 15):
                    v = row.iloc[j] if j < len(row) else None
                    if pd.notna(v):
                        try:
                            months.append(int(float(v)))
                        except (ValueError, TypeError):
                            months.append(None)
                    else:
                        months.append(None)
                self.results[(current_sub, target)] = months
            return True, f"결과: {len(self.results)}개 선로"
        except Exception as e:
            return False, f"결과 오류: {e}"

    # ─── 일별 부하 데이터 파싱 (★ 신규) ───
    def _parse_daily_load(self, filepath, available) -> tuple[bool, str]:
        """
        '일일 최대부하' + '전환선로 부하' 시트를 파싱한다.
        두 시트의 구조는 동일:
          - Row 0~2: 빈 행
          - Row 3:   변전소명 | 회선명 | (빈칸) | '2025년01월 일일 사용량' | ... (월 헤더)
          - Row 4:   (빈)     | (빈)   | (빈)   | 1 | 2 | 3 | ... | 31 | 1 | 2 | ... (일 번호)
          - Row 5~:  데이터 행 (변전소명 | 회선명 | (빈) | 1일값 | 2일값 | ...)
        각 월은 실제 일수만큼의 컬럼을 사용 (1월=31, 2월=28/29, ...)
        """
        if self.SHEET_DAILY_TARGET not in available:
            return False, f"'{self.SHEET_DAILY_TARGET}' 시트 없음"
        if self.SHEET_DAILY_TRANSFER not in available:
            return False, f"'{self.SHEET_DAILY_TRANSFER}' 시트 없음"

        try:
            # ① 대상선로 일일 부하 파싱
            df_t = pd.read_excel(filepath, sheet_name=self.SHEET_DAILY_TARGET, header=None)
            self._detect_month_offsets(df_t)
            self.daily_target = {}
            self._parse_daily_rows(df_t, self.daily_target)

            # ② 전환선로 부하 파싱 (구조 동일, 같은 month_offsets 사용)
            df_tr = pd.read_excel(filepath, sheet_name=self.SHEET_DAILY_TRANSFER, header=None)
            self.daily_transfer = {}
            self._parse_daily_rows(df_tr, self.daily_transfer)

            n = len(self.daily_target)
            return True, f"일별부하: {n}건"
        except Exception as e:
            return False, f"일별 부하 오류: {e}"

    def _detect_month_offsets(self, df):
        """
        Row 3의 헤더에서 '20XX년XX월 일일 사용량' 패턴을 찾아
        각 월의 시작 컬럼과 실제 일수를 자동 감지한다.
        예: [(3, 31), (34, 28), (62, 31), ...]  → (시작컬럼, 일수) × 12개월
        """
        self.month_offsets = []
        row3 = df.iloc[3]
        starts = []
        for j in range(df.shape[1]):
            v = row3.iloc[j]
            if pd.notna(v) and isinstance(v, str) and "월" in v and "사용량" in v:
                starts.append(j)

        for i, s in enumerate(starts):
            if i + 1 < len(starts):
                num_days = starts[i + 1] - s
            else:
                num_days = df.shape[1] - s
            self.month_offsets.append((s, num_days))

    def _parse_daily_rows(self, df, storage: dict):
        """
        데이터 행(Row 5~)을 순회하며 일별 부하값을 storage에 저장한다.
        Key: (변전소명, 회선명, month_idx)  →  0-indexed 월 인덱스
        Value: [day1, day2, ..., day31]  →  31개로 패딩 (부족한 일수는 None)
        """
        current_sub = None
        for i in range(5, len(df)):
            sub_val = df.iloc[i, 0]
            line_val = df.iloc[i, 1]

            # 변전소명이 명시된 행은 갱신, NaN이면 이전 값 유지
            if pd.notna(sub_val):
                current_sub = str(sub_val).strip()
            # 회선명이 없거나 변전소 미확정이면 건너뜀
            if pd.isna(line_val) or current_sub is None:
                continue

            line = str(line_val).strip()

            for m_idx, (col_start, num_days) in enumerate(self.month_offsets):
                days = []
                for d in range(num_days):
                    col = col_start + d
                    if col < df.shape[1]:
                        v = df.iloc[i, col]
                        if pd.notna(v):
                            try:
                                days.append(round(float(v), 2))
                            except (ValueError, TypeError):
                                days.append(None)
                        else:
                            days.append(None)
                    else:
                        days.append(None)
                # 31일로 패딩 (28~30일인 월의 나머지는 None)
                while len(days) < 31:
                    days.append(None)
                storage[(current_sub, line, m_idx)] = days

    # ─── 조회 헬퍼 ───
    def get_substation_list(self) -> list[str]:
        return sorted(self.substations.keys())

    def get_target_lines(self, sub: str) -> list[str]:
        return [t for t, _ in self.substations.get(sub, [])]

    def get_transfer_line(self, sub: str, target: str) -> str:
        for t, tr in self.substations.get(sub, []):
            if t == target:
                return tr
        return ""

    def get_monthly_days(self, sub: str, target: str) -> list:
        return self.results.get((sub, target), [None] * 12)

    def get_all_lines_data(self, sub: str) -> list[tuple[str, str, list]]:
        lines = self.substations.get(sub, [])
        result = []
        for target, transfer in lines:
            monthly = self.results.get((sub, target), [None] * 12)
            result.append((target, transfer, monthly))
        return result

    def get_daily_data(self, sub: str, target: str, month_idx: int) -> tuple[list, list]:
        """
        ★★★ 핵심 데이터 매칭 규칙 ★★★

        [대상선로 부하]
          → '일일 최대부하' 시트에서 (변전소명, 대상선로명)으로 행을 검색

        [전환선로 부하]  ← ★★★ 주의 ★★★
          → '전환선로 부하' 시트에서도 동일하게 (변전소명, '대상선로명')으로 행을 검색
          → 전환선로명(예: '동X상')이 아니라, 대상선로명(예: '김X해')으로 찾는다!
          → 이유: '전환선로 부하' 시트는 "대상선로 기준으로 해당 전환선로의 부하"를
                  정리한 구조이기 때문

        (예시)
          대상선로='김X해', 전환선로='동X상' 일 때
          - '일일 최대부하'  시트에서 회선명='김X해'인 행 → 대상선로 부하
          - '전환선로 부하'  시트에서 회선명='김X해'인 행 → 전환선로 부하  (O)
          - '전환선로 부하'  시트에서 회선명='동X상'인 행 → (X) 절대 아님!
        """
        target_days = self.daily_target.get((sub, target, month_idx), [None] * 31)
        transfer_days = self.daily_transfer.get((sub, target, month_idx), [None] * 31)
        return target_days, transfer_days

    def get_month_actual_days(self, month_idx: int) -> int:
        """month_idx(0-indexed)에 해당하는 월의 실제 일수 반환"""
        if month_idx < len(self.month_offsets):
            return self.month_offsets[month_idx][1]
        # 월 오프셋이 없으면 기본값
        return [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][month_idx]

    def has_daily_data(self) -> bool:
        """일별 부하 데이터가 로드되었는지 확인"""
        return len(self.daily_target) > 0 or len(self.daily_transfer) > 0


# ═══════════════════════════════════════════════════════════════
#  색상 판정 유틸리티
# ═══════════════════════════════════════════════════════════════
def get_level_info(days, C: dict) -> dict:
    """일수에 따른 색상/텍스트 정보를 딕셔너리로 반환"""
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
        self.is_dark = False
        self.C = LIGHT
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        self._setup_window()
        self._build_ui()
        # 현재 조회 중인 선로 정보 (팝업에서 사용)
        self._current_sub = ""
        self._current_target = ""
        self._current_transfer = ""

    def _setup_window(self):
        self.title("배전선로 휴전 가능월 검토 프로그램")
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

        # ═══ 데이터 로드 카드 (탭 바깥 — 항상 표시) ═══
        load_card = ctk.CTkFrame(self, fg_color=C["card_bg"], corner_radius=10)
        load_card.pack(fill="x", padx=14, pady=(8, 4))

        ctk.CTkLabel(
            load_card, text="  데이터 로드",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=C["text"],
        ).grid(row=0, column=0, padx=14, pady=(8, 2), sticky="w", columnspan=3)

        # 종합 결과 엑셀
        ctk.CTkLabel(
            load_card, text="종합 결과 엑셀 :",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12), text_color=C["text_sub"],
        ).grid(row=1, column=0, padx=(14, 6), pady=4, sticky="e")

        self.excel_var = ctk.StringVar(value="파일을 선택하세요")
        ctk.CTkEntry(
            load_card, textvariable=self.excel_var, width=620,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            state="readonly", fg_color=C["entry_bg"], text_color=C["text"],
        ).grid(row=1, column=1, padx=4, pady=4, sticky="w")

        ctk.CTkButton(
            load_card, text="파일 선택", width=100, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            command=self._on_select_excel,
        ).grid(row=1, column=2, padx=8, pady=4)

        # 사용량 폴더 (비활성)
        ctk.CTkLabel(
            load_card, text="사용량 폴더 :",
            font=ctk.CTkFont(family=FONT_FAMILY, size=12), text_color=C["nodata"],
        ).grid(row=2, column=0, padx=(14, 6), pady=4, sticky="e")

        ctk.CTkEntry(
            load_card, width=620,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            state="disabled", fg_color=C["entry_bg"],
            placeholder_text="(추후 일별 분석 시 사용 예정)",
        ).grid(row=2, column=1, padx=4, pady=4, sticky="w")

        ctk.CTkButton(
            load_card, text="폴더 선택", width=100, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            state="disabled",
        ).grid(row=2, column=2, padx=8, pady=4)

        # 로드 상태
        self.status_var = ctk.StringVar(value="")
        ctk.CTkLabel(
            load_card, textvariable=self.status_var,
            font=ctk.CTkFont(family=FONT_FAMILY, size=11), text_color=C["ok"],
        ).grid(row=3, column=0, columnspan=3, padx=14, pady=(0, 6), sticky="w")

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

        # ── 조건 선택 영역 ──
        filter_fr = ctk.CTkFrame(parent, fg_color=C["card_bg"], corner_radius=8)
        filter_fr.pack(fill="x", padx=6, pady=(6, 4))

        ctk.CTkLabel(
            filter_fr, text="  조건 선택",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            text_color=C["text"],
        ).grid(row=0, column=0, padx=10, pady=(8, 4), sticky="w", columnspan=8)

        # 변전소
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

        # 휴전선로
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

        # 전환선로
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

        # 결과 조회 버튼
        ctk.CTkButton(
            filter_fr, text="  결과 조회", width=130, height=36,
            font=ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            fg_color=C["accent"], hover_color=C["accent_hover"],
            command=self._t1_on_run,
        ).grid(row=1, column=6, padx=(20, 10), pady=8)

        filter_fr.columnconfigure(7, weight=1)

        # ── 결과 영역 (범례 + 대시보드) ──
        result_fr = ctk.CTkFrame(parent, fg_color=C["card_bg"], corner_radius=8)
        result_fr.pack(fill="both", expand=True, padx=6, pady=(4, 6))

        # 범례 + 요약 행
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

        # 대시보드 컨테이너
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

        # ── 조건 선택 ──
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

        # 요약
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

        # 클릭 안내 라벨
        tk.Label(leg, text="※ 대상선로를 클릭하면 상세 조회로 이동합니다",
                 font=(FONT_FAMILY, 10), bg=C["card_bg"], fg=C["accent"],
        ).pack(side="right", padx=8)

        # ── 고정 헤더 영역 (스크롤 밖 — 틀고정) ──
        self.t2_header_fr = tk.Frame(parent, bg=C["card_bg"])
        self.t2_header_fr.pack(fill="x", padx=(6, 6), pady=(4, 0))

        # ── 결과 그리드 영역 (ScrollableFrame — 데이터만 스크롤) ──
        self.t2_scroll = ctk.CTkScrollableFrame(
            parent, fg_color=C["card_bg"], corner_radius=8,
            label_text="", label_fg_color=C["card_bg"],
        )
        self.t2_scroll.pack(fill="both", expand=True, padx=6, pady=(0, 6))

        # 초기 안내
        self.t2_placeholder = ctk.CTkLabel(
            self.t2_scroll,
            text="변전소를 선택한 후  [ 전체 조회 ]  버튼을 눌러주세요.",
            font=ctk.CTkFont(family=FONT_FAMILY, size=13), text_color=C["text_sub"],
        )
        self.t2_placeholder.pack(pady=40)

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
        sub    = self.t1_sub_combo.get()
        target = self.t1_target_combo.get()
        transfer = self.t1_transfer_var.get()

        if not self.dm.substations:
            messagebox.showwarning("알림", "종합 결과 엑셀 파일을 먼저 로드해주세요.")
            return
        if target in ("(선로 없음)", "변전소를 선택하세요"):
            messagebox.showwarning("알림", "휴전선로(대상선로)를 선택해주세요.")
            return

        monthly = self.dm.get_monthly_days(sub, target)
        # 현재 선로 정보 저장 (팝업에서 참조)
        self._current_sub = sub
        self._current_target = target
        self._current_transfer = transfer
        self._t1_render(monthly, sub, target, transfer)

    # ── 탭 1 대시보드 렌더링 ──
    def _t1_render(self, monthly: list, sub: str, target: str, transfer: str):
        C = self.C
        for w in self.t1_dashboard.winfo_children():
            w.destroy()

        # 요약
        ok_c  = sum(1 for d in monthly if d is not None and d >= LEVEL_HIGH)
        wa_c  = sum(1 for d in monthly if d is not None and LEVEL_MID <= d < LEVEL_HIGH)
        ng_c  = sum(1 for d in monthly if d is not None and d < LEVEL_MID)
        nd_c  = sum(1 for d in monthly if d is None)
        self.t1_summary_var.set(
            f"{sub}  |  {target} → {transfer}      "
            f"[ 유력 {ok_c}  /  고려 {wa_c}  /  불가 {ng_c}"
            + (f"  /  없음 {nd_c}" if nd_c else "") + " ]"
        )

        # ── 절체 가능 일수 테이블 + 🔍 상세 버튼 ──
        self._t1_render_table(monthly, sub, target, transfer)

    def _t1_render_table(self, monthly: list, sub: str, target: str, transfer: str):
        C = self.C

        # ── "절체 가능 일수" 타이틀 ──
        title_bar = tk.Frame(self.t1_dashboard, bg=C["accent"])
        title_bar.pack(fill="x", padx=2, pady=(6, 0))
        tk.Label(title_bar, text="  절체 가능 일수", font=(FONT_FAMILY, 12, "bold"),
                 bg=C["accent"], fg="#ffffff",
        ).pack(side="left", padx=4, pady=3)
        tk.Label(title_bar, text="※ 🔍 버튼을 클릭하면 일별 상세 사용량을 조회할 수 있습니다  ",
                 font=(FONT_FAMILY, 10), bg=C["accent"], fg="#d0e8ff",
        ).pack(side="right", padx=4, pady=3)

        # ── Grid 테이블 ──
        wrap = tk.Frame(self.t1_dashboard, bg=C["card_bg"])
        wrap.pack(fill="both", expand=True, padx=2, pady=(0, 2))

        grid = tk.Frame(wrap, bg=C["card_bg"])
        grid.pack(fill="both", expand=True, padx=4, pady=4)

        # 헤더 행
        headers = [("월", 6), ("절체 가능 일수", 14), ("판정", 10), ("상세", 6)]
        for col, (txt, w) in enumerate(headers):
            tk.Label(grid, text=txt, font=(FONT_FAMILY, 11, "bold"),
                     bg=C["tree_hdr_bg"], fg=C["tree_hdr_fg"],
                     width=w, height=2, relief="flat",
            ).grid(row=0, column=col, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # 데이터 행 (1~12월)
        for i in range(12):
            row_num = i + 1
            d = monthly[i]
            info = get_level_info(d, C)
            d_str = f"{d}일" if d is not None else "—"
            row_bg = info["bg"]

            # 월
            tk.Label(grid, text=f"{i+1}월", font=(FONT_FAMILY, 11, "bold"),
                     bg=row_bg, fg=C["text"],
                     width=6, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")

            # 절체 가능 일수
            tk.Label(grid, text=d_str, font=(FONT_FAMILY, 13, "bold"),
                     bg=row_bg, fg=info["color"],
                     width=14, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

            # 판정
            tk.Label(grid, text=f'{info["icon"]}  {info["level"]}',
                     font=(FONT_FAMILY, 11, "bold"),
                     bg=row_bg, fg=info["color"],
                     width=10, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=2, padx=(0, 1), pady=(0, 1), sticky="nsew")

            # 🔍 상세 버튼 — 클릭 시 일별 상세 사용량 팝업 호출
            btn = tk.Button(
                grid, text="🔍 상세", font=(FONT_FAMILY, 10),
                bg=C["accent"], fg="#ffffff", activebackground=C["accent_hover"],
                activeforeground="#ffffff", relief="flat", cursor="hand2",
                command=lambda m=i, s=sub, t=target, tr=transfer:
                    self._open_daily_popup(s, t, tr, m),
            )
            btn.grid(row=row_num, column=3, padx=(2, 1), pady=(1, 1), sticky="nsew")

        # 열 균등 분배
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
            messagebox.showwarning("알림", "종합 결과 엑셀 파일을 먼저 로드해주세요.")
            return
        if sub == "데이터를 먼저 로드하세요":
            messagebox.showwarning("알림", "변전소를 선택해주세요.")
            return

        all_data = self.dm.get_all_lines_data(sub)
        self._t2_render(sub, all_data)

    # ── 탭 2 종합 Grid 렌더링 ──
    def _t2_render(self, sub: str, all_data: list):
        C = self.C

        # 기존 위젯 클리어 (고정 헤더 + 스크롤 데이터 모두)
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

        self.t2_summary_var.set(f"{sub} 변전소  —  총 {len(all_data)}개 선로")

        # ══════════════════════════════════
        #  고정 헤더 (스크롤 밖 — 틀고정)
        # ══════════════════════════════════
        hdr_fr = tk.Frame(self.t2_header_fr, bg=C["card_bg"])
        hdr_fr.pack(fill="x", padx=(0, 14), pady=(0, 0))

        # ── 상단 스패닝 헤더: "절체 가능 일수" ──
        tk.Label(hdr_fr, text="", font=(FONT_FAMILY, 10),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 height=1, relief="flat",
        ).grid(row=0, column=0, columnspan=2, padx=(0, 1), pady=(0, 0), sticky="nsew")

        tk.Label(hdr_fr, text="절체 가능 일수", font=(FONT_FAMILY, 12, "bold"),
                 bg=C["accent"], fg="#ffffff",
                 height=1, relief="flat",
        ).grid(row=0, column=2, columnspan=12, padx=(0, 1), pady=(0, 0), sticky="nsew")

        # 대상선로 헤더
        tk.Label(hdr_fr, text="대상선로", font=(FONT_FAMILY, 11, "bold"),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 width=10, height=2, relief="flat",
        ).grid(row=1, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # 전환선로 헤더
        tk.Label(hdr_fr, text="전환선로", font=(FONT_FAMILY, 11, "bold"),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 width=10, height=2, relief="flat",
        ).grid(row=1, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # 월 헤더 (1~12월)
        for m in range(12):
            tk.Label(hdr_fr, text=f"{m+1}월", font=(FONT_FAMILY, 11, "bold"),
                     bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                     width=6, height=2, relief="flat",
            ).grid(row=1, column=m+2, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # 헤더 열 균등 분배
        hdr_fr.columnconfigure(0, weight=2, uniform="g")
        hdr_fr.columnconfigure(1, weight=2, uniform="g")
        for m in range(12):
            hdr_fr.columnconfigure(m+2, weight=1, uniform="g")

        # ══════════════════════════════════
        #  데이터 행 (스크롤 영역 안)
        # ══════════════════════════════════
        grid_fr = tk.Frame(self.t2_scroll, bg=C["card_bg"])
        grid_fr.pack(fill="x", padx=0, pady=(0, 2))

        for r_idx, (target, transfer, monthly) in enumerate(all_data):
            row_num = r_idx
            stripe = C["card_bg"] if r_idx % 2 == 0 else C["grid_line_bg"]

            # ★ 대상선로명 — 클릭 가능한 라벨 (클릭 시 탭 1로 이동)
            target_lbl = tk.Label(
                grid_fr, text=target, font=(FONT_FAMILY, 11, "bold"),
                bg=stripe, fg=C["link"],
                width=10, height=2, anchor="center", relief="flat",
                cursor="hand2",
            )
            target_lbl.grid(row=row_num, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")
            # 클릭 이벤트 바인딩: 대상선로 클릭 → 탭 1 자동 전환 및 조회
            target_lbl.bind(
                "<Button-1>",
                lambda e, s=sub, t=target: self._t2_navigate_to_t1(s, t),
            )
            # 호버 효과
            target_lbl.bind(
                "<Enter>",
                lambda e, lbl=target_lbl: lbl.configure(fg=C["link_hover"], font=(FONT_FAMILY, 11, "bold underline")),
            )
            target_lbl.bind(
                "<Leave>",
                lambda e, lbl=target_lbl: lbl.configure(fg=C["link"], font=(FONT_FAMILY, 11, "bold")),
            )

            # 전환선로명
            tk.Label(grid_fr, text=transfer, font=(FONT_FAMILY, 10),
                     bg=stripe, fg=C["highlight"],
                     width=10, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

            # 1~12월 데이터 셀
            for m in range(12):
                days = monthly[m] if m < len(monthly) else None
                info = get_level_info(days, C)

                cell_bg = info["cell"]
                cell_fg = info["color"]
                cell_text = info["text"] if days is not None else "—"

                lbl = tk.Label(
                    grid_fr, text=cell_text,
                    font=(FONT_FAMILY, 11, "bold"),
                    bg=cell_bg, fg=cell_fg,
                    width=6, height=2, anchor="center", relief="flat",
                )
                lbl.grid(row=row_num, column=m+2, padx=(0, 1), pady=(0, 1), sticky="nsew")

                # 툴팁
                tip = (f"{target} → {transfer}\n"
                       f"{m+1}월: {cell_text}일\n판정: {info['level']}")
                self._bind_tooltip(lbl, tip)

        # 데이터 열 균등 분배 (헤더와 동일한 비율)
        grid_fr.columnconfigure(0, weight=2, uniform="g")
        grid_fr.columnconfigure(1, weight=2, uniform="g")
        for m in range(12):
            grid_fr.columnconfigure(m+2, weight=1, uniform="g")

    # ══════════════════════════════════
    #  ★ 탭 2 → 탭 1 자동 전환 (요구사항 1)
    # ══════════════════════════════════
    def _t2_navigate_to_t1(self, sub: str, target: str):
        """
        변전소 종합 조회(탭 2)에서 대상선로 클릭 시:
        1) 탭 1의 변전소 콤보박스를 해당 변전소로 설정
        2) 대상선로 콤보박스를 클릭한 선로로 설정
        3) 전환선로 표시 갱신
        4) 탭 1("단일 선로 상세 조회")로 자동 전환
        5) 해당 선로의 결과를 즉시 조회
        """
        # ① 변전소 콤보박스 설정
        self.t1_sub_combo.set(sub)
        targets = self.dm.get_target_lines(sub)
        self.t1_target_combo.configure(values=targets or ["(선로 없음)"])

        # ② 대상선로 콤보박스 설정
        self.t1_target_combo.set(target)

        # ③ 전환선로 갱신
        transfer = self.dm.get_transfer_line(sub, target)
        self.t1_transfer_var.set(transfer if transfer else "—")

        # ④ 탭 전환
        self.tabview.set("단일 선로 상세 조회")

        # ⑤ 즉시 결과 조회
        self._t1_on_run()

    # ══════════════════════════════════
    #  ★ 월 카드 클릭 → 일별 상세 팝업 (요구사항 2, 3)
    # ══════════════════════════════════
    def _bind_card_click(self, widget, month_idx: int, sub: str, target: str, transfer: str):
        """월 카드 위젯에 클릭 이벤트를 바인딩하여 일별 상세 팝업을 연다."""
        widget.configure(cursor="hand2")
        widget.bind(
            "<Button-1>",
            lambda e, m=month_idx, s=sub, t=target, tr=transfer:
                self._open_daily_popup(s, t, tr, m),
        )

    def _open_daily_popup(self, sub: str, target: str, transfer: str, month_idx: int):
        """
        ★★★ 일별 상세 사용량 팝업 (요구사항 2, 3) ★★★

        해당 월의 1일~말일까지 일별 부하를 표로 표시한다.
        [일자 | 대상선로 부하 | 전환선로 부하 | 합산 부하]

        데이터 매칭 규칙 (★★★매우 중요★★★):
        - 대상선로 부하: '일일 최대부하' 시트에서 (변전소, 대상선로)로 검색
        - 전환선로 부하: '전환선로 부하' 시트에서도 (변전소, '대상선로')로 검색
          → 전환선로명이 아닌 대상선로명으로 행을 찾아야 한다!
        - 합산 부하 > 10 → 빨간색 행으로 강조 표시
        """
        C = self.C
        month_label = f"{month_idx + 1}월"

        # ── 팝업 창 생성 ──
        popup = ctk.CTkToplevel(self)
        popup.title(f"일별 상세 사용량 — {sub} | {target} → {transfer} | {month_label}")
        popup.geometry("720x680")
        popup.resizable(True, True)
        popup.transient(self)
        popup.grab_set()

        # 중앙 배치
        popup.update_idletasks()
        px = self.winfo_x() + (self.winfo_width() - 720) // 2
        py = self.winfo_y() + (self.winfo_height() - 680) // 2
        popup.geometry(f"+{px}+{py}")

        # ── 헤더 영역 ──
        header_fr = ctk.CTkFrame(popup, fg_color=C["title_bg"], corner_radius=0, height=50)
        header_fr.pack(fill="x")
        header_fr.pack_propagate(False)

        ctk.CTkLabel(
            header_fr,
            text=f"  {month_label} 일별 상세 사용량",
            font=ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            text_color=C["title_fg"],
        ).pack(side="left", padx=14)

        ctk.CTkLabel(
            header_fr,
            text=f"{sub}  |  {target} → {transfer}",
            font=ctk.CTkFont(family=FONT_FAMILY, size=11),
            text_color=C["subtitle_fg"],
        ).pack(side="left", padx=14)

        # ── 범례 ──
        legend_fr = tk.Frame(popup, bg=C["card_bg"])
        legend_fr.pack(fill="x", padx=14, pady=(8, 2))

        tk.Label(
            legend_fr,
            text=f"※ 합산 부하 > {OVERLOAD_THRESHOLD} 시 빨간색으로 표시됩니다",
            font=(FONT_FAMILY, 10), bg=C["card_bg"], fg=C["ng"],
        ).pack(side="left")

        # ── 일별 부하 데이터 가져오기 ──
        if not self.dm.has_daily_data():
            # 일별 데이터가 로드되지 않은 경우
            ctk.CTkLabel(
                popup,
                text="일별 부하 데이터가 로드되지 않았습니다.\n"
                     "엑셀 파일에 '일일 최대부하' 및 '전환선로 부하' 시트가 필요합니다.",
                font=ctk.CTkFont(family=FONT_FAMILY, size=13),
                text_color=C["text_sub"],
            ).pack(expand=True)
            return

        # ★★★ 핵심: 대상선로 부하와 전환선로 부하 모두 '대상선로명'으로 검색 ★★★
        target_days, transfer_days = self.dm.get_daily_data(sub, target, month_idx)
        actual_days = self.dm.get_month_actual_days(month_idx)

        # ── Treeview 테이블 ──
        table_fr = tk.Frame(popup, bg=C["card_bg"])
        table_fr.pack(fill="both", expand=True, padx=14, pady=(4, 8))

        style = ttk.Style()
        style.theme_use("clam")

        # 고유 스타일 이름으로 충돌 방지
        sname = "Popup.Treeview"
        style.configure(sname, font=(FONT_FAMILY, 11), rowheight=26,
                        background=C["card_bg"], fieldbackground=C["card_bg"],
                        foreground=C["text"], borderwidth=0)
        style.configure(f"{sname}.Heading", font=(FONT_FAMILY, 11, "bold"),
                        background=C["tree_hdr_bg"], foreground=C["tree_hdr_fg"],
                        borderwidth=0, relief="flat")
        style.map(sname, background=[("selected", C["tree_sel"])])

        cols = ("day", "target_load", "transfer_load", "total_load")
        tree = ttk.Treeview(table_fr, columns=cols, show="headings",
                            style=sname, height=min(actual_days, 20))
        tree.heading("day",           text="일자")
        tree.heading("target_load",   text=f"대상선로 부하 ({target})")
        tree.heading("transfer_load", text=f"전환선로 부하 ({transfer})")
        tree.heading("total_load",    text="합산 부하")
        tree.column("day",           width=70,  anchor="center")
        tree.column("target_load",   width=180, anchor="center")
        tree.column("transfer_load", width=180, anchor="center")
        tree.column("total_load",    width=150, anchor="center")

        # 행 스타일: 정상 / 초과(빨간색)
        tree.tag_configure("normal",   background=C["card_bg"])
        tree.tag_configure("overload", background=C["overload_bg"], foreground=C["overload_fg"])
        tree.tag_configure("stripe",   background=C["grid_line_bg"])

        # ── 일별 데이터 삽입 ──
        overload_count = 0
        for d in range(actual_days):
            t_val = target_days[d]
            tr_val = transfer_days[d]

            # 값 표시: 데이터 있으면 소수점 1자리, 없으면 "—" (0 또는 빈칸 처리)
            t_str = f"{t_val:.1f}" if t_val is not None else "0"
            tr_str = f"{tr_val:.1f}" if tr_val is not None else "0"

            # 합산 부하 계산 (None은 0으로 처리)
            t_num = t_val if t_val is not None else 0.0
            tr_num = tr_val if tr_val is not None else 0.0
            total = t_num + tr_num
            total_str = f"{total:.1f}"

            # ★ 합산 부하 > 10 이면 빨간색 행 강조
            if total > OVERLOAD_THRESHOLD:
                tag = "overload"
                overload_count += 1
            else:
                tag = "stripe" if d % 2 == 1 else "normal"

            tree.insert("", "end",
                        values=(f"{d + 1}일", t_str, tr_str, total_str),
                        tags=(tag,))

        sb = ttk.Scrollbar(table_fr, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # ── 하단 요약 ──
        summary_fr = tk.Frame(popup, bg=C["card_bg"])
        summary_fr.pack(fill="x", padx=14, pady=(0, 10))

        summary_text = f"총 {actual_days}일"
        if overload_count > 0:
            summary_text += f"  |  합산 부하 초과({OVERLOAD_THRESHOLD} 초과): {overload_count}일"
        tk.Label(
            summary_fr, text=summary_text,
            font=(FONT_FAMILY, 11, "bold"), bg=C["card_bg"],
            fg=C["ng"] if overload_count > 0 else C["text_sub"],
        ).pack(side="left")

        # 닫기 버튼
        ctk.CTkButton(
            summary_fr, text="닫기", width=80, height=30,
            font=ctk.CTkFont(family=FONT_FAMILY, size=12),
            command=popup.destroy,
        ).pack(side="right")

    # ══════════════════════════════════
    #  공통 이벤트
    # ══════════════════════════════════
    def _on_select_excel(self):
        path = filedialog.askopenfilename(
            title="종합 결과 엑셀 파일 선택",
            filetypes=[("Excel", "*.xlsx *.xls"), ("All", "*.*")],
        )
        if not path:
            return
        self.excel_var.set(path)
        ok, msg = self.dm.load_excel(path)

        if ok:
            self.status_var.set("  " + msg)
            subs = self.dm.get_substation_list()

            # 탭 1 콤보 갱신
            self.t1_sub_combo.configure(values=subs)
            if subs:
                self.t1_sub_combo.set(subs[0])
                self._t1_on_sub(subs[0])

            # 탭 2 콤보 갱신
            self.t2_sub_combo.configure(values=subs)
            if subs:
                self.t2_sub_combo.set(subs[0])
        else:
            self.status_var.set("  " + msg)
            messagebox.showerror("오류", msg)

    # ──────────────────────────────────
    #  다크/라이트 토글
    # ──────────────────────────────────
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

    # ──────────────────────────────────
    #  툴팁
    # ──────────────────────────────────
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
