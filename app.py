"""
=============================================================================
  배전선로 휴전 가능월 검토 프로그램 v4.0
  Distribution Line Suspension Monthly Feasibility Review

  - 종합 결과 엑셀 파일 1개에서 두 시트를 읽어 UI에 표출
    ① '전환선로' 시트 → 콤보박스 매핑
    ② '절체가능여부 판단결과' 시트 → 월별 절체 가능 일수
  - 탭 1: 단일 선로 상세 조회 (카드형 대시보드)
  - 탭 2: 변전소 종합 조회 (전 선로 히트맵 Grid)
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
}


# ═══════════════════════════════════════════════════════════════
#  데이터 매니저
# ═══════════════════════════════════════════════════════════════
class DataManager:
    """종합 결과 엑셀에서 '전환선로' + '절체가능여부 판단결과' 시트를 파싱"""

    SHEET_MAPPING = "전환선로"
    SHEET_RESULT  = "절체가능여부 판단결과"

    def __init__(self):
        self.substations: dict = {}   # {변전소: [(대상선로, 전환선로), ...]}
        self.results: dict     = {}   # {(변전소, 대상선로): [1월~12월 가능일수]}

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
        return ok1, "  |  ".join(msgs)

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
        """
        변전소 전체 선로의 (대상선로, 전환선로, [1~12월 가능일수]) 리스트 반환
        """
        lines = self.substations.get(sub, [])
        result = []
        for target, transfer in lines:
            monthly = self.results.get((sub, target), [None] * 12)
            result.append((target, transfer, monthly))
        return result


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

        # ── 결과 그리드 영역 (ScrollableFrame) ──
        self.t2_scroll = ctk.CTkScrollableFrame(
            parent, fg_color=C["card_bg"], corner_radius=8,
            label_text="", label_fg_color=C["card_bg"],
        )
        self.t2_scroll.pack(fill="both", expand=True, padx=6, pady=(2, 6))

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

        # ── 카드 그리드 (4열×3행) ──
        card_area = tk.Frame(self.t1_dashboard, bg=C["card_bg"])
        card_area.pack(fill="both", expand=True, pady=(0, 4))

        for idx in range(12):
            row, col = divmod(idx, 4)
            info = get_level_info(monthly[idx], C)

            outer = tk.Frame(card_area, bg=info["border"], padx=3, pady=3)
            outer.grid(row=row, column=col, padx=7, pady=7, sticky="nsew")

            cell = tk.Frame(outer, bg=info["bg"])
            cell.pack(fill="both", expand=True)

            tk.Label(cell, text=f"{idx+1}월", font=(FONT_FAMILY, 12, "bold"),
                     bg=info["bg"], fg=C["text"]).pack(anchor="w", padx=12, pady=(10, 1))

            day_text = f'{info["text"]}일' if monthly[idx] is not None else "—"
            tk.Label(cell, text=day_text, font=(FONT_FAMILY, 26, "bold"),
                     bg=info["bg"], fg=info["color"]).pack(pady=(2, 1))

            tk.Label(cell, text=f'  {info["icon"]}  {info["level"]}  ',
                     font=(FONT_FAMILY, 11, "bold"),
                     bg=info["badge"], fg="#ffffff").pack(pady=(1, 10))

            tip = f'{idx+1}월 — {target} → {transfer}\n가능 일수: {day_text}\n판정: {info["level"]}'
            for w in [outer, cell] + list(cell.winfo_children()):
                self._bind_tooltip(w, tip)

        for c in range(4):
            card_area.columnconfigure(c, weight=1, uniform="c")
        for r in range(3):
            card_area.rowconfigure(r, weight=1, uniform="c")

        # ── 하단 테이블 ──
        self._t1_render_table(monthly, target)

    def _t1_render_table(self, monthly: list, target: str):
        C = self.C
        wrap = tk.Frame(self.t1_dashboard, bg=C["card_bg"])
        wrap.pack(fill="x", padx=2, pady=(2, 2))

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("T1.Treeview", font=(FONT_FAMILY, 11), rowheight=28,
                         background=C["card_bg"], fieldbackground=C["card_bg"],
                         foreground=C["text"], borderwidth=0)
        style.configure("T1.Treeview.Heading", font=(FONT_FAMILY, 11, "bold"),
                         background=C["tree_hdr_bg"], foreground=C["tree_hdr_fg"],
                         borderwidth=0)
        style.map("T1.Treeview", background=[("selected", C["tree_sel"])])

        cols = ("month", "days", "level")
        tree = ttk.Treeview(wrap, columns=cols, show="headings",
                            style="T1.Treeview", height=5)
        tree.heading("month", text="월")
        tree.heading("days",  text="절체 가능 일수")
        tree.heading("level", text="판정")
        tree.column("month", width=80,  anchor="center")
        tree.column("days",  width=180, anchor="center")
        tree.column("level", width=100, anchor="center")

        tree.tag_configure("ok",   background=C["ok_light"])
        tree.tag_configure("warn", background=C["warn_light"])
        tree.tag_configure("ng",   background=C["ng_light"])
        tree.tag_configure("nd",   background=C["nodata_light"])

        for i in range(12):
            d = monthly[i]
            info = get_level_info(d, C)
            d_str = f"{d}일" if d is not None else "—"
            tag = "ok" if d is not None and d >= LEVEL_HIGH else \
                  "warn" if d is not None and d >= LEVEL_MID else \
                  "ng" if d is not None else "nd"
            tree.insert("", "end",
                        values=(f"{i+1}월", d_str, f'{info["icon"]} {info["level"]}'),
                        tags=(tag,))

        sb = ttk.Scrollbar(wrap, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

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

        # 기존 위젯 클리어
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

        # ── 헤더 행 ──
        grid_fr = tk.Frame(self.t2_scroll, bg=C["card_bg"])
        grid_fr.pack(fill="x", padx=2, pady=(4, 2))

        # 대상선로 헤더
        tk.Label(grid_fr, text="대상선로", font=(FONT_FAMILY, 11, "bold"),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 width=10, height=2, relief="flat",
        ).grid(row=0, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # 전환선로 헤더
        tk.Label(grid_fr, text="전환선로", font=(FONT_FAMILY, 11, "bold"),
                 bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                 width=10, height=2, relief="flat",
        ).grid(row=0, column=1, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # 월 헤더 (1~12월)
        for m in range(12):
            tk.Label(grid_fr, text=f"{m+1}월", font=(FONT_FAMILY, 11, "bold"),
                     bg=C["grid_hdr"], fg=C["grid_hdr_fg"],
                     width=6, height=2, relief="flat",
            ).grid(row=0, column=m+2, padx=(0, 1), pady=(0, 1), sticky="nsew")

        # ── 데이터 행 ──
        for r_idx, (target, transfer, monthly) in enumerate(all_data):
            row_num = r_idx + 1
            stripe = C["card_bg"] if r_idx % 2 == 0 else C["grid_line_bg"]

            # 대상선로명
            tk.Label(grid_fr, text=target, font=(FONT_FAMILY, 11, "bold"),
                     bg=stripe, fg=C["text"],
                     width=10, height=2, anchor="center", relief="flat",
            ).grid(row=row_num, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")

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

        # 열 균등 분배
        grid_fr.columnconfigure(0, weight=2, uniform="g")
        grid_fr.columnconfigure(1, weight=2, uniform="g")
        for m in range(12):
            grid_fr.columnconfigure(m+2, weight=1, uniform="g")

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
