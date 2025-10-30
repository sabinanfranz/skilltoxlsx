# -*- coding: utf-8 -*-
import streamlit as st
from io import BytesIO
from pathlib import Path
from typing import List, Dict, Any, Tuple
import json
import re
import zipfile

from openpyxl import load_workbook
from openpyxl.styles import Alignment

# ==========================
# 상수 / 경로
# ==========================
TEMPLATE_DIR = Path(__file__).parent / "templates"
DEFAULT_TEMPLATE_NAME = "Non Track_Paper Interview_상위조직명_직무명(포맷).xlsx"

TASK_START_ROW, TASK_END_ROW = 5, 14   # Task: A(이름), C(설명)
SKILL_START_ROW, SKILL_END_ROW = 5, 11 # Skill: A(related_tasks bullets), B(name), D(def), F(tech_stack)

# ==========================
# 유틸 (파일명/텍스트/정렬)
# ==========================
INVALID_WIN_CHARS = r'<>:"/\\|?*'
INVALID_WIN_PATTERN = re.compile(f"[{re.escape(INVALID_WIN_CHARS)}]+")

def sanitize_filename_component(s: str, fallback: str = "untitled") -> str:
    if not s:
        return fallback
    s = INVALID_WIN_PATTERN.sub(" ", s).strip().strip(".")
    return s if s else fallback

def title_tokens(stem: str) -> List[str]:
    return [t.strip() for t in stem.split("_") if t.strip()]

def is_trailing_excluded(token: str) -> bool:
    # 끝부분에서 제외할 토큰: 'skill', 'HC 제외'
    t = token.lower().replace(" ", "")
    return t in {"skill", "hc제외"}

def parse_org_role_from_filename(filename: str) -> Tuple[str, str, str]:
    """
    {상위조직명} = 첫 번째 토큰
    {직무명} = 두 번째 토큰부터, 끝에서 'skill'/'HC 제외' 제거한 나머지
    반환: org_display, role_display, role_for_filename
    """
    stem = Path(filename).stem
    toks = title_tokens(stem)
    if not toks:
        return "unknown", "", ""
    org = toks[0]

    end = len(toks)
    while end > 1 and is_trailing_excluded(toks[end - 1]):
        end -= 1
    role_tokens = toks[1:end]
    if not role_tokens:
        role_tokens = toks[1:] if len(toks) > 1 else [""]

    role_display = " ".join(role_tokens)  # 셀에는 공백 연결
    role_for_filename = " ".join(role_tokens)  # 파일명에도 동일하게 사용
    return org, role_display, role_for_filename

# 줄바꿈만 활성화하고 기존 정렬/서식은 보존
def with_wrap(cell):
    a = cell.alignment or Alignment()
    return Alignment(
        horizontal=a.horizontal,
        vertical=a.vertical,
        text_rotation=a.text_rotation,
        wrap_text=True,
        shrink_to_fit=a.shrink_to_fit,
        indent=a.indent
    )

def set_text(ws, coord: str, text: str, wrap: bool = True):
    cell = ws[coord]
    cell.value = text
    if wrap:
        cell.alignment = with_wrap(cell)

# ==========================
# JSON/텍스트 처리
# ==========================
CITE_PATTERN = re.compile(r'\s*\[\s*cite\s*:\s*.*?\]\s*', flags=re.IGNORECASE | re.DOTALL)

def strip_citations(text: str) -> str:
    if not text:
        return text
    cleaned = CITE_PATTERN.sub(' ', str(text))
    cleaned = re.sub(r'[ \t]+', ' ', cleaned).strip()
    return cleaned

def load_json_from_txt_bytes(b: bytes) -> Dict[str, Any]:
    """
    TXT 내에 전후 설명이 섞여 있어도 {} 블록만 추출 시도
    """
    txt = b.decode("utf-8-sig", errors="ignore")
    try:
        return json.loads(txt)
    except json.JSONDecodeError:
        start = txt.find("{")
        end = txt.rfind("}")
        if start != -1 and end != -1 and start < end:
            return json.loads(txt[start:end+1])
        raise

def collect_tasks(obj: Dict[str, Any]) -> List[Dict[str, Any]]:
    return obj.get("tasks") or []

def iter_skills(obj: Dict[str, Any]):
    skills = obj.get("skills") or []
    for item in skills:
        if isinstance(item, dict) and "skill" in item:
            s = item.get("skill") or {}
            name = s.get("name", "")
            definition = s.get("definition", "")
            stack = s.get("tech_stack", {})
            related = item.get("related_tasks") or s.get("related_tasks") or []
        else:
            s = item if isinstance(item, dict) else {}
            name = s.get("name", "")
            definition = s.get("definition", "")
            stack = s.get("tech_stack", {})
            related = s.get("related_tasks") or []
        yield {
            "name": name,
            "definition": definition,
            "tech_stack": stack,
            "related_tasks": related
        }

def normalize_list(val) -> List[str]:
    if val is None:
        return []
    if isinstance(val, (list, tuple, set)):
        return [str(x).strip() for x in val if str(x).strip()]
    s = str(val).strip()
    if not s:
        return []
    parts = []
    for chunk in s.replace(";", ",").replace("/", ",").split(","):
        chunk = chunk.strip()
        if chunk:
            parts.append(chunk)
    return parts

def extract_tech_lines(tech_stack: Dict[str, Any]) -> str:
    if not isinstance(tech_stack, dict):
        tech_stack = {}
    lower_map = {str(k).lower(): v for k, v in tech_stack.items()}
    languages = normalize_list(lower_map.get("language") or lower_map.get("languages"))
    os_list   = normalize_list(lower_map.get("os") or lower_map.get("platform") or lower_map.get("operating_system"))
    tools     = normalize_list(lower_map.get("tools") or lower_map.get("tool"))

    lines = []
    if languages:
        lines.append(f"* language: {', '.join(languages)}")
    if os_list:
        lines.append(f"* os: {', '.join(os_list)}")
    if tools:
        lines.append(f"* tools: {', '.join(tools)}")
    return "\n".join(lines)

def bullet_lines(items: List[str]) -> str:
    items = [str(i).strip() for i in items if str(i).strip()]
    return "\n".join(f"* {i}" for i in items)

def related_task_names(related_tasks: List[Dict[str, Any]], task_id_to_name: Dict[str, str]) -> List[str]:
    names = []
    for rt in related_tasks or []:
        name = (rt.get("task_name") or "").strip()
        if not name:
            tid = (rt.get("task_id") or "").strip()
            if tid and tid in task_id_to_name:
                name = task_id_to_name[tid]
        if name:
            names.append(name)
    return names

# ==========================
# 엑셀 생성 (서식 유지)
# ==========================
def get_ws_case_insensitive(wb, name: str):
    if name in wb.sheetnames:
        return wb[name]
    lower_map = {s.lower(): s for s in wb.sheetnames}
    key = name.lower()
    if key in lower_map:
        return wb[lower_map[key]]
    raise KeyError(f"시트를 찾을 수 없습니다: {name}")

def build_workbook_from_template(
    template_bytes: bytes,
    org: str,
    role: str,
    data: Dict[str, Any]
) -> BytesIO:
    """
    템플릿을 로드해 값만 주입(행/열 크기, 서식 유지), 결과를 메모리 바이트로 반환
    """
    wb = load_workbook(BytesIO(template_bytes))
    ws_task = get_ws_case_insensitive(wb, "Task")
    ws_skill = get_ws_case_insensitive(wb, "Skill")

    # --- Task 시트 ---
    set_text(ws_task, "B1", org)
    set_text(ws_task, "B2", role)

    tasks = collect_tasks(data)
    # task_id -> name
    task_id_to_name = {}
    for t in tasks:
        tid = str(t.get("task_id") or "").strip()
        tname = str(t.get("task_name") or "").strip()
        if tid and tname:
            task_id_to_name[tid] = tname

    row = TASK_START_ROW
    for t in tasks[: (TASK_END_ROW - TASK_START_ROW + 1) ]:
        t_name = str(t.get("task_name") or "").strip()
        t_desc = str(t.get("task_description") or "").strip()
        set_text(ws_task, f"A{row}", t_name)
        set_text(ws_task, f"C{row}", t_desc)
        row += 1
    for r in range(row, TASK_END_ROW + 1):
        set_text(ws_task, f"A{r}", "")
        set_text(ws_task, f"C{r}", "")

    # --- Skill 시트 ---
    set_text(ws_skill, "B1", org)
    set_text(ws_skill, "B2", role)

    processed = 0
    max_rows = SKILL_END_ROW - SKILL_START_ROW + 1
    for s in iter_skills(data):
        if processed >= max_rows:
            break
        r = SKILL_START_ROW + processed

        rel_names = related_task_names(s.get("related_tasks"), task_id_to_name)
        set_text(ws_skill, f"A{r}", bullet_lines(rel_names) if rel_names else "")

        set_text(ws_skill, f"B{r}", str(s.get("name") or "").strip())

        # 스킬 정의에서 [cite: ...] 제거
        definition = strip_citations(s.get("definition"))
        set_text(ws_skill, f"D{r}", definition)

        tech_text = extract_tech_lines(s.get("tech_stack"))
        set_text(ws_skill, f"F{r}", tech_text)

        processed += 1

    for r in range(SKILL_START_ROW + processed, SKILL_END_ROW + 1):
        set_text(ws_skill, f"A{r}", "")
        set_text(ws_skill, f"B{r}", "")
        set_text(ws_skill, f"D{r}", "")
        set_text(ws_skill, f"F{r}", "")

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ==========================
# 변환 파이프라인
# ==========================
def process_uploaded_txt(uploaded_file, template_bytes: bytes):
    """
    단일 TXT 업로드 파일을 처리하여 (출력파일명, 엑셀바이트) 반환
    """
    org, role_display, role_for_filename = parse_org_role_from_filename(uploaded_file.name)
    # 파일명 구성
    safe_org = sanitize_filename_component(org, "org")
    safe_role = sanitize_filename_component(role_for_filename, "role")
    out_name = f"Non Track_Paper Interview_{safe_org}_{safe_role}.xlsx"

    try:
        data = load_json_from_txt_bytes(uploaded_file.read())
    except Exception as e:
        raise RuntimeError(f"JSON 파싱 실패: {uploaded_file.name} ({e})")

    # 템플릿 주입
    wb_bytes = build_workbook_from_template(
        template_bytes=template_bytes,
        org=org,
        role=role_display,
        data=data
    )
    return out_name, wb_bytes

def zip_bytes(files: Dict[str, BytesIO]) -> bytes:
    mem = BytesIO()
    with zipfile.ZipFile(mem, mode="w", compression=zipfile.ZIP_DEFLATED) as z:
        for name, bio in files.items():
            z.writestr(name, bio.getvalue())
    mem.seek(0)
    return mem.read()

# ==========================
# Streamlit UI
# ==========================
st.set_page_config(page_title="TXT → Excel 변환기 (현대자동차 스킬 컨설팅)", layout="wide")

st.title("TXT(JSON) → Excel 변환기")
st.caption("템플릿 서식을 유지하고, 파일명 규칙/Task·Skill 시트 채움, 스킬 정의의 [cite: ...] 제거, 줄바꿈 표시까지 자동 처리합니다.")

with st.sidebar:
    st.header("템플릿 설정")
    tpl_upload = st.file_uploader("템플릿 업로드 (.xlsx) — (선택)", type=["xlsx"], accept_multiple_files=False)
    if tpl_upload is None:
        default_tpl_path = TEMPLATE_DIR / DEFAULT_TEMPLATE_NAME
        if not default_tpl_path.exists():
            st.error(f"기본 템플릿이 없습니다: {default_tpl_path}")
        else:
            st.success(f"기본 템플릿 사용: {DEFAULT_TEMPLATE_NAME}")
            template_bytes = default_tpl_path.read_bytes()
    else:
        template_bytes = tpl_upload.read()
        st.success(f"업로드한 템플릿 사용: {tpl_upload.name}")

    st.divider()
    st.markdown("**고정 옵션**")
    st.markdown("- 기존 **열너비/행높이/서식** 유지")
    st.markdown("- **줄바꿈** 표시(wrap_text=True)")
    st.markdown("- 스킬 정의의 **[cite: ...] 제거**")

st.subheader("1) TXT(JSON) 파일 업로드")
uploaded_files = st.file_uploader("여러 파일을 동시에 올릴 수 있습니다.", type=["txt"], accept_multiple_files=True)

# 미리보기 표
if uploaded_files:
    st.write("**파일명 파싱 미리보기**")
    preview_rows = []
    for f in uploaded_files:
        org, role_display, role_for_filename = parse_org_role_from_filename(f.name)
        out = f"Non Track_Paper Interview_{sanitize_filename_component(org)}_{sanitize_filename_component(role_for_filename)}.xlsx"
        preview_rows.append({"원본 파일": f.name, "상위조직명": org, "직무명": role_display, "생성될 엑셀": out})
    st.dataframe(preview_rows, use_container_width=True)

run = st.button("변환 실행", type="primary", disabled=not uploaded_files)

if run and uploaded_files:
    if "template_bytes" not in locals():
        st.error("템플릿을 찾을 수 없습니다. 사이드바에서 템플릿을 업로드하거나 기본 템플릿을 확인하세요.")
    else:
        results: Dict[str, BytesIO] = {}
        errors = []
        with st.spinner("변환 중..."):
            for uf in uploaded_files:
                try:
                    name, bio = process_uploaded_txt(uf, template_bytes)
                    results[name] = bio
                except Exception as e:
                    errors.append(f"{uf.name} → 실패: {e}")

        st.subheader("2) 변환 결과")
        col1, col2 = st.columns([2, 1])

        with col1:
            if results:
                st.success(f"{len(results)}개 파일 생성 완료")
                for fname, bio in results.items():
                    st.download_button(
                        label=f"⬇️ {fname} 다운로드",
                        data=bio.getvalue(),
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )

        with col2:
            if results and len(results) > 1:
                z = zip_bytes(results)
                st.download_button(
                    label="📦 전체 ZIP 다운로드",
                    data=z,
                    file_name="converted_excels.zip",
                    mime="application/zip",
                    use_container_width=True
                )

        if errors:
            st.warning("일부 파일 변환 중 오류가 발생했습니다.")
            for msg in errors:
                st.write(f"• {msg}")

st.divider()
st.markdown(
"""
**규칙 요약**
- 파일명 규칙  
  - `{상위조직명}` = 파일명을 `_`로 분할했을 때 첫 번째 토큰  
  - `{직무명}` = 두 번째 토큰부터, 끝에서 `'skill'`, `'HC 제외'`를 제거한 나머지  
- Task 시트  
  - `B1={상위조직명}`, `B2={직무명}`  
  - `A5..A14 = tasks[*].task_name`, `C5..C14 = tasks[*].task_description`  
- Skill 시트  
  - `B1={상위조직명}`, `B2={직무명}`  
  - `A5..A11 = related_tasks[*].task_name`을 `* {이름}` 줄바꿈 목록  
  - `B5..B11 = skill.name`, `D5..D11 = skill.definition( cite 제거 )`, `F5..F11 = tech_stack(language/os/tools)`
- **서식**: 템플릿의 열너비/행높이/폰트/테두리/병합 등을 **그대로 유지**하며 값만 주입합니다.
"""
)
