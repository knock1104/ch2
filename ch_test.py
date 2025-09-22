# -*- coding: utf-8 -*-
"""
Streamlit 예배 자료 업로드 + Word 저장 (개편+저장/제출 버전)
요청 반영:
1) 예배 구분 직접기입
2) 자료 순서 조절
3) 라이브 프리뷰 삭제
4) 이미지 여러 장 업로드
5) 설명 강조(**굵게**, ==형광펜==) -> Word 반영
6) 자료 유형 '설교 전문' 추가
7) 임시 저장/불러오기/제출(GitHub Contents API)
"""

# ---------------------------
# 페이지 설정 (최상단)
# ---------------------------
import streamlit as st
st.set_page_config(
    page_title="설교 자료 업로드",
    page_icon="🙏",
    layout="wide"
)

# ---------------------------
# 표준 라이브러리 / 서드파티
# ---------------------------
import io
import os
import uuid
import base64
import tempfile
import re
import json
import requests
from datetime import date, datetime, timezone
from typing import List, Dict, Any

# python-docx / PIL
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
except Exception:
    st.warning("python-docx가 설치되지 않았습니다. 터미널에서: pip install python-docx")
    Document = None

try:
    from PIL import Image
except Exception:
    st.warning("Pillow가 설치되지 않았습니다. 터미널에서: pip install pillow")
    Image = None

# ---------------------------
# 스타일
# ---------------------------
st.markdown(
    """
    <style>
    .small-note { color:#666; font-size:0.9rem; }
    .section-title { font-weight:700; font-size:1.1rem; margin-top:0.5rem; }
    .pill {
        display:inline-block; padding:4px 10px; border-radius:999px; background:#f0f2f6; margin-right:6px;
        font-size:0.85rem; color:#333; border:1px solid #e5e7eb;
    }
    .landing-card {
        padding: 16px; border: 1px solid #e5e7eb; border-radius: 12px; background: #fff;
        box-shadow: 0 4px 12px rgba(0,0,0,0.04);
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ---------------------------
# 세션 상태 초기화
# ---------------------------
if "materials" not in st.session_state:
    # 각 항목: {id, kind, files, file, verse_text, description, full_text}
    st.session_state.materials: List[Dict[str, Any]] = []

if "preview_idx" not in st.session_state:
    st.session_state.preview_idx = 0  # (프리뷰 제거, 호환 유지)

# 권한/사용자 상태
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "role" not in st.session_state:
    st.session_state.role = None       # "교역자" | "미디어부"
if "user_name" not in st.session_state:
    st.session_state.user_name = ""
if "position" not in st.session_state:
    st.session_state.position = ""
if "can_edit" not in st.session_state:
    st.session_state.can_edit = False

# 예배 구분(옵션/선택) 상태
BASE_SERVICES = ["1부", "2부", "3부", "오후예배"]
if "services_options" not in st.session_state:
    st.session_state.services_options = BASE_SERVICES.copy()
if "services_selected" not in st.session_state:
    st.session_state.services_selected: List[str] = []

# ---------------------------
# 0) 랜딩 페이지 (권한/접근 제어)
# ---------------------------
def render_landing():
    st.title("Ch2 설교 자료 업로더")
    st.markdown(
        "<div class='landing-card'>"
        "<b>역할을 선택하고 입장하세요.</b><br>"
        "교역자는 작성/수정이 가능하며, 미디어부는 확인만 가능합니다.<br>"
        "<span class='small-note'>[테스트 안내] 현재는 모든 액세스 코드가 <b>0001</b>이면 입장 가능합니다.</span>"
        "</div>",
        unsafe_allow_html=True
    )
    st.write("")

    with st.form("landing_form"):
        role = st.radio("역할 선택", ["교역자", "미디어부"], horizontal=True)
        cols = st.columns(2)
        with cols[0]:
            user_name = st.text_input("이름")
        with cols[1]:
            position = st.selectbox(
                "직분 선택",
                ['원로목사', "담임목사", "부목사", '강도사', "전도사", "미디어부"],
                index=2
            )
        access_code = st.text_input("개인 액세스 코드", type="password", placeholder="예) 0001")
        submitted = st.form_submit_button("입장")

    if submitted:
        if access_code == "0001":
            st.session_state.authenticated = True
            st.session_state.role = role
            st.session_state.user_name = user_name.strip()
            st.session_state.position = position
            st.session_state.can_edit = (role == "교역자")
            st.success("입장되었습니다.")
            st.rerun()
        else:
            st.error("액세스 코드가 올바르지 않습니다. (테스트: 0001)")

if not st.session_state.authenticated:
    render_landing()
    st.stop()

# ---------------------------
# 상단 사용자/권한 표시
# ---------------------------
can_edit = st.session_state.get("can_edit", False)
role_badge = "🟢 편집 가능" if can_edit else "🔒 읽기 전용(확인만)"
st.markdown(
    f"**접속자:** {st.session_state.user_name or '이름 미입력'} "
    f"({st.session_state.position or '직분 미선택'}) · "
    f"{st.session_state.role} · {role_badge}"
)

# ---------------------------
# 유틸 함수 (자료 편집)
# ---------------------------
def add_material():
    st.session_state.materials.append({
        "id": str(uuid.uuid4()),
        "kind": "성경 구절",     # "성경 구절" | "이미지" | "기타 파일" | "설교 전문"
        "files": [],             # 이미지 다중 업로드용
        "file": None,            # 과거 호환(단일 파일)
        "verse_text": "",
        "description": "",
        "full_text": ""          # 설교 전문
    })

def remove_material(mid: str):
    st.session_state.materials = [m for m in st.session_state.materials if m["id"] != mid]

def move_material(mid: str, direction: str):
    """direction: 'up' or 'down'"""
    mats = st.session_state.materials
    idx = next((i for i, m in enumerate(mats) if m["id"] == mid), None)
    if idx is None:
        return
    if direction == "up" and idx > 0:
        mats[idx-1], mats[idx] = mats[idx], mats[idx-1]
        st.session_state.materials = mats
        st.rerun()
    elif direction == "down" and idx < len(mats)-1:
        mats[idx+1], mats[idx] = mats[idx], mats[idx+1]
        st.session_state.materials = mats
        st.rerun()

def add_rich_text(paragraph, text: str):
    """
    간단 마크업 → Word 서식 변환:
      **굵게**  => bold
      ==형광펜== => highlight(YELLOW)
    그 외는 일반 텍스트로 추가.
    """
    if not text:
        return
    pattern = r'(\*\*.*?\*\*|==.*?==)'
    parts = re.split(pattern, text)
    for part in parts:
        if not part:
            continue
        if part.startswith("**") and part.endswith("**"):
            run = paragraph.add_run(part[2:-2])
            run.bold = True
        elif part.startswith("==") and part.endswith("=="):
            run = paragraph.add_run(part[2:-2])
            try:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            except Exception:
                pass
        else:
            paragraph.add_run(part)

def build_docx(
    worship_date: date,
    services: List[str],
    materials: List[Dict[str, Any]],
    user_name: str,
    position: str,
    role: str,
) -> bytes:
    """문서(.docx) 생성 후 바이트로 반환 (설교 전문/강조 마크업 반영 버전)"""
    if Document is None:
        raise RuntimeError("python-docx가 설치되지 않았습니다. 'pip install python-docx' 실행 후 다시 시도해주세요.")

    doc = Document()

    # 글꼴 기본값(선택)
    style = doc.styles['Normal']
    style.font.name = '맑은 고딕'
    style.font.size = Pt(11)

    # 제목
    title = doc.add_paragraph()
    run = title.add_run("설교 자료")
    run.bold = True
    run.font.size = Pt(20)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # 메타 정보
    meta = doc.add_paragraph()
    meta.add_run(f"날짜: {worship_date.strftime('%Y-%m-%d')}\n").bold = True
    if services:
        meta.add_run("예배 구분: " + ", ".join(services) + "\n").bold = True
    else:
        meta.add_run("예배 구분: (미선택)\n").bold = True
    if user_name or position or role:
        meta.add_run(f"작성자/권한: {user_name or '(미입력)'} ({position or '직분 미선택'}) - {role or '권한 미지정'}").bold = True

    doc.add_paragraph("")

    # 자료(스토리보드)
    doc.add_heading("자료 (스토리보드)", level=1)
    if not materials:
        doc.add_paragraph("(추가된 자료가 없습니다)")
    else:
        for idx, item in enumerate(materials, start=1):
            kind = item.get("kind", "")
            verse_text = item.get("verse_text", "") or ""
            description = item.get("description", "") or ""
            full_text = item.get("full_text", "") or ""
            files = item.get("files", []) or []
            single_file = item.get("file", None)  # 과거 호환

            doc.add_heading(f"{idx}. {kind}", level=2)

            if kind == "성경 구절":
                if verse_text.strip():
                    for line in verse_text.splitlines():
                        p = doc.add_paragraph()
                        add_rich_text(p, line)
                    doc.add_paragraph("")
                else:
                    doc.add_paragraph("(성경 구절 미입력)")

            elif kind == "이미지":
                # 다중 업로드(현행)
                if files and Image is not None:
                    for f in files:
                        try:
                            with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(f.name)[1]) as tmp:
                                tmp.write(f.getvalue())
                                tmp.flush()
                                doc.add_picture(tmp.name, width=Inches(5))
                        except Exception:
                            doc.add_paragraph(f"(이미지 삽입 실패) 파일명: {getattr(f, 'name', 'unknown')}")
                # 단일 파일(과거 호환)
                elif single_file is not None and Image is not None:
                    try:
                        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(single_file.name)[1]) as tmp:
                            tmp.write(single_file.getvalue())
                            tmp.flush()
                            doc.add_picture(tmp.name, width=Inches(5))
                    except Exception:
                        doc.add_paragraph(f"(이미지 삽입 실패) 파일명: {single_file.name}")
                else:
                    doc.add_paragraph("(이미지 파일 없음)")

            elif kind == "기타 파일":
                if single_file is not None:
                    doc.add_paragraph(f"첨부 파일: {single_file.name} (문서에 직접 삽입되지 않습니다)")
                else:
                    doc.add_paragraph("(첨부 파일 없음)")

            elif kind == "설교 전문":
                if full_text.strip():
                    for line in full_text.splitlines():
                        p = doc.add_paragraph()
                        add_rich_text(p, line)
                else:
                    doc.add_paragraph("(설교 전문 미입력)")

            # 공통: 설명(스토리보드) — 마크업 반영
            p = doc.add_paragraph()
            p.add_run("설명(스토리보드): ")
            if description.strip():
                add_rich_text(p, description)
            else:
                p.add_run("(미입력)")

            doc.add_paragraph("")

    # 최종 저장
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.read()

# ---------------------------
# GitHub 저장/불러오기 유틸
# ---------------------------
def _gh_headers():
    return {
        "Authorization": f"token {st.secrets['GITHUB_TOKEN']}",
        "Accept": "application/vnd.github+json",
    }

def _gh_api_base():
    owner = st.secrets["GITHUB_OWNER"]
    repo = st.secrets["GITHUB_REPO"]
    return f"https://api.github.com/repos/{owner}/{repo}"

def gh_put_bytes(path: str, content_bytes: bytes, message: str):
    """
    GitHub Contents API로 파일 생성/업데이트
    """
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    # 기존 sha 조회(업데이트 대비)
    get = requests.get(url, headers=_gh_headers())
    sha = get.json().get("sha") if get.status_code == 200 else None

    b64 = base64.b64encode(content_bytes).decode("utf-8")
    payload = {
        "message": message,
        "content": b64,
        "branch": st.secrets.get("GITHUB_BRANCH", "main"),
    }
    if sha:
        payload["sha"] = sha

    r = requests.put(url, headers=_gh_headers(), json=payload)
    if r.status_code not in (200, 201):
        raise RuntimeError(f"GitHub 업로드 실패: {r.status_code} {r.text}")
    return r.json()

def gh_get_bytes(path: str) -> bytes:
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    r = requests.get(url, headers=_gh_headers())
    if r.status_code != 200:
        raise FileNotFoundError(f"GitHub 파일 없음: {path}")
    content = r.json()["content"]
    return base64.b64decode(content)

def gh_list_dir(path: str):
    api = _gh_api_base()
    url = f"{api}/contents/{path}"
    r = requests.get(url, headers=_gh_headers())
    if r.status_code != 200:
        return []  # 폴더 없을 수 있음(초기)
    return r.json()  # list of items (name, path, type, sha ...)

# 제출 데이터 직렬화/역직렬화
def serialize_submission():
    return {
        "worship_date": str(st.session_state.get("worship_date")),
        "services": st.session_state.get("services_selected", []),
        "materials": st.session_state.get("materials", []),
        "user_name": st.session_state.get("user_name"),
        "position": st.session_state.get("position"),
        "role": st.session_state.get("role"),
        "saved_at": datetime.now(timezone.utc).isoformat()
    }

def load_into_session(payload: dict):
    st.session_state.worship_date = date.fromisoformat(payload.get("worship_date"))
    st.session_state.services_selected = payload.get("services", [])
    st.session_state.materials = payload.get("materials", [])
    # 작성자/권한은 현재 세션값 유지

# 경로 규칙 (draft / submitted)
def gh_paths(user_name: str, worship_date: date, submission_id: str = None):
    base = st.secrets.get("GITHUB_BASE_DIR", "worship_submissions")
    d = worship_date.strftime("%Y-%m-%d")
    safe_user = (user_name or "unknown").strip().replace("/", "_")
    sub_id = "draft" if submission_id is None else submission_id
    folder = f"{base}/{d}/{safe_user}/{sub_id}"
    return {
        "json": f"{folder}/submission.json",
        "docx": f"{folder}/submission.docx"
    }

# ---------------------------
# ① 날짜/예배 선택
# ---------------------------
st.markdown("<div class='section-title'>① 날짜/예배 선택</div>", unsafe_allow_html=True)

# 세션의 worship_date 초기화(안전 가드)
if "worship_date" not in st.session_state:
    st.session_state.worship_date = date.today()

worship_date = st.date_input(
    "예배 날짜",
    value=st.session_state.worship_date,
    format="YYYY-MM-DD",
    disabled=not can_edit
)

# 최신 선택값을 세션에 반영 (← 여기 순서가 매우 중요!)
st.session_state.worship_date = worship_date

# 예배 구분: 기본 옵션 + 직접 입력 추가
col_sv1, col_sv2 = st.columns([2, 1])
with col_sv1:
    st.session_state.services_selected = st.multiselect(
        "예배 구분 선택",
        options=st.session_state.services_options,
        default=st.session_state.services_selected,
        help="해당 날짜에 해당되는 예배를 모두 선택하세요.",
        disabled=not can_edit
    )
with col_sv2:
    new_service = st.text_input("직접 입력", placeholder="예: 청년예배 / 새벽기도", disabled=not can_edit)
    add_new = st.button("추가", disabled=not can_edit)
    if add_new and new_service.strip():
        if new_service not in st.session_state.services_options:
            st.session_state.services_options.append(new_service.strip())
        if new_service not in st.session_state.services_selected:
            st.session_state.services_selected.append(new_service.strip())
        st.rerun()

services = st.session_state.services_selected

# ---------------------------
# ③ 자료 추가 (순서 조절 포함)
# ---------------------------
st.markdown("<div class='section-title'>② 자료 추가 (성경/이미지/기타/설교 전문)</div>", unsafe_allow_html=True)
st.caption("• 설명(스토리보드)에서 **굵게**, ==형광펜== 으로 강조하면 Word에 그대로 반영됩니다.")

add_btn = st.button("+ 자료 추가", disabled=not can_edit)
if add_btn and can_edit:
    add_material()

to_remove: List[str] = []
for i, item in enumerate(st.session_state.materials):
    with st.container(border=True):
        # 상단: 유형 선택 + 순서/삭제
        top_cols = st.columns([1.2, 0.2, 0.2, 0.2])
        with top_cols[0]:
            item["kind"] = st.selectbox(
                "자료 유형",
                ["성경 구절", "이미지", "기타 파일", "설교 전문"],
                index=["성경 구절", "이미지", "기타 파일", "설교 전문"].index(item.get("kind", "성경 구절")),
                key=f"kind_{item['id']}",
                disabled=not can_edit
            )
        with top_cols[1]:
            st.write("")
            st.button("▲", key=f"up_{item['id']}", disabled=(not can_edit or i == 0), on_click=move_material, args=(item["id"], "up"))
        with top_cols[2]:
            st.write("")
            st.button("▼", key=f"down_{item['id']}", disabled=(not can_edit or i == len(st.session_state.materials)-1), on_click=move_material, args=(item["id"], "down"))
        with top_cols[3]:
            st.write("")
            if st.button("삭제", key=f"del_{item['id']}", disabled=not can_edit):
                to_remove.append(item["id"])

        # 본문 입력 영역
        if item["kind"] == "성경 구절":
            item["verse_text"] = st.text_area(
                "성경 구절 입력 (예: 요한복음 3:16)",
                value=item.get("verse_text", ""),
                key=f"verse_{item['id']}",
                height=120,
                disabled=not can_edit
            )
            item["files"] = []
            item["file"] = None

        elif item["kind"] == "이미지":
            # 다중 업로드
            item["files"] = st.file_uploader(
                "이미지 업로드 (PNG/JPG) — 여러 장 선택 가능",
                type=["png", "jpg", "jpeg"],
                key=f"files_{item['id']}",
                accept_multiple_files=True,
                disabled=not can_edit
            )
            item["verse_text"] = ""
            item["file"] = None  # 단일 파일은 사용 안 함(과거 호환만)

        elif item["kind"] == "기타 파일":
            item["file"] = st.file_uploader(
                "기타 파일 업로드",
                type=None,
                key=f"file_{item['id']}",
                accept_multiple_files=False,
                disabled=not can_edit
            )
            item["verse_text"] = ""
            item["files"] = []

        elif item["kind"] == "설교 전문":
            item["full_text"] = st.text_area(
                "설교 전문 입력 (줄바꿈 유지 / **굵게**, ==형광펜== 지원)",
                value=item.get("full_text", ""),
                key=f"full_{item['id']}",
                height=300,
                disabled=not can_edit
            )
            item["verse_text"] = ""
            item["files"] = []
            item["file"] = None

        # 설명(스토리보드): 마크업 안내
        item["description"] = st.text_area(
            "설명(스토리보드)",
            value=item.get("description", ""),
            key=f"desc_{item['id']}",
            height=100,
            placeholder="노출 타이밍, 강조 부분 등. **굵게**, ==형광펜== 으로 강조 가능합니다.",
            disabled=not can_edit
        )

# 실제 삭제 처리
if to_remove and can_edit:
    for rid in to_remove:
        remove_material(rid)

st.divider()

# ---------------------------
# ② 저장/불러오기/제출 (순서 변경)
# ---------------------------
st.markdown("#### 저장/제출")
btn_cols = st.columns([1, 1, 1, 2])
with btn_cols[0]:
    save_draft = st.button("💾 임시 저장", disabled=not can_edit)
with btn_cols[1]:
    load_draft = st.button("↩️ 불러오기(임시저장)", disabled=not can_edit)
with btn_cols[2]:
    submit_now = st.button("✅ 제출", disabled=not can_edit)  # 제출 후 미디어부가 확인

# 제출 ID(제출 시 고정)
if "submission_id" not in st.session_state:
    st.session_state.submission_id = None

if save_draft and can_edit:
    try:
        data = serialize_submission()
        p = gh_paths(st.session_state.user_name, worship_date)  # draft
        gh_put_bytes(
            p["json"],
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
            message=f"[draft] {st.session_state.user_name} {worship_date} 저장"
        )
        st.success("임시 저장되었습니다. (GitHub)")
    except Exception as e:
        st.error(f"임시 저장 실패: {e}")

if load_draft and can_edit:
    try:
        p = gh_paths(st.session_state.user_name, worship_date)  # draft
        draft_bytes = gh_get_bytes(p["json"])
        payload = json.loads(draft_bytes.decode("utf-8"))
        load_into_session(payload)
        st.success("임시 저장본을 불러왔습니다.")
        st.rerun()
    except Exception as e:
        st.error(f"불러오기 실패 또는 저장본 없음: {e}")

if submit_now and can_edit:
    try:
        # 1) docx 생성
        docx_bytes = build_docx(
            worship_date=worship_date,
            services=st.session_state.services_selected,
            materials=st.session_state.materials,
            user_name=st.session_state.user_name,
            position=st.session_state.position,
            role=st.session_state.role
        )
        # 2) JSON + DOCX 업로드 (제출용 고유 ID 생성)
        sub_id = st.session_state.submission_id or datetime.now().strftime("%H%M%S") + "-" + uuid.uuid4().hex[:6]
        st.session_state.submission_id = sub_id
        p = gh_paths(st.session_state.user_name, worship_date, submission_id=sub_id)

        data = serialize_submission()
        data["status"] = "submitted"
        data["submission_id"] = sub_id

        gh_put_bytes(
            p["json"],
            json.dumps(data, ensure_ascii=False).encode("utf-8"),
            message=f"[submit] {st.session_state.user_name} {worship_date} 제출"
        )
        gh_put_bytes(
            p["docx"],
            docx_bytes,
            message=f"[submit-docx] {st.session_state.user_name} {worship_date} DOCX"
        )

        st.success("제출 완료! 미디어부 화면에서 확인 가능합니다.")
    except Exception as e:
        st.error(f"제출 실패: {e}")

st.divider()
# ---------------------------
# ④ 업로드(Word 저장)
# ---------------------------
col1, col2 = st.columns([1, 2])
with col1:
    do_save = st.button("📄 업로드 하기 (Word 저장)", type="primary", disabled=not can_edit)

if do_save and can_edit:
    try:
        docx_bytes = build_docx(
            worship_date=worship_date,
            services=services,
            materials=st.session_state.materials,
            user_name=st.session_state.user_name,
            position=st.session_state.position,
            role=st.session_state.role
        )
        filename = f"설교자료_{worship_date.strftime('%Y%m%d')}_{'-'.join(services) if services else '미지정'}.docx"
        st.success("Word 파일이 생성되었습니다. 아래 버튼으로 다운로드하세요.")
        st.download_button(
            "⬇️ Word 파일 다운로드",
            data=docx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
    except Exception as e:
        st.error(f"문서 생성 중 오류가 발생했습니다: {e}")

# ---------------------------
# ⑤ 미디어부 전용 제출함 (유틸/날짜 정의 이후여야 함)
# ---------------------------
if st.session_state.role == "미디어부":
    st.divider()
    st.markdown("### 📬 제출함(미디어부) — 날짜별/제출자별 목록")

    base = st.secrets.get("GITHUB_BASE_DIR", "worship_submissions")
    days = gh_list_dir(base)
    if not days:
        st.info("아직 제출된 자료가 없습니다.")
    else:
        day_names = sorted([d["name"] for d in days if d["type"] == "dir"], reverse=True)
        sel_day = st.selectbox("날짜 선택", options=day_names)
        if sel_day:
            day_dir = f"{base}/{sel_day}"
            users = gh_list_dir(day_dir) or []
            for u in users:
                if u.get("type") != "dir":
                    continue
                with st.expander(f"👤 {u['name']} — {sel_day} 제출물들"):
                    subs = gh_list_dir(u["path"]) or []  # submission_id 디렉토리들
                    for s in subs:
                        if s.get("type") != "dir":
                            continue
                        files = gh_list_dir(s["path"]) or []
                        json_item = next((f for f in files if f.get("name") == "submission.json"), None)
                        docx_item = next((f for f in files if f.get("name") == "submission.docx"), None)

                        cols = st.columns([2, 1, 2])
                        with cols[0]:
                            st.markdown(f"**제출 ID:** {s['name']}")
                        with cols[1]:
                            if json_item:
                                try:
                                    payload = json.loads(gh_get_bytes(json_item["path"]).decode("utf-8"))
                                    info = (
                                        f"- 예배: {', '.join(payload.get('services', [])) or '(미지정)'}\n"
                                        f"- 자료개수: {len(payload.get('materials', []))}\n"
                                        f"- 제출시각(UTC): {payload.get('saved_at','')}\n"
                                    )
                                    st.caption(info)
                                except Exception:
                                    st.caption("메타 로드 실패")
                            else:
                                st.caption("메타 없음")

                        with cols[2]:
                            if docx_item:
                                try:
                                    docx_bytes = gh_get_bytes(docx_item["path"])
                                    st.download_button(
                                        "⬇️ Word 다운로드",
                                        data=docx_bytes,
                                        file_name=f"설교자료_{sel_day}_{u['name']}_{s['name']}.docx",
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        key=f"dl_{sel_day}_{u['name']}_{s['name']}"
                                    )
                                except Exception as e:
                                    st.error(f"다운로드 오류: {e}")
                            else:
                                # JSON만 있고 DOCX 없는 제출물도 고려 → 즉석 변환
                                if json_item and Document is not None:
                                    if st.button("📄 즉석 Word 생성", key=f"mk_{sel_day}_{u['name']}_{s['name']}"):
                                        try:
                                            payload = json.loads(gh_get_bytes(json_item["path"]).decode("utf-8"))
                                            docx_bytes2 = build_docx(
                                                worship_date=date.fromisoformat(sel_day),
                                                services=payload.get("services", []),
                                                materials=payload.get("materials", []),
                                                user_name=payload.get("user_name"),
                                                position=payload.get("position"),
                                                role=payload.get("role")
                                            )
                                            st.download_button(
                                                "⬇️ Word 다운로드(즉석)",
                                                data=docx_bytes2,
                                                file_name=f"설교자료_{sel_day}_{u['name']}_{s['name']}.docx",
                                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                                key=f"dl2_{sel_day}_{u['name']}_{s['name']}"
                                            )
                                        except Exception as e:
                                            st.error(f"생성 오류: {e}")

# ---------------------------
# 풋터
# ---------------------------
st.markdown(
    """
    <hr/>
    <div class='small-note'>
    ⚙️ 이미지 외의 기타 파일은 Word에 직접 삽입되지 않으며, 파일명과 설명이 기록됩니다.<br>
    🖨️ 출력은 Word에서 페이지 여백/서식을 조정해 인쇄하면 보기 좋습니다.<br>
    ✍️ 강조법: **굵게**, ==형광펜== (Word 변환 시 자동 적용)
    </div>
    """,
    unsafe_allow_html=True
)
