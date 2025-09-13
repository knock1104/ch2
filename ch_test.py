# -*- coding: utf-8 -*-
"""
Streamlit 예배 자료 업로드 + Word 저장 (개편 버전)
요청 반영:
1) 예배 구분 직접기입
2) 자료 순서 조절
3) 라이브 프리뷰 삭제
4) 이미지 여러 장 업로드
5) 설명 강조(**굵게**, ==형광펜==) -> Word 반영
6) 자료 유형 '설교 전문' 추가
"""

import io
import os
import uuid
import base64
import tempfile
import re
from datetime import date
from typing import List, Dict, Any

import streamlit as st

# python-docx 관련 모듈 로드 (미설치 시 안내)
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
# 페이지 설정 & 스타일
# ---------------------------
st.set_page_config(
    page_title="설교 자료 업로드",
    page_icon="🙏",
    layout="wide"
)

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
    st.session_state.preview_idx = 0  # (프리뷰는 제거되었지만 호환을 위해 유지)

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
# 유틸 함수
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
# ① 날짜/예배 선택
# ---------------------------
st.markdown("<div class='section-title'>① 날짜/예배 선택</div>", unsafe_allow_html=True)
worship_date = st.date_input(
    "예배 날짜",
    value=date.today(),
    format="YYYY-MM-DD",
    disabled=not can_edit
)

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

st.divider()

# ---------------------------
# ② 자료 추가 (순서 조절 포함)
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
# 업로드(Word 저장)
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

# 풋터
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
