# -*- coding: utf-8 -*-
"""
Streamlit 예배 자료 업로드 + 라이브 프리뷰(Stage) 앱 (권한/랜딩페이지 + 개편 버전)

기능 개요
0) 랜딩 페이지에서 역할/이름/직분/액세스 코드 입력 후 입장
   - 교역자: 작성 및 수정 가능
   - 미디어부: 작성 및 수정 불가(읽기 전용 확인만 가능)
   - [테스트 안내] 현재는 모든 액세스 코드 0001로 입장 가능
1) 달력에서 날짜 선택
2) 자료 추가(성경 구절 / 이미지 / 기타) + 각 자료에 대한 설명(스토리보드)
   - 설명(스토리보드) placeholder: "해당 자료의 노출 타이밍, 강조를 원하시는 부분 등을 적어주세요."
3) "업로드 하기" 클릭 시 Word(.docx) 파일 생성 및 다운로드
4) 🔎 입력 내용을 4608:2240 자막 화면으로 미리보기(이미지 배경 + 하단 자막)

로컬 실행 방법 (VS Code 권장)
- Python 3.9+ 권장
- pip install streamlit python-docx pillow
- streamlit run app.py
"""

import io
import os
import uuid
import base64
import tempfile
from datetime import date
from typing import List, Dict, Any

import streamlit as st

# python-docx 관련 모듈 로드 (미설치 시 안내)
try:
    from docx import Document
    from docx.shared import Inches, Pt
    from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    # 각 항목: {id, kind, file, verse_text, description}
    st.session_state.materials: List[Dict[str, Any]] = []

if "preview_idx" not in st.session_state:
    st.session_state.preview_idx = 0

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

# ---------------------------
# 유틸 함수
# ---------------------------
def add_material():
    st.session_state.materials.append({
        "id": str(uuid.uuid4()),
        "kind": "성경 구절",
        "file": None,
        "verse_text": "",
        "description": ""
    })

def remove_material(mid: str):
    st.session_state.materials = [m for m in st.session_state.materials if m["id"] != mid]

def _b64_of_upload(file) -> str:
    """UploadedFile -> data:image/...;base64,.... 문자열"""
    if file is None:
        return ""
    mime = "image/png"
    ext = os.path.splitext(file.name)[1].lower()
    if ext in [".jpg", ".jpeg"]:
        mime = "image/jpeg"
    b64 = base64.b64encode(file.getvalue()).decode("utf-8")
    return f"data:{mime};base64,{b64}"

def build_preview_frames(materials):
    """
    프리뷰(자막)용 '장면' 리스트 구성.
    - 성경 구절: 그 문장 자체가 자막
    - 이미지: 이미지를 배경으로, 설명(스토리보드)을 자막
    - 기타 파일: 배경 없음, 설명을 자막
    """
    frames = []

    for item in materials:
        kind = item.get("kind")
        desc = (item.get("description") or "").strip()
        verse = (item.get("verse_text") or "").strip()
        file  = item.get("file")

        if kind == "성경 구절":
            lines = [ln.strip() for ln in verse.splitlines() if ln.strip()]
            if not lines:
                lines = ["(성경 구절 미입력)"]
            for ln in lines:
                frames.append({"bg":"", "caption": ln})
            if desc:
                frames.append({"bg":"", "caption": desc})

        elif kind == "이미지":
            bg = _b64_of_upload(file) if file else ""
            cap = desc or "(설명 없음)"
            frames.append({"bg": bg, "caption": cap})

        else:  # 기타 파일
            cap = (f"[첨부] {file.name} — " if file else "") + (desc or "(설명 없음)")
            frames.append({"bg": "", "caption": cap})

    if not frames:
        frames = [{"bg":"", "caption":"(자막 미리보기 없음) 자료나 텍스트를 입력하세요."}]

    return frames

def build_docx(
    worship_date: date,
    services: List[str],
    materials: List[Dict[str, Any]],
    user_name: str,
    position: str,
    role: str,
) -> bytes:
    """문서(.docx) 생성 후 바이트로 반환 (설교 전문 섹션 삭제된 버전)"""
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
            verse_text = item.get("verse_text", "")
            description = item.get("description", "")
            file = item.get("file")

            doc.add_heading(f"{idx}. {kind}", level=2)

            if kind == "성경 구절":
                if verse_text.strip():
                    for line in verse_text.splitlines():
                        doc.add_paragraph(line)
                    doc.add_paragraph("")
                else:
                    doc.add_paragraph("(성경 구절 미입력)")

            elif kind == "이미지":
                if file is not None and Image is not None:
                    try:
                        # 업로드된 이미지 바이트를 임시 파일로 저장 후 삽입
                        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.name)[1]) as tmp:
                            tmp.write(file.getvalue())
                            tmp.flush()
                            # 가로 너비 5인치로 리사이즈(비율 유지)
                            doc.add_picture(tmp.name, width=Inches(5))
                    except Exception:
                        doc.add_paragraph(f"(이미지 삽입 실패) 파일명: {file.name}")
                else:
                    doc.add_paragraph("(이미지 파일 없음)")

            else:  # 기타 파일
                if file is not None:
                    doc.add_paragraph(f"첨부 파일: {file.name} (문서에 직접 삽입되지 않습니다)")
                else:
                    doc.add_paragraph("(첨부 파일 없음)")

            if description.strip():
                doc.add_paragraph("설명(스토리보드): " + description)
            else:
                doc.add_paragraph("설명(스토리보드): (미입력)")

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
# 상단 사용자/권한 표시 (1열 구성 시작)
# ---------------------------
can_edit = st.session_state.get("can_edit", False)
role_badge = "🟢 편집 가능" if can_edit else "🔒 읽기 전용(확인만)"
st.markdown(
    f"**접속자:** {st.session_state.user_name or '이름 미입력'} "
    f"({st.session_state.position or '직분 미선택'}) · "
    f"{st.session_state.role} · {role_badge}"
)

# ---------------------------
# ① 날짜/예배 선택 (1열)
# ---------------------------
st.markdown("<div class='section-title'>① 날짜/예배 선택</div>", unsafe_allow_html=True)
worship_date = st.date_input(
    "예배 날짜",
    value=date.today(),
    format="YYYY-MM-DD",
    disabled=not can_edit
)

services = st.multiselect(
    "예배 구분 선택",
    options=["1부", "2부", "3부", "오후예배"],
    default=[],
    help="해당 날짜에 해당되는 예배를 모두 선택하세요.",
    disabled=not can_edit
)

st.divider()

# ---------------------------
# ② 자료 추가 (1열)  ← 요청 반영: 순서 변경 & 단일 컬럼
# ---------------------------
st.markdown("<div class='section-title'>② 자료 추가 (성경/이미지/기타)</div>", unsafe_allow_html=True)
st.caption("설교에 사용하실 자료를 업로드하세요. 각 자료별로 설명(스토리보드)을 작성할 수 있습니다.")

add_btn = st.button("+ 자료 추가", disabled=not can_edit)
if add_btn and can_edit:
    add_material()

# 각 자료 입력 블록 (세로로 나열)
to_remove: List[str] = []
for i, item in enumerate(st.session_state.materials):
    with st.container(border=True):
        # 유형 선택 & 삭제 버튼을 한 줄로 배치하고 싶으면 columns 사용 가능
        top_cols = st.columns([1, 0.2])
        with top_cols[0]:
            item["kind"] = st.selectbox(
                "자료 유형",
                ["성경 구절", "이미지", "기타 파일"],
                index=["성경 구절", "이미지", "기타 파일"].index(item.get("kind", "성경 구절")),
                key=f"kind_{item['id']}",
                disabled=not can_edit
            )
        with top_cols[1]:
            st.write("")  # spacing
            if st.button("삭제", key=f"del_{item['id']}", disabled=not can_edit):
                to_remove.append(item["id"])

        if item["kind"] == "성경 구절":
            item["verse_text"] = st.text_area(
                "성경 구절 입력 (예: 요한복음 3:16)",
                value=item.get("verse_text", ""),
                key=f"verse_{item['id']}",
                height=120,
                disabled=not can_edit
            )
            item["file"] = None

        elif item["kind"] == "이미지":
            item["file"] = st.file_uploader(
                "이미지 업로드 (PNG/JPG)",
                type=["png", "jpg", "jpeg"],
                key=f"file_{item['id']}",
                accept_multiple_files=False,
                disabled=not can_edit
            )
            item["verse_text"] = ""

        else:  # 기타 파일
            item["file"] = st.file_uploader(
                "기타 파일 업로드",
                type=None,
                key=f"file_{item['id']}",
                accept_multiple_files=False,
                disabled=not can_edit
            )
            item["verse_text"] = ""

        item["description"] = st.text_area(
            "설명(스토리보드)",
            value=item.get("description", ""),
            key=f"desc_{item['id']}",
            height=100,
            placeholder="해당 자료의 노출 타이밍, 강조를 원하시는 부분 등을 적어주세요.",
            disabled=not can_edit
        )

# 실제 삭제 처리
if to_remove and can_edit:
    for rid in to_remove:
        remove_material(rid)

st.divider()

# ---------------------------
# 🔎 라이브 프리뷰 (4608:2240)
# ---------------------------
st.markdown("<div class='section-title'>🔎 라이브 프리뷰 (4608:2240 자막 화면)</div>", unsafe_allow_html=True)

# 프리뷰 스타일 슬라이더 (읽기전용이라도 조작은 가능하도록 유지하려면 disabled=False로 두세요)
colA, colB = st.columns([1, 1])
with colA:
    fs = st.slider("글자 크기(px)", 24, 96, 48, help="자막 폰트 크기")
    lh = st.slider("줄간격(line-height)", 1.0, 2.0, 1.3, 0.1)
with colB:
    bottom_padding = st.slider("하단 여백(px)", 0, 200, 40)
    max_width = st.slider("자막 최대 너비(%)", 40, 100, 80)

# 장면 생성
preview_frames = build_preview_frames(st.session_state.materials)

# 네비게이션
nav1, nav2, nav3 = st.columns([1, 2, 1])
with nav1:
    if st.button("◀ 이전"):
        st.session_state.preview_idx = (st.session_state.preview_idx - 1) % len(preview_frames)
with nav3:
    if st.button("다음 ▶"):
        st.session_state.preview_idx = (st.session_state.preview_idx + 1) % len(preview_frames)

cur = st.session_state.preview_idx
st.caption(f"장면 {cur+1} / {len(preview_frames)}")

# 4608:2240 스테이지 (CSS aspect-ratio 사용)
stage = preview_frames[cur]
bg_style = "background:#000;"  # 기본 검은 배경
if stage["bg"]:
    bg_style = f"background: url({stage['bg']}) center/cover no-repeat, #000;"

stage_html = f"""
<div style="
  position:relative;
  width:100%;
  max-width: 1920px; /* 표시 폭 제한(필요 시 조정) */
  aspect-ratio: 4608 / 2240;
  margin: 0 auto;
  border-radius: 12px;
  overflow:hidden;
  box-shadow: 0 8px 24px rgba(0,0,0,0.35);
  {bg_style}
">
  <!-- 자막 박스 -->
  <div style="
    position:absolute; left:0; right:0; bottom:{bottom_padding}px;
    display:flex; justify-content:center;
  ">
    <div style="
      max-width:{max_width}%;
      padding: 12px 18px;
      background: rgba(0,0,0,0.6);
      color:#fff;
      font-weight:600;
      font-size:{fs}px;
      line-height:{lh};
      text-align:center;
      border-radius: 10px;
      text-shadow: 0 2px 8px rgba(0,0,0,0.85);
      white-space:pre-wrap;
    ">{stage["caption"]}</div>
  </div>
</div>
"""
st.markdown(stage_html, unsafe_allow_html=True)

with st.expander("전체 화면으로 띄우는 팁"):
    st.info("브라우저에서 [F11] 전체화면을 활용하거나, 앱을 새 창으로 열어 프리뷰만 크게 띄워서 송출 화면처럼 사용할 수 있습니다.")

st.divider()

# ---------------------------
# 업로드(Word 저장) 실행
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
    ⚙️ 팁: 이미지 외의 기타 파일은 Word에 직접 삽입되지 않으며, 파일명과 설명이 기록됩니다. 필요 시 Zip 묶음으로 함께 배포하세요.<br>
    🖨️ 출력은 Word에서 페이지 여백/서식을 조정해 인쇄하면 보기 좋습니다.
    </div>
    """,
    unsafe_allow_html=True
)
