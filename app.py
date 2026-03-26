import os
import streamlit as st
from src.optimizer import optimize_document, parse_slides
from src.image_generator import generate_slide_image
from src.pdf_builder import build_pdf
from src.ppt_generator import (
    generate_slide_code, build_single_slide_pptx, build_full_pptx,
    save_slide_code, load_slide_code, load_all_slide_codes,
    get_single_pptx_path,
)

PROJECTS_DIR = os.path.join(os.path.dirname(__file__), "projects")
os.makedirs(PROJECTS_DIR, exist_ok=True)

st.set_page_config(page_title="Our's NotebookLM", layout="wide")
st.title("Our's NotebookLM")

# ── Sidebar: project management ──────────────────────────────────────
with st.sidebar:
    st.header("项目管理")

    existing = sorted(
        d for d in os.listdir(PROJECTS_DIR)
        if os.path.isdir(os.path.join(PROJECTS_DIR, d))
    )

    new_name = st.text_input("新建项目名称")
    if st.button("创建项目") and new_name:
        proj = os.path.join(PROJECTS_DIR, new_name)
        for sub in ["原文档/images", "优化PP页文档", "生成的图片", "最终文档"]:
            os.makedirs(os.path.join(proj, sub), exist_ok=True)
        st.session_state["selected_project"] = new_name
        st.rerun()

    if not existing:
        st.info("请先创建一个项目")
        st.stop()

    default_idx = 0
    if "selected_project" in st.session_state and st.session_state["selected_project"] in existing:
        default_idx = existing.index(st.session_state["selected_project"])

    project_name = st.selectbox("选择项目", existing, index=default_idx)
    proj_dir = os.path.join(PROJECTS_DIR, project_name)

    st.divider()
    st.markdown("""
**🚀 功能说明**

1. **智能优化** — 粘贴原始文本，AI 自动拆页、提炼要点
2. **风格生成** — 根据内容主题生成统一视觉风格
3. **信息图渲染** — 逐页生成专业信息图幻灯片
4. **导出** — 合并为 PDF / AI 生成可编辑 PPT
    """)

    text_model = "gemini-3-flash-preview"
    image_model = "gemini-3-pro-image-preview"
    ppt_model = "gemini-3-pro-preview"

# ── Helper paths ─────────────────────────────────────────────────────
raw_path = os.path.join(proj_dir, "原文档", "原稿.md")
opt_path = os.path.join(proj_dir, "优化PP页文档", "优化稿.md")
style_path = os.path.join(proj_dir, "优化PP页文档", "ppt样式风格描述.md")
img_dir = os.path.join(proj_dir, "生成的图片")
pdf_path = os.path.join(proj_dir, "最终文档", f"{project_name}.pdf")
ppt_path = os.path.join(proj_dir, "最终文档", f"{project_name}.pptx")


def read_file(path: str) -> str:
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return ""


def write_file(path: str, content: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)


# ── Step 1: Raw document ─────────────────────────────────────────────
st.header("Step 1: 原稿编辑")

raw_text = st.text_area(
    "输入或粘贴原始文档 (Markdown)",
    value=read_file(raw_path),
    height=300,
    key="raw_text",
)

col1, col2 = st.columns(2)
with col1:
    if st.button("保存原稿"):
        write_file(raw_path, raw_text)
        st.success("已保存")

with col2:
    if st.button("生成优化稿", type="primary"):
        if not raw_text.strip():
            st.warning("请先输入原稿内容")
        else:
            write_file(raw_path, raw_text)
            with st.spinner("正在生成优化稿和风格描述..."):
                opt_md, sty_md = optimize_document(raw_text, model=text_model)
            write_file(opt_path, opt_md)
            write_file(style_path, sty_md)
            st.success("优化稿生成完成！")
            st.rerun()

# ── Step 2: Optimized document & style ───────────────────────────────
st.header("Step 2: 优化稿 & 风格描述")

tab_opt, tab_style = st.tabs(["优化稿", "风格描述"])

with tab_opt:
    opt_text = st.text_area(
        "优化稿 (可编辑)",
        value=read_file(opt_path),
        height=400,
        key="opt_text",
    )
    if st.button("保存优化稿"):
        write_file(opt_path, opt_text)
        st.success("已保存")

with tab_style:
    style_text = st.text_area(
        "PPT样式风格描述 (可编辑)",
        value=read_file(style_path),
        height=300,
        key="style_text",
    )
    if st.button("保存风格描述"):
        write_file(style_path, style_text)
        st.success("已保存")

# ── Step 3: Generate images ──────────────────────────────────────────
st.header("Step 3: 生成信息图")

current_opt = read_file(opt_path)
current_style = read_file(style_path)

if current_opt:
    slides = parse_slides(current_opt)
    st.info(f"共解析出 {len(slides)} 页幻灯片")

    if st.button("一键生成所有图片", type="primary"):
        os.makedirs(img_dir, exist_ok=True)
        progress = st.progress(0)
        for i, slide in enumerate(slides):
            with st.spinner(f"正在生成第 {i+1}/{len(slides)} 页..."):
                img_bytes = generate_slide_image(
                    slide, current_style, i + 1, len(slides), model=image_model
                )
                if img_bytes:
                    img_path = os.path.join(img_dir, f"{i+1:02d}.jpg")
                    with open(img_path, "wb") as f:
                        f.write(img_bytes)
            progress.progress((i + 1) / len(slides))
        st.success("所有图片生成完成！")
        st.rerun()

    # Show existing images and allow per-page regeneration
    existing_imgs = sorted(
        f for f in os.listdir(img_dir)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    ) if os.path.exists(img_dir) else []

    if existing_imgs:
        cols_per_row = 3
        for row_start in range(0, len(existing_imgs), cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col in enumerate(cols):
                idx = row_start + j
                if idx >= len(existing_imgs):
                    break
                img_file = existing_imgs[idx]
                img_full = os.path.join(img_dir, img_file)
                with col:
                    st.image(img_full, caption=img_file, use_container_width=True)
                    page_idx = idx  # index into slides list
                    if page_idx < len(slides):
                        with st.expander(f"重新生成 {img_file}"):
                            custom_prompt = st.text_area(
                                "自定义该页内容（可选）",
                                value=slides[page_idx],
                                key=f"regen_{idx}",
                                height=150,
                            )
                            if st.button("重新生成", key=f"btn_regen_{idx}"):
                                with st.spinner("重新生成中..."):
                                    new_bytes = generate_slide_image(
                                        custom_prompt,
                                        current_style,
                                        page_idx + 1,
                                        len(slides),
                                        model=image_model,
                                    )
                                    if new_bytes:
                                        with open(img_full, "wb") as f:
                                            f.write(new_bytes)
                                        st.success("已重新生成")
                                        st.rerun()
else:
    st.info("请先生成优化稿")

# ── Step 4: Export (PDF / PPT) ───────────────────────────────────────
st.header("Step 4: 导出")

has_images = (
    os.path.exists(img_dir)
    and any(f.endswith((".jpg", ".png")) for f in os.listdir(img_dir))
)

tab_pdf, tab_ppt = st.tabs(["合并为 PDF", "生成 PPT"])

# ── Tab 1: PDF ──
with tab_pdf:
    if has_images:
        if st.button("合并为 PDF", type="primary"):
            with st.spinner("正在合并..."):
                build_pdf(img_dir, pdf_path)
            st.success("PDF 生成完成！")
            st.rerun()

        if os.path.exists(pdf_path):
            with open(pdf_path, "rb") as f:
                st.download_button(
                    "下载 PDF",
                    data=f.read(),
                    file_name=f"{project_name}.pdf",
                    mime="application/pdf",
                )
    else:
        st.info("请先生成图片")

# ── Tab 2: PPT ──
with tab_ppt:
    if not has_images:
        st.info("请先生成图片")
    else:
        # Collect image files
        ppt_img_files = sorted(
            f for f in os.listdir(img_dir)
            if f.lower().endswith((".jpg", ".jpeg", ".png"))
        )
        total_pages = len(ppt_img_files)

        # Load existing slide codes
        saved_codes = load_all_slide_codes(proj_dir)

        st.caption(
            f"共 {total_pages} 页信息图。AI 将逐页分析图片并生成 python-pptx 代码，"
            "最终合并为可编辑的 PPTX 文件。"
        )

        # ── One-click full generation ──
        if st.button("一键生成完整 PPT", type="primary", key="btn_full_ppt"):
            progress = st.progress(0)
            status = st.empty()
            all_codes = {}

            for i, img_file in enumerate(ppt_img_files):
                page = i + 1
                status.text(f"正在分析第 {page}/{total_pages} 页...")
                img_full = os.path.join(img_dir, img_file)
                code = generate_slide_code(
                    image_path=img_full,
                    page_num=page,
                    total_pages=total_pages,
                    model=ppt_model,
                )
                save_slide_code(proj_dir, page, code)
                all_codes[page] = code

                # Also build single-page pptx
                single_path = get_single_pptx_path(proj_dir, page)
                build_single_slide_pptx(code, single_path)

                progress.progress(page / total_pages)

            status.text("正在合并为完整 PPTX...")
            success, error = build_full_pptx(all_codes, ppt_path)
            status.empty()
            if success:
                st.success("完整 PPT 生成完成！")
                st.rerun()
            else:
                st.error("PPT 生成失败")
                with st.expander("错误详情"):
                    st.code(error)

        st.divider()

        # ── Per-page grid: checkbox + status + regenerate ──
        st.subheader("逐页管理")

        cols_per_row = 3
        for row_start in range(0, total_pages, cols_per_row):
            cols = st.columns(cols_per_row)
            for j, col in enumerate(cols):
                idx = row_start + j
                if idx >= total_pages:
                    break
                page = idx + 1
                img_file = ppt_img_files[idx]
                img_full = os.path.join(img_dir, img_file)
                has_code = load_slide_code(proj_dir, page) is not None
                single_pptx = get_single_pptx_path(proj_dir, page)
                has_pptx = os.path.exists(single_pptx)

                with col:
                    st.image(img_full, caption=f"第 {page} 页", use_container_width=True)

                    if has_pptx:
                        st.caption("已生成")
                    else:
                        st.caption("未生成")

                    # AI generate / regenerate
                    if st.button(
                        "AI 重新生成代码" if has_code else "AI 生成代码",
                        key=f"btn_ppt_page_{page}",
                    ):
                        with st.spinner(f"正在生成第 {page} 页..."):
                            code = generate_slide_code(
                                image_path=img_full,
                                page_num=page,
                                total_pages=total_pages,
                                model=ppt_model,
                            )
                            save_slide_code(proj_dir, page, code)
                            ok, err = build_single_slide_pptx(
                                code, single_pptx
                            )
                        if ok:
                            st.success(f"第 {page} 页生成完成")
                            st.rerun()
                        else:
                            st.error(f"第 {page} 页生成失败")
                            with st.expander("错误详情"):
                                st.code(err)

                    # Editable code + run from code
                    if has_code:
                        with st.expander("编辑代码 / 生成 PPT 页"):
                            edited = st.text_area(
                                "代码",
                                value=load_slide_code(proj_dir, page),
                                height=300,
                                key=f"code_edit_{page}",
                            )
                            c1, c2 = st.columns(2)
                            with c1:
                                if st.button("保存代码", key=f"btn_save_code_{page}"):
                                    save_slide_code(proj_dir, page, edited)
                                    st.success("已保存")
                            with c2:
                                if st.button("生成 PPT 页", key=f"btn_run_code_{page}", type="primary"):
                                    save_slide_code(proj_dir, page, edited)
                                    with st.spinner("正在生成..."):
                                        ok, err = build_single_slide_pptx(
                                            edited, single_pptx
                                        )
                                    if ok:
                                        st.success(f"第 {page} 页 PPT 生成完成")
                                        st.rerun()
                                    else:
                                        st.error("生成失败")
                                        st.code(err)

        st.divider()

        # ── Merge existing single-page codes into full PPTX ──
        saved_codes = load_all_slide_codes(proj_dir)
        if saved_codes:
            st.subheader("合并为完整 PPT")
            st.info(f"已有 {len(saved_codes)}/{total_pages} 页代码")

            if st.button("合并已有页面为完整 PPT", key="btn_merge_ppt"):
                with st.spinner("正在合并..."):
                    success, error = build_full_pptx(saved_codes, ppt_path)
                if success:
                    st.success("合并完成！")
                    st.rerun()
                else:
                    st.error("合并失败")
                    with st.expander("错误详情"):
                        st.code(error)

        # ── Download ──
        if os.path.exists(ppt_path):
            with open(ppt_path, "rb") as f:
                st.download_button(
                    "下载完整 PPT",
                    data=f.read(),
                    file_name=f"{project_name}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    key="btn_download_full_ppt",
                )
