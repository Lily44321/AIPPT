import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from openai import OpenAI
import json
import os
import tempfile
import docx
import openpyxl
import pandas as pd
from pptx import Presentation as PPTXReader
import PyPDF2
import io

# ==================== 1. 配置 DeepSeek 客户端 ====================
client = OpenAI(
    api_key=st.secrets["DEEPSEEK_API_KEY"],   # 从 secrets 读取
    base_url="https://api.deepseek.com"
)

# ==================== 2. 领域与细分映射 ====================
SUB_OPTIONS = {
    "具身智能": ["数据采集", "装调装配", "维修", "仿真训练", "场景开发"],
    "人工智能": ["AI Agent 开发", "算法", "数据标注与模型训练", "模型部署"],
    "低空经济": ["无人机教学", "无人机应用", "飞行控制", "无人机部署", "无人机维修"]
}

# ==================== 3. 文件解析函数（基于文件大小判断截断） ====================
def extract_text_from_file(uploaded_file):
    """根据文件扩展名提取文本内容，如果文件超过200MB则截断前3000字符"""
    ext = os.path.splitext(uploaded_file.name)[1].lower()
    text = ""
    try:
        if ext == ".txt":
            text = uploaded_file.getvalue().decode("utf-8")
        elif ext == ".docx":
            doc = docx.Document(uploaded_file)
            text = "\n".join([para.text for para in doc.paragraphs])
        elif ext in [".xls", ".xlsx"]:
            df = pd.read_excel(uploaded_file)
            text = df.to_string(index=False)
        elif ext == ".pptx":
            prs = PPTXReader(uploaded_file)
            slides_text = []
            for slide in prs.slides:
                slide_text = []
                for shape in slide.shapes:
                    if hasattr(shape, "text"):
                        slide_text.append(shape.text)
                slides_text.append("\n".join(slide_text))
            text = "\n\n".join(slides_text)
        elif ext == ".pdf":
            reader = PyPDF2.PdfReader(uploaded_file)
            text = "\n".join([page.extract_text() for page in reader.pages if page.extract_text()])
        elif ext == ".csv":
            df = pd.read_csv(uploaded_file)
            text = df.to_string(index=False)
        else:
            st.error(f"不支持的文件格式：{ext}")
            return None
    except Exception as e:
        st.error(f"文件读取失败：{e}")
        return None

    # ========== 修改点：根据文件大小决定是否截断 ==========
    file_size = uploaded_file.size  # 字节数
    max_size = 200 * 1024 * 1024    # 200MB
    if file_size > max_size:
        st.info("文件超过200MB，内容过长，已自动截取前3000字符作为参考。")
        if len(text) > 3000:
            text = text[:3000] + "...(内容已截断)"
    # 否则返回完整文本（不截断）
    return text

# ==================== 4. 网页标题 ====================
st.title("智能PPT生成器")
st.markdown("填写以下信息，AI将为您自动生成定制化PPT。")

# ==================== 5. 用户输入区 ====================
col1, col2 = st.columns(2)

with col1:
    school_name = st.text_input("学校名称", placeholder="请输入学校名称")
    school_level = st.selectbox(
        "学校层次",
        ["985", "211", "双一流", "重点一本", "公办本科", "民办本科", "专科"],
        index=None,
        placeholder="请选择学校层次"
    )

with col2:
    pass

domain = st.selectbox(
    "领域选择",
    list(SUB_OPTIONS.keys()),
    index=None,
    placeholder="请选择领域"
)

sub_options = SUB_OPTIONS.get(domain, []) if domain else []
sub_domain = st.selectbox(
    "细分方向",
    sub_options,
    index=None,
    placeholder="请先选择领域" if not domain else "请选择细分方向"
)

need_nvidia = st.selectbox(
    "是否体现英伟达技术/合作",
    ["是", "否"],
    index=None,
    placeholder="请选择"
)

budget = st.number_input("项目预算（万元）", min_value=0, value=None, step=10, placeholder="请输入预算")
cost = st.number_input("目标成本控制（万元）", min_value=0, value=None, step=10, placeholder="请输入成本")

uploaded_image = st.file_uploader("场地地形图片（可选）", type=["png", "jpg", "jpeg"])
if uploaded_image:
    st.image(uploaded_image, caption="预览上传图片", width=300)

uploaded_files = st.file_uploader(
    "上传参考资料（Word/PPT/Excel/PDF/TXT/CSV，可多选）",
    type=["docx", "pptx", "xls", "xlsx", "pdf", "txt", "csv"],
    accept_multiple_files=True
)

reference_texts = []
if uploaded_files:
    for file in uploaded_files:
        with st.spinner(f"正在读取 {file.name} ..."):
            content = extract_text_from_file(file)
            if content:
                reference_texts.append(f"【文件：{file.name}】\n{content}")

page_num = st.number_input(
    "生成页数", min_value=3, max_value=80, value=None, step=1, placeholder="请输入PPT页数"
)

# ==================== 6. 构造发送给 AI 的指令 ====================
def build_user_prompt(school_name, school_level, need_nvidia,
                      domain, sub_domain, budget, cost, page_num,
                      reference_texts=None):
    nvidia_text = need_nvidia if need_nvidia else "否"
    prompt = f"""请为以下需求生成一份PPT大纲：
- 学校名称：{school_name}
- 学校层次：{school_level if school_level else '未指定'}
- 核心领域：{domain if domain else '未指定'} - {sub_domain if sub_domain else '未指定'}
- 项目预算：{budget if budget is not None else '未指定'}万元，目标成本控制在{cost if cost is not None else '未指定'}万元
- 是否需要提及英伟达技术或合作：{nvidia_text}
- PPT页数要求：{page_num}页（包含封面和尾页）

请根据学校层次调整内容深度和专业方向。层次越高越强调科研合作与前沿技术，层次较低则侧重技能培训、就业对接等。
如果提到了英伟达，请自然融入相关技术或合作点。
"""
    if reference_texts:
        combined_ref = "\n\n".join(reference_texts)
        prompt += f"\n\n【以下为用户提供的参考资料，请严格基于这些内容生成PPT】\n{combined_ref}"
    return prompt

# ==================== 7. 调用 DeepSeek 生成大纲 ====================
def generate_ppt_outline(full_user_prompt):
    system_prompt = """你是一位专业的企业方案演示专家。请根据用户的详细需求，生成一份结构清晰、内容贴合的PPT大纲。
要求：
1. 返回严格的JSON格式，只包含JSON，不要有任何其他文字。
2. JSON结构：{"title": "...", "slides": [{"heading": "...", "content": ["要点1","要点2"]}, ...]}
3. slides的数量必须严格等于用户要求的页数。
4. 内容要针对学校层次、预算成本、领域细分进行定制，若需体现英伟达则自然融入。
5. 必须包含一页"场地地形"，heading为"场地地形"，content可以为"（此处插入场地地形图片）"等简要说明。
"""
    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": full_user_prompt}
            ],
            temperature=0.7,
            max_tokens=2500
        )
        raw = response.choices[0].message.content
        if raw.startswith("```json"):
            raw = raw[7:-3].strip()
        elif raw.startswith("```"):
            raw = raw[3:-3].strip()
        return json.loads(raw)
    except json.JSONDecodeError:
        st.error("AI 返回内容无法解析为 JSON，请重试。")
        st.code(raw)
        return None
    except Exception as e:
        st.error(f"调用 DeepSeek 出错: {e}")
        return None

# ==================== 8. 根据 JSON 生成 PPT 对象 ====================
def build_pptx_from_json(data):
    prs = Presentation()
    for i, slide_data in enumerate(data["slides"]):
        if i == 0:
            layout = prs.slide_layouts[0]
        else:
            layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(layout)
        slide.shapes.title.text = slide_data["heading"]
        if len(slide.placeholders) > 1:
            body = slide.placeholders[1]
            tf = body.text_frame
            tf.clear()
            for point in slide_data["content"]:
                p = tf.add_paragraph()
                p.text = point
                p.level = 1
    return prs

# ==================== 9. 在 PPT 中插入场地地形图片 ====================
def add_image_slide(prs, image_file):
    layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(layout)

    left, top, width, height = Inches(1), Inches(0.5), Inches(8), Inches(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = "场地地形"
    p.font.size = Pt(28)
    p.font.bold = True

    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(image_file.read())
        tmp_path = tmp.name

    slide.shapes.add_picture(tmp_path, Inches(1), Inches(1.8), width=Inches(8))
    return prs

# ==================== 10. 生成按钮与主逻辑 ====================
if st.button("生成 PPT"):
    if not school_name.strip():
        st.warning("请输入学校名称")
    elif page_num is None:
        st.warning("请输入PPT页数")
    elif page_num < 3:
        st.warning("页数至少为3页")
    else:
        with st.spinner("正在为您定制内容..."):
            user_prompt = build_user_prompt(
                school_name, school_level, need_nvidia,
                domain, sub_domain, budget, cost, page_num,
                reference_texts=reference_texts if uploaded_files else None
            )
            outline = generate_ppt_outline(user_prompt)

        if outline is not None:
            st.success("内容生成完毕，正在打包 PPT...")
            prs = build_pptx_from_json(outline)

            if uploaded_image is not None:
                prs = add_image_slide(prs, uploaded_image)

            ppt_path = "final_output.pptx"
            prs.save(ppt_path)

            with open(ppt_path, "rb") as f:
                st.download_button(
                    label="下载PPT",
                    data=f,
                    file_name=f"{school_name}_{domain}方案.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )