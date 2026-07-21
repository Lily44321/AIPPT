import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from openai import OpenAI
import json
import os
import tempfile

# ==================== 1. 配置 DeepSeek 客户端 ====================
client = OpenAI(
    api_key="sk-fba280a58b9d4d5b8e9fbcac4f62c6f5",  # 替换成你自己的 key
    base_url="https://api.deepseek.com"
)

# ==================== 2. 领域与细分映射 ====================
SUB_OPTIONS = {
    "具身智能": ["数据采集", "装调装配", "维修", "仿真训练", "场景开发"],
    "人工智能": ["数据标注", "模型训练", "模型部署", "AI安全", "AI应用开发"],
    "低空经济": ["无人机巡检", "低空物流", "空域管理", "飞控系统", "低空通信"]
}

# ==================== 3. 网页标题 ====================
st.title("智能 PPT 生成器")
st.markdown("填写以下信息，AI将为您自动生成定制化PPT。")

# ==================== 4. 用户输入区 ====================
# ==================== 4. 用户输入区 ====================
col1, col2 = st.columns(2)

with col1:
    school_name = st.text_input("学校名称")
    school_level = st.selectbox(
        "学校层次",
        ["985", "211", "双一流", "重点一本", "公办本科", "民办本科", "专科"]
    )

with col2:
    need_nvidia = st.checkbox("内容中体现英伟达技术/合作")
    # 页数输入框暂时不放在这里，后面单独放

domain = st.selectbox("领域选择", list(SUB_OPTIONS.keys()))
sub_domain = st.selectbox("细分方向", SUB_OPTIONS[domain])

budget = st.number_input("项目预算（万元）", min_value=0, value=None, step=10, placeholder="请输入预算")
cost = st.number_input("目标成本控制（万元）", min_value=0, value=None, step=10, placeholder="请输入成本")

uploaded_image = st.file_uploader("场地地形图片（可选）", type=["png", "jpg", "jpeg"])
if uploaded_image:
    st.image(uploaded_image, caption="预览上传图片", width=300)

# 页数选择放到这里（按钮上方）
page_num = st.number_input(
    "生成页数", min_value=3, max_value=20, value=5, step=1
)


# ==================== 5. 构造发送给 AI 的指令 ====================
def build_user_prompt(school_name, school_level, need_nvidia,
                      domain, sub_domain, budget, cost, page_num):
    nvidia_text = "是" if need_nvidia else "否"
    prompt = f"""请为以下需求生成一份PPT大纲：
- 学校名称：{school_name}
- 学校层次：{school_level}
- 核心领域：{domain} - {sub_domain}
- 项目预算：{budget}万元，目标成本控制在{cost}万元
- 是否需要提及英伟达技术或合作：{nvidia_text}
- PPT页数要求：{page_num}页（包含封面和尾页）

请根据学校层次调整内容深度和专业方向。层次越高越强调科研合作与前沿技术，层次较低则侧重技能培训、就业对接等。
如果提到了英伟达，请自然融入相关技术或合作点。
"""
    return prompt


# ==================== 6. 调用 DeepSeek 生成大纲 ====================
def generate_ppt_outline(full_user_prompt):
    system_prompt = """你是一位专业的企业方案演示专家。请根据用户的详细需求，生成一份结构清晰、内容贴合的PPT大纲。
要求：
1. 返回严格的JSON格式，只包含JSON，不要有任何其他文字。
2. JSON结构：{"title": "...", "slides": [{"heading": "...", "content": ["要点1","要点2"]}, ...]}
3. slides的数量必须严格等于用户要求的页数。
4. 内容要针对学校层次、预算成本、领域细分进行定制，若需体现英伟达则自然融入。
5. 必须包含一页“场地地形”，heading为“场地地形”，content可以为“（此处插入场地地形图片）”等简要说明。
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
        # 清理 Markdown 代码块
        if raw.startswith("```json"):
            raw = raw[7:-3].strip()
        elif raw.startswith("```"):
            raw = raw[3:-3].strip()
        return json.loads(raw)
    except json.JSONDecodeError:
        st.error("AI返回内容无法解析为JSON，请重试。")
        st.code(raw)
        return None
    except Exception as e:
        st.error(f"调用出错: {e}")
        return None


# ==================== 7. 根据 JSON 生成 PPT 对象 ====================
def build_pptx_from_json(data):
    prs = Presentation()
    for i, slide_data in enumerate(data["slides"]):
        if i == 0:
            layout = prs.slide_layouts[0]  # 封面
        else:
            layout = prs.slide_layouts[1]  # 标题+内容
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


# ==================== 8. 在 PPT 中插入场地地形图片 ====================
def add_image_slide(prs, image_file):
    """在 PPT 最后插入一张带标题的图片页"""
    layout = prs.slide_layouts[6]  # 空白版式
    slide = prs.slides.add_slide(layout)

    # 添加标题
    left, top, width, height = Inches(1), Inches(0.5), Inches(8), Inches(1)
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    p = tf.add_paragraph()
    p.text = "场地地形"
    p.font.size = Pt(28)
    p.font.bold = True

    # 保存上传的图片到临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmp:
        tmp.write(image_file.read())
        tmp_path = tmp.name

    # 插入图片
    slide.shapes.add_picture(tmp_path, Inches(1), Inches(1.8), width=Inches(8))
    return prs


# ==================== 9. 生成按钮与主逻辑 ====================
if st.button("生成 PPT"):
    # 简单校验
    if not school_name.strip():
        st.warning("请输入学校名称")
    elif page_num < 3:
        st.warning("页数至少为3页")
    else:
        with st.spinner("正在为您定制内容..."):
            user_prompt = build_user_prompt(
                school_name, school_level, need_nvidia,
                domain, sub_domain, budget, cost, page_num
            )
            outline = generate_ppt_outline(user_prompt)

        if outline is not None:
            st.success("内容生成完毕，正在打包 PPT...")
            prs = build_pptx_from_json(outline)

            # 如果上传了图片，插入图片页
            if uploaded_image is not None:
                prs = add_image_slide(prs, uploaded_image)

            # 保存最终 PPT
            ppt_path = "final_output.pptx"
            prs.save(ppt_path)

            with open(ppt_path, "rb") as f:
                st.download_button(
                    label="下载PPT",
                    data=f,
                    file_name=f"{school_name}_{domain}方案.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )