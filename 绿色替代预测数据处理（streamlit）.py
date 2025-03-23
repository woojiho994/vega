import pandas as pd
import streamlit as st
import io

# 设置页面配置
st.set_page_config(
    page_title="模型评分处理应用",
    page_icon="📊",
    layout="wide"
)

# 设置页面样式
st.markdown("""
    <style>
    .main-header {
        font-size: 36px;
        font-weight: bold;
        color: #2E7D32;
        text-align: center;
        margin-bottom: 30px;
    }
    .sub-header {
        font-size: 24px;
        font-weight: 600;
        color: #1565C0;
        margin-top: 20px;
        margin-bottom: 10px;
    }
    .success-text {
        color: #2E7D32;
        font-weight: bold;
    }
    </style>
    """, unsafe_allow_html=True)

# 应用标题
st.markdown('<div class="main-header">VEGA模型评分处理系统</div>', unsafe_allow_html=True)

# 添加简介
st.markdown("""
    此应用可帮助您处理VEGA模型的评分数据，根据可靠性计算结果并进行分级。
    请上传您的Excel文件并选择相应的VEGA类别进行处理。
""")

# 创建侧边栏
with st.sidebar:
    import os
    logo_path = os.path.join(os.path.dirname(__file__), "华南所logo-011.jpg")
    st.image(logo_path)
    st.markdown("## 操作面板")
    st.markdown("---")
    # 文件上传器
    uploaded_file = st.file_uploader("上传Excel文件", type=["xlsx"])
    
    if uploaded_file is not None:
        # 用户选择类别
        category = st.selectbox(
            "请选择VEGA类别",
            ["I类", "II类"],
            help="根据不同的VEGA类别，评分标准有所不同"
        )
        
        st.markdown("---")
        st.markdown("### 关于分级标准")
        if category == "I类":
            st.markdown("""
            - **L级**: 1 ≤ 结果 ≤ 2
            - **M级**: 2 < 结果 ≤ 3
            """)
        else:
            st.markdown("""
            - **L级**: 1 ≤ 结果 ≤ 1.67
            - **M级**: 1.67 < 结果 ≤ 2.33
            - **H级**: 2.33 < 结果 ≤ 3
            """)

# 主界面内容
if uploaded_file is not None:
    # 显示加载提示
    with st.spinner('正在处理数据，请稍候...'):
        # 加载上传的Excel文件
        df = pd.read_excel(uploaded_file)
        
        # 提取评分和可靠性列
        score_columns = [col for col in df.columns if 'Score' in col and 'Reliability' not in col]
        reliability_columns = [col for col in df.columns if 'Reliability' in col]
        
        # 将评分和可靠性列转换为数值
        for col in score_columns + reliability_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # 计算Results列
        def calculate_results(row):
            # 提取可靠性和评分值
            reliabilities = row[reliability_columns].values
            scores = row[score_columns].values
            
            # 找到最大可靠性值
            max_reliability = reliabilities.max()
            
            # 获取对应于最大可靠性的评分
            relevant_scores = [scores[i] for i in range(len(scores)) if reliabilities[i] == max_reliability]
            
            # 计算相关评分的平均值
            if len(relevant_scores) > 0:
                return sum(relevant_scores) / len(relevant_scores)
            else:
                return None
        
        # 应用计算到每一行
        df['Results'] = df.apply(calculate_results, axis=1)
        
        # 根据类别和结果进行分级的函数
        def grade_result(result, category):
            if pd.isna(result):
                return None
            if category == "I类":
                if 1 <= result <= 2:
                    return "L"
                elif 2 < result <= 3:
                    return "M"
            elif category == "II类":
                if 1 <= result <= 1.67:
                    return "L"
                elif 1.67 < result <= 2.33:
                    return "M"
                elif 2.33 < result <= 3:
                    return "H"
            return None
        
        # 应用分级函数
        df['Grade'] = df['Results'].apply(lambda x: grade_result(x, category))
    
    # 显示处理后的数据
    st.markdown('<div class="sub-header">处理结果</div>', unsafe_allow_html=True)
    
    # 创建三个标签页
    tab1, tab2, tab3 = st.tabs(["📊 数据表格", "📈 结果统计", "📥 下载数据"])
    
    with tab1:
        # 显示数据表格
        st.dataframe(df, use_container_width=True)
    
    with tab2:
        # 显示统计信息
        col1, col2 = st.columns(2)
        
        with col1:
            # 计算和显示每个等级的数量
            grade_counts = df['Grade'].value_counts().reset_index()
            grade_counts.columns = ['等级', '数量']
            st.markdown("### 分级统计")
            st.dataframe(grade_counts, use_container_width=True)
            
            # 计算总体平均分
            avg_score = df['Results'].mean()
            st.markdown(f"### 总体平均分: **{avg_score:.2f}**")
        
        with col2:
            # 结果分布饼图
            if not grade_counts.empty:
                st.markdown("### 等级分布")
                st.bar_chart(grade_counts.set_index('等级'))
    
    with tab3:
        # 导出处理后的数据到Excel
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='openpyxl')
        processed_data = output.getvalue()
        
        # 提供下载链接
        st.markdown('<div class="success-text">数据处理完成，可以下载结果文件</div>', unsafe_allow_html=True)
        st.download_button(
            label="📥 下载处理后的数据",
            data=processed_data,
            file_name=f"VEGA_{category}_处理结果.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="download-excel"
        )
        
        # 显示文件预览提示
        st.info("💡 提示：文件将保留原始数据，并添加计算的'Results'和'Grade'列")

else:
    # 显示等待上传的提示信息
    st.info("👆 请使用左侧面板上传Excel文件开始处理")
    
    # # 显示示例图像
    # col1, col2, col3 = st.columns([1, 2, 1])
    # with col2:
    #     st.image("https://img.icons8.com/clouds/400/000000/excel.png", width=200)
