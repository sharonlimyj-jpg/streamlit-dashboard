from datetime import datetime
import os
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import base64

# 매뉴얼 HTML 내용
MANUAL_HTML = """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Streamlit 대시보드 사용 매뉴얼</title>
    <style>
        @page {
            size: A4;
            margin: 2cm;
        }
        
        body {
            font-family: 'Malgun Gothic', '맑은 고딕', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 210mm;
            margin: 0 auto;
            padding: 20px;
            background: #fff;
        }
        
        .page {
            page-break-after: always;
            min-height: 297mm;
            padding: 40px;
            box-sizing: border-box;
        }
        
        .page:last-child {
            page-break-after: auto;
        }
        
        h1 {
            color: #2c3e50;
            border-bottom: 4px solid #3498db;
            padding-bottom: 15px;
            margin-bottom: 30px;
            font-size: 32px;
        }
        
        h2 {
            color: #34495e;
            margin-top: 30px;
            margin-bottom: 20px;
            font-size: 24px;
            border-left: 5px solid #3498db;
            padding-left: 15px;
        }
        
        h3 {
            color: #555;
            margin-top: 20px;
            margin-bottom: 15px;
            font-size: 20px;
        }
        
        .intro-box {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            border-radius: 15px;
            margin: 30px 0;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        
        .intro-box h1 {
            color: white;
            border-bottom: 3px solid rgba(255,255,255,0.3);
            margin-top: 0;
        }
        
        .info-box {
            background: #f8f9fa;
            border-left: 5px solid #3498db;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
        }
        
        .success-box {
            background: #d4edda;
            border-left: 5px solid #28a745;
            padding: 20px;
            margin: 20px 0;
            border-radius: 5px;
        }
        
        .feature-grid {
            display: grid;
            grid-template-columns: repeat(2, 1fr);
            gap: 20px;
            margin: 20px 0;
        }
        
        .feature-card {
            background: white;
            border: 2px solid #e0e0e0;
            border-radius: 10px;
            padding: 20px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
        }
        
        .feature-card h3 {
            margin-top: 0;
            color: #3498db;
        }
        
        .step-box {
            background: #fff3cd;
            border-left: 5px solid #ffc107;
            padding: 15px;
            margin: 15px 0;
            border-radius: 5px;
        }
        
        ul {
            margin: 15px 0;
            padding-left: 25px;
        }
        
        li {
            margin: 10px 0;
            line-height: 1.8;
        }
        
        .emoji {
            font-size: 1.3em;
            margin-right: 8px;
        }
        
        .highlight {
            background: #fff3cd;
            padding: 2px 8px;
            border-radius: 3px;
            font-weight: bold;
        }
        
        .url-box {
            background: #e7f3ff;
            padding: 15px;
            border-radius: 8px;
            margin: 20px 0;
            font-family: 'Consolas', monospace;
            word-break: break-all;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            margin: 20px 0;
        }
        
        th, td {
            border: 1px solid #ddd;
            padding: 12px;
            text-align: left;
        }
        
        th {
            background: #3498db;
            color: white;
            font-weight: bold;
        }
        
        tr:nth-child(even) {
            background: #f8f9fa;
        }
        
        .footer {
            margin-top: 40px;
            padding-top: 20px;
            border-top: 2px solid #e0e0e0;
            text-align: center;
            color: #777;
            font-size: 14px;
        }
        
        @media print {
            body {
                print-color-adjust: exact;
                -webkit-print-color-adjust: exact;
            }
            
            .page {
                page-break-after: always;
            }
        }
    </style>
</head>
<body>

<!-- 페이지 1: 소개 및 시작하기 -->
<div class="page">
    <div class="intro-box">
        <h1>📊 Streamlit 대시보드 사용 매뉴얼</h1>
        <p style="font-size: 18px; margin: 15px 0;">2025 영업 실적 분석 대시보드 완벽 가이드</p>
        <p style="font-size: 14px; opacity: 0.9;">버전 1.0 | 최종 업데이트: 2025년</p>
    </div>

    <h2>🎯 대시보드 개요</h2>
    <p>이 대시보드는 2025년 영업 실적 데이터를 실시간으로 분석하고 시각화하는 도구입니다. 두 가지 주요 데이터 파일을 통해 다양한 관점에서 영업 성과를 분석할 수 있습니다.</p>

    <div class="feature-grid">
        <div class="feature-card">
            <h3>📈 파일 1: 영업채널 분석</h3>
            <ul style="padding-left: 20px; font-size: 14px;">
                <li>핵심 KPI 메트릭</li>
                <li>월별 실적 추이</li>
                <li>영업채널별 분석</li>
                <li>제품별 실적</li>
            </ul>
        </div>
        <div class="feature-card">
            <h3>⏱️ 파일 2: 약정기간 분석</h3>
            <ul style="padding-left: 20px; font-size: 14px;">
                <li>약정기간별 분석</li>
                <li>리스구분별 분석</li>
                <li>비용구분별 분석</li>
                <li>상세 데이터 테이블</li>
            </ul>
        </div>
    </div>

    <h2>🚀 시작하기</h2>
    
    <div class="step-box">
        <h3>1️⃣ 대시보드 접속</h3>
        <div class="url-box">
            🔗 https://my-dashboard-coway.streamlit.app
        </div>
        <p>위 URL을 웹 브라우저(Chrome, Edge, Safari 등)에서 접속합니다.</p>
    </div>

    <div class="success-box">
        <h3>✅ 기본 파일 자동 로드 기능</h3>
        <p><strong>별도의 파일 업로드 없이 바로 사용할 수 있습니다!</strong></p>
        <ul>
            <li><span class="emoji">📦</span> 대시보드 실행 시 기본 데이터 파일이 자동으로 로드됩니다</li>
            <li><span class="emoji">☑️</span> "기본 파일 사용" 체크박스가 자동으로 선택되어 있습니다</li>
            <li><span class="emoji">⚡</span> 접속 즉시 모든 분석 기능을 사용할 수 있습니다</li>
        </ul>
    </div>

    <div class="info-box">
        <h3>💡 다른 파일을 사용하고 싶다면?</h3>
        <p><strong>Step 1:</strong> 사이드바의 "기본 파일 사용" 체크박스를 해제합니다</p>
        <p><strong>Step 2:</strong> 파일 업로더가 나타나면 원하는 파일을 선택합니다</p>
        <p><strong>Step 3:</strong> 업로드가 완료되면 새로운 데이터로 분석이 시작됩니다</p>
    </div>

    <div class="footer">
        <p>📄 페이지 1 / 3</p>
    </div>
</div>

<!-- 페이지 2: 각 섹션 상세 설명 -->
<div class="page">
    <h1>📋 대시보드 섹션별 가이드</h1>

    <h2>📈 파일 1: 영업채널 분석</h2>
    
    <h3>1. 핵심 KPI 대시보드 (상단)</h3>
    <div class="info-box">
        <p>대시보드 상단에 4개의 주요 지표가 큰 숫자로 표시됩니다:</p>
        <table>
            <tr>
                <th style="width: 30%;">지표명</th>
                <th>확인할 수 있는 내용</th>
            </tr>
            <tr>
                <td><span class="emoji">📊</span> 총 렌탈 (건)</td>
                <td>선택한 기간 동안의 전체 렌탈 건수</td>
            </tr>
            <tr>
                <td><span class="emoji">🆕</span> 신규 (건)</td>
                <td>신규 렌탈 계약 건수</td>
            </tr>
            <tr>
                <td><span class="emoji">🔄</span> 재렌탈 (건)</td>
                <td>재계약 렌탈 건수</td>
            </tr>
            <tr>
                <td><span class="emoji">🏠</span> 홈케어 비중 (%)</td>
                <td>전체 렌탈 중 홈케어 제품이 차지하는 비율</td>
            </tr>
        </table>
    </div>

    <h3>2. 월별 실적 추이 (차트)</h3>
    <div class="info-box">
        <p><span class="emoji">📉</span> <strong>라인 차트로 표시:</strong> 월별로 총 렌탈, 신규, 재렌탈 건수의 변화를 한눈에 파악할 수 있습니다</p>
        <p><span class="emoji">🔍</span> <strong>활용 방법:</strong> 특정 월에 마우스를 올리면 상세 수치를 확인할 수 있으며, 성수기/비수기 패턴을 분석할 수 있습니다</p>
    </div>

    <h3>3. 영업채널별 분석</h3>
    <div class="info-box">
        <p><span class="emoji">📊</span> <strong>막대 차트로 표시:</strong> 각 영업채널(직영, 대리점, 온라인 등)별 실적을 비교합니다</p>
        <p><span class="emoji">💡</span> <strong>확인 가능한 정보:</strong> 어떤 채널이 가장 성과가 좋은지, 채널별 신규/재렌탈 비율은 어떤지 비교 분석할 수 있습니다</p>
    </div>

    <h3>4. 제품별 실적 & 렌탈 유형 분석</h3>
    <div class="info-box">
        <p><span class="emoji">🎯</span> <strong>제품계층구조별 차트:</strong> 제품 카테고리별 렌탈 건수를 확인할 수 있습니다</p>
        <p><span class="emoji">🔄</span> <strong>렌탈 유형 분포:</strong> 신규 vs 재렌탈 비율을 파이 차트로 시각화하여 보여줍니다</p>
    </div>

    <h2>⏱️ 파일 2: 약정기간/리스구분 분석</h2>

    <h3>1. 약정기간별 분석</h3>
    <div class="info-box">
        <p><span class="emoji">📅</span> <strong>확인 내용:</strong> 12개월, 24개월, 36개월 등 약정기간별로 렌탈 건수를 비교합니다</p>
        <p><span class="emoji">📊</span> <strong>차트 형식:</strong> 막대 차트로 표시되어 어떤 약정기간이 가장 선호되는지 한눈에 파악할 수 있습니다</p>
    </div>

    <h3>2. 리스구분별 분석</h3>
    <div class="info-box">
        <p><span class="emoji">🏷️</span> <strong>확인 내용:</strong> 금융리스, 운용리스 등 리스 유형별 계약 건수를 분석합니다</p>
        <p><span class="emoji">📈</span> <strong>활용 방법:</strong> 리스 구조별 선호도와 트렌드를 파악할 수 있습니다</p>
    </div>

    <h3>3. 비용구분별 분석 & 상세 데이터</h3>
    <div class="info-box">
        <p><span class="emoji">💰</span> <strong>비용 구조 분석:</strong> 비용 유형별로 렌탈 건수를 구분하여 표시합니다</p>
        <p><span class="emoji">📋</span> <strong>상세 테이블:</strong> 페이지 하단에 필터링된 모든 데이터를 테이블 형식으로 제공하여 세부 정보를 확인할 수 있습니다</p>
    </div>

    <div class="footer">
        <p>📄 페이지 2 / 3</p>
    </div>
</div>

<!-- 페이지 3: 필터링 및 데이터 다운로드 -->
<div class="page">
    <h1>🎛️ 필터 및 데이터 활용 가이드</h1>

    <h2>🔍 사이드바 필터 사용하기</h2>
    <p>대시보드 왼쪽 사이드바에서 다양한 필터를 설정하여 원하는 데이터만 볼 수 있습니다.</p>

    <div class="step-box">
        <h3>📅 1. 연도 선택</h3>
        <p><span class="emoji">📌</span> 드롭다운 메뉴에서 분석하고 싶은 연도를 선택합니다</p>
        <p><span class="emoji">💡</span> 기본값은 최신 연도(2025년)로 설정되어 있습니다</p>
    </div>

    <div class="step-box">
        <h3>📆 2. 월 선택 (다중 선택 가능)</h3>
        <p><span class="emoji">✅</span> 체크박스 형식으로 여러 개의 월을 동시에 선택할 수 있습니다</p>
        <p><span class="emoji">📊</span> <strong>예시:</strong> 1월, 2월, 3월을 선택하면 1분기 데이터만 분석할 수 있습니다</p>
        <p><span class="emoji">⚡</span> <strong>팁:</strong> 특정 기간을 비교하고 싶을 때 유용합니다 (상반기 vs 하반기 등)</p>
    </div>

    <div class="step-box">
        <h3>🏢 3. 영업채널 선택</h3>
        <p><span class="emoji">🎯</span> 분석하고 싶은 영업채널을 선택합니다 (직영, 대리점, 온라인 등)</p>
        <p><span class="emoji">🔄</span> 여러 채널을 선택하여 비교 분석도 가능합니다</p>
    </div>

    <div class="step-box">
        <h3>📦 4. 제품계층구조1 선택</h3>
        <p><span class="emoji">🎁</span> 제품 카테고리를 선택하여 특정 제품군의 실적만 확인할 수 있습니다</p>
        <p><span class="emoji">💼</span> 홈케어, 정수기, 공기청정기 등 대분류 제품별 분석이 가능합니다</p>
    </div>

    <h2>💾 데이터 다운로드 및 활용</h2>

    <div class="success-box">
        <h3>📥 CSV 파일 다운로드</h3>
        <p><strong>필터링된 결과를 엑셀로 저장할 수 있습니다!</strong></p>
        <ol>
            <li><span class="emoji">🔍</span> 원하는 필터 조건을 모두 설정합니다</li>
            <li><span class="emoji">📊</span> 화면에서 원하는 데이터가 표시되는지 확인합니다</li>
            <li><span class="emoji">⬇️</span> "CSV 다운로드" 버튼을 클릭합니다</li>
            <li><span class="emoji">📁</span> 다운로드된 CSV 파일을 엑셀에서 열어 추가 분석을 진행합니다</li>
        </ol>
    </div>

    <div class="info-box">
        <h3>💡 활용 팁</h3>
        <ul>
            <li><span class="emoji">🎯</span> <strong>비교 분석:</strong> 서로 다른 필터 조건으로 여러 번 다운로드하여 비교 자료를 만들 수 있습니다</li>
            <li><span class="emoji">📈</span> <strong>보고서 작성:</strong> 다운로드한 데이터를 활용하여 맞춤형 보고서를 작성할 수 있습니다</li>
            <li><span class="emoji">📊</span> <strong>추가 분석:</strong> 엑셀의 피벗 테이블 기능을 활용하여 더 깊이 있는 분석이 가능합니다</li>
            <li><span class="emoji">🔄</span> <strong>정기 리포트:</strong> 매월 동일한 필터 조건으로 다운로드하여 트렌드를 추적할 수 있습니다</li>
        </ul>
    </div>

    <h2>📱 모바일 사용</h2>
    <div class="info-box">
        <p><span class="emoji">📱</span> <strong>반응형 디자인:</strong> 대시보드는 모바일 기기에서도 정상적으로 작동합니다</p>
        <p><span class="emoji">👆</span> <strong>터치 조작:</strong> 모바일에서는 차트를 터치하여 상세 정보를 확인할 수 있습니다</p>
        <p><span class="emoji">🔄</span> <strong>화면 회전:</strong> 가로 모드로 전환하면 차트를 더 크게 볼 수 있습니다</p>
    </div>

    <div class="success-box" style="margin-top: 40px;">
        <h3>🎉 이제 대시보드를 자유롭게 활용하세요!</h3>
        <p>궁금한 사항이 있으시면 언제든지 홈케어기획팀 임영주에게 문의해주세요.</p>
        <p style="text-align: center; margin-top: 20px; font-size: 18px;">
            <strong>📊 데이터 기반 의사결정의 시작! 🚀</strong>
        </p>
    </div>

    <div class="footer">
        <p>📄 페이지 3 / 3</p>
        <p style="margin-top: 10px;">© 2025 Streamlit Dashboard | 내부용 매뉴얼</p>
    </div>
</div>

</body>
</html>
"""



def load_and_clean_dataframe(uploaded_file, file_label="파일"):
    """파일을 로드하고 컬럼명을 정리하여 반환"""
    try:
        if uploaded_file is None:
            st.warning(f"⚠️ {file_label}이 업로드되지 않았습니다.")
            return pd.DataFrame()
        
        # 파일 이름 확인
        st.info(f"🔍 {file_label} 파일명: {uploaded_file.name if hasattr(uploaded_file, 'name') else '기본 파일'}")
        
        # 파일 포인터를 처음으로 이동
        try:
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
        except Exception as seek_error:
            st.warning(f"⚠️ 파일 포인터 이동 실패: {str(seek_error)}")
        
        # 파일 읽기
        df = None
        file_name = uploaded_file if isinstance(uploaded_file, str) else uploaded_file.name
        
        if file_name.endswith('.csv'):
            # CSV 파일 처리
            try:
                df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
                st.success(f"✅ {file_label} CSV 파일 읽기 성공 (utf-8-sig)")
            except UnicodeDecodeError:
                try:
                    if hasattr(uploaded_file, 'seek'):
                        uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, encoding='cp949')
                    st.success(f"✅ {file_label} CSV 파일 읽기 성공 (cp949)")
                except Exception as e:
                    st.error(f"❌ {file_label} CSV 읽기 실패: {str(e)}")
                    return pd.DataFrame()
            except Exception as e:
                st.error(f"❌ {file_label} CSV 읽기 중 오류: {str(e)}")
                return pd.DataFrame()
        else:
            # Excel 파일 처리
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                st.success(f"✅ {file_label} Excel 파일 읽기 성공")
            except Exception as e:
                st.error(f"❌ {file_label} Excel 읽기 실패: {str(e)}")
                return pd.DataFrame()
        
        # DataFrame이 비어있는지 확인
        if df is None or df.empty:
            st.error(f"❌ {file_label} 파일이 비어있거나 읽을 수 없습니다.")
            return pd.DataFrame()
        
        # 컬럼명 정리: 공백, 줄바꿈, 특수문자 제거
        df.columns = df.columns.astype(str).str.strip().str.replace('\n', '').str.replace('\r', '')
        
        st.success(f"✅ {file_label} 로드 성공: {len(df)}행 × {len(df.columns)}열")
        
        # 디버깅 정보 표시
        with st.expander(f"🔍 {file_label} 컬럼 및 데이터 미리보기"):
            st.write("**컬럼 목록:**")
            for idx, col in enumerate(df.columns, 1):
                st.write(f"{idx}. `{col}` (길이: {len(col)}자, repr: {repr(col)})")
            st.write("**데이터 샘플 (첫 3행):**")
            st.dataframe(df.head(3))
        
        return df
    
    except Exception as e:
        st.error(f"❌ {file_label} 로드 중 예상치 못한 오류: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return pd.DataFrame()


# 페이지 설정
st.set_page_config(
    page_title="2025 영업 실적 대시보드",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 타이틀 및 매뉴얼 버튼
# 507-548번째 줄을 아래 코드로 교체
col_title, col_manual = st.columns([6, 1])

with col_title:
    st.title("📊 2025년 영업 실적 분석 대시보드")

with col_manual:
    st.markdown("""
    <style>
    .manual-button {
        display: inline-block;
        background-color: #4CAF50;
        color: white !important;
        padding: 10px 20px;
        text-align: center;
        text-decoration: none;
        font-size: 16px;
        margin-top: 20px;
        cursor: pointer;
        border-radius: 8px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    .manual-button:hover {
        background-color: #45a049;
        transform: scale(1.05);
    }
    </style>
    <div style="text-align: right;">
        <a class="manual-button" href="https://sharonlimyj-jpg.github.io/streamlit-dashboard/dashboard_manual.html" target="_blank">
            📖 매뉴얼
        </a>
    </div>
    """, unsafe_allow_html=True)


st.markdown("---")

# 기본 파일 경로 설정
DEFAULT_FILE1 = "data/2025년_영업실적.xlsx"
DEFAULT_FILE2 = "data/2025년_비용약정2.csv"

# 파일 업로드 섹션
st.markdown("### 📁 데이터 파일 설정")

# 기본 파일 존재 여부 확인
file1_exists = os.path.exists(DEFAULT_FILE1)
file2_exists = os.path.exists(DEFAULT_FILE2)

if file1_exists or file2_exists:
    st.success("✅ 기본 데이터 파일이 감지되었습니다. 바로 대시보드를 사용할 수 있습니다!")
    
col_upload1, col_upload2 = st.columns(2)

with col_upload1:
    st.markdown("**파일 1: 영업채널 분석용**")
    if file1_exists:
        use_default_1 = st.checkbox("기본 파일 사용 (2025년_영업실적.xlsx)", value=True, key="use_default_1")
    else:
        use_default_1 = False
        st.info("기본 파일이 없습니다. 파일을 업로드해주세요.")
    
    if not use_default_1:
        uploaded_file = st.file_uploader(
            "영업채널 중심 데이터 (필수)",
            type=['xlsx', 'xls', 'csv'],
            key="file1"
        )
    else:
        uploaded_file = DEFAULT_FILE1

with col_upload2:
    st.markdown("**파일 2: 약정기간/리스구분 분석용**")
    if file2_exists:
        use_default_2 = st.checkbox("기본 파일 사용 (2025년_비용약정2.csv)", value=True, key="use_default_2")
    else:
        use_default_2 = False
        st.info("기본 파일이 없습니다. 파일을 업로드해주세요.")
    
    if not use_default_2:
        uploaded_file2 = st.file_uploader(
            "약정기간/리스구분 중심 데이터 (선택)",
            type=['xlsx', 'xls', 'csv'],
            key="file2"
        )
    else:
        uploaded_file2 = DEFAULT_FILE2

st.markdown("---")

# ========================================
# 파일 1 분석 (기존 코드)
# ========================================
if uploaded_file is not None:
    # 파일 읽기
    df = load_and_clean_dataframe(uploaded_file, "파일1")
    
    if not df.empty:
        try:
            # 필수 컬럼 매핑 (유연한 컬럼명 처리)
            column_mapping = {}

            # 연도 컬럼 찾기
            if '연도' in df.columns:
                column_mapping['연도'] = '연도'
            else:
                st.error("'연도' 컬럼을 찾을 수 없습니다.")
                st.stop()

            # 월 컬럼 찾기
            if '월' in df.columns:
                column_mapping['월'] = '월'
            else:
                st.error("'월' 컬럼을 찾을 수 없습니다.")
                st.stop()

            # 영업채널 컬럼 찾기
            if '영업채널' in df.columns:
                column_mapping['영업채널'] = '영업채널'
            else:
                st.error("'영업채널' 컬럼을 찾을 수 없습니다.")
                st.stop()

            # 제품 컬럼 찾기
            for col in df.columns:
                if '제품계층구조1' in col or '제품계층구조 1' in col:
                    column_mapping['제품계층구조1'] = col
                if '제품계층구조2' in col or '제품계층구조 2' in col:
                    column_mapping['제품계층구조2'] = col
                if '제품명' in col:
                    column_mapping['제품명'] = col

            # 렌탈 건수 컬럼 찾기
            for col in df.columns:
                if '총렌탈' in col and '건' in col:
                    column_mapping['총렌탈(건)'] = col
                elif col == '렌탈(건)' or (('렌탈' in col or '신규' in col) and '건' in col and '총' not in col and '재' not in col):
                    column_mapping['렌탈(건)'] = col
                elif '재렌탈' in col and '건' in col:
                    column_mapping['재렌탈(건)'] = col

            # 필수 컬럼 확인
            required_keys = ['연도', '월', '영업채널', '제품계층구조1', '제품계층구조2', '제품명',
                             '총렌탈(건)', '렌탈(건)', '재렌탈(건)']

            missing_keys = [key for key in required_keys if key not in column_mapping]

            if missing_keys:
                st.error(f"필수 컬럼을 찾을 수 없습니다: {', '.join(missing_keys)}")
                st.info(f"현재 파일의 컬럼명: {', '.join(df.columns.tolist())}")
                st.stop()

            # 컬럼명 표준화
            df_renamed = df.rename(columns={v: k for k, v in column_mapping.items()})

            # 데이터 전처리
            # 연도 처리
            df_renamed['연도'] = df_renamed['연도'].astype(str).str.replace('년', '').str.strip()
            
            # 월 처리  
            df_renamed['월'] = df_renamed['월'].astype(str).str.replace('월', '').str.strip()
            df_renamed['월_숫자'] = pd.to_numeric(df_renamed['월'], errors='coerce')
            df_renamed = df_renamed.dropna(subset=['월_숫자'])
            df_renamed['월_숫자'] = df_renamed['월_숫자'].astype(int)
            
            # 렌탈 건수 숫자 변환
            for col in ['총렌탈(건)', '렌탈(건)', '재렌탈(건)']:
                df_renamed[col] = pd.to_numeric(df_renamed[col], errors='coerce').fillna(0)
            
            # 사이드바 필터
            st.sidebar.header("🔍 필터 설정 (파일1)")
            
            # 연도 필터
            years = sorted(df_renamed['연도'].unique())
            selected_year = st.sidebar.selectbox("연도 선택", years, index=len(years)-1 if years else 0)
            
            # 월 필터
            months = sorted(df_renamed[df_renamed['연도'] == selected_year]['월_숫자'].unique())
            selected_months = st.sidebar.multiselect(
                "월 선택",
                months,
                default=months
            )
            
            # 영업채널 필터
            channels = sorted(df_renamed['영업채널'].unique())
            selected_channels = st.sidebar.multiselect(
                "영업채널 선택",
                channels,
                default=channels
            )
            
            # 제품계층구조1 필터
            product1 = sorted(df_renamed['제품계층구조1'].unique())
            selected_product1 = st.sidebar.multiselect(
                "제품계층구조1 선택",
                product1,
                default=product1
            )
            
            # 데이터 필터링
            filtered_df = df_renamed[
                (df_renamed['연도'] == selected_year) &
                (df_renamed['월_숫자'].isin(selected_months)) &
                (df_renamed['영업채널'].isin(selected_channels)) &
                (df_renamed['제품계층구조1'].isin(selected_product1))
            ].copy()
            
            # 이전 월 데이터 (전월 대비용)
            if len(selected_months) > 0:
                prev_month = max(selected_months) - 1
                if prev_month > 0:
                    prev_month_df = df_renamed[
                        (df_renamed['연도'] == selected_year) &
                        (df_renamed['월_숫자'] == prev_month) &
                        (df_renamed['영업채널'].isin(selected_channels)) &
                        (df_renamed['제품계층구조1'].isin(selected_product1))
                    ].copy()
                else:
                    prev_month_df = pd.DataFrame()
            else:
                prev_month_df = pd.DataFrame()

            # ========== Section 1: 핵심 KPI 메트릭 ==========
            st.markdown("## 📈 핵심 성과 지표 (KPI)")

            if filtered_df.empty:
                st.warning("⚠️ 선택한 필터 조건에 해당하는 데이터가 없습니다.")
            else:
                col1, col2, col3, col4 = st.columns(4)

                # 총 렌탈 건수
                total_rental = filtered_df['총렌탈(건)'].sum()
                prev_total_rental = prev_month_df['총렌탈(건)'].sum() if not prev_month_df.empty else 0
                delta_total = total_rental - prev_total_rental
                delta_pct_total = (delta_total / prev_total_rental * 100) if prev_total_rental > 0 else 0

                with col1:
                    st.metric(
                        label="총 렌탈 건수",
                        value=f"{int(total_rental):,}건",
                        delta=f"{delta_pct_total:+.1f}% ({int(delta_total):+,}건)" if prev_total_rental > 0 else "N/A"
                    )

                # 신규 렌탈 건수
                new_rental = filtered_df['렌탈(건)'].sum()
                prev_new_rental = prev_month_df['렌탈(건)'].sum() if not prev_month_df.empty else 0
                delta_new = new_rental - prev_new_rental
                delta_pct_new = (delta_new / prev_new_rental * 100) if prev_new_rental > 0 else 0

                with col2:
                    st.metric(
                        label="신규 렌탈 건수",
                        value=f"{int(new_rental):,}건",
                        delta=f"{delta_pct_new:+.1f}% ({int(delta_new):+,}건)" if prev_new_rental > 0 else "N/A"
                    )

                # 재렌탈 건수
                re_rental = filtered_df['재렌탈(건)'].sum()
                prev_re_rental = prev_month_df['재렌탈(건)'].sum() if not prev_month_df.empty else 0
                delta_re = re_rental - prev_re_rental
                delta_pct_re = (delta_re / prev_re_rental * 100) if prev_re_rental > 0 else 0

                with col3:
                    st.metric(
                        label="재렌탈 건수",
                        value=f"{int(re_rental):,}건",
                        delta=f"{delta_pct_re:+.1f}% ({int(delta_re):+,}건)" if prev_re_rental > 0 else "N/A"
                    )

                # 홈케어 채널 비중
                homecare_rental = filtered_df[filtered_df['영업채널'] == '홈케어']['총렌탈(건)'].sum()
                homecare_ratio = (homecare_rental / total_rental * 100) if total_rental > 0 else 0

                prev_homecare_rental = prev_month_df[prev_month_df['영업채널'] == '홈케어']['총렌탈(건)'].sum() if not prev_month_df.empty else 0
                prev_homecare_ratio = (prev_homecare_rental / prev_total_rental * 100) if prev_total_rental > 0 else 0
                delta_homecare = homecare_ratio - prev_homecare_ratio

                with col4:
                    st.metric(
                        label="홈케어 채널 비중",
                        value=f"{homecare_ratio:.1f}%",
                        delta=f"{delta_homecare:+.1f}%p" if prev_total_rental > 0 else "N/A"
                    )

            st.markdown("---")

            # ========== Section 2: 월별 추이 분석 ==========
            st.markdown("## 📊 월별 실적 추이")

            col1, col2 = st.columns(2)

            with col1:
                # 월별 영업채널별 총렌탈 건수
                monthly_channel = filtered_df.groupby(['월_숫자', '영업채널'], as_index=False)['총렌탈(건)'].sum()

                if not monthly_channel.empty:
                    # 월별 전체 합계 계산 (백분율용)
                    monthly_total = monthly_channel.groupby('월_숫자')['총렌탈(건)'].sum().reset_index()
                    monthly_total.columns = ['월_숫자', '월별합계']
                    monthly_channel = monthly_channel.merge(monthly_total, on='월_숫자')
                    monthly_channel['비중(%)'] = (monthly_channel['총렌탈(건)'] / monthly_channel['월별합계'] * 100).round(1)
                    monthly_channel = monthly_channel.sort_values('월_숫자')

                    fig1 = px.bar(
                        monthly_channel,
                        x='월_숫자',
                        y='총렌탈(건)',
                        color='영업채널',
                        title="월별 영업채널별 총렌탈 건수",
                        labels={'월_숫자': '월', '총렌탈(건)': '총렌탈 건수'},
                        text='총렌탈(건)',
                        height=400,
                        hover_data={
                            '총렌탈(건)': ':,',
                            '비중(%)': ':.1f',
                            '월_숫자': False
                        }
                    )
                    fig1.update_traces(
                        texttemplate='%{text:,.0f}',
                        textposition='inside',
                        hovertemplate='<b>%{fullData.name}</b><br>' +
                                      '월: %{x}월<br>' +
                                      '총렌탈: %{y:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    )
                    fig1.update_layout(
                        xaxis_type='category',
                        xaxis_title="월",
                        yaxis_title="총렌탈 건수"
                    )
                    st.plotly_chart(fig1, use_container_width=True)
                else:
                    st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            with col2:
                # 월별 렌탈 유형별 건수
                monthly_type = filtered_df.groupby('월_숫자', as_index=False).agg({
                    '렌탈(건)': 'sum',
                    '재렌탈(건)': 'sum'
                })

                if not monthly_type.empty:
                    monthly_type['총렌탈'] = monthly_type['렌탈(건)'] + monthly_type['재렌탈(건)']
                    monthly_type['신규비중(%)'] = (monthly_type['렌탈(건)'] / monthly_type['총렌탈'] * 100).round(1)
                    monthly_type['재렌탈비중(%)'] = (monthly_type['재렌탈(건)'] / monthly_type['총렌탈'] * 100).round(1)
                    monthly_type = monthly_type.sort_values('월_숫자')

                    fig2 = go.Figure()
                    fig2.add_trace(go.Bar(
                        x=monthly_type['월_숫자'],
                        y=monthly_type['렌탈(건)'],
                        name='신규 렌탈',
                        text=monthly_type['렌탈(건)'],
                        texttemplate='%{text:,.0f}',
                        textposition='inside',
                        customdata=monthly_type[['신규비중(%)']],
                        hovertemplate='<b>신규 렌탈</b><br>' +
                                      '월: %{x}월<br>' +
                                      '건수: %{y:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    ))
                    fig2.add_trace(go.Bar(
                        x=monthly_type['월_숫자'],
                        y=monthly_type['재렌탈(건)'],
                        name='재렌탈',
                        text=monthly_type['재렌탈(건)'],
                        texttemplate='%{text:,.0f}',
                        textposition='inside',
                        customdata=monthly_type[['재렌탈비중(%)']],
                        hovertemplate='<b>재렌탈</b><br>' +
                                      '월: %{x}월<br>' +
                                      '건수: %{y:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    ))

                    fig2.update_layout(
                        title="월별 렌탈 유형별 건수 (신규 vs 재렌탈)",
                        xaxis_title="월",
                        yaxis_title="건수",
                        barmode='group',
                        height=400,
                        xaxis_type='category'
                    )
                    st.plotly_chart(fig2, use_container_width=True)
                else:
                    st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            st.markdown("---")

            # ========== Section 3: 채널별 심층 분석 ==========
            st.markdown("## 🎯 영업채널별 분석")

            col1, col2 = st.columns(2)

            with col1:
                # 영업채널별 실적 비중
                channel_total = filtered_df.groupby('영업채널', as_index=False)['총렌탈(건)'].sum()

                if not channel_total.empty and channel_total['총렌탈(건)'].sum() > 0:
                    channel_total['비중(%)'] = (channel_total['총렌탈(건)'] / channel_total['총렌탈(건)'].sum() * 100).round(1)
                    channel_total = channel_total.sort_values('총렌탈(건)', ascending=False)

                    fig3 = px.pie(
                        channel_total,
                        values='총렌탈(건)',
                        names='영업채널',
                        title="영업채널별 실적 비중",
                        hole=0.4,
                        height=400
                    )
                    fig3.update_traces(
                        textposition='inside',
                        textinfo='percent+label',
                        hovertemplate='<b>%{label}</b><br>' +
                                      '건수: %{value:,}건<br>' +
                                      '비중: %{percent}<br>' +
                                      '<extra></extra>'
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            with col2:
                # 영업채널별 성장 추세
                monthly_channel_growth = filtered_df.groupby(['월_숫자', '영업채널'], as_index=False)['총렌탈(건)'].sum()

                if not monthly_channel_growth.empty:
                    # 각 채널별 월별 비중 계산
                    monthly_totals = monthly_channel_growth.groupby('월_숫자')['총렌탈(건)'].sum().reset_index()
                    monthly_totals.columns = ['월_숫자', '월별합계']
                    monthly_channel_growth = monthly_channel_growth.merge(monthly_totals, on='월_숫자')
                    monthly_channel_growth['비중(%)'] = (monthly_channel_growth['총렌탈(건)'] / monthly_channel_growth['월별합계'] * 100).round(1)
                    monthly_channel_growth = monthly_channel_growth.sort_values('월_숫자')

                    fig4 = px.line(
                        monthly_channel_growth,
                        x='월_숫자',
                        y='총렌탈(건)',
                        color='영업채널',
                        title="영업채널별 월별 성장 추세",
                        markers=True,
                        labels={'월_숫자': '월', '총렌탈(건)': '총렌탈 건수'},
                        height=400,
                        hover_data={
                            '총렌탈(건)': ':,',
                            '비중(%)': ':.1f',
                            '월_숫자': False
                        }
                    )
                    fig4.update_traces(
                        hovertemplate='<b>%{fullData.name}</b><br>' +
                                      '월: %{x}월<br>' +
                                      '총렌탈: %{y:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    )
                    fig4.update_layout(
                        xaxis_type='category',
                        xaxis_title="월",
                        yaxis_title="총렌탈 건수"
                    )
                    st.plotly_chart(fig4, use_container_width=True)
                else:
                    st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            st.markdown("---")

            # ========== Section 4: 제품 분석 ==========
            st.markdown("## 🔝 제품별 분석")

            col1, col2 = st.columns(2)

            with col1:
                # 제품계층구조1별 매출 비중
                product1_total = filtered_df.groupby('제품계층구조1', as_index=False)['총렌탈(건)'].sum()

                if not product1_total.empty and product1_total['총렌탈(건)'].sum() > 0:
                    product1_total['비중(%)'] = (product1_total['총렌탈(건)'] / product1_total['총렌탈(건)'].sum() * 100).round(1)
                    product1_total = product1_total.sort_values('총렌탈(건)', ascending=True)

                    fig5 = px.bar(
                        product1_total,
                        x='총렌탈(건)',
                        y='제품계층구조1',
                        orientation='h',
                        title="제품계층구조1별 실적",
                        text='총렌탈(건)',
                        height=400,
                        hover_data={
                            '총렌탈(건)': ':,',
                            '비중(%)': ':.1f'
                        }
                    )
                    fig5.update_traces(
                        texttemplate='%{text:,.0f}',
                        textposition='outside',
                        hovertemplate='<b>%{y}</b><br>' +
                                      '건수: %{x:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    )
                    fig5.update_layout(
                        xaxis_title="총렌탈 건수",
                        yaxis_title="제품계층구조1"
                    )
                    st.plotly_chart(fig5, use_container_width=True)
                else:
                    st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            with col2:
                # Top 10 제품명 실적
                if not filtered_df.empty and filtered_df['총렌탈(건)'].sum() > 0:
                    top_products = filtered_df.groupby('제품명', as_index=False)['총렌탈(건)'].sum()
                    top_products['비중(%)'] = (top_products['총렌탈(건)'] / filtered_df['총렌탈(건)'].sum() * 100).round(1)
                    top_products = top_products.sort_values('총렌탈(건)', ascending=False).head(10)
                    top_products = top_products.sort_values('총렌탈(건)', ascending=True)

                    if not top_products.empty:
                        fig6 = px.bar(
                            top_products,
                            x='총렌탈(건)',
                            y='제품명',
                            orientation='h',
                            title="Top 10 제품명 실적",
                            text='총렌탈(건)',
                            height=400,
                            hover_data={
                                '총렌탈(건)': ':,',
                                '비중(%)': ':.1f'
                            }
                        )
                        fig6.update_traces(
                            texttemplate='%{text:,.0f}',
                            textposition='outside',
                            hovertemplate='<b>%{y}</b><br>' +
                                          '건수: %{x:,}건<br>' +
                                          '비중: %{customdata[0]:.1f}%<br>' +
                                          '<extra></extra>'
                        )
                        fig6.update_layout(
                            xaxis_title="총렌탈 건수",
                            yaxis_title="제품명"
                        )
                        st.plotly_chart(fig6, use_container_width=True)
                    else:
                        st.warning("제품명 데이터가 없습니다.")
                else:
                    st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            st.markdown("---")

            # ========== Section 5: 영업채널별 렌탈 유형 비중 (수정됨) ==========
            st.markdown("## 🔄 영업채널별 렌탈 유형 분석")

            channel_type = filtered_df.groupby('영업채널', as_index=False).agg({
                '총렌탈(건)': 'sum',
                '렌탈(건)': 'sum',
                '재렌탈(건)': 'sum'
            })

            if not channel_type.empty:
                # 비중 계산
                channel_type['신규비중(%)'] = channel_type.apply(
                    lambda x: (x['렌탈(건)'] / x['총렌탈(건)'] * 100) if x['총렌탈(건)'] > 0 else 0,
                    axis=1
                ).round(1)
                channel_type['재렌탈비중(%)'] = channel_type.apply(
                    lambda x: (x['재렌탈(건)'] / x['총렌탈(건)'] * 100) if x['총렌탈(건)'] > 0 else 0,
                    axis=1
                ).round(1)

                # 2열 레이아웃: 왼쪽에 차트, 오른쪽에 표
                col1, col2 = st.columns([1.2, 0.8])
                
                with col1:
                    # 세로 누적 막대 차트
                    fig7 = go.Figure()
                    fig7.add_trace(go.Bar(
                        x=channel_type['영업채널'],
                        y=channel_type['렌탈(건)'],
                        name='신규 렌탈',
                        text=channel_type['렌탈(건)'],
                        texttemplate='%{text:,.0f}',
                        textposition='inside',
                        customdata=channel_type[['신규비중(%)']],
                        hovertemplate='<b>신규 렌탈</b><br>' +
                                      '채널: %{x}<br>' +
                                      '건수: %{y:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    ))
                    fig7.add_trace(go.Bar(
                        x=channel_type['영업채널'],
                        y=channel_type['재렌탈(건)'],
                        name='재렌탈',
                        text=channel_type['재렌탈(건)'],
                        texttemplate='%{text:,.0f}',
                        textposition='inside',
                        customdata=channel_type[['재렌탈비중(%)']],
                        hovertemplate='<b>재렌탈</b><br>' +
                                      '채널: %{x}<br>' +
                                      '건수: %{y:,}건<br>' +
                                      '비중: %{customdata[0]:.1f}%<br>' +
                                      '<extra></extra>'
                    ))

                    fig7.update_layout(
                        title="영업채널별 렌탈 유형 비중 (신규 vs 재렌탈)",
                        xaxis_title="영업채널",
                        yaxis_title="건수",
                        barmode='stack',
                        height=500
                    )
                    st.plotly_chart(fig7, use_container_width=True)
                
                with col2:
                    # 표 생성 (열합계, 행합계, 백분율 포함)
                    st.markdown("#### 📊 영업채널별 집계표")
                    
                    # 표 데이터 준비
                    table_data = channel_type[['영업채널', '렌탈(건)', '재렌탈(건)', '총렌탈(건)']].copy()
                    
                    # 비중 계산 (백분율)
                    total_sum = table_data['총렌탈(건)'].sum()
                    table_data['신규(%)'] = (table_data['렌탈(건)'] / table_data['총렌탈(건)'] * 100).round(1)
                    table_data['재렌탈(%)'] = (table_data['재렌탈(건)'] / table_data['총렌탈(건)'] * 100).round(1)
                    table_data['비중(%)'] = (table_data['총렌탈(건)'] / total_sum * 100).round(1)
                    
                    # 열합계 행 추가
                    sum_row = pd.DataFrame({
                        '영업채널': ['합계'],
                        '렌탈(건)': [table_data['렌탈(건)'].sum()],
                        '재렌탈(건)': [table_data['재렌탈(건)'].sum()],
                        '총렌탈(건)': [table_data['총렌탈(건)'].sum()],
                        '신규(%)': [(table_data['렌탈(건)'].sum() / table_data['총렌탈(건)'].sum() * 100)],
                        '재렌탈(%)': [(table_data['재렌탈(건)'].sum() / table_data['총렌탈(건)'].sum() * 100)],
                        '비중(%)': [100.0]
                    })
                    
                    table_data = pd.concat([table_data, sum_row], ignore_index=True)
                    
                    # 표 표시
                    st.dataframe(
                        table_data.style.format({
                            '렌탈(건)': '{:,.0f}',
                            '재렌탈(건)': '{:,.0f}',
                            '총렌탈(건)': '{:,.0f}',
                            '신규(%)': '{:.1f}%',
                            '재렌탈(%)': '{:.1f}%',
                            '비중(%)': '{:.1f}%'
                        }),
                        use_container_width=True,
                        height=500
                    )
            else:
                st.warning("선택한 필터 조건에 해당하는 데이터가 없습니다.")

            st.markdown("---")
            
            # ========== Section 6: 상세 데이터 테이블 ==========
            st.markdown("## 📋 상세 데이터")
            
            if not filtered_df.empty:
                # 표시할 컬럼 선택
                display_columns = ['연도', '월', '영업채널', '제품계층구조1', '제품계층구조2',
                                 '제품명', '총렌탈(건)', '렌탈(건)', '재렌탈(건)']
                
                # 월 컬럼을 문자열로 변환 (표시용)
                filtered_df_display = filtered_df.copy()
                filtered_df_display['월'] = filtered_df_display['월_숫자'].astype(str) + '월'
                
                # 먼저 정렬한 후 컬럼 선택
                filtered_df_sorted = filtered_df_display.sort_values(['월_숫자', '총렌탈(건)'], ascending=[True, False])
                
                st.dataframe(
                    filtered_df_sorted[display_columns],
                    use_container_width=True,
                    height=400
                )
                
                # CSV 다운로드
                csv = filtered_df_sorted[display_columns].to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="⬇️ 필터링된 데이터 다운로드 (CSV)",
                    data=csv,
                    file_name=f"영업실적_필터링_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
                
                # 데이터 요약 정보
                total_rental_sum = filtered_df['총렌탈(건)'].sum()
                new_rental_sum = filtered_df['렌탈(건)'].sum()
                re_rental_sum = filtered_df['재렌탈(건)'].sum()
                
                st.info(
                    f"📊 필터링된 데이터: 총 {len(filtered_df):,}건 | 총 렌탈: {int(total_rental_sum):,}건 | 신규: {int(new_rental_sum):,}건 | 재렌탈: {int(re_rental_sum):,}건")
            else:
                st.warning("표시할 데이터가 없습니다.")
            
            st.markdown("---")

        except Exception as e:
            st.error(f"❌ 오류 발생: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
else:
    st.info("👆 파일 1을 선택하거나 업로드하여 시작하세요.")


# ========================================
# 파일 2 분석 (약정기간/리스구분/비용구분) - 수정된 버전
# ========================================
if uploaded_file2 is not None:
    st.markdown("---")
    st.markdown("---")
    st.markdown("# 📅 약정기간 & 리스구분 & 비용구분 분석 (파일2)")
    st.markdown("---")
    
    # 파일 읽기
    df2 = load_and_clean_dataframe(uploaded_file2, "파일2")
    
    if not df2.empty:
        try:
            # 필수 컬럼 목록
            required_cols_file2 = [
                '연도', '월', '제품계층구조1', '제품계층구조2', '제품계층구조3',
                '제품코드', '제품명', '약정기간', '리스구분', '비용구분',
                '총렌탈(건)', '렌탈(건)', '재렌탈(건)', '일시불 건'
            ]
            
            # 실제로 없는 컬럼 찾기
            missing_cols = []
            for required_col in required_cols_file2:
                if required_col not in df2.columns:
                    missing_cols.append(required_col)
            
            if missing_cols:
                st.error(f"❌ 파일2에 필수 컬럼이 없습니다: {', '.join(missing_cols)}")
                st.stop()
            
            st.success("✅ 모든 필수 컬럼이 확인되었습니다!")
            
            # 데이터 전처리
            df2['연도'] = df2['연도'].astype(str).str.replace('년', '').str.strip()
            df2['월'] = df2['월'].astype(str).str.replace('월', '').str.strip()
            df2['월_숫자'] = pd.to_numeric(df2['월'], errors='coerce')
            
            # NaN 체크
            nan_count = df2['월_숫자'].isna().sum()
            if nan_count > 0:
                st.warning(f"⚠️ 월 데이터 변환 중 {nan_count}개 행 제외됨")
            
            df2 = df2.dropna(subset=['월_숫자'])
            df2['월_숫자'] = df2['월_숫자'].astype(int)
            
            # 렌탈 건수 컬럼 숫자 변환
            for col in ['총렌탈(건)', '렌탈(건)', '재렌탈(건)', '일시불 건']:
                df2[col] = pd.to_numeric(df2[col], errors='coerce').fillna(0)
            
            # '지정되지 않음' 값 처리
            for col in ['약정기간', '리스구분', '비용구분']:
                df2[col] = df2[col].astype(str).str.strip()
                df2[col] = df2[col].replace(['지정되지 않음', 'nan', 'NaN', 'None', ''], '미지정')
            
            # 사이드바 필터 (파일2용)
            st.sidebar.markdown("---")
            st.sidebar.header("🔍 필터 설정 (파일2)")
            
            # 연도 필터
            years_f2 = sorted(df2['연도'].unique())
            selected_year_f2 = st.sidebar.selectbox(
                "연도 선택 (파일2)", 
                years_f2,
                index=len(years_f2)-1 if years_f2 else 0,
                key="year_f2"
            )
            
            # 월 필터
            months_f2 = sorted(df2[df2['연도'] == selected_year_f2]['월_숫자'].unique())
            selected_months_f2 = st.sidebar.multiselect(
                "월 선택 (파일2)",
                months_f2,
                default=months_f2,
                key="months_f2"
            )
            
            # 제품계층구조1 필터
            product1_f2 = sorted(df2['제품계층구조1'].unique())
            selected_product1_f2 = st.sidebar.multiselect(
                "제품계층구조1 선택 (파일2)",
                product1_f2,
                default=product1_f2,
                key="product1_f2"
            )
            
            # 제품명 검색 필터 추가
            st.sidebar.markdown("---")
            st.sidebar.subheader("🔍 제품명 검색")
            
            # 검색어 입력
            search_query = st.sidebar.text_input(
                "제품명 검색 (일부 입력)",
                "",
                key="product_search",
                help="제품명의 일부를 입력하면 포함된 제품들을 선택할 수 있습니다."
            )
            
            # 검색 결과에 따른 제품명 필터링
            if search_query:
                matching_products = sorted([p for p in df2['제품명'].unique() if search_query.lower() in str(p).lower()])
                if matching_products:
                    st.sidebar.success(f"🔍 {len(matching_products)}개 제품 발견")
                    selected_products_f2 = st.sidebar.multiselect(
                        "제품명 선택",
                        matching_products,
                        default=matching_products,
                        key="selected_products_f2"
                    )
                else:
                    st.sidebar.warning("⚠️ 일치하는 제품이 없습니다.")
                    selected_products_f2 = []
            else:
                # 검색어가 없으면 전체 선택
                selected_products_f2 = df2['제품명'].unique().tolist()
            
            # 데이터 필터링
            filtered_df2 = df2[
                (df2['연도'] == selected_year_f2) &
                (df2['월_숫자'].isin(selected_months_f2)) &
                (df2['제품계층구조1'].isin(selected_product1_f2)) &
                (df2['제품명'].isin(selected_products_f2))
            ].copy()
            
            if filtered_df2.empty:
                st.warning("⚠️ 선택한 필터 조건에 해당하는 데이터가 없습니다.")
            else:
                # ========== 리스구분 × 약정기간 크로스 분석 (수정됨) ==========
                st.markdown("## 📊 리스구분 × 약정기간 크로스 분석")
                
                col1, col2 = st.columns([1.2, 0.8])
                
                with col1:
                    # 구분 필터 추가
                    st.markdown("#### 📌 필터 옵션")
                    
                    # 리스구분 필터
                    lease_types = sorted(filtered_df2['리스구분'].unique())
                    selected_lease = st.multiselect(
                        "리스구분 선택",
                        lease_types,
                        default=lease_types,
                        key="lease_filter"
                    )
                    
                    # 약정기간 필터
                    contract_periods = sorted(filtered_df2['약정기간'].unique())
                    selected_periods = st.multiselect(
                        "약정기간 선택",
                        contract_periods,
                        default=contract_periods,
                        key="period_filter"
                    )
                    
                    # 필터링된 데이터
                    filtered_cross = filtered_df2[
                        (filtered_df2['리스구분'].isin(selected_lease)) &
                        (filtered_df2['약정기간'].isin(selected_periods))
                    ]
                    
                    # 월별 리스구분 x 약정기간 크로스 데이터
                    cross_monthly = filtered_cross.groupby(['월_숫자', '리스구분', '약정기간'], as_index=False)['총렌탈(건)'].sum()
                    cross_monthly = cross_monthly[cross_monthly['총렌탈(건)'] > 0]
                    
                    if not cross_monthly.empty:
                        # 리스구분+약정기간 조합 컬럼 생성
                        cross_monthly['구분'] = cross_monthly['리스구분'] + ' - ' + cross_monthly['약정기간']
                        
                        # 세로 누적 막대 그래프
                        fig_cross = px.bar(
                            cross_monthly,
                            x='월_숫자',
                            y='총렌탈(건)',
                            color='구분',
                            title="월별 리스구분 × 약정기간 실적 (누적)",
                            labels={'월_숫자': '월', '총렌탈(건)': '총렌탈 건수'},
                            text='총렌탈(건)',
                            height=550,
                            barmode='stack'  # 누적 모드
                        )
                        fig_cross.update_traces(
                            texttemplate='%{text:,.0f}',
                            textposition='inside'
                        )
                        fig_cross.update_layout(
                            xaxis_type='category',
                            xaxis_title="월",
                            yaxis_title="총렌탈 건수"
                        )
                        st.plotly_chart(fig_cross, use_container_width=True)
                    else:
                        st.warning("크로스 데이터가 없습니다.")
                
                with col2:
                    # 크로스 테이블 (리스구분 x 약정기간) - 백분율 포함
                    cross_table = filtered_cross.groupby(['리스구분', '약정기간'], as_index=False)['총렌탈(건)'].sum()
                    
                    if not cross_table.empty:
                        # 피벗 테이블 생성
                        pivot_table = cross_table.pivot_table(
                            index='리스구분',
                            columns='약정기간',
                            values='총렌탈(건)',
                            fill_value=0,
                            aggfunc='sum'
                        )
                        
                        # 행 합계 추가
                        pivot_table['행합계'] = pivot_table.sum(axis=1)
                        
                        # 열 합계 추가
                        pivot_table.loc['열합계'] = pivot_table.sum()
                        
                        # 정수형으로 변환
                        pivot_table_count = pivot_table.astype(int)
                        
                        # 백분율 테이블 생성
                        total_sum = pivot_table.loc['열합계', '행합계']
                        pivot_table_pct = (pivot_table / total_sum * 100).round(1)
                        
                        st.markdown("#### 📋 집계표 (건수)")
                        st.dataframe(
                            pivot_table_count.style.format("{:,}"),
                            use_container_width=True,
                            height=250
                        )
                        
                        st.markdown("#### 📊 집계표 (비중 %)")
                        st.dataframe(
                            pivot_table_pct.style.format("{:.1f}%"),
                            use_container_width=True,
                            height=250
                        )
                    else:
                        st.warning("크로스 테이블 데이터가 없습니다.")
                
                st.markdown("---")
                
                # ========== 비용구분별 분석 ==========
                st.markdown("## 💰 비용구분별 분석")
                
                col1, col2 = st.columns([1, 1])
                
                with col1:
                    # 비용구분별 실적 (원형 그래프)
                    cost_total = filtered_df2.groupby('비용구분', as_index=False)['총렌탈(건)'].sum()
                    cost_total = cost_total[cost_total['총렌탈(건)'] > 0]
                    
                    if not cost_total.empty:
                        cost_total['비중(%)'] = (cost_total['총렌탈(건)'] / cost_total['총렌탈(건)'].sum() * 100).round(1)
                        cost_total = cost_total.sort_values('총렌탈(건)', ascending=False)
                        
                        fig_cost = px.pie(
                            cost_total,
                            values='총렌탈(건)',
                            names='비용구분',
                            title="비용구분별 실적 비중",
                            hole=0.4,
                            height=500
                        )
                        fig_cost.update_traces(
                            textposition='inside',
                            textinfo='percent+label'
                        )
                        st.plotly_chart(fig_cost, use_container_width=True)
                    else:
                        st.warning("비용구분 데이터가 없습니다.")
                
                with col2:
                    # 비용구분별 실적 테이블 (행합계, 열합계, 비중 포함)
                    if not cost_total.empty:
                        st.markdown("#### 📋 비용구분별 실적 비중표")
                        
                        # 테이블 생성
                        cost_display = cost_total[['비용구분', '총렌탈(건)', '비중(%)']].copy()
                        
                        # 합계 행 추가
                        total_row = pd.DataFrame({
                            '비용구분': ['합계'],
                            '총렌탈(건)': [cost_display['총렌탈(건)'].sum()],
                            '비중(%)': [100.0]
                        })
                        cost_display = pd.concat([cost_display, total_row], ignore_index=True)
                        
                        # 스타일 적용
                        st.dataframe(
                            cost_display.style.format({
                                '총렌탈(건)': '{:,.0f}',
                                '비중(%)': '{:.1f}%'
                            }),
                            use_container_width=True,
                            height=500
                        )
                    else:
                        st.warning("비용구분 데이터가 없습니다.")
                
                st.markdown("---")
                
                # ========== 상세 데이터 테이블 ==========
                st.markdown("## 📋 상세 데이터 (파일2)")
                
                if not filtered_df2.empty:
                    display_columns_f2 = ['연도', '월', '제품계층구조1', '제품계층구조2', '제품명',
                                         '약정기간', '리스구분', '비용구분', '총렌탈(건)', '렌탈(건)', '재렌탈(건)']
                    
                    filtered_df2_display = filtered_df2.copy()
                    filtered_df2_display['월'] = filtered_df2_display['월_숫자'].astype(str) + '월'
                    
                    filtered_df2_sorted = filtered_df2_display.sort_values(['월_숫자', '총렌탈(건)'], ascending=[True, False])
                    
                    st.dataframe(
                        filtered_df2_sorted[display_columns_f2],
                        use_container_width=True,
                        height=400
                    )
                    
                    # CSV 다운로드
                    csv2 = filtered_df2_sorted[display_columns_f2].to_csv(index=False, encoding='utf-8-sig')
                    st.download_button(
                        label="⬇️ 필터링된 데이터 다운로드 (CSV)",
                        data=csv2,
                        file_name=f"약정기간_리스구분_필터링_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                        mime="text/csv",
                        key="download_f2"
                    )
                else:
                    st.warning("표시할 데이터가 없습니다.")
            
        except Exception as e:
            st.error(f"❌ 오류 발생: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
else:
    st.info("👆 파일 2를 선택하거나 업로드하여 약정기간/리스구분 분석을 시작하세요.")

# 푸터
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: gray; padding: 20px;'>
    <p>📊 2025 영업 실적 대시보드 | Powered by Streamlit</p>
</div>
""", unsafe_allow_html=True)
