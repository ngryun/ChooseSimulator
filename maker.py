#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
과목 선택 시뮬레이션 HTML 생성기
엑셀 파일을 읽어서 인터랙티브한 과목 선택 시뮬레이션 HTML을 생성합니다.
"""

import pandas as pd
import json
import os
import sys
import unicodedata
import re
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import webbrowser

# macOS 한글 경로 문제 해결
if sys.platform == 'darwin':  # macOS
    try:
        import locale
        locale.setlocale(locale.LC_ALL, 'ko_KR.UTF-8')
    except:
        pass

class CourseSimulatorGenerator:
    def __init__(self):
        self.df = None
        self.school_name = ""
        self.group_limits = {}  # 그룹별 선택 제한 정보 (키: "학기_선택그룹명")
        self.html_template = self._get_html_template()
        self.available_columns_map = {} # 엑셀의 실제 컬럼명 매핑

    def select_excel_file(self):
        """엑셀 파일 선택 대화상자"""
        try:
            root = tk.Tk()
            root.withdraw()  # 메인 창 숨기기
            initial_dir = os.path.expanduser("~/Desktop")
            file_path = filedialog.askopenfilename(
                title="과목 데이터 엑셀 파일을 선택하세요",
                initialdir=initial_dir,
                filetypes=[
                    ("Excel files", "*.xlsx *.xls"),
                    ("All files", "*.*")
                ]
            )
            root.destroy()
            if file_path:
                file_path = os.path.normpath(file_path)
                print(f"📁 선택된 파일: {file_path}")
            return file_path
        except Exception as e:
            print(f"❌ 파일 선택 중 오류: {e}")
            return None

    def load_excel_data(self, file_path):
        """엑셀 파일 로드 및 데이터 검증"""
        try:
            if file_path.endswith('.xlsx'):
                self.df = pd.read_excel(file_path, engine='openpyxl', dtype=str) 
            else:
                self.df = pd.read_excel(file_path, dtype=str) 
            
            for col in self.df.columns:
                if self.df[col].apply(type).eq(str).all(): 
                     self.df[col] = self.df[col].str.strip()

            print(f"✅ 엑셀 파일 로드 완료 (공백 제거 적용): {len(self.df)}개 행")
            print(f"📋 원본 컬럼: {list(self.df.columns)}")

            filename = Path(file_path).stem.strip() 
            if '고등학교' in filename or '고' in filename:
                self.school_name = filename.split('_')[0].strip() if '_' in filename else filename
            else:
                self.school_name = filename if filename else "고등학교" 
            
            if not self.school_name: 
                self.school_name = "고등학교"

            return True
        except Exception as e:
            print(f"❌ 엑셀 파일 로드 실패: {e}")
            return False

    def parse_group_limits(self, selection_info):
        """그룹 내 선택수 정보 파싱"""
        if pd.isna(selection_info) or not str(selection_info).strip():
            return None, None
        selection_str = str(selection_info).strip() 
        pattern = r'택(\d+)'
        match = re.search(pattern, selection_str)
        if match:
            limit = int(match.group(1))
            group_name = selection_str.split('택')[0].strip()
            if not group_name:
                group_name = "선택그룹" 
            return group_name, limit
        return None, None

    def _map_columns(self):
        """표준 컬럼명과 실제 엑셀 컬럼명 매핑"""
        standard_to_korean_map = {
            'year': '입학년도', 'semester': '학기', 'type': '유형',
            'name': '과목명', 'credits': '학점', 'required': '지정여부',
            'classes': '개설반수', 'subject': '담당과목', 'period': '수업시기',
            'group': '교과(군)', 'selection_count': '그룹 내 선택수'
        }
        self.available_columns_map = {}
        if self.df is None:
            return

        self.df.columns = [col.strip() for col in self.df.columns]
        df_columns_normalized = {unicodedata.normalize('NFC', col): col for col in self.df.columns}

        for std_name, kor_default_name in standard_to_korean_map.items():
            kor_default_name_normalized = unicodedata.normalize('NFC', kor_default_name.strip())
            if kor_default_name_normalized in df_columns_normalized:
                self.available_columns_map[std_name] = df_columns_normalized[kor_default_name_normalized]
            else: 
                 if std_name in df_columns_normalized: 
                     self.available_columns_map[std_name] = df_columns_normalized[std_name]
        print(f"🔍 인식된 컬럼 매핑: {self.available_columns_map}")

    def get_col_name(self, standard_name):
        return self.available_columns_map.get(standard_name)

    def process_data(self):
        if self.df is None:
            return False
        try:
            self._map_columns() 

            required_std_cols = ['semester', 'name', 'credits', 'required', 'group']
            missing_cols = [std_col for std_col in required_std_cols if not self.get_col_name(std_col)]
            if missing_cols:
                print(f"❌ 필수 컬럼 누락 (표준명 기준): {missing_cols}")
                return False

            name_col = self.get_col_name('name')
            credits_col = self.get_col_name('credits')
            
            self.df = self.df.dropna(subset=[name_col])
            self.df[credits_col] = self.df[credits_col].astype(str).str.strip()
            self.df[credits_col] = pd.to_numeric(self.df[credits_col], errors='coerce').fillna(0)

            self.group_limits = {}
            selection_count_col = self.get_col_name('selection_count')
            semester_col = self.get_col_name('semester')

            if selection_count_col:
                for _, row in self.df.iterrows():
                    selection_info = str(row.get(selection_count_col, '')).strip()
                    parsed_group_name, limit = self.parse_group_limits(selection_info)

                    if parsed_group_name and limit:
                        semester = str(row.get(semester_col, '')).strip() 
                        if not semester: continue 

                        key = f"{semester}_{parsed_group_name}" # Key is based on semester and selection group name

                        if key not in self.group_limits:
                            self.group_limits[key] = {
                                'semester': semester,
                                'group_name': parsed_group_name, # This is the "선택그룹명"
                                'limit': limit
                            }
                print(f"🎯 그룹별 선택 제한 정보: {len(self.group_limits)}개")
                for key, info in self.group_limits.items():
                    print(f"   - {info['semester']} / '{info['group_name']}' 그룹: 최대 {info['limit']}개 선택 (Key: {key})")
            else:
                print("⚠️ '그룹 내 선택수' 컬럼이 없어 그룹 선택 제한 기능을 사용하지 않습니다.")

            print(f"✅ 데이터 처리 완료: {len(self.df)}개 과목")
            return True
        except Exception as e:
            print(f"❌ 데이터 처리 실패: {e}")
            return False

    def generate_course_data(self):
        try:
            course_list = []
            if self.df is None or not self.available_columns_map:
                print("❌ 데이터프레임 또는 컬럼 매핑 정보가 없습니다.")
                return [], []

            name_col = self.get_col_name('name')
            semester_col = self.get_col_name('semester')
            type_col = self.get_col_name('type')
            credits_col = self.get_col_name('credits')
            required_col = self.get_col_name('required')
            subject_col = self.get_col_name('subject') 
            group_col = self.get_col_name('group') # 교과(군)
            selection_count_col = self.get_col_name('selection_count')

            for _, row in self.df.iterrows():
                parsed_group_name, selection_limit = None, None
                if selection_count_col:
                    selection_info = str(row.get(selection_count_col, '')).strip()
                    parsed_group_name, selection_limit = self.parse_group_limits(selection_info)

                course_name_val = str(row.get(name_col, '')).strip()
                semester_val = str(row.get(semester_col, '')).strip()
                
                if not course_name_val or not semester_val: 
                    continue

                course = {
                    'semester': semester_val,
                    'type': str(row.get(type_col, '')).strip(),
                    'name': course_name_val,
                    'credits': int(row.get(credits_col, 0)) if pd.notna(row.get(credits_col, 0)) else 0,
                    'required': str(row.get(required_col, '')).strip(),
                    'subject': str(row.get(subject_col, '')).strip(), 
                    'group': str(row.get(group_col, '')).strip(), # 교과(군) for display
                    'selection_group': parsed_group_name if parsed_group_name else None, # 선택그룹명
                    'selection_limit': selection_limit if selection_limit else None     
                }
                course_list.append(course)

            print(f"✅ {len(course_list)}개 과목 데이터 생성")
            semesters = sorted(list(set(course['semester'] for course in course_list if course['semester'])))
            print(f"📅 학기 목록: {semesters}")
            return course_list, semesters
        except Exception as e:
            print(f"❌ 과목 데이터 생성 실패: {e}")
            return [], []

    def generate_html(self, output_path=None):
        try:
            course_data, semesters = self.generate_course_data()
            if not course_data: 
                print("❌ 생성할 과목 데이터가 없습니다 (generate_course_data 실패).")
                return False

            js_course_data = json.dumps(course_data, ensure_ascii=False, indent=2)
            js_group_limits = json.dumps(self.group_limits, ensure_ascii=False, indent=2)

            display_school_name = self.school_name if self.school_name else "고등학교"

            html_content = self.html_template.format(
                school_name=display_school_name,
                course_data=js_course_data,
                group_limits=js_group_limits
            )

            final_file_path_to_use = ""
            if output_path: 
                final_file_path_to_use = output_path
            else: 
                current_school_name_for_file = self.school_name
                safe_school_filename_part = re.sub(r'[\\/*?:"<>|\'\"]', "", current_school_name_for_file) 
                safe_school_filename_part = re.sub(r'\s+', "_", safe_school_filename_part) 
                safe_school_filename_part = safe_school_filename_part.strip('_') 
                
                if not safe_school_filename_part: 
                    safe_school_filename_part = "학교" 
                final_file_path_to_use = f"{safe_school_filename_part}_과목선택시뮬레이션.html"
            
            output_abs_path = os.path.abspath(final_file_path_to_use)
            os.makedirs(os.path.dirname(output_abs_path), exist_ok=True)

            with open(output_abs_path, 'w', encoding='utf-8') as f:
                f.write(html_content)

            print(f"✅ HTML 파일 생성 완료: {output_abs_path}")
            print(f"📊 총 {len(course_data)}개 과목, {len(semesters)}개 학기")
            return output_abs_path
        except KeyError as ke:
            # This is where the error '' \n            infoText += ` | 담당'' would be caught if it's a Python format key error
            print(f"❌ HTML 생성 중 KeyError 발생: 키 '{ke}'를 찾을 수 없습니다. HTML 템플릿의 {{...}} 사용을 확인하거나, 데이터 또는 컬럼명을 확인해주세요.")
            return False
        except OSError as oe:
            print(f"❌ HTML 파일 저장 중 OSError 발생: {oe}. 파일 경로 또는 권한을 확인해주세요.")
            return False
        except Exception as e:
            print(f"❌ HTML 생성 실패 (기타 오류): {e}")
            return False

    def _get_html_template(self):
        # Ensure this template string is exactly as intended.
        # Python's .format() uses {key}. JavaScript uses ${expression}.
        # Literal braces in CSS/JS that Python might misinterpret as placeholders should be escaped: {{ for { and }} for }.
        return '''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>{school_name} 과목 선택 시뮬레이션</title>
    <style>
        * {{ /* CSS uses single braces, Python .format() needs these escaped if they are not for JS template literals */
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}

        html, body {{
            height: 100%;
            overflow-x: hidden;
        }}

        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 10px;
            font-size: 14px;
        }}

        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            overflow: hidden;
            min-height: calc(100vh - 20px);
        }}

        .header {{
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            padding: 20px 15px;
            text-align: center;
        }}

        .header h1 {{
            font-size: 1.8em;
            margin-bottom: 8px;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
            word-break: keep-all;
        }}

        .header p {{
            font-size: 0.9em;
            opacity: 0.9;
        }}

        .tabs {{
            display: flex;
            background: #f8f9fa;
            border-bottom: 1px solid #dee2e6;
            overflow-x: auto;
            -webkit-overflow-scrolling: touch;
        }}

        .tabs::-webkit-scrollbar {{
            height: 3px;
        }}

        .tabs::-webkit-scrollbar-track {{
            background: #f1f1f1;
        }}

        .tabs::-webkit-scrollbar-thumb {{
            background: #4facfe;
            border-radius: 3px;
        }}

        .tab {{
            padding: 12px 16px;
            cursor: pointer;
            border: none;
            background: none;
            font-size: 0.9em;
            transition: all 0.3s ease;
            white-space: nowrap;
            border-bottom: 3px solid transparent;
            min-width: 100px;
            flex-shrink: 0;
        }}

        .tab:hover {{
            background: #e9ecef;
        }}

        .tab.active {{
            background: white;
            border-bottom-color: #4facfe;
            color: #4facfe;
            font-weight: bold;
        }}

        .semester-content {{
            display: none;
            padding: 15px;
        }}

        .semester-content.active {{
            display: block;
        }}

        .semester-info {{
            background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
            padding: 15px;
            border-radius: 10px;
            margin-bottom: 20px;
            text-align: center;
        }}

        .semester-info h2 {{
            font-size: 1.3em;
            margin-bottom: 8px;
            word-break: keep-all;
        }}

        .semester-info p {{
            font-size: 0.9em;
            color: #666;
        }}

        .course-section, .selection-group-wrapper {{ 
            margin-bottom: 25px;
        }}
        
        .section-title {{ 
            font-size: 1.1em;
            font-weight: bold;
            color: #333;
            margin-bottom: 12px;
            padding: 8px 12px;
            background: linear-gradient(90deg, #e0e0e0, #f0f0f0); 
            border-left: 4px solid #667eea;
            border-radius: 4px;
        }}

        .course-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
            gap: 12px;
            margin-bottom: 15px;
        }}

        .course-card {{
            background: #f8f9fa;
            border: 2px solid #e9ecef;
            border-radius: 10px;
            padding: 12px;
            transition: all 0.3s ease;
            min-height: 120px; 
            display: flex;
            flex-direction: column;
            justify-content: space-between;
        }}

        .course-card:hover {{
            transform: translateY(-1px);
            box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        }}

        .course-card.required {{
            background: linear-gradient(135deg, #ffeaef 0%, #fdeff9 100%); 
            border-color: #ffacc5;
        }}

        .course-card.selected {{
            background: linear-gradient(135deg, #e6f7ff 0%, #f0faff 100%); 
            border-color: #4facfe;
            box-shadow: 0 3px 10px rgba(79, 172, 254, 0.3);
        }}

        .course-card.disabled {{
            background: #f1f1f1;
            border-color: #ccc;
            opacity: 0.6;
            cursor: not-allowed;
        }}

        .course-header {{
            display: flex;
            justify-content: space-between;
            align-items: flex-start;
            margin-bottom: 10px;
            gap: 8px;
        }}

        .course-name {{
            font-weight: bold;
            font-size: 1em;
            color: #333;
            line-height: 1.3;
            word-break: keep-all; 
            flex: 1;
        }}

        .course-credit {{
            background: #667eea;
            color: white;
            padding: 3px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            font-weight: bold;
            white-space: nowrap;
            flex-shrink: 0;
        }}

        .course-info {{
            color: #666;
            font-size: 0.8em;
            margin-bottom: 10px;
            line-height: 1.4; 
        }}

        .course-checkbox {{
            margin-top: auto; 
            padding-top: 10px;
        }}

        .course-checkbox input {{
            margin-right: 8px;
            transform: scale(1.3);
            cursor: pointer;
            vertical-align: middle; 
        }}

        .course-checkbox input:disabled {{
            cursor: not-allowed;
        }}

        .course-checkbox label {{
            cursor: pointer;
            font-weight: 500;
            font-size: 0.9em;
            user-select: none;
            -webkit-user-select: none;
            vertical-align: middle; 
        }}

        .selection-group-wrapper {{ 
            background: #fff9e6; 
            border: 1px solid #ffecb3;
            border-radius: 8px;
            padding: 15px;
            margin: 15px 0;
        }}

        .selection-group-title {{ 
            font-size: 1.05em; 
            font-weight: bold;
            color: #854d0e; 
            margin-bottom: 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 8px 0;
            border-bottom: 2px solid #ffdd80;
        }}

        .selection-count {{
            background: rgba(133, 77, 14, 0.1);
            padding: 4px 10px; 
            border-radius: 15px; 
            font-size: 0.85em; 
            color: #854d0e;
            font-weight: 500;
        }}

        .selection-group-wrapper.selection-limit-reached {{
            background: #ffebee; 
            border-color: #ffcdd2;
        }}

        .selection-group-wrapper.selection-limit-reached .selection-group-title {{
            color: #c62828; 
            border-bottom-color: #ef9a9a;
        }}
        
        .required-notice {{ 
            color: #1b5e20; 
            font-weight: bold;
            margin-top: 10px;
            font-size: 0.85em;
        }}


        .summary {{
            position: sticky;
            top: 10px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 15px;
            border-radius: 10px;
            margin-top: 20px;
            margin-bottom: 20px;
            z-index: 10; 
        }}

        .summary h3 {{
            margin-bottom: 12px;
            text-align: center;
            font-size: 1.1em;
        }}

        .selected-courses {{
            max-height: 250px;
            overflow-y: auto;
            -webkit-overflow-scrolling: touch;
        }}

        .selected-courses::-webkit-scrollbar {{
            width: 4px;
        }}

        .selected-courses::-webkit-scrollbar-track {{
            background: rgba(255,255,255,0.1);
        }}

        .selected-courses::-webkit-scrollbar-thumb {{
            background: rgba(255,255,255,0.3);
            border-radius: 2px;
        }}

        .selected-course-item {{
            background: rgba(255,255,255,0.1);
            padding: 8px 10px;
            margin: 4px 0;
            border-radius: 5px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            font-size: 0.9em;
        }}

        .selected-course-item span:first-child {{
            flex: 1;
            margin-right: 8px;
            word-break: keep-all;
        }}

        .total-credits {{
            text-align: center;
            font-size: 1.1em;
            font-weight: bold;
            margin-top: 12px;
            padding: 8px;
            background: rgba(255,255,255,0.1);
            border-radius: 5px;
        }}

        .group-credits {{
            margin-top: 10px;
            font-size: 0.9em;
            line-height: 1.4;
        }}

        .group-summary-title {{
            font-weight: bold;
            margin-top: 15px;
            margin-bottom: 5px;
        }}

        .group-summary-section {{
            margin-bottom: 10px;
        }}

        /* 모바일 최적화 */
        @media (max-width: 768px) {{
            body {{
                padding: 5px;
                font-size: 13px;
            }}
            
            .container {{
                border-radius: 10px;
                min-height: calc(100vh - 10px);
            }}
            
            .header {{
                padding: 15px 10px;
            }}
            
            .header h1 {{
                font-size: 1.5em;
            }}
            
            .tab {{
                padding: 10px 12px;
                font-size: 0.85em;
                min-width: 90px;
            }}
            
            .semester-content {{
                padding: 10px;
            }}
            
            .course-grid {{
                grid-template-columns: 1fr; 
                gap: 10px;
            }}
            
            .course-card {{
                padding: 10px;
            }}
            
            .course-name {{
                font-size: 0.95em;
            }}
            
            .course-credit {{
                font-size: 0.75em;
                padding: 2px 6px;
            }}
            
            .summary {{
                position: static; 
                margin-top: 15px;
                padding: 12px;
            }}
            
            .selected-courses {{
                max-height: 200px;
            }}
        }}

        @media (max-width: 480px) {{
            body {{
                font-size: 12px;
            }}
            
            .header h1 {{
                font-size: 1.3em;
            }}
            
            .tab {{
                padding: 8px 10px;
                font-size: 0.8em;
                min-width: 80px;
            }}
            
            .course-card {{
                padding: 8px;
            }}
            
            .course-name {{
                font-size: 0.9em;
            }}
            
            .course-checkbox input {{
                transform: scale(1.4); 
            }}
            
            .course-checkbox label {{
                font-size: 0.85em;
            }}
        }}

        @media (pointer: coarse) {{
            .tab {{
                min-height: 44px; 
            }}
            
            .course-checkbox {{
                padding: 8px 0; 
            }}
            
            .course-checkbox input {{
                min-width: 24px; 
                min-height: 24px;
            }}
        }}

    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎓 {school_name}</h1>
            <p>과목 선택 시뮬레이션</p>
        </div>

        <div class="tabs" id="tabsContainer">
            <!-- 탭들이 동적으로 생성됩니다 -->
        </div>

        <div id="semesterContents">
            <!-- 학기별 콘텐츠가 동적으로 생성됩니다 -->
        </div>

        <div class="summary">
            <h3>📋 선택 현황 요약</h3>
            <div class="selected-courses" id="summaryList"></div>
            <div class="total-credits" id="totalCredits">총 학점: 0학점</div>
            <div class="group-credits" id="groupCredits"></div>
        </div>
    </div>

    <script>
        const courseData = {course_data};
        const groupLimits = {group_limits}; // Key: "학기_선택그룹명"
        
        let selectedCourses = {};
        let semesterList = [];
        let selectionGroups = {}; // Key: "학기_선택그룹명", Value: { semester, name, limit, selected: [] }
        const groupTabName = '교과군별';
        let allTabs = [];

        document.addEventListener('DOMContentLoaded', function() {{
            initializeSimulator();
        }});

        function initializeSimulator() {{
            try {{
                semesterList = [...new Set(courseData.map(course => course.semester))].filter(s => s && String(s).trim() !== "").sort();
                semesterList.forEach(semester => {
                    selectedCourses[semester] = [];
                });
                allTabs = [...semesterList, groupTabName];

                initializeSelectionGroups(); 

                generateTabs();
                generateSemesterContents(); // This will now build the new structure

                if (semesterList.length > 0) {{
                    showSemester(semesterList[0]); 
                }} else {{
                    document.getElementById('semesterContents').innerHTML = '<p style="text-align:center; padding:20px;">표시할 학기 정보가 없습니다. 엑셀 파일의 학기 데이터를 확인해주세요.</p>';
                    updateSummary(); 
                }}

                console.log('시뮬레이션 초기화 완료:', {{
                    총과목수: courseData.length,
                    학기목록: semesterList,
                    선택그룹정의_fromPython: groupLimits,
                    활성선택그룹_JS: selectionGroups
                }});
                
            }} catch (error) {{
                console.error('초기화 오류:', error);
                alert('시뮬레이션 초기화 중 오류가 발생했습니다. 개발자 콘솔을 확인해주세요.');
            }}
        }}

        function initializeSelectionGroups() {{
            selectionGroups = {{}}; 
            for (const key in groupLimits) {{ // groupLimits has "학기_선택그룹명" as key
                if (groupLimits.hasOwnProperty(key)) {{
                    const limitInfo = groupLimits[key];
                    selectionGroups[key] = {{ // JS selectionGroups also uses "학기_선택그룹명" as key
                        semester: limitInfo.semester,
                        name: limitInfo.group_name, // This is the "선택그룹명"
                        limit: limitInfo.limit,
                        selected: [] 
                    }};
                }}
            }}

            courseData.forEach(course => {{
                if (course.required === '지정') {{ 
                    if (selectedCourses[course.semester] && !selectedCourses[course.semester].find(c => c.name === course.name)) {{
                         selectedCourses[course.semester].push(course);
                    }}
                    
                    if (course.selection_group) {{ 
                        const sgKey = `${{course.semester}}_${{course.selection_group}}`; // "학기_선택그룹명"
                        if (selectionGroups[sgKey]) {{
                            if (!selectionGroups[sgKey].selected.find(c => c.name === course.name)) {{
                                selectionGroups[sgKey].selected.push(course);
                            }}
                        }} else {{
                            // This case means a course has a selection_group, but that group is not defined in groupLimits
                            // This might happen if "그룹 내 선택수" column is missing for some courses with a selection_group name.
                            console.warn(`선택 그룹 '${{sgKey}}' (과목: ${{course.name}})이 groupLimits에 정의되지 않았습니다. '그룹 내 선택수' 컬럼을 확인해주세요.`);
                        }}
                    }}
                }}
            }});
        }}

        function generateTabs() {{
            const tabsContainer = document.getElementById('tabsContainer');
            tabsContainer.innerHTML = '';
            allTabs.forEach((semester, index) => {
                const tab = document.createElement('button');
                tab.className = `tab ${{index === 0 ? 'active' : ''}}`;
                tab.textContent = semester;
                
                const semesterString = String(semester).replace(/'/g, "\\'"); 
                tab.setAttribute('onclick', `showSemester('${{semesterString}}')`); 
                
                tab.addEventListener('touchstart', (e) => { e.preventDefault(); showSemester(semester); }, {passive: false});
                tabsContainer.appendChild(tab);
            });
        }}

        function generateSemesterContents() {{
            const contentsContainer = document.getElementById('semesterContents');
            contentsContainer.innerHTML = '';

            semesterList.forEach((semester, index) => {{
                const semesterDiv = document.createElement('div');
                const safeSemesterId = String(semester).replace(/[^a-zA-Z0-9-_]/g, '');
                semesterDiv.className = `semester-content ${{index === 0 ? 'active' : ''}}`;
                semesterDiv.id = `semester-${{safeSemesterId}}`;

                const semesterCourses = courseData.filter(course => course.semester === semester);
                const requiredCourses = semesterCourses.filter(course => course.required === '지정');
                const optionalCourses = semesterCourses.filter(course => course.required !== '지정');

                const infoDiv = document.createElement('div');
                infoDiv.className = 'semester-info';
                infoDiv.innerHTML = `<h2>${{semester}}</h2><p>지정과목: ${{requiredCourses.length}}개, 선택과목: ${{optionalCourses.length}}개</p>`;
                semesterDiv.appendChild(infoDiv);

                // 1. 지정 과목 섹션
                if (requiredCourses.length > 0) {{
                    const requiredSection = document.createElement('div');
                    requiredSection.className = 'course-section'; 
                    requiredSection.innerHTML = `<div class="section-title">📚 지정과목</div><div class="course-grid" id="required-${{safeSemesterId}}"></div>`;
                    semesterDiv.appendChild(requiredSection);
                }}

                // 2. 선택 그룹별 과목 (선택 제한이 있는 그룹)
                //    Iterate over selectionGroups that match the current semester
                const processedSelectionGroupNames = new Set(); // Track processed group names to avoid duplicate wrappers
                for (const sgKey in selectionGroups) {{
                    if (selectionGroups.hasOwnProperty(sgKey) && selectionGroups[sgKey].semester === semester) {{
                        const groupInfo = selectionGroups[sgKey]; // name here is the "선택그룹명"
                        const selectionGroupName = groupInfo.name;

                        if (processedSelectionGroupNames.has(selectionGroupName)) continue; // Already created a wrapper for this group name

                        const safeSelectionGroupName = String(selectionGroupName).replace(/[^a-zA-Z0-9-_]/g, '');
                        const wrapperId = `wrapper-${{safeSemesterId}}-${{safeSelectionGroupName}}`;
                        const gridId = `grid-${{safeSemesterId}}-${{safeSelectionGroupName}}`;
                        const countId = `count-${{safeSemesterId}}-${{safeSelectionGroupName}}`;

                        const sgWrapper = document.createElement('div');
                        sgWrapper.className = 'selection-group-wrapper';
                        sgWrapper.id = wrapperId; 

                        // Title uses the selectionGroupName. 교과(군) is not part of the main title here.
                        sgWrapper.innerHTML = `
                            <div class="selection-group-title">
                                <span>🎯 ${{selectionGroupName}}</span>
                                <span class="selection-count" id="${{countId}}">${{groupInfo.selected.length}} / ${{groupInfo.limit}}개 선택</span>
                            </div>
                            <div class="course-grid" id="${{gridId}}"></div>`;
                        semesterDiv.appendChild(sgWrapper);
                        processedSelectionGroupNames.add(selectionGroupName);
                    }}
                }}
                
                // 3. 일반 선택 과목 (선택 그룹명이 없거나, 있어도 groupLimits에 정의되지 않은 과목)
                const generalOptionalCourses = optionalCourses.filter(course => {{
                    if (!course.selection_group) return true; // No selection group name
                    const sgKey = `${{course.semester}}_${{course.selection_group}}`;
                    return !selectionGroups[sgKey]; // Selection group name exists, but not in defined selectionGroups
                }});

                if (generalOptionalCourses.length > 0) {{
                    const 교과군들ForGeneral = [...new Set(generalOptionalCourses.map(course => course.group || '기타'))].sort();
                    교과군들ForGeneral.forEach(교과군_이름 => {{
                        const safe교과군 = String(교과군_이름).replace(/[^a-zA-Z0-9-_]/g, '');
                        const sectionId = `section-general-${{safeSemesterId}}-${{safe교과군}}`;
                        const gridId = `grid-general-${{safeSemesterId}}-${{safe교과군}}`;
                        
                        const sectionDiv = document.createElement('div');
                        sectionDiv.className = 'course-section';
                        sectionDiv.id = sectionId;
                        sectionDiv.innerHTML = `
                            <div class="section-title">📖 ${{교과군_이름}} (일반선택)</div>
                            <div class="course-grid" id="${{gridId}}"></div>`;
                        semesterDiv.appendChild(sectionDiv);
                    }});
                }}
                contentsContainer.appendChild(semesterDiv);
            }});
            const groupDiv = document.createElement('div');
            groupDiv.className = "semester-content";
            groupDiv.id = "semester-group";
            groupDiv.innerHTML = "<div id=\"groupSummary\"></div>";
            contentsContainer.appendChild(groupDiv);
            
            // Initial UI update for selection limits after structure is built
            semesterList.forEach(semester => {{
                for (const sgKey in selectionGroups) {{
                    if (selectionGroups.hasOwnProperty(sgKey) && selectionGroups[sgKey].semester === semester) {{
                        const groupInfo = selectionGroups[sgKey];
                        updateSelectionLimitUI(semester, groupInfo.name); // Pass selectionGroupName
                    }}
                }}
            }});
        }}

        function showSemester(semester) {{
            const safeSemesterId = String(semester).replace(/[^a-zA-Z0-9-_]/g, '');
            document.querySelectorAll('.tab').forEach(tab => tab.classList.remove('active'));
            const activeTab = Array.from(document.querySelectorAll('.tab')).find(tab => tab.textContent === semester);
            if(activeTab) activeTab.classList.add('active');

            document.querySelectorAll('.semester-content').forEach(content => content.classList.remove('active'));
            if (semester === groupTabName) {{
                const groupContent = document.getElementById('semester-group');
                if(groupContent) {{
                    groupContent.classList.add('active');
                    renderGroupSummary();
                }}
            }} else {{
                const semesterContent = document.getElementById(`semester-${{safeSemesterId}}`);
                if (semesterContent) {{
                    semesterContent.classList.add('active');
                    renderCourses(semester);
                }}
            }}
            updateSummary();
        }}

        function renderCourses(semester) {{
            const semesterCourses = courseData.filter(course => course.semester === semester);
            const safeSemesterId = String(semester).replace(/[^a-zA-Z0-9-_]/g, '');

            // 1. Render 지정 과목
            const requiredContainer = document.getElementById(`required-${{safeSemesterId}}`);
            if (requiredContainer) {{
                requiredContainer.innerHTML = '';
                semesterCourses.filter(c => c.required === '지정').forEach(course => {{
                    requiredContainer.appendChild(createCourseCard(course, true));
                }});
            }}

            const optionalCourses = semesterCourses.filter(course => course.required !== '지정');

            // 2. Render 과목 in 선택 그룹 (선택 제한 있는 그룹)
            const processedSelectionGroupNames = new Set();
            for (const sgKey in selectionGroups) {{
                if (selectionGroups.hasOwnProperty(sgKey) && selectionGroups[sgKey].semester === semester) {{
                    const groupInfo = selectionGroups[sgKey];
                    const selectionGroupName = groupInfo.name;
                    if (processedSelectionGroupNames.has(selectionGroupName)) continue;

                    const safeSelectionGroupName = String(selectionGroupName).replace(/[^a-zA-Z0-9-_]/g, '');
                    const gridId = `grid-${{safeSemesterId}}-${{safeSelectionGroupName}}`;
                    const gridContainer = document.getElementById(gridId);

                    if (gridContainer) {{
                        gridContainer.innerHTML = '';
                        // Find all courses for this selection group name in this semester
                        const coursesForThisGroup = optionalCourses.filter(c => c.selection_group === selectionGroupName);
                        coursesForThisGroup.forEach(course => {{
                            gridContainer.appendChild(createCourseCard(course, false));
                        }});
                        updateSelectionLimitUI(semester, selectionGroupName);
                    }} else {{
                        // console.warn(`선택 그룹 그리드 컨테이너 '${{gridId}}'를 찾을 수 없습니다.`);
                    }}
                    processedSelectionGroupNames.add(selectionGroupName);
                }}
            }}
            
            // 3. Render 일반 선택 과목
            const generalOptionalCourses = optionalCourses.filter(course => {{
                if (!course.selection_group) return true;
                const sgKey = `${{course.semester}}_${{course.selection_group}}`;
                return !selectionGroups[sgKey];
            }});

            if (generalOptionalCourses.length > 0) {{
                const 교과군들ForGeneral = [...new Set(generalOptionalCourses.map(course => course.group || '기타'))].sort();
                교과군들ForGeneral.forEach(교과군_이름 => {{
                    const safe교과군 = String(교과군_이름).replace(/[^a-zA-Z0-9-_]/g, '');
                    const gridId = `grid-general-${{safeSemesterId}}-${{safe교과군}}`;
                    const gridContainer = document.getElementById(gridId);
                    if (gridContainer) {{
                        gridContainer.innerHTML = '';
                        const coursesForThis교과군 = generalOptionalCourses.filter(c => (c.group || '기타') === 교과군_이름);
                        coursesForThis교과군.forEach(course => {{
                            gridContainer.appendChild(createCourseCard(course, false));
                        }});
                    }} else {{
                        // console.warn(`일반 선택용 그리드 컨테이너 '${{gridId}}'를 찾을 수 없습니다.`);
                    }}
                }});
            }}
        }}


        function renderGroupSummary() {
            const container = document.getElementById('groupSummary');
            if (!container) return;
            container.innerHTML = '';
            const groupMap = {};
            semesterList.forEach(sem => {
                (selectedCourses[sem] || []).forEach(course => {
                    const g = course.group || '기타';
                    if (!groupMap[g]) groupMap[g] = [];
                    groupMap[g].push(course);
                });
            });
            const groups = Object.keys(groupMap).sort();
            if (groups.length === 0) {
                container.innerHTML = '<p style="text-align:center; opacity:0.7;">선택된 과목이 없습니다.</p>';
                return;
            }
            groups.forEach(g => {
                const section = document.createElement('div');
                section.className = 'group-summary-section';
                const total = groupMap[g].reduce((s,c) => s + (Number(c.credits) || 0), 0);
                const title = document.createElement('div');
                title.className = 'group-summary-title';
                title.textContent = `📚 ${g} - ${total}학점`;
                section.appendChild(title);
                groupMap[g].forEach(course => {
                    const item = document.createElement('div');
                    item.className = 'selected-course-item';
                    item.innerHTML = `<span>${course.name}</span><span>${course.credits}학점</span>`;
                    section.appendChild(item);
                });
                container.appendChild(section);
            });
        }

        function createCourseCard(course, isRequired) {{
            const card = document.createElement('div');
            card.className = 'course-card';
            if (isRequired) card.classList.add('required');

            const isSelected = selectedCourses[course.semester]?.some(c => c.name === course.name);
            if (isSelected) card.classList.add('selected');
            
            const safeSemester = String(course.semester).replace(/[^a-zA-Z0-9-_]/g, '');
            const safeCourseName = String(course.name).replace(/[^a-zA-Z0-9가-힣-_]/g, '');
            const courseId = `course-${{safeSemester}}-${{safeCourseName}}`;
            let isDisabled = false;

            if (!isRequired && course.selection_group) {{ // Check if it belongs to any selection_group
                const sgKey = `${{course.semester}}_${{course.selection_group}}`;
                const groupInfo = selectionGroups[sgKey]; // Check if this group is a defined limited group
                if (groupInfo && groupInfo.selected.length >= groupInfo.limit && !isSelected) {{
                    isDisabled = true;
                }}
            }}
            if (isDisabled) card.classList.add('disabled');

            let infoText = `${{course.group || '미분류'}} | ${{course.type || '정보없음'}}`;
            // Display selection_group name if it exists, regardless of whether it's a limited group
            if (course.selection_group) {{ 
                 infoText += ` | 그룹: ${{course.selection_group}}`;
            }}
            if(course.subject) {{ 
                infoText += ` | 담당: ${{course.subject}}`;
            }}

            const escapedSemester = String(course.semester).replace(/'/g, "\\\\'");
            const escapedCourseName = String(course.name).replace(/'/g, "\\\\'");  

            card.innerHTML = `
                <div> 
                    <div class="course-header">
                        <div class="course-name">${{course.name}}</div>
                        <div class="course-credit">${{course.credits}}학점</div>
                    </div>
                    <div class="course-info">${{infoText}}</div>
                </div>
                <div> 
                ${{isRequired ? `<div class="required-notice">✓ 지정과목 (자동 선택)</div>` : `
                    <div class="course-checkbox">
                        <input type="checkbox" id="${{courseId}}" 
                               ${{isSelected ? 'checked' : ''}} 
                               ${{isDisabled ? 'disabled' : ''}}
                               onchange="toggleCourse('${{escapedSemester}}', '${{escapedCourseName}}', this)">
                        <label for="${{courseId}}">선택</label>
                    </div>
                `}}
                </div>
            `;
            return card;
        }}

        function toggleCourse(semester, courseName, checkbox) {{
            const course = courseData.find(c => c.semester === semester && c.name === courseName);
            if (!course) return;

            const isCurrentlySelected = selectedCourses[semester]?.some(c => c.name === courseName);

            if (checkbox.checked && !isCurrentlySelected) {{ 
                if (course.selection_group) {{ // If the course belongs to a selection_group
                    const sgKey = `${{semester}}_${{course.selection_group}}`;
                    const groupInfo = selectionGroups[sgKey]; // Check if it's a defined limited group
                    if (groupInfo && groupInfo.selected.length >= groupInfo.limit) {{
                        checkbox.checked = false; 
                        alert(`'${{groupInfo.name}}' 그룹은 최대 ${{groupInfo.limit}}개까지만 선택할 수 있습니다.`);
                        return;
                    }}
                    if (groupInfo) groupInfo.selected.push(course); // Add to selectionGroups only if it's a defined limited group
                }}
                selectedCourses[semester].push(course);
            }} else if (!checkbox.checked && isCurrentlySelected) {{ 
                if (course.selection_group) {{
                    const sgKey = `${{semester}}_${{course.selection_group}}`;
                    const groupInfo = selectionGroups[sgKey];
                    if (groupInfo) {{ // Remove from selectionGroups only if it's a defined limited group
                        groupInfo.selected = groupInfo.selected.filter(c => c.name !== courseName);
                    }}
                }}
                selectedCourses[semester] = selectedCourses[semester].filter(c => c.name !== courseName);
            }}
            
            checkbox.closest('.course-card').classList.toggle('selected', checkbox.checked);

            if (course.selection_group) {{ 
                 // Update UI for this specific selection group name
                 updateSelectionLimitUI(semester, course.selection_group);
            }}
            
            renderCourses(semester); // Re-render to update disabled states of other cards
            updateSummary();
        }}

        // Updated: 교과군_이름 parameter is removed as it's not needed to identify the selection group UI elements
        function updateSelectionLimitUI(semester, selectionGroupName) {{
            if (!selectionGroupName) return; 

            const sgDataKey = `${{semester}}_${{selectionGroupName}}`; 
            const groupInfo = selectionGroups[sgDataKey]; // Get info for this selection group
            if (!groupInfo) {{ // Not a defined limited group, or no limit info
                return;
            }}
            
            const safeSemesterId = String(semester).replace(/[^a-zA-Z0-9-_]/g, '');
            const safeSelectionGroupName = String(selectionGroupName).replace(/[^a-zA-Z0-9-_]/g, '');
            
            // DOM IDs are now based on semester and selectionGroupName only
            const countId = `count-${{safeSemesterId}}-${{safeSelectionGroupName}}`;
            const wrapperId = `wrapper-${{safeSemesterId}}-${{safeSelectionGroupName}}`;

            const countElement = document.getElementById(countId);
            const wrapperElement = document.getElementById(wrapperId); 

            if (countElement) {{
                countElement.textContent = `${{groupInfo.selected.length}} / ${{groupInfo.limit}}개 선택`;
            }}
            if (wrapperElement) {{
                wrapperElement.classList.toggle('selection-limit-reached', groupInfo.selected.length >= groupInfo.limit);
            }}
        }}

        function updateSummary() {{
            const summaryList = document.getElementById('summaryList');
            const totalCreditsElement = document.getElementById('totalCredits');
            const groupCreditsElement = document.getElementById('groupCredits');
            summaryList.innerHTML = '';
            if (groupCreditsElement) groupCreditsElement.innerHTML = '';
            let totalCredits = 0;
            const groupTotals = {{}};

            semesterList.forEach(semester => {{
                const courses = selectedCourses[semester] || [];
                if (courses.length > 0) {{
                    const semesterHeader = document.createElement('div');
                    semesterHeader.innerHTML = `<strong>${{semester}} (${{courses.length}}과목)</strong>`;
                    semesterHeader.style.cssText = `margin-top: 10px; padding-bottom: 5px; border-bottom: 1px solid rgba(255,255,255,0.2); font-size: 0.95em;`;
                    if (summaryList.children.length > 0) semesterHeader.style.marginTop = '15px'; 
                    summaryList.appendChild(semesterHeader);

                    courses.forEach(course => {{
                        const item = document.createElement('div');
                        item.className = 'selected-course-item';
                        item.innerHTML = `<span>${{course.name}}</span><span>${{course.credits}}학점</span>`;
                        summaryList.appendChild(item);
                        totalCredits += Number(course.credits) || 0;
                        const gName = course.group || '기타';
                        groupTotals[gName] = (groupTotals[gName] || 0) + (Number(course.credits) || 0);
                    }});
                }}
            }});
            totalCreditsElement.textContent = `총 선택 학점: ${{totalCredits}}학점`;
            if (groupCreditsElement) {
                Object.keys(groupTotals).sort().forEach(g => {
                    const div = document.createElement('div');
                    div.textContent = `${g}: ${groupTotals[g]}학점`;
                    groupCreditsElement.appendChild(div);
                });
            }
            if (summaryList.innerHTML === '') {{
                summaryList.innerHTML = '<p style="text-align:center; opacity:0.7; padding:10px 0;">선택된 과목이 없습니다.</p>';
            }}
        }}
        
        document.addEventListener('touchstart', function() {{}}, {{passive: true}});
    </script>
</body>
</html>'''

def create_gui():
    """GUI 인터페이스 생성"""
    root = tk.Tk()
    root.title("과목 선택 시뮬레이션 생성기")
    root.geometry("550x350") 
    root.configure(bg='#eef2f3') 

    style = ttk.Style()
    style.theme_use('clam') 
    style.configure('TLabel', background='#eef2f3', font=('Helvetica', 10))
    style.configure('TButton', font=('Helvetica', 10, 'bold'), padding=10)
    style.configure('Accent.TButton', foreground='white', background='#5c6bc0') 

    main_frame = ttk.Frame(root, padding="25 30") 
    main_frame.pack(fill=tk.BOTH, expand=True)

    title_label = ttk.Label(main_frame, text="🎓 과목 선택 시뮬레이션 HTML 생성기",
                           font=('Helvetica', 18, 'bold'), foreground='#3f51b5')
    title_label.pack(pady=(0, 25)) 

    desc_label = ttk.Label(main_frame,
                          text="엑셀 파일을 선택하여 인터랙티브한 과목 선택 HTML을 생성합니다.\n필수 컬럼: 학기, 과목명, 학점, 지정여부, 교과(군)",
                          justify=tk.CENTER, font=('Helvetica', 10), wraplength=450) 
    desc_label.pack(pady=(0, 30))

    button_frame = ttk.Frame(main_frame, style='TFrame') 
    button_frame.pack(pady=20)

    select_btn = ttk.Button(button_frame, text="📁 엑셀 파일 선택 및 변환",
                           command=lambda: process_file(root, status_label), 
                           style='Accent.TButton', width=30) 
    select_btn.pack(pady=10)

    status_label = ttk.Label(main_frame, text="파일을 선택해주세요.",
                            font=('Helvetica', 9), foreground='gray', justify=tk.CENTER)
    status_label.pack(pady=(15,0), fill=tk.X, expand=True)
    
    try:
        root.eval('tk::PlaceWindow . center') 
    except tk.TclError: 
        pass 
    return root

def process_file(parent_window, status_label_widget): 
    generator = CourseSimulatorGenerator()
    file_path = generator.select_excel_file()
    if not file_path:
        status_label_widget.config(text="파일 선택이 취소되었습니다.")
        return

    status_label_widget.config(text=f"파일 처리 중... {Path(file_path).name}")
    parent_window.update_idletasks() 

    if not generator.load_excel_data(file_path):
        messagebox.showerror("오류", "엑셀 파일 로드에 실패했습니다.\n파일 형식이나 내용을 확인해주세요.")
        status_label_widget.config(text="엑셀 파일 로드 실패.")
        return

    if not generator.process_data():
        messagebox.showerror("오류", "데이터 처리에 실패했습니다.\n필수 컬럼이 모두 있는지, 데이터 형식이 올바른지 확인해주세요.\n(콘솔 로그에서 상세 오류 확인 가능)")
        status_label_widget.config(text="데이터 처리 실패.")
        return

    output_html_path = generator.generate_html()
    if output_html_path:
        status_label_widget.config(text=f"성공! {Path(output_html_path).name} 생성됨")
        result = messagebox.askyesno("완료",
            f"HTML 파일이 성공적으로 생성되었습니다!\n\n경로: {output_html_path}\n\n지금 파일을 열어보시겠습니까?")
        if result:
            try:
                url = Path(output_html_path).as_uri()
                webbrowser.open(url)
            except Exception as e:
                messagebox.showwarning("파일 열기 실패", f"브라우저를 자동으로 여는데 실패했습니다.\n오류: {e}\n파일 탐색기에서 직접 열어주세요:\n{output_html_path}")
    else:
        messagebox.showerror("오류", "HTML 파일 생성에 실패했습니다.\n콘솔 로그에서 상세 오류를 확인해주세요.")
        status_label_widget.config(text="HTML 생성 실패.")


def main():
    """메인 함수"""
    print("🎓 과목 선택 시뮬레이션 HTML 생성기")
    print("=" * 50)

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        if not os.path.exists(file_path):
            print(f"❌ 지정된 파일을 찾을 수 없습니다: {file_path}")
            return

        generator = CourseSimulatorGenerator()
        print(f"커맨드라인 모드로 실행: {file_path}")
        if generator.load_excel_data(file_path) and \
           generator.process_data():
            output_path = generator.generate_html()
            if output_path:
                print(f"\n🎉 완료! 생성된 HTML 파일: {output_path}")
            else:
                print("❌ HTML 파일 생성 실패.") 
        else:
            print("❌ 파일 처리 또는 데이터 분석 실패.") 
    else:
        try:
            root = create_gui()
            root.mainloop()
        except ImportError: 
            print("GUI를 실행할 수 없습니다. tkinter 라이브러리가 설치되어 있는지 확인해주세요.")
            print("커맨드라인 모드를 사용하려면 다음 형식으로 실행하세요:")
            print(f"python {os.path.basename(__file__)} <엑셀파일_경로>")
        except Exception as e:
            print(f"GUI 실행 중 예상치 못한 오류 발생: {e}")
            try:
                messagebox.showerror("치명적 오류", f"GUI 실행 중 오류 발생:\n{e}\n프로그램을 종료합니다.")
            except tk.TclError: 
                pass

if __name__ == "__main__":
    if os.name == 'nt':
        try:
            sys.stdout.reconfigure(encoding='utf-8')
            sys.stderr.reconfigure(encoding='utf-8')
        except AttributeError: 
            pass
            
    main()
