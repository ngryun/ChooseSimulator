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
        self.group_limits = {}  # 그룹별 선택 제한 정보
        self.html_template = self._get_html_template()
    
    def select_excel_file(self):
        """엑셀 파일 선택 대화상자"""
        try:
            root = tk.Tk()
            root.withdraw()  # 메인 창 숨기기
            
            # macOS에서 한글 폴더 처리
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
            
            # 경로 정규화
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
            # 엑셀 파일 읽기
            if file_path.endswith('.xlsx'):
                self.df = pd.read_excel(file_path, engine='openpyxl')
            else:
                self.df = pd.read_excel(file_path)
            
            print(f"✅ 엑셀 파일 로드 완료: {len(self.df)}개 행")
            print(f"📋 컬럼: {list(self.df.columns)}")
            
            # 파일명에서 학교명 추출 (선택사항)
            filename = Path(file_path).stem
            if '고등학교' in filename or '고' in filename:
                self.school_name = filename.split('_')[0] if '_' in filename else filename
            else:
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
        
        # "택{N}" 패턴 찾기
        pattern = r'택(\d+)'
        match = re.search(pattern, selection_str)
        
        if match:
            limit = int(match.group(1))
            # 그룹명 추출 (택{N} 앞부분)
            group_name = selection_str.split('택')[0].strip()
            return group_name, limit
        
        return None, None
    
    def process_data(self):
        """데이터 처리 및 검증"""
        if self.df is None:
            return False
        
        try:
            # 컬럼명 정규화
            column_mapping = {
                '입학년도': 'year',
                '학기': 'semester', 
                '유형': 'type',
                '과목명': 'name',
                '학점': 'credits',
                '지정여부': 'required',
                '개설반수': 'classes',
                '담당과목': 'subject',
                '수업시기': 'period',
                '교과(군)': 'group',
                '그룹 내 선택수': 'selection_count'
            }
            
            # 컬럼 매핑
            available_columns = {}
            for korean_col, english_col in column_mapping.items():
                if korean_col in self.df.columns:
                    available_columns[english_col] = korean_col
            
            print(f"🔍 인식된 컬럼: {available_columns}")
            
            # 필수 컬럼 확인
            required_columns = ['semester', 'name', 'credits', 'required', 'group']
            missing_columns = [col for col in required_columns if col not in available_columns]
            
            if missing_columns:
                print(f"❌ 필수 컬럼 누락: {missing_columns}")
                return False
            
            # 데이터 정제
            self.df = self.df.dropna(subset=[available_columns['name']])  # 과목명이 없는 행 제거
            self.df['credits'] = pd.to_numeric(self.df[available_columns['credits']], errors='coerce').fillna(0)
            
            # 그룹별 선택 제한 정보 추출
            if 'selection_count' in available_columns:
                self.group_limits = {}
                for _, row in self.df.iterrows():
                    selection_info = row.get(available_columns['selection_count'])
                    group_name, limit = self.parse_group_limits(selection_info)
                    
                    if group_name and limit:
                        semester = str(row.get(available_columns['semester'], ''))
                        group = str(row.get(available_columns['group'], ''))
                        key = f"{semester}_{group}_{group_name}"
                        
                        if key not in self.group_limits:
                            self.group_limits[key] = {
                                'semester': semester,
                                'group': group,
                                'group_name': group_name,
                                'limit': limit
                            }
                
                print(f"🎯 그룹별 선택 제한 정보: {len(self.group_limits)}개")
                for key, info in self.group_limits.items():
                    print(f"   - {info['semester']} / {info['group']} / {info['group_name']}: 최대 {info['limit']}개 선택")
            
            print(f"✅ 데이터 처리 완료: {len(self.df)}개 과목")
            return True
            
        except Exception as e:
            print(f"❌ 데이터 처리 실패: {e}")
            return False
    
    def generate_course_data(self):
        """과목 데이터를 JavaScript 형태로 변환"""
        try:
            course_list = []
            
            for _, row in self.df.iterrows():
                # 그룹 내 선택수 정보 파싱
                selection_info = row.get('그룹 내 선택수')
                group_name, selection_limit = self.parse_group_limits(selection_info)
                
                course = {
                    'semester': str(row.get('학기', '')),
                    'type': str(row.get('유형', '')), 
                    'name': str(row.get('과목명', '')),
                    'credits': int(row.get('학점', 0)) if pd.notna(row.get('학점', 0)) else 0,
                    'required': str(row.get('지정여부', '')),
                    'subject': str(row.get('담당과목', '')),
                    'group': str(row.get('교과(군)', '')),
                    'selection_group': group_name if group_name else None,
                    'selection_limit': selection_limit if selection_limit else None
                }
                
                # 빈 값 체크
                if course['name'] and course['semester']:
                    course_list.append(course)
            
            print(f"✅ {len(course_list)}개 과목 데이터 생성")
            
            # 학기 목록 추출
            semesters = sorted(list(set(course['semester'] for course in course_list)))
            print(f"📅 학기 목록: {semesters}")
            
            return course_list, semesters
            
        except Exception as e:
            print(f"❌ 과목 데이터 생성 실패: {e}")
            return [], []
    
    def generate_html(self, output_path=None):
        """HTML 파일 생성"""
        try:
            course_data, semesters = self.generate_course_data()
            
            if not course_data:
                print("❌ 생성할 과목 데이터가 없습니다.")
                return False
            
            # JavaScript 데이터 생성
            js_course_data = json.dumps(course_data, ensure_ascii=False, indent=2)
            js_group_limits = json.dumps(self.group_limits, ensure_ascii=False, indent=2)
            
            # HTML 템플릿에 데이터 삽입
            html_content = self.html_template.format(
                school_name=self.school_name,
                course_data=js_course_data,
                group_limits=js_group_limits
            )
            
            # 출력 파일 경로 설정
            if not output_path:
                output_path = f"{self.school_name}_과목선택시뮬레이션.html"
            
            # HTML 파일 저장
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(html_content)
            
            print(f"✅ HTML 파일 생성 완료: {output_path}")
            print(f"📊 총 {len(course_data)}개 과목, {len(semesters)}개 학기")
            
            return output_path
            
        except Exception as e:
            print(f"❌ HTML 생성 실패: {e}")
            return False
    
    def _get_html_template(self):
        """모바일 최적화된 HTML 템플릿 반환"""
        return '''<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no">
    <title>{school_name} 과목 선택 시뮬레이션</title>
    <style>
        * {{
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

        .course-section {{
            margin-bottom: 25px;
        }}

        .section-title {{
            font-size: 1.1em;
            font-weight: bold;
            color: #333;
            margin-bottom: 12px;
            padding: 8px 12px;
            background: linear-gradient(90deg, #667eea, #764ba2);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            background-clip: text;
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
        }}

        .course-card:hover {{
            transform: translateY(-1px);
            box-shadow: 0 3px 10px rgba(0,0,0,0.1);
        }}

        .course-card.required {{
            background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);
            border-color: #ff6b9d;
        }}

        .course-card.selected {{
            background: linear-gradient(135deg, #a8edea 0%, #fed6e3 100%);
            border-color: #4facfe;
            box-shadow: 0 3px 10px rgba(79, 172, 254, 0.3);
        }}

        .course-card.disabled {{
            background: #f1f1f1;
            border-color: #ccc;
            opacity: 0.6;
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
            line-height: 1.2;
        }}

        .course-checkbox {{
            margin-top: 10px;
        }}

        .course-checkbox input {{
            margin-right: 8px;
            transform: scale(1.3);
            cursor: pointer;
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
        }}

        .selection-group {{
            background: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 8px;
            padding: 12px;
            margin: 15px 0;
        }}

        .selection-group-title {{
            font-weight: bold;
            color: #8b4513;
            margin-bottom: 12px;
            text-align: center;
            font-size: 0.95em;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}

        .selection-count {{
            background: rgba(139, 69, 19, 0.1);
            padding: 4px 8px;
            border-radius: 12px;
            font-size: 0.8em;
            color: #8b4513;
        }}

        .selection-limit-reached {{
            background: #ffe6e6;
            border-color: #ffb3b3;
        }}

        .selection-limit-reached .selection-group-title {{
            color: #d63384;
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
                min-height: 100px;
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
                min-height: 90px;
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

        /* 터치 친화적 스타일 */
        @media (pointer: coarse) {{
            .tab {{
                min-height: 44px;
            }}
            
            .course-card {{
                min-height: 120px;
            }}
            
            .course-checkbox {{
                padding: 5px 0;
            }}
            
            .course-checkbox input {{
                min-width: 20px;
                min-height: 20px;
            }}
        }}

        /* 다크모드 대응 */
        @media (prefers-color-scheme: dark) {{
            .course-info {{
                color: #888;
            }}
            
            .selection-group {{
                background: #2a2a2a;
                border-color: #444;
            }}
            
            .selection-group-title {{
                color: #ffd700;
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
        </div>
    </div>

    <script>
        // 과목 데이터
        const courseData = {course_data};
        
        // 그룹별 선택 제한 정보
        const groupLimits = {group_limits};
        
        // 전역 변수
        let selectedCourses = {{}};
        let semesterList = [];
        let selectionGroups = {{}};

        // 초기화
        document.addEventListener('DOMContentLoaded', function() {{
            initializeSimulator();
        }});

        function initializeSimulator() {{
            try {{
                // 학기 목록 추출
                semesterList = [...new Set(courseData.map(course => course.semester))].sort();
                
                // selectedCourses 초기화
                semesterList.forEach(semester => {{
                    selectedCourses[semester] = [];
                }});

                // 선택 그룹 초기화
                initializeSelectionGroups();

                // 탭 생성
                generateTabs();
                
                // 각 학기별 콘텐츠 생성
                generateSemesterContents();
                
                // 첫 번째 학기 활성화
                if (semesterList.length > 0) {{
                    showSemester(semesterList[0]);
                }}

                console.log('시뮬레이션 초기화 완료:', {{
                    총과목수: courseData.length,
                    학기수: semesterList.length,
                    학기목록: semesterList,
                    선택그룹: selectionGroups
                }});
                
            }} catch (error) {{
                console.error('초기화 오류:', error);
            }}
        }}

        function initializeSelectionGroups() {{
            selectionGroups = {{}};
            
            courseData.forEach(course => {{
                if (course.selection_group && course.selection_limit) {{
                    const key = `${{course.semester}}_${{course.group}}_${{course.selection_group}}`;
                    if (!selectionGroups[key]) {{
                        selectionGroups[key] = {{
                            semester: course.semester,
                            group: course.group,
                            name: course.selection_group,
                            limit: course.selection_limit,
                            selected: []
                        }};
                    }}
                }}
            }});
        }}

        function generateTabs() {{
            const tabsContainer = document.getElementById('tabsContainer');
            tabsContainer.innerHTML = '';

            semesterList.forEach((semester, index) => {{
                const tab = document.createElement('button');
                tab.className = `tab ${{index === 0 ? 'active' : ''}}`;
                tab.textContent = semester;
                tab.onclick = () => showSemester(semester);
                
                // 터치 이벤트 추가
                tab.addEventListener('touchstart', () => showSemester(semester));
                
                tabsContainer.appendChild(tab);
            }});
        }}

        function generateSemesterContents() {{
            const contentsContainer = document.getElementById('semesterContents');
            contentsContainer.innerHTML = '';

            semesterList.forEach((semester, index) => {{
                const semesterDiv = document.createElement('div');
                semesterDiv.className = `semester-content ${{index === 0 ? 'active' : ''}}`;
                semesterDiv.id = `semester-${{semester}}`;

                const semesterCourses = courseData.filter(course => course.semester === semester);
                const requiredCourses = semesterCourses.filter(course => course.required === '지정');
                const optionalCourses = semesterCourses.filter(course => course.required !== '지정');

                // 학기 정보
                const infoDiv = document.createElement('div');
                infoDiv.className = 'semester-info';
                infoDiv.innerHTML = `
                    <h2>${{semester}}</h2>
                    <p>지정과목: ${{requiredCourses.length}}개, 선택과목: ${{optionalCourses.length}}개</p>
                `;
                semesterDiv.appendChild(infoDiv);

                // 지정과목
                if (requiredCourses.length > 0) {{
                    const requiredSection = document.createElement('div');
                    requiredSection.className = 'course-section';
                    requiredSection.innerHTML = `
                        <div class="section-title">📚 지정과목</div>
                        <div class="course-grid" id="required-${{semester}}"></div>
                    `;
                    semesterDiv.appendChild(requiredSection);
                }}

                // 선택과목 (그룹별로 구분)
                if (optionalCourses.length > 0) {{
                    const groups = [...new Set(optionalCourses.map(course => course.required))];
                    
                    groups.forEach(group => {{
                        const groupCourses = optionalCourses.filter(course => course.required === group);
                        
                        // 선택 그룹별로 분리
                        const selectionGroupMap = {{}};
                        groupCourses.forEach(course => {{
                            const groupKey = course.selection_group || 'default';
                            if (!selectionGroupMap[groupKey]) {{
                                selectionGroupMap[groupKey] = [];
                            }}
                            selectionGroupMap[groupKey].push(course);
                        }});

                        Object.keys(selectionGroupMap).forEach(selectionGroupKey => {{
                            const courses = selectionGroupMap[selectionGroupKey];
                            const hasLimit = courses[0].selection_limit;
                            
                            const groupSection = document.createElement('div');
                            groupSection.className = `selection-group ${{hasLimit ? 'has-limit' : ''}}`;
                            groupSection.id = `group-${{semester}}-${{group}}-${{selectionGroupKey}}`;
                            
                            let titleContent = `🎯 ${{group}} 선택과목`;
                            if (hasLimit) {{
                                titleContent = `🎯 ${{courses[0].selection_group}} 선택과목`;
                            }}
                            
                            groupSection.innerHTML = `
                                <div class="selection-group-title">
                                    <span>${{titleContent}}</span>
                                    ${{hasLimit ? `<span class="selection-count" id="count-${{semester}}-${{group}}-${{selectionGroupKey}}">0 / ${{courses[0].selection_limit}}개 선택</span>` : ''}}
                                </div>
                                <div class="course-grid" id="optional-${{group}}-${{selectionGroupKey}}-${{semester}}"></div>
                            `;
                            semesterDiv.appendChild(groupSection);
                        }});
                    }});
                }}

                contentsContainer.appendChild(semesterDiv);

                // 지정과목 자동 선택
                selectedCourses[semester] = [...requiredCourses];
                // 선택 제한 그룹에 지정과목 반영
                requiredCourses.forEach(course => {{
                    if (course.selection_group && course.selection_limit) {{
                        const key = `${{semester}}_${{course.group}}_${{course.selection_group}}`;
                        if (selectionGroups[key] && !selectionGroups[key].selected.find(c => c.name === course.name)) {{
                            selectionGroups[key].selected.push(course);
                        }}
                    }}
                }});
            }});
            Object.keys(selectionGroups).forEach(key => {{
                const info = selectionGroups[key];
                updateSelectionLimit(info.semester, info.group, info.name);
            }});
        }}

        function showSemester(semester) {{
            // 모든 탭 비활성화
            document.querySelectorAll('.tab').forEach(tab => {{
                tab.classList.remove('active');
                if (tab.textContent === semester) {{
                    tab.classList.add('active');
                }}
            }});

            // 모든 콘텐츠 숨기기
            document.querySelectorAll('.semester-content').forEach(content => {{
                content.classList.remove('active');
            }});

            // 선택된 학기 콘텐츠 표시
            const semesterContent = document.getElementById(`semester-${{semester}}`);
            if (semesterContent) {{
                semesterContent.classList.add('active');
                renderCourses(semester);
            }}
        }}

        function renderCourses(semester) {{
            const semesterCourses = courseData.filter(course => course.semester === semester);
            const requiredCourses = semesterCourses.filter(course => course.required === '지정');
            const optionalCourses = semesterCourses.filter(course => course.required !== '지정');

            // 지정과목 렌더링
            const requiredContainer = document.getElementById(`required-${{semester}}`);
            if (requiredContainer) {{
                requiredContainer.innerHTML = '';
                requiredCourses.forEach(course => {{
                    requiredContainer.appendChild(createCourseCard(course, semester, true));
                }});
            }}

            // 선택과목 렌더링 (그룹별)
            const groups = [...new Set(optionalCourses.map(course => course.required))];
            groups.forEach(group => {{
                const groupCourses = optionalCourses.filter(course => course.required === group);
                
                // 선택 그룹별로 분리
                const selectionGroupMap = {{}};
                groupCourses.forEach(course => {{
                    const groupKey = course.selection_group || 'default';
                    if (!selectionGroupMap[groupKey]) {{
                        selectionGroupMap[groupKey] = [];
                    }}
                    selectionGroupMap[groupKey].push(course);
                }});

                Object.keys(selectionGroupMap).forEach(selectionGroupKey => {{
                    const courses = selectionGroupMap[selectionGroupKey];
                    const groupContainer = document.getElementById(`optional-${{group}}-${{selectionGroupKey}}-${{semester}}`);
                    
                    if (groupContainer) {{
                        groupContainer.innerHTML = '';
                        courses.forEach(course => {{
                            groupContainer.appendChild(createCourseCard(course, semester, false));
                        }});
                        
                        // 선택 제한 업데이트
                        updateSelectionLimit(semester, group, selectionGroupKey);
                    }}
                }});
            }});

            updateSummary();
        }}

        function createCourseCard(course, semester, isRequired) {{
            const card = document.createElement('div');
            card.className = `course-card ${{isRequired ? 'required' : ''}}`;
            
            const isSelected = selectedCourses[semester] && 
                              selectedCourses[semester].some(c => c.name === course.name);
            
            if (isSelected) {{
                card.classList.add('selected');
            }}

            const courseId = `course-${{course.name.replace(/[^a-zA-Z0-9가-힣]/g, '_')}}-${{semester}}`;
            
            // 선택 제한 확인
            let isDisabled = false;
            if (!isRequired && course.selection_group && course.selection_limit) {{
                const groupKey = `${{semester}}_${{course.group}}_${{course.selection_group}}`;
                const groupInfo = selectionGroups[groupKey];
                if (groupInfo && groupInfo.selected.length >= groupInfo.limit && !isSelected) {{
                    isDisabled = true;
                    card.classList.add('disabled');
                }}
            }}

            card.innerHTML = `
                <div class="course-header">
                    <div class="course-name">${{course.name}}</div>
                    <div class="course-credit">${{course.credits}}학점</div>
                </div>
                <div class="course-info">
                    ${{course.group}} | ${{course.type}}
                    ${{course.selection_group ? ` | ${{course.selection_group}}` : ''}}
                </div>
                ${{!isRequired ? `
                <div class="course-checkbox">
                    <input type="checkbox" id="${{courseId}}" ${{isSelected ? 'checked' : ''}} 
                           ${{isDisabled ? 'disabled' : ''}}
                           onchange="toggleCourse('${{semester}}', '${{course.name}}', this)"
                           ontouchstart="this.focus()">
                    <label for="${{courseId}}">선택</label>
                </div>
                ` : '<div style="color: #28a745; font-weight: bold; margin-top: 10px;">✓ 지정과목</div>'}}
            `;

            return card;
        }}

        function toggleCourse(semester, courseName, checkbox) {{
            if (!selectedCourses[semester]) {{
                selectedCourses[semester] = [];
            }}

            const course = courseData.find(c => c.name === courseName);
            const isCurrentlySelected = selectedCourses[semester].some(c => c.name === courseName);

            if (checkbox.checked && !isCurrentlySelected) {{
                // 선택 제한 확인
                if (course.selection_group && course.selection_limit) {{
                    const groupKey = `${{semester}}_${{course.group}}_${{course.selection_group}}`;
                    const groupInfo = selectionGroups[groupKey];
                    
                    if (groupInfo && groupInfo.selected.length >= groupInfo.limit) {{
                        checkbox.checked = false;
                        alert(`${{course.selection_group}} 그룹에서는 최대 ${{groupInfo.limit}}개까지만 선택할 수 있습니다.`);
                        return;
                    }}
                    
                    // 선택 그룹에 추가
                    if (groupInfo) {{
                        groupInfo.selected.push(course);
                    }}
                }}
                
                selectedCourses[semester].push(course);
                checkbox.closest('.course-card').classList.add('selected');
                
            }} else if (!checkbox.checked && isCurrentlySelected) {{
                // 선택 그룹에서 제거
                if (course.selection_group && course.selection_limit) {{
                    const groupKey = `${{semester}}_${{course.group}}_${{course.selection_group}}`;
                    const groupInfo = selectionGroups[groupKey];
                    
                    if (groupInfo) {{
                        groupInfo.selected = groupInfo.selected.filter(c => c.name !== courseName);
                    }}
                }}
                
                selectedCourses[semester] = selectedCourses[semester].filter(c => c.name !== courseName);
                checkbox.closest('.course-card').classList.remove('selected');
            }}

            // 선택 제한 UI 업데이트
            if (course.selection_group && course.selection_limit) {{
                updateSelectionLimit(semester, course.group, course.selection_group);
            }}

            // 다른 과목들의 비활성화 상태 업데이트
            renderCourses(semester);
            
            updateSummary();
        }}

        function updateSelectionLimit(semester, group, selectionGroup) {{
            const groupKey = `${{semester}}_${{group}}_${{selectionGroup}}`;
            const groupInfo = selectionGroups[groupKey];
            
            if (!groupInfo) return;
            
            const countElement = document.getElementById(`count-${{semester}}-${{group}}-${{selectionGroup}}`);
            const groupElement = document.getElementById(`group-${{semester}}-${{group}}-${{selectionGroup}}`);
            
            if (countElement) {{
                const selectedCount = groupInfo.selected.length;
                const limit = groupInfo.limit;
                countElement.textContent = `${{selectedCount}} / ${{limit}}개 선택`;
                
                // 선택 제한 도달 시 스타일 변경
                if (groupElement) {{
                    if (selectedCount >= limit) {{
                        groupElement.classList.add('selection-limit-reached');
                    }} else {{
                        groupElement.classList.remove('selection-limit-reached');
                    }}
                }}
            }}
        }}

        function updateSummary() {{
            const summaryList = document.getElementById('summaryList');
            const totalCreditsElement = document.getElementById('totalCredits');
            
            summaryList.innerHTML = '';
            let totalCredits = 0;

            semesterList.forEach(semester => {{
                const courses = selectedCourses[semester] || [];
                if (courses.length > 0) {{
                    const semesterDiv = document.createElement('div');
                    semesterDiv.innerHTML = `<strong>${{semester}}</strong>`;
                    semesterDiv.style.marginTop = '12px';
                    semesterDiv.style.borderBottom = '1px solid rgba(255,255,255,0.3)';
                    semesterDiv.style.paddingBottom = '4px';
                    semesterDiv.style.fontSize = '0.9em';
                    summaryList.appendChild(semesterDiv);

                    courses.forEach(course => {{
                        const courseItem = document.createElement('div');
                        courseItem.className = 'selected-course-item';
                        courseItem.innerHTML = `
                            <span>${{course.name}}</span>
                            <span>${{course.credits}}학점</span>
                        `;
                        summaryList.appendChild(courseItem);
                        totalCredits += course.credits;
                    }});
                }}
            }});

            totalCreditsElement.textContent = `총 학점: ${{totalCredits}}학점`;
        }}

        // 모바일 터치 최적화
        document.addEventListener('touchstart', function() {{}}, false);
        
        // iOS Safari 뷰포트 높이 문제 해결
        function updateViewportHeight() {{
            const vh = window.innerHeight * 0.01;
            document.documentElement.style.setProperty('--vh', `${{vh}}px`);
        }}
        
        window.addEventListener('resize', updateViewportHeight);
        updateViewportHeight();
    </script>
</body>
</html>'''

def create_gui():
    """GUI 인터페이스 생성"""
    root = tk.Tk()
    root.title("과목 선택 시뮬레이션 생성기")
    root.geometry("500x300")
    root.configure(bg='#f0f0f0')
    
    # 스타일 설정
    style = ttk.Style()
    style.theme_use('clam')
    
    # 메인 프레임
    main_frame = ttk.Frame(root, padding="20")
    main_frame.pack(fill=tk.BOTH, expand=True)
    
    # 제목
    title_label = ttk.Label(main_frame, text="🎓 과목 선택 시뮬레이션 생성기", 
                           font=('Arial', 16, 'bold'))
    title_label.pack(pady=(0, 20))
    
    # 설명
    desc_label = ttk.Label(main_frame, 
                          text="엑셀 파일을 선택하여 인터랙티브한 과목 선택 시뮬레이션을 생성합니다.",
                          font=('Arial', 10))
    desc_label.pack(pady=(0, 30))
    
    # 버튼 프레임
    button_frame = ttk.Frame(main_frame)
    button_frame.pack(pady=20)
    
    # 파일 선택 버튼
    select_btn = ttk.Button(button_frame, text="📁 엑셀 파일 선택", 
                           command=lambda: process_file(root),
                           style='Accent.TButton')
    select_btn.pack(pady=10)
    
    # 상태 라벨
    status_label = ttk.Label(main_frame, text="파일을 선택해주세요.", 
                            font=('Arial', 9), foreground='gray')
    status_label.pack(pady=10)
    
    def process_file(parent):
        generator = CourseSimulatorGenerator()
        
        # 파일 선택
        file_path = generator.select_excel_file()
        if not file_path:
            return
        
        status_label.config(text=f"선택된 파일: {Path(file_path).name}")
        parent.update()
        
        # 데이터 로드
        if not generator.load_excel_data(file_path):
            messagebox.showerror("오류", "엑셀 파일 로드에 실패했습니다.")
            return
        
        # 데이터 처리
        if not generator.process_data():
            messagebox.showerror("오류", "데이터 처리에 실패했습니다.\\n필수 컬럼을 확인해주세요.")
            return
        
        # HTML 생성
        output_path = generator.generate_html()
        if output_path:
            result = messagebox.askyesno("완료", 
                f"HTML 파일이 생성되었습니다!\\n\\n파일: {output_path}\\n\\n지금 열어보시겠습니까?")
            
            if result:
                webbrowser.open(f"file://{os.path.abspath(output_path)}")
        else:
            messagebox.showerror("오류", "HTML 생성에 실패했습니다.")
    
    return root

def main():
    """메인 함수"""
    print("🎓 과목 선택 시뮬레이션 HTML 생성기")
    print("=" * 50)
    
    if len(sys.argv) > 1:
        # 커맨드라인 모드
        file_path = sys.argv[1]
        generator = CourseSimulatorGenerator()
        
        if generator.load_excel_data(file_path) and generator.process_data():
            output_path = generator.generate_html()
            if output_path:
                print(f"\\n🎉 완료! HTML 파일: {output_path}")
        else:
            print("❌ 파일 처리 실패")
    else:
        # GUI 모드
        try:
            root = create_gui()
            root.mainloop()
        except ImportError:
            print("GUI가 지원되지 않습니다. 커맨드라인 모드를 사용하세요:")
            print("python course_generator.py <엑셀파일경로>")

if __name__ == "__main__":
    main()