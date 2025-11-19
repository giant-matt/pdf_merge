import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, Listbox, ttk
import os
import re
from pathlib import Path
from pypdf import PdfWriter
from PIL import Image
import threading
import time
import win32com.client
import pythoncom

class PdfMergerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 변환 & 취합 프로그램")
        self.root.geometry("750x750")
        self.root.resizable(True, True)

        # 아이콘 설정
        try:
            import sys
            if getattr(sys, 'frozen', False):
                # PyInstaller로 빌드된 경우
                base_path = sys._MEIPASS
            else:
                # 개발 환경
                base_path = os.path.dirname(os.path.abspath(__file__))
            icon_path = os.path.join(base_path, 'assests', 'pdf_merge_icon.ico')
            self.root.iconbitmap(icon_path)
        except:
            pass

        # 심플하고 조화로운 컬러 팔레트
        self.colors = {
            'bg': '#F5F5F5',
            'card_bg': '#FFFFFF',
            'primary': '#5B7EFF',
            'primary_hover': '#4A6FEE',
            'success': '#34C759',
            'success_hover': '#2DB84D',
            'text': '#000000',
            'text_secondary': '#6B6B6B',
            'border': '#E5E5E5',
            'guide_bg': '#F8F9FA',
            'guide_border': '#DEE2E6',
            'guide_text': '#2C2C2C',
            'button_text': '#FFFFFF'
        }

        # 폰트 폴백 설정 (Noto Sans KR 우선, 없으면 맑은 고딕)
        import tkinter.font as tkfont
        available_fonts = tkfont.families()

        # Noto Sans KR 여러 이름으로 확인
        font_family = '맑은 고딕'  # 기본값
        for font_name in ['Noto Sans KR', 'Noto Sans Korean', 'NotoSansKR']:
            if font_name in available_fonts:
                font_family = font_name
                break

        self.fonts = {
            'title': (font_family, 16, 'bold'),
            'heading': (font_family, 11, 'bold'),
            'body': (font_family, 9),
            'button': (font_family, 10, 'bold'),
            'small': (font_family, 9)
        }

        # 백그라운드 색상 설정
        self.root.configure(bg=self.colors['bg'])

        self.folder_path = tk.StringVar()
        self.image_extensions = ['.jpg', '.jpeg', '.png']
        self.doc_extensions = ['.docx', '.doc', '.hwp', '.hwpx', '.xlsx', '.xls', '.pptx', '.ppt']

        # --- 상수 정의 ---
        self.TEMP_PREFIX = "__temp_"
        self.MERGED_SUFFIX = "_merged.pdf"
        self.ORIGINALS_DIR = "원본"

        # --- GUI 구성 요소 ---

        # 메인 컨테이너
        main_container = tk.Frame(root, bg=self.colors['bg'], padx=20, pady=20)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 타이틀
        title_label = tk.Label(
            main_container,
            text="PDF 변환 & 취합 프로그램",
            font=self.fonts['title'],
            fg=self.colors['text'],
            bg=self.colors['bg']
        )
        title_label.pack(pady=(0, 15))

        # 카드 프레임: 프로그램 소개
        guide_card = tk.Frame(
            main_container,
            bg=self.colors['guide_bg'],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors['guide_border']
        )
        guide_card.pack(fill=tk.X, pady=(0, 15))

        guide_inner = tk.Frame(guide_card, bg=self.colors['guide_bg'], padx=20, pady=15)
        guide_inner.pack(fill=tk.X)

        # 가이드 타이틀
        guide_title_frame = tk.Frame(guide_inner, bg=self.colors['guide_bg'])
        guide_title_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(
            guide_title_frame,
            text="📖 프로그램 소개",
            font=self.fonts['heading'],
            fg=self.colors['text'],
            bg=self.colors['guide_bg']
        ).pack(side=tk.LEFT)

        # 접기/펼치기 버튼
        self.guide_visible = tk.BooleanVar(value=False)
        self.toggle_btn = tk.Button(
            guide_title_frame,
            text="▼ 펼치기",
            command=self.toggle_guide,
            font=self.fonts['small'],
            bg=self.colors['guide_bg'],
            fg=self.colors['text_secondary'],
            relief=tk.FLAT,
            cursor='hand2',
            borderwidth=0
        )
        self.toggle_btn.pack(side=tk.RIGHT)

        # 가이드 내용 프레임 (처음에는 숨김)
        self.guide_content = tk.Frame(guide_inner, bg=self.colors['guide_bg'])

        # 가이드 내용
        guide_text = tk.Text(
            self.guide_content,
            font=self.fonts['small'],
            bg=self.colors['guide_bg'],
            fg=self.colors['guide_text'],
            relief=tk.FLAT,
            highlightthickness=0,
            borderwidth=0,
            wrap=tk.WORD,
            height=18,
            state='normal',
            cursor='arrow'
        )
        guide_text.pack(fill=tk.X)

        guide_content_text = """
📌 주요 기능

이 프로그램은 여러 형식의 증빙자료를 하나의 PDF 파일로 자동 병합합니다.

✅ 지원 파일 형식
   • PDF 파일 (.pdf)
   • 이미지 파일 (.jpg, .jpeg, .png)
   • 워드 문서 (.doc, .docx)
   • 엑셀 파일 (.xls, .xlsx)
   • 파워포인트 (.ppt, .pptx)
   • 한글 문서 (.hwp, .hwpx)

🔄 작동 방식
   1. 폴더를 선택하면 자동으로 파일 목록이 표시됩니다
   2. 파일 순서를 원하는 대로 조정할 수 있습니다
   3. '병합 시작'을 클릭하면 모든 파일이 하나의 PDF로 병합됩니다
   4. 이미지, 오피스, 한글 파일은 자동으로 PDF로 변환됩니다

💡 유의사항
   • Office 파일 변환을 위해서는 MS Office가 설치되어 있어야 합니다
   • 한글 문서 변환을 위해서는 한/글 프로그램이 설치되어 있어야 합니다
   • Office 2016, O365 모두 지원됩니다
   • 변환 실패한 파일은 병합에서 제외되며 상세 내역이 로그에 표시됩니다
"""
        guide_text.insert('1.0', guide_content_text)
        guide_text.config(state='disabled')

        # 카드 프레임: 폴더 선택
        folder_card = tk.Frame(
            main_container,
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        folder_card.pack(fill=tk.X, pady=(0, 15))

        folder_inner = tk.Frame(folder_card, bg=self.colors['card_bg'], padx=20, pady=15)
        folder_inner.pack(fill=tk.X)

        tk.Label(
            folder_inner,
            text="병합할 폴더",
            font=self.fonts['heading'],
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, pady=(0, 8))

        folder_entry_frame = tk.Frame(folder_inner, bg=self.colors['card_bg'])
        folder_entry_frame.pack(fill=tk.X)

        self.folder_entry = tk.Entry(
            folder_entry_frame,
            textvariable=self.folder_path,
            font=self.fonts['body'],
            state='readonly',
            relief=tk.FLAT,
            bg='#F1F3F5',
            fg=self.colors['text'],
            highlightthickness=1,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['primary']
        )
        self.folder_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=8, padx=(0, 10))

        select_btn = tk.Button(
            folder_entry_frame,
            text="찾아보기",
            command=self.select_folder,
            font=self.fonts['button'],
            bg=self.colors['primary'],
            fg=self.colors['button_text'],
            relief=tk.FLAT,
            cursor='hand2',
            padx=20,
            pady=8
        )
        select_btn.pack(side=tk.LEFT)
        select_btn.bind('<Enter>', lambda e: select_btn.config(bg=self.colors['primary_hover']))
        select_btn.bind('<Leave>', lambda e: select_btn.config(bg=self.colors['primary']))

        # 카드 프레임: 파일 목록
        list_card = tk.Frame(
            main_container,
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        list_card.pack(expand=True, fill=tk.BOTH, pady=(0, 15))

        list_inner = tk.Frame(list_card, bg=self.colors['card_bg'], padx=20, pady=15)
        list_inner.pack(expand=True, fill=tk.BOTH)

        tk.Label(
            list_inner,
            text="파일 목록",
            font=self.fonts['heading'],
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, pady=(0, 8))

        tk.Label(
            list_inner,
            text="파일을 선택 후 ▲▼ 버튼으로 순서 조정, ✕ 버튼으로 제외할 수 있습니다",
            font=self.fonts['small'],
            fg=self.colors['text_secondary'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, pady=(0, 10))

        # 파일 목록과 버튼을 담는 프레임
        listbox_frame = tk.Frame(list_inner, bg=self.colors['card_bg'])
        listbox_frame.pack(expand=True, fill=tk.BOTH)

        # 파일 목록 리스트박스
        self.file_listbox = Listbox(
            listbox_frame,
            selectmode=tk.SINGLE,
            font=self.fonts['heading'],
            bg='#F8F9FA',
            fg=self.colors['text'],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors['border'],
            highlightcolor=self.colors['primary'],
            selectbackground=self.colors['primary'],
            selectforeground='white',
            activestyle='none',
            borderwidth=0
        )
        self.file_listbox.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, padx=(0, 10))

        # 순서 조정 버튼 프레임
        button_sub_frame = tk.Frame(listbox_frame, bg=self.colors['card_bg'])
        button_sub_frame.pack(side=tk.LEFT, fill=tk.Y)

        up_btn = tk.Button(
            button_sub_frame,
            text="▲",
            command=self.move_up,
            font=self.fonts['button'],
            bg=self.colors['card_bg'],
            fg=self.colors['text'],
            relief=tk.FLAT,
            cursor='hand2',
            width=4,
            height=2,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        up_btn.pack(pady=(0, 5))
        up_btn.bind('<Enter>', lambda e: up_btn.config(bg=self.colors['border']))
        up_btn.bind('<Leave>', lambda e: up_btn.config(bg=self.colors['card_bg']))

        down_btn = tk.Button(
            button_sub_frame,
            text="▼",
            command=self.move_down,
            font=self.fonts['button'],
            bg=self.colors['card_bg'],
            fg=self.colors['text'],
            relief=tk.FLAT,
            cursor='hand2',
            width=4,
            height=2,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        down_btn.pack(pady=(0, 5))
        down_btn.bind('<Enter>', lambda e: down_btn.config(bg=self.colors['border']))
        down_btn.bind('<Leave>', lambda e: down_btn.config(bg=self.colors['card_bg']))

        remove_btn = tk.Button(
            button_sub_frame,
            text="✕",
            command=self.remove_file,
            font=self.fonts['button'],
            bg=self.colors['card_bg'],
            fg='#E74C3C',  # 빨간색
            relief=tk.FLAT,
            cursor='hand2',
            width=4,
            height=2,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        remove_btn.pack()
        remove_btn.bind('<Enter>', lambda e: remove_btn.config(bg='#FFEBEE'))
        remove_btn.bind('<Leave>', lambda e: remove_btn.config(bg=self.colors['card_bg']))

        # 실행 버튼
        self.merge_button = tk.Button(
            main_container,
            text="병합 시작",
            command=self.start_merge_thread,
            font=self.fonts['button'],
            bg=self.colors['success'],
            fg=self.colors['button_text'],
            relief=tk.FLAT,
            cursor='hand2',
            pady=12
        )
        self.merge_button.pack(fill=tk.X, pady=(0, 15))
        self.merge_button.bind('<Enter>', lambda e: self.merge_button.config(bg=self.colors['success_hover']))
        self.merge_button.bind('<Leave>', lambda e: self.merge_button.config(bg=self.colors['success']))

        # 진행률 바
        progress_frame = tk.Frame(main_container, bg=self.colors['bg'])
        progress_frame.pack(fill=tk.X, pady=(0, 15))

        self.progress_label = tk.Label(
            progress_frame,
            text="대기 중",
            font=self.fonts['small'],
            fg=self.colors['text_secondary'],
            bg=self.colors['bg']
        )
        self.progress_label.pack(anchor=tk.W, pady=(0, 5))

        self.progress_bar = ttk.Progressbar(
            progress_frame,
            mode='determinate',
            length=100
        )
        self.progress_bar.pack(fill=tk.X)

        # 카드 프레임: 로그
        log_card = tk.Frame(
            main_container,
            bg=self.colors['card_bg'],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightbackground=self.colors['border']
        )
        log_card.pack(fill=tk.BOTH, expand=False)

        log_inner = tk.Frame(log_card, bg=self.colors['card_bg'], padx=20, pady=15)
        log_inner.pack(fill=tk.BOTH, expand=True)

        tk.Label(
            log_inner,
            text="진행 상황",
            font=self.fonts['heading'],
            fg=self.colors['text'],
            bg=self.colors['card_bg']
        ).pack(anchor=tk.W, pady=(0, 8))

        self.log_area = scrolledtext.ScrolledText(
            log_inner,
            height=6,
            state='disabled',
            font=self.fonts['small'],
            bg='#F8F9FA',
            fg=self.colors['text'],
            relief=tk.FLAT,
            highlightthickness=0,
            borderwidth=0,
            wrap=tk.WORD
        )
        self.log_area.pack(fill=tk.BOTH, expand=True)

    def toggle_guide(self):
        """가이드 내용을 접거나 펼칩니다."""
        if self.guide_visible.get():
            # 숨기기
            self.guide_content.pack_forget()
            self.toggle_btn.config(text="▼ 펼치기")
            self.guide_visible.set(False)
        else:
            # 보이기
            self.guide_content.pack(fill=tk.X, pady=(0, 5))
            self.toggle_btn.config(text="▲ 접기")
            self.guide_visible.set(True)

    def update_progress(self, value, text):
        """진행률 바와 레이블을 업데이트합니다."""
        self.root.after(0, self._progress_update, value, text)

    def _progress_update(self, value, text):
        self.progress_bar['value'] = value
        self.progress_label.config(text=f"{text} ({int(value)}%)")

    def log(self, message):
        """로그 영역에 메시지를 추가합니다."""
        self.root.after(0, self._log_update, message)

    def _log_update(self, message):
        self.log_area.config(state='normal')
        self.log_area.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_area.see(tk.END)
        self.log_area.config(state='disabled')

    def select_folder(self):
        """폴더 선택 대화상자를 엽니다."""
        directory = filedialog.askdirectory()
        if directory:
            self.folder_path.set(directory)
            self.update_file_list()
            self.log(f"선택된 폴더: {directory}")

    def update_file_list(self):
        """리스트박스에 파일 목록을 업데이트합니다."""
        self.file_listbox.delete(0, tk.END)
        folder = Path(self.folder_path.get())
        if folder.is_dir():
            def natural_sort_key(s):
                return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', s)]

            # 모든 파일 확인
            all_files = []
            for f in folder.iterdir():
                if f.is_file():
                    ext = f.suffix.lower()
                    # PDF, 이미지, 문서 파일 모두 포함
                    if ext == '.pdf' or ext in self.image_extensions or ext in self.doc_extensions:
                        all_files.append(f.name)

            # 정렬 후 리스트에 추가
            files = sorted(all_files, key=natural_sort_key)
            for file_name in files:
                self.file_listbox.insert(tk.END, file_name)

    def move_up(self): self.move_item(-1)
    def move_down(self): self.move_item(1)

    def move_item(self, direction):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices: return
        idx = selected_indices[0]
        new_idx = idx + direction
        if 0 <= new_idx < self.file_listbox.size():
            item = self.file_listbox.get(idx)
            self.file_listbox.delete(idx)
            self.file_listbox.insert(new_idx, item)
            self.file_listbox.selection_set(new_idx)
            self.file_listbox.activate(new_idx)

    def remove_file(self):
        """선택된 파일을 목록에서 제거합니다."""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("경고", "제거할 파일을 선택해주세요.")
            return
        idx = selected_indices[0]
        file_name = self.file_listbox.get(idx)

        # 확인 메시지
        result = messagebox.askyesno("확인", f"'{file_name}'을(를) 병합 목록에서 제외하시겠습니까?\n\n(파일은 삭제되지 않습니다)")
        if result:
            self.file_listbox.delete(idx)
            self.log(f"목록에서 제외: {file_name}")
            # 다음 항목 선택 (있다면)
            if self.file_listbox.size() > 0:
                new_idx = min(idx, self.file_listbox.size() - 1)
                self.file_listbox.selection_set(new_idx)
                self.file_listbox.activate(new_idx)

    def start_merge_thread(self):
        if not self.folder_path.get():
            messagebox.showerror("오류", "먼저 병합할 파일이 있는 폴더를 선택해주세요.")
            return
        self.merge_button.config(state='disabled', text="병합 중...")
        thread = threading.Thread(target=self.merge_files)
        thread.daemon = True
        thread.start()

    def convert_doc_to_pdf(self, doc_path, output_pdf_path):
        """Word/Excel/PowerPoint/HWP 문서를 PDF로 변환"""
        pythoncom.CoInitialize()
        try:
            ext = doc_path.suffix.lower()

            # 절대 경로로 변환하고 문자열로 변환
            input_path = os.path.abspath(str(doc_path))
            output_path = os.path.abspath(str(output_pdf_path))

            if ext in ['.docx', '.doc']:
                # Word 문서 변환
                try:
                    word = win32com.client.Dispatch("Word.Application")
                except Exception as e:
                    raise Exception(f"MS Word를 찾을 수 없습니다. Word가 설치되어 있는지 확인하세요. ({str(e)})")

                word.Visible = False
                try:
                    doc = word.Documents.Open(input_path)
                    doc.SaveAs(output_path, FileFormat=17)  # 17 = PDF
                    doc.Close()
                finally:
                    try:
                        word.Quit()
                    except:
                        pass

            elif ext in ['.xlsx', '.xls']:
                # Excel 문서 변환
                try:
                    excel = win32com.client.Dispatch("Excel.Application")
                except Exception as e:
                    raise Exception(f"MS Excel을 찾을 수 없습니다. Excel이 설치되어 있는지 확인하세요. ({str(e)})")

                excel.Visible = False
                excel.DisplayAlerts = False
                try:
                    workbook = excel.Workbooks.Open(input_path)
                    # PDF 형식으로 저장 (0 = xlTypePDF)
                    workbook.ExportAsFixedFormat(0, output_path)
                    workbook.Close(SaveChanges=False)
                finally:
                    try:
                        excel.Quit()
                    except:
                        pass

            elif ext in ['.pptx', '.ppt']:
                # PowerPoint 문서 변환
                try:
                    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
                except Exception as e:
                    raise Exception(f"MS PowerPoint를 찾을 수 없습니다. PowerPoint가 설치되어 있는지 확인하세요. ({str(e)})")

                try:
                    presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
                    # PDF 형식으로 저장 (32 = ppSaveAsPDF)
                    presentation.SaveAs(output_path, 32)
                    presentation.Close()
                finally:
                    try:
                        powerpoint.Quit()
                    except:
                        pass

            elif ext in ['.hwp', '.hwpx']:
                # 한글 문서 변환
                self.log(f"    [디버그] 한/글 프로그램 초기화 중...")
                try:
                    hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
                except Exception as e:
                    raise Exception(f"한/글 프로그램을 찾을 수 없습니다. 한/글이 설치되어 있는지 확인하세요. ({str(e)})")

                try:
                    # 보안 경고 무시 설정
                    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
                    hwp.SetMessageBoxMode(0x00010000)  # 메시지 박스 자동 확인

                    # 파일 열기
                    self.log(f"    [디버그] 파일 열기 시도: {input_path}")
                    result = hwp.Open(input_path, "HWP", "forceopen:true")
                    if not result:
                        raise Exception("파일 열기 실패 (hwp.Open 반환값: False)")

                    self.log(f"    [디버그] 파일 열기 성공")

                    # PDF로 저장 - HAction 사용 방식
                    self.log(f"    [디버그] PDF 저장 시도: {output_path}")

                    # HAction을 이용한 PDF 저장
                    act = hwp.CreateAction("FileSaveAs")
                    pset = act.CreateSet()
                    act.GetDefault(pset)
                    pset.SetItem("Format", "PDF")
                    pset.SetItem("FileName", output_path)
                    result = act.Execute(pset)

                    if not result:
                        raise Exception("PDF 저장 실패 (HAction.Execute 반환값: False)")

                    self.log(f"    [디버그] PDF 저장 완료")

                    # 파일 닫기
                    hwp.Clear(1)  # 1 = 저장하지 않고 닫기
                except Exception as e:
                    raise Exception(f"한글 변환 중 오류: {str(e)}")
                finally:
                    try:
                        hwp.Quit()
                    except:
                        pass

        except Exception as e:
            raise Exception(f"{doc_path.name} 변환 실패: {str(e)}")
        finally:
            pythoncom.CoUninitialize()

    def merge_files(self):
        temp_pdf_paths = []
        # 성공/실패 추적
        successfully_merged = []
        failed_files = []

        try:
            self.update_progress(0, "시작 중")

            source_folder = Path(self.folder_path.get())
            original_files_to_process = [source_folder / f for f in self.file_listbox.get(0, tk.END)]

            if not original_files_to_process:
                self.log("병합할 파일이 목록에 없습니다.")
                messagebox.showinfo("완료", "병합할 파일이 목록에 없습니다.")
                return

            total_files = len(original_files_to_process)
            current_file = 0

            # 이미지 파일 변환
            image_files = [f for f in original_files_to_process if f.suffix.lower() in self.image_extensions]
            if image_files:
                self.log("이미지 파일을 PDF로 변환 시작...")

            for img_path in image_files:
                self.log(f"  -> 변환 중: {img_path.name}")
                self.update_progress((current_file / total_files) * 50, "이미지 변환 중")
                try:
                    image = Image.open(img_path).convert("RGB")
                    temp_pdf_path = source_folder / f"{self.TEMP_PREFIX}{img_path.stem}.pdf"
                    image.save(temp_pdf_path)
                    temp_pdf_paths.append(temp_pdf_path)
                    successfully_merged.append(img_path.name)
                    self.log(f"  ✓ 변환 성공: {img_path.name}")
                except Exception as e:
                    failed_files.append((img_path.name, str(e)))
                    self.log(f"  ⚠️ 변환 실패: {img_path.name} - {str(e)}")
                current_file += 1

            # 문서 파일 변환
            doc_files = [f for f in original_files_to_process if f.suffix.lower() in self.doc_extensions]
            if doc_files:
                self.log("문서 파일을 PDF로 변환 시작...")

            for doc_path in doc_files:
                self.log(f"  -> 변환 중: {doc_path.name}")
                self.update_progress(50 + (current_file / total_files) * 30, "문서 변환 중")
                temp_pdf_path = source_folder / f"{self.TEMP_PREFIX}{doc_path.stem}.pdf"
                try:
                    self.convert_doc_to_pdf(doc_path, temp_pdf_path)
                    if temp_pdf_path.exists():
                        temp_pdf_paths.append(temp_pdf_path)
                        successfully_merged.append(doc_path.name)
                        self.log(f"  ✓ 변환 성공: {doc_path.name}")
                    else:
                        failed_files.append((doc_path.name, "PDF 파일이 생성되지 않음"))
                        self.log(f"  ⚠️ 변환 실패: {doc_path.name} (PDF 파일이 생성되지 않음)")
                except Exception as e:
                    failed_files.append((doc_path.name, str(e)))
                    self.log(f"  ⚠️ 변환 실패: {doc_path.name} - {str(e)}")
                    # 변환 실패해도 계속 진행
                current_file += 1

            self.update_progress(80, "PDF 병합 준비 중")
            self.log("PDF 병합을 시작합니다...")
            merger = PdfWriter()
            all_pdf_files = []
            for f in original_files_to_process:
                if f.suffix.lower() in self.image_extensions or f.suffix.lower() in self.doc_extensions:
                    # 이미지나 문서 파일은 변환된 PDF 사용
                    temp_pdf = source_folder / f"{self.TEMP_PREFIX}{f.stem}.pdf"
                    if temp_pdf.exists():  # 변환 성공한 파일만 추가
                        all_pdf_files.append(temp_pdf)
                    else:
                        self.log(f"  ⚠️ 건너뛰기: {f.name} (변환 실패)")
                        # 이미 failed_files에 추가되어 있음
                else:
                    # 원본 PDF 사용
                    if f.exists():
                        all_pdf_files.append(f)
                        successfully_merged.append(f.name)
                    else:
                        failed_files.append((f.name, "파일을 찾을 수 없음"))
                        self.log(f"  ⚠️ 건너뛰기: {f.name} (파일을 찾을 수 없음)")

            total_pdfs = len(all_pdf_files)
            for idx, pdf_path in enumerate(all_pdf_files):
                self.log(f"  -> 추가: {pdf_path.name.replace(self.TEMP_PREFIX, '')}")
                self.update_progress(80 + (idx / total_pdfs) * 15, "PDF 병합 중")
                merger.append(str(pdf_path))

            self.update_progress(95, "파일 저장 중")
            output_filename = f"{source_folder.name}{self.MERGED_SUFFIX}"
            output_path = source_folder / output_filename
            with open(output_path, "wb") as output_file:
                merger.write(output_file)
            merger.close()

            self.update_progress(100, "완료!")

            # 최종 요약 메시지
            self.log("\n" + "="*60)
            self.log("📊 병합 완료 요약")
            self.log("="*60)
            self.log(f"총 파일 수: {total_files}개")
            self.log(f"성공적으로 병합된 파일: {len(successfully_merged)}개")
            self.log(f"실패한 파일: {len(failed_files)}개")

            if successfully_merged:
                self.log("\n✅ 병합에 포함된 파일:")
                for idx, file_name in enumerate(successfully_merged, 1):
                    self.log(f"  {idx}. {file_name}")

            if failed_files:
                self.log("\n⚠️ 병합에서 제외된 파일:")
                for idx, (file_name, reason) in enumerate(failed_files, 1):
                    self.log(f"  {idx}. {file_name} - {reason}")

            self.log(f"\n💾 저장된 파일: {output_path}")
            self.log("="*60 + "\n")

            # 메시지 박스 내용도 요약 포함
            summary_msg = f"병합이 완료되었습니다!\n\n"
            summary_msg += f"총 {total_files}개 중 {len(successfully_merged)}개 파일 병합 성공\n"

            if failed_files:
                summary_msg += f"\n⚠️ {len(failed_files)}개 파일 실패:\n"
                for file_name, reason in failed_files:
                    # 이유가 너무 길면 축약
                    short_reason = reason if len(reason) < 50 else reason[:47] + "..."
                    summary_msg += f"  • {file_name}\n    ({short_reason})\n"
                summary_msg += "\n자세한 내용은 아래 '진행 상황'을 확인하세요.\n"

            summary_msg += f"\n저장된 파일:\n{output_path}"

            messagebox.showinfo("성공", summary_msg)

        except Exception as e:
            self.log(f"\n❌ 오류: 병합 중 문제가 발생했습니다. {e}")
            messagebox.showerror("오류", f"병합 중 오류가 발생했습니다:\n{e}")

        finally:
            self.log("임시 파일을 삭제합니다.")
            for temp_path in temp_pdf_paths:
                if temp_path.exists():
                    os.remove(temp_path)
            
            self.root.after(0, self._finalize_ui)

    def _finalize_ui(self):
        self.merge_button.config(state='normal', text="병합 시작")
        self.update_progress(0, "대기 중")
        self.update_file_list()

if __name__ == "__main__":
    root = tk.Tk()
    app = PdfMergerApp(root)
    root.mainloop()