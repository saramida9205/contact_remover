import sys
import os
import shutil
import pandas as pd
import openpyxl
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import msoffcrypto
import io
import time
import socket
import json
import uuid
from win32_setctime import setctime

# 연락처 컬럼을 삭제하는 파이썬 핵심 로직 클래스 (UI와 분리)
class ContactRemoverCore:
    def __init__(self, stop_event, update_callback):
        self.stop_event = stop_event
        self.update_callback = update_callback  # (index, total, filename, status_text)
        self.target_keywords = ['연락처', '전화번호', '휴대폰', '핸드폰', '전화']

    def process_files(self, files_to_process):
        success_count = 0
        error_messages = []
        success_messages = []
        total_files = len(files_to_process)

        for i, file_path in enumerate(files_to_process):
            if self.stop_event.is_set():
                break

            filename = os.path.basename(file_path)
            self.update_callback(i + 1, total_files, filename, f"처리 중...")

            if not os.path.exists(file_path):
                continue

            # 1. 파일 시간 정보 백업
            try:
                stat = os.stat(file_path)
                c_time, m_time, a_time = stat.st_ctime, stat.st_mtime, stat.st_atime
            except:
                c_time, m_time, a_time = None, None, None

            # 2. 암호화 체크
            if self._is_encrypted(file_path):
                error_messages.append(f"[{filename}] 암호화된 파일입니다.")
                continue

            # 3. 파일 처리
            try:
                result = self._process_single_file(file_path)
                if result == "SUCCESS":
                    success_count += 1
                    success_messages.append(f"[{filename}] 완료")
                    # 시간 복원 (새 파일에 적용)
                    if c_time is not None:
                        # process_single_file에서 원본이 교체되었으므로 file_path는 새로운 파일임
                        try:
                            os.utime(file_path, (a_time, m_time))
                            setctime(file_path, c_time)
                        except: pass
                elif result == "NO_TARGET":
                    error_messages.append(f"[{filename}] '연락처' 연관 열 없음.")
                else:
                    error_messages.append(f"[{filename}] {result}")
            except Exception as e:
                error_messages.append(f"[{filename}] 오류: {str(e)}")

        return success_count, success_messages, error_messages

    def _is_encrypted(self, file_path):
        try:
            with open(file_path, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                return office_file.is_encrypted()
        except:
            return False

    def _process_single_file(self, file_path):
        dfs = {}
        # 가변 엔진 시도 (calamine -> html -> csv)
        try:
            dfs = pd.read_excel(file_path, engine='calamine', sheet_name=None)
        except Exception as e:
            err = str(e).lower()
            if 'encrypted' in err or 'password' in err: return "암호화된 파일"
            # 가짜 엑셀 대응
            try:
                from_html = pd.read_html(file_path, encoding='cp949')
                dfs = {"Sheet1": from_html[0]}
            except:
                try:
                    dfs = {"Sheet1": pd.read_csv(file_path, encoding='cp949', sep=None, engine='python')}
                except:
                    return "지원하지 않는 파일 형식"

        processed_dfs = {}
        any_dropped = False

        for sheet_name, df in dfs.items():
            if df is None or df.empty:
                processed_dfs[sheet_name] = df
                continue

            # 헤더 정규화
            df.columns = [str(col).strip() for col in df.columns]
            
            # 지능형 헤더 매칭 (상단 15행 스캔)
            header_found = False
            if not any(kw in "".join(df.columns).replace(' ', '') for kw in self.target_keywords):
                for idx, row in df.head(15).iterrows():
                    row_vals = [str(x).strip() for x in row.values]
                    if any(kw in "".join(row_vals).replace(' ', '') for kw in self.target_keywords):
                        df.columns = row_vals
                        df = df.iloc[idx+1:].reset_index(drop=True)
                        header_found = True
                        break
            
            # 삭제 대상 열 찾기
            cols_to_drop = [col for col in df.columns if any(kw in str(col).replace(' ', '') for kw in self.target_keywords)]
            if cols_to_drop:
                df = df.drop(columns=cols_to_drop)
                any_dropped = True
            
            processed_dfs[sheet_name] = df

        if not any_dropped:
            return "NO_TARGET"

        # 원자적 저장 (Atomic Save)
        dir_name = os.path.dirname(file_path)
        base_name = os.path.basename(file_path)
        temp_name = f"~tmp_{uuid.uuid4().hex}.xlsx"
        temp_path = os.path.join(dir_name, temp_name)

        try:
            with pd.ExcelWriter(temp_path, engine='openpyxl') as writer:
                for sheet_name, df in processed_dfs.items():
                    if df is not None:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            # 성공 시 교체
            os.remove(file_path)
            # 원본 확장자가 .xls 였다면 .xlsx로 변경될 수 있음 (코드 단순화를 위해 .xlsx로 통합 권장)
            final_path = os.path.splitext(file_path)[0] + ".xlsx"
            if os.path.exists(final_path) and final_path != file_path:
                os.remove(final_path) # 충돌 방지
            os.rename(temp_path, final_path)
            return "SUCCESS"
        except Exception as e:
            if os.path.exists(temp_path): os.remove(temp_path)
            return f"저장 실패: {str(e)}"

class ContactRemoverApp:
    def __init__(self, root, target_files=None, auto_run=False):
        self.root = root
        self.root.title("엑셀 연락처 삭제 도구 (Pro)")
        self.root.geometry("450x300")
        
        self.target_files = target_files if target_files else []
        self.stop_event = threading.Event()
        self.is_running = False

        # UI 구성
        tk.Label(root, text="작업할 엑셀 파일:", font=("맑은 고딕", 9)).pack(pady=(15, 0))
        
        self.target_file_var = tk.StringVar()
        self._update_file_label()

        entry_frame = tk.Frame(root)
        entry_frame.pack(fill="x", padx=20, pady=5)
        self.file_entry = tk.Entry(entry_frame, textvariable=self.target_file_var, state='readonly')
        self.file_entry.pack(side="left", expand=True, fill="x")
        self.browse_btn = tk.Button(entry_frame, text="파일 선택", command=self.browse_file)
        self.browse_btn.pack(side="right", padx=(5, 0))
        
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=15)
        
        self.start_btn = tk.Button(btn_frame, text="작업시작", command=self.on_start_click, width=12, height=2, bg="#4CAF50", fg="white", font=("맑은 고딕", 10, "bold"))
        self.start_btn.pack(side="left", padx=10)
        
        self.stop_btn = tk.Button(btn_frame, text="작업중단", command=self.on_stop_click, width=12, height=2, state="disabled", bg="#f44336", fg="white")
        self.stop_btn.pack(side="left", padx=10)

        self.close_btn = tk.Button(btn_frame, text="닫기", command=self.root.destroy, width=10, height=2)
        self.close_btn.pack(side="left", padx=10)

        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
        self.progress_bar.pack(pady=5)
        
        self.status_label = tk.Label(root, text="대기 중...", fg="blue", font=("맑은 고딕", 9))
        self.status_label.pack()

        # IPC & AutoRun
        if auto_run:
            self.start_ipc_listen()
            self.root.after(500, self.auto_start_logic)

    def _update_file_label(self):
        if len(self.target_files) == 1:
            self.target_file_var.set(self.target_files[0])
        elif len(self.target_files) > 1:
            self.target_file_var.set(f"{len(self.target_files)}개 파일 선택됨")
        else:
            self.target_file_var.set("")

    def auto_start_logic(self):
        self.status_label.config(text="병합 대기 중... (우클릭 입력을 기다립니다)")
        # 0.5초 후 시작 버튼 활성화 (완전 자동 보다는 클릭 유도)
        self.root.after(500, lambda: self.status_label.config(text=f"총 {len(self.target_files)}개 파일 준비됨. 시작을 눌러주세요."))

    def browse_file(self):
        files = filedialog.askopenfilenames(title='엑셀 파일 선택', filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')])
        if files:
            self.target_files = list(files)
            self._update_file_label()
            self.status_label.config(text="작업시작 버튼을 클릭해주세요.")

    def on_start_click(self):
        if not self.target_files:
            messagebox.showwarning("경고", "파일을 선택해주세요.")
            return
        
        self.is_running = True
        self.stop_event.clear()
        self.start_btn.config(state="disabled")
        self.browse_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        # 코어 로직 스레드 시작
        thread = threading.Thread(target=self.run_core_logic, daemon=True)
        thread.start()

    def on_stop_click(self):
        if messagebox.askyesno("확인", "정말 작업을 중단하시겠습니까?"):
            self.stop_event.set()
            self.status_label.config(text="중지 요청 중...")
            self.stop_btn.config(state="disabled")

    def run_core_logic(self):
        # 파일 수집 (폴더 포함)
        files_to_process = []
        for path in self.target_files:
            if os.path.isdir(path):
                for r, d, f in os.walk(path):
                    for file in f:
                        if file.lower().endswith(('.xls', '.xlsx')):
                            files_to_process.append(os.path.join(r, file))
            elif path.lower().endswith(('.xls', '.xlsx')):
                files_to_process.append(path)

        if not files_to_process:
            self.root.after(0, lambda: messagebox.showwarning("경고", "처리할 파일이 없습니다."))
            self._reset_ui()
            return

        core = ContactRemoverCore(self.stop_event, self.update_ui_from_thread)
        success_count, success_msgs, error_msgs = core.process_files(files_to_process)
        
        # 로그 저장 및 결과 보고 (메인 스레드에서 실행)
        self.root.after(0, self.finalize_job, success_count, success_msgs, error_msgs, files_to_process)

    def update_ui_from_thread(self, current, total, filename, text):
        def update():
            self.progress_bar["maximum"] = total
            self.progress_bar["value"] = current
            self.status_label.config(text=f"[{current}/{total}] {filename} {text}")
        self.root.after(0, update)

    def finalize_job(self, success_count, success_msgs, error_msgs, processed_list):
        # 로그 파일 저장
        if processed_list and (success_msgs or error_msgs):
            log_dir = os.path.dirname(processed_list[0])
            if not log_dir: log_dir = os.getcwd()
            
            if success_msgs:
                p = os.path.join(log_dir, "작업성공_로그.txt")
                with open(p, "w", encoding="utf-8") as f:
                    f.write(f"--- 작업 성공 (총 {len(success_msgs)}건) ---\n" + "\n".join(success_msgs))
                os.startfile(p)
            
            if error_msgs:
                p = os.path.join(log_dir, "작업실패_로그.txt")
                with open(p, "w", encoding="utf-8") as f:
                    f.write(f"--- 작업 실패 (총 {len(error_msgs)}건) ---\n" + "\n".join(error_msgs))
                os.startfile(p)

        # 알림창
        if self.stop_event.is_set():
            messagebox.showwarning("중단됨", f"사용자에 의해 작업이 중단되었습니다.\n(성공: {success_count}건 / 에러: {len(error_msgs)}건)")
        else:
            if error_msgs:
                messagebox.showerror("완료", f"작업 완료: {success_count}건 패스\n에러: {len(error_msgs)}건 (로그 참조)")
            else:
                messagebox.showinfo("완료", f"총 {success_count}개 파일 처리가 완벽하게 끝났습니다.")

        self._reset_ui()

    def _reset_ui(self):
        self.is_running = False
        self.target_files = []
        self._update_file_label()
        self.progress_bar["value"] = 0
        self.status_label.config(text="대기 중...")
        self.start_btn.config(state="normal")
        self.browse_btn.config(state="normal")
        self.stop_btn.config(state="disabled")

    # IPC 리스너 (기존과 동일)
    def start_ipc_listen(self):
        def listen():
            try:
                s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                s.bind(('127.0.0.1', 48194))
                s.listen(20)
                while True:
                    conn, addr = s.accept()
                    data = conn.recv(4096)
                    if data:
                        new_files = json.loads(data.decode('utf-8'))
                        self.root.after(0, self._append_files, new_files)
                    conn.close()
            except: pass
        threading.Thread(target=listen, daemon=True).start()

    def _append_files(self, files):
        for f in files:
            if f not in self.target_files: self.target_files.append(f)
        self._update_file_label()
        self.status_label.config(text=f"새 파일 {len(files)}개가 추가되었습니다.")

def try_send_to_master(args):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(0.3)
        s.connect(('127.0.0.1', 48194))
        s.sendall(json.dumps(args).encode('utf-8'))
        s.close()
        return True
    except: return False

def main():
    target_files = []
    auto_run = False
    if len(sys.argv) > 1:
        target_files = sys.argv[1:]; auto_run = True
        if try_send_to_master(target_files): sys.exit(0)

    root = tk.Tk()
    app = ContactRemoverApp(root, target_files, auto_run)
    if not auto_run:
        root.lift()
        root.attributes('-topmost', True)
        root.after_idle(root.attributes, '-topmost', False)
    root.mainloop()

if __name__ == "__main__":
    main()
