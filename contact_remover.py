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
from win32_setctime import setctime

# 연락처 컬럼을 삭제하는 파이썬 스크립트

class ContactRemoverApp:
    def __init__(self, root, target_files=None, auto_run=False):
        self.root = root
        self.root.title("엑셀 연락처 삭제 도구")
        self.root.geometry("400x250")
        
        self.target_files = target_files if target_files else []
        self.target_file_var = tk.StringVar()
        if len(self.target_files) == 1:
            self.target_file_var.set(self.target_files[0])
        elif len(self.target_files) > 1:
            self.target_file_var.set(f"{len(self.target_files)}개 파일 선택됨")
        
        tk.Label(root, text="작업할 엑셀 파일:").pack(pady=(10, 0))
        
        entry_frame = tk.Frame(root)
        entry_frame.pack(fill="x", padx=10, pady=5)
        
        self.file_entry = tk.Entry(entry_frame, textvariable=self.target_file_var, state='readonly', width=45)
        self.file_entry.pack(side="left", expand=True)
        
        self.browse_btn = tk.Button(entry_frame, text="파일 선택", command=self.browse_file)
        self.browse_btn.pack(side="right", padx=(5, 0))
        
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=10)
        
        self.start_btn = tk.Button(btn_frame, text="작업시작", command=self.start_job, width=15, height=2, bg="#4CAF50", fg="white", font=("맑은 고딕", 10, "bold"))
        self.start_btn.pack(side="left", padx=10)
        
        self.close_btn = tk.Button(btn_frame, text="닫기", command=self.root.destroy, width=15, height=2)
        self.close_btn.pack(side="left", padx=10)

        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=350, mode="determinate")
        self.progress_bar.pack(pady=5)
        
        self.status_label = tk.Label(root, text="대기 중...", fg="blue")
        self.status_label.pack()
        
        self.auto_run = auto_run
        if auto_run:
            self.start_btn.config(state="disabled")
            self.browse_btn.config(state="disabled")
            self.status_label.config(text="다른 파일들의 입력을 대기 중...")
            self.ipc_timer = None
            self.start_ipc_listen()
            self.schedule_start()
            
    def start_ipc_listen(self):
        try:
            self.server_socket = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.server_socket.bind(('127.0.0.1', 48194))
            self.server_socket.listen(20)
        except:
            return

        def listen_loop():
            while True:
                try:
                    conn, addr = self.server_socket.accept()
                    data = conn.recv(4096)
                    if data:
                        args = json.loads(data.decode('utf-8'))
                        self.root.after(0, self.add_ipc_files, args)
                    conn.close()
                except:
                    break
        
        t = threading.Thread(target=listen_loop, daemon=True)
        t.start()

    def add_ipc_files(self, args):
        for arg in args:
            if arg not in self.target_files:
                self.target_files.append(arg)
        
        if len(self.target_files) == 1:
            self.target_file_var.set(self.target_files[0])
        else:
            self.target_file_var.set(f"{len(self.target_files)}개 파일 선택됨")
            
        self.schedule_start()

    def schedule_start(self):
        if hasattr(self, 'ipc_timer') and self.ipc_timer:
            self.root.after_cancel(self.ipc_timer)
        self.ipc_timer = self.root.after(500, self.finish_ipc_wait)

    def finish_ipc_wait(self):
        self.start_btn.config(state="normal")
        self.browse_btn.config(state="normal")
        count = len(self.target_files)
        if count > 0:
            self.status_label.config(text=f"총 {count}개의 명단이 병합 대기 중입니다. 작업시작을 클릭해주세요.")
        else:
            self.status_label.config(text="대기 중...")

    def browse_file(self):
        filetypes = (
            ('Excel files', '*.xlsx *.xls'),
            ('All files', '*.*')
        )
        filenames = filedialog.askopenfilenames(
            title='엑셀 파일 선택 (다중 선택 가능)',
            initialdir='/',
            filetypes=filetypes
        )
        
        if filenames:
            self.target_files = list(filenames)
            if len(self.target_files) == 1:
                self.target_file_var.set(self.target_files[0])
            else:
                self.target_file_var.set(f"{len(self.target_files)}개 파일 선택됨")
            self.status_label.config(text="작업시작 버튼을 클릭해주세요.")

    def check_encryption(self, file_path):
        try:
            with open(file_path, "rb") as f:
                office_file = msoffcrypto.OfficeFile(f)
                if office_file.is_encrypted():
                    return True
            return False
        except Exception:
            return False

    def start_job(self):
        if self.target_files is None or len(self.target_files) == 0:
            messagebox.showwarning("경고", "파일을 선택해주세요.")
            return

        if hasattr(self, 'server_socket'):
            try:
                self.server_socket.close()
            except:
                pass

        raw_files = self.target_files
        files_to_process = []
        for path in raw_files:
            if not path or not os.path.exists(path):
                continue
            if os.path.isdir(path):
                for root_dir, dirs, files in os.walk(path):
                    for file in files:
                        if file.lower().endswith(('.xls', '.xlsx')):
                            files_to_process.append(os.path.join(root_dir, file))
            else:
                if path.lower().endswith(('.xls', '.xlsx')):
                    files_to_process.append(path)

        if not files_to_process:
            messagebox.showwarning("경고", "처리할 엑셀 파일이 없습니다.")
            return

        success_count = 0
        error_messages = []
        success_messages = []
        target_keywords = ['연락처', '전화번호', '휴대폰', '핸드폰', '전화']

        total_files = len(files_to_process)

        self.start_btn.config(state="disabled")
        self.browse_btn.config(state="disabled")
        self.progress_bar["maximum"] = total_files

        for i, file_path in enumerate(files_to_process):
            self.status_label.config(text=f"처리 중: {i+1} / {total_files} ({os.path.basename(file_path)})")
            self.progress_bar["value"] = i + 1
            
            self.root.update_idletasks()
            self.root.update()

            if not file_path or not os.path.exists(file_path):
                continue
                
            try:
                stat = os.stat(file_path)
                c_time = stat.st_ctime
                m_time = stat.st_mtime
                a_time = stat.st_atime
            except Exception:
                c_time, m_time, a_time = None, None, None

            if self.check_encryption(file_path):
                error_messages.append(f"[{os.path.basename(file_path)}] 암호화된 파일입니다.")
                continue
                
            try:
                dfs = {}
                try:
                    dfs = pd.read_excel(file_path, engine='calamine', sheet_name=None)
                except Exception as e:
                    error_str = str(e).lower()
                    if 'encrypted' in error_str or 'password' in error_str or 'workbook is encrypted' in error_str:
                        error_messages.append(f"[{os.path.basename(file_path)}] 암호화된 파일입니다.")
                        continue
                    elif 'engine' in error_str or 'format' in error_str or 'calamine' in error_str:
                        try:
                            from_html = pd.read_html(file_path, encoding='cp949')
                            dfs = {"Sheet1": from_html[0]}
                        except:
                            try:
                                from_html = pd.read_html(file_path, encoding='utf-8')
                                dfs = {"Sheet1": from_html[0]}
                            except:
                                try:
                                    dfs = {"Sheet1": pd.read_csv(file_path, encoding='cp949', sep=None, engine='python')}
                                except:
                                    try:
                                        dfs = {"Sheet1": pd.read_csv(file_path, encoding='utf-8', sep=None, engine='python')}
                                    except Exception as e2:
                                        error_messages.append(f"[{os.path.basename(file_path)}] 알 수 없는 형식입니다.")
                                        continue
                    else:
                        error_messages.append(f"[{os.path.basename(file_path)}] 형식 오류({e}).")
                        continue
                
                processed_dfs = {}
                any_contact_dropped = False

                for sheet_name, df in dfs.items():
                    if df is None or df.empty:
                        processed_dfs[sheet_name] = df
                        continue
                        
                    if isinstance(df.columns, pd.MultiIndex):
                        df.columns = [str(col[-1]).strip() if isinstance(col, tuple) else str(col).strip() for col in df.columns]
                    else:
                        df.columns = [str(col).strip() for col in df.columns]

                    if not any(kw in str(col).replace(' ', '') for col in df.columns for kw in target_keywords):
                        for idx, row in df.head(15).iterrows():
                            row_vals = [str(x).strip() for x in row.values]
                            if any(kw in val.replace(' ', '') for val in row_vals for kw in target_keywords):
                                df.columns = row_vals
                                df = df.iloc[idx+1:].reset_index(drop=True)
                                break

                    found_cols_to_drop = [col for col in df.columns if any(kw in str(col).replace(' ', '') for kw in target_keywords)]
                    
                    if found_cols_to_drop:
                        df = df.drop(columns=found_cols_to_drop)
                        any_contact_dropped = True

                    processed_dfs[sheet_name] = df

                if any_contact_dropped:
                    base_name, ext = os.path.splitext(file_path)
                    ext = ext.lower()
                    
                    if ext == ".xlsx":
                        new_file_path = f"{base_name}1.xlsx"
                    else:
                        new_file_path = f"{base_name}.xlsx"
                    
                    with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
                        for sheet_name, df in processed_dfs.items():
                            if df is not None:
                                df.to_excel(writer, sheet_name=sheet_name, index=False)
                    
                    if c_time is not None:
                        try:
                            os.utime(new_file_path, (a_time, m_time))
                            setctime(new_file_path, c_time)
                        except Exception:
                            pass
                    
                    try:
                        os.remove(file_path)
                    except Exception as e:
                        error_messages.append(f"[{os.path.basename(file_path)}] 원본 삭제 실패 확인 요망: {e}")
                    
                    success_count += 1
                    success_messages.append(f"[{os.path.basename(file_path)}] 완료")
                else:
                    error_messages.append(f"[{os.path.basename(file_path)}] '연락처' 연관 열 없음.")
                    
            except Exception as e:
                error_messages.append(f"[{os.path.basename(file_path)}] 알 수 없는 오류({e}).")

        if len(files_to_process) > 0 and (error_messages or success_messages):
            first_file_dir = os.path.dirname(files_to_process[0])
            if not first_file_dir: first_file_dir = os.getcwd()
            
            if error_messages:
                err_log_path = os.path.join(first_file_dir, "작업실패_로그.txt")
                try:
                    with open(err_log_path, "w", encoding="utf-8") as f:
                        f.write(f"--- 작업 실패 내역 (총 {len(error_messages)}건) ---\n\n")
                        f.write("\n".join(error_messages))
                    os.startfile(err_log_path)
                except:
                    pass
            
            if success_messages:
                succ_log_path = os.path.join(first_file_dir, "작업성공_로그.txt")
                try:
                    with open(succ_log_path, "w", encoding="utf-8") as f:
                        f.write(f"--- 작업 성공 내역 (총 {len(success_messages)}건) ---\n\n")
                        f.write("\n".join(success_messages))
                    os.startfile(succ_log_path)
                except:
                    pass

        if error_messages:
            msg = f"작업 완료: {success_count}건 성공\n오류: {len(error_messages)}건 (상세 내역은 '작업실패_로그.txt' 참조)\n\n"
            msg += "\n".join(error_messages[:10])
            if len(error_messages) > 10:
                msg += f"\n...외 {len(error_messages) - 10}개 오류 발생"
            messagebox.showerror("작업 완료 알림", msg)
        else:
            messagebox.showinfo("작업 완료", f"총 {success_count}개의 파일 처리가 완료되었습니다.\n(성공 내역은 '작업성공_로그.txt' 참조)")

        self.target_files = []
        self.target_file_var.set("")
        self.progress_bar["value"] = 0
        self.status_label.config(text="작업이 완료되었습니다. 닫기 버튼을 누르거나 추가 파일을 선택하세요.")
        self.start_btn.config(state="normal")
        self.browse_btn.config(state="normal")

def try_send_to_master(args):
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(0.5)
        s.connect(('127.0.0.1', 48194))
        s.sendall(json.dumps(args).encode('utf-8'))
        s.close()
        return True
    except:
        return False

def main():
    target_files = []
    auto_run = False
    
    if len(sys.argv) > 1:
        target_files = sys.argv[1:]
        auto_run = True
        
        if try_send_to_master(target_files):
            sys.exit(0)

    root = tk.Tk()
    app = ContactRemoverApp(root, target_files, auto_run)
    
    if not auto_run:
        root.lift()
        root.attributes('-topmost', True)
        root.after_idle(root.attributes, '-topmost', False)
    
    root.mainloop()

if __name__ == "__main__":
    main()
