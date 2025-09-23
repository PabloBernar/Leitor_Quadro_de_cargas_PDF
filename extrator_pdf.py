# Bibliotecas necessárias:
# pip install customtkinter
# pip install pdfplumber
# pip install pandas

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pdfplumber, re, pandas as pd
import datetime
import os
import webbrowser
from threading import Thread

# Configuração global do CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# --------------------------- 
# Funções auxiliares do parser (mantidas do código original)
# --------------------------- 
def find_header_columns(page):
    """Encontra e define os limites das colunas do cabeçalho na página."""
    words = page.extract_words(use_text_flow=True)
    lines = {}
    for w in words:
        y = round(w["top"], 1)
        lines.setdefault(y, []).append(w)
    header_tokens = None
    horas_token = None
    for y, items in sorted(lines.items()):
        for it in items:
            if it["text"].strip().lower() == "horas":
                horas_token = it
        text = " ".join(it["text"] for it in sorted(items, key=lambda k: k["x0"]))
        if "Tipo de aparelho" in text:
            header_tokens = sorted(items, key=lambda k: k["x0"])
            break
    if not header_tokens:
        return None
    cols = {}
    def find_index(tokens, text):
        for idx,t in enumerate(tokens):
            if t["text"]==text:
                return idx
        return None
    tokens = header_tokens
    idx_tipo = find_index(tokens, "Tipo")
    if idx_tipo is not None:
        for j in range(idx_tipo, min(idx_tipo+6, len(tokens))):
            if tokens[j]["text"].lower().startswith("aparelho"):
                cols["tipo"] = (tokens[idx_tipo]["x0"], tokens[j]["x1"]); break
    idx_sub = find_index(tokens, "Subtipo")
    if idx_sub is not None:
        for j in range(idx_sub, min(idx_sub+6, len(tokens))):
            if tokens[j]["text"].lower().startswith("aparelho"):
                cols["subtipo"] = (tokens[idx_sub]["x0"], tokens[j]["x1"]); break
    idx_inicio = find_index(tokens, "Início")
    if idx_inicio is not None:
        for j in range(idx_inicio-1, max(idx_inicio-3,-1), -1):
            if tokens[j]["text"]=="Data":
                cols["data_ini"] = (tokens[j]["x0"], tokens[idx_inicio]["x1"]); break
    idx_fim = find_index(tokens, "Fim")
    if idx_fim is not None:
        for j in range(idx_fim-1, max(idx_fim-3,-1), -1):
            if tokens[j]["text"]=="Data":
                cols["data_fim"] = (tokens[j]["x0"], tokens[idx_fim]["x1"]); break
    idx_dias = find_index(tokens, "Dias"); idx_oper = find_index(tokens, "Oper")
    if idx_dias is not None: cols["dias"] = (tokens[idx_dias]["x0"], tokens[idx_dias]["x1"])
    if idx_oper is not None: cols["oper"] = (tokens[idx_oper]["x0"], tokens[idx_oper]["x1"])
    for idx,t in enumerate(tokens):
        if "Qtd.Aparelhos" in t["text"] or "Qtd.Apare" in t["text"]:
            cols["qtd_aparelhos"] = (t["x0"], t["x1"]); break
    for idx,t in enumerate(tokens):
        if "Qtd.Potência" in t["text"] or "Potência" in t["text"]:
            cols["qtd_potencia"] = (t["x0"], t["x1"]); break
    for idx,t in enumerate(tokens):
        if t["text"].upper()=="DIC":
            cols["dic"] = (t["x0"], t["x1"]); break
    idx_qtd = None
    for idx,t in enumerate(tokens):
        if t["text"] in ("Qtd.","Qtd"):
            idx_qtd = idx; break
    if idx_qtd is not None:
        for j in range(idx_qtd+1, min(idx_qtd+4, len(tokens))):
            if tokens[j]["text"] in ("Fat.","Fat"):
                cols["qtd_fat"] = (tokens[idx_qtd]["x0"], tokens[j]["x1"]); break
    if horas_token:
        cols["horas"] = (horas_token["x0"], horas_token["x1"])
    col_list = []
    for name, span in cols.items():
        col_list.append((name, span[0], span[1], (span[0]+span[1])/2.0))
    col_list_sorted = sorted(col_list, key=lambda x: x[3])
    boundaries = []
    for i,col in enumerate(col_list_sorted):
        left = -1e6 if i==0 else (col_list_sorted[i-1][2] + col[1]) / 2.0
        right = 1e6 if i==len(col_list_sorted)-1 else (col[2] + col_list_sorted[i+1][1]) / 2.0
        boundaries.append((col[0], left, right, col[3]))
    return boundaries

def assign_tokens_using_boundaries(tokens, boundaries):
    """Atribui tokens às colunas com base nos limites."""
    assigned = {name: [] for name,_,_,_ in boundaries}
    for t in tokens:
        center = (t["x0"] + t["x1"]) / 2.0
        found=False
        for name,left,right,_ in boundaries:
            if center >= left and center < right:
                assigned[name].append(t["text"]); found=True; break
        if not found and boundaries:
            nearest = min(boundaries, key=lambda it: abs(center - it[3]))
            assigned[nearest[0]].append(t["text"])
    return {k: " ".join(v) for k,v in assigned.items()}

def valid_tipo(t):
    """Verifica se o tipo é válido."""
    if not t: return False
    if not re.search(r"[A-Za-zÀ-ÿ]", t): return False
    return True

def parse_pdf_to_csv(pdf_path, csv_out, log_func):
    """Função principal que extrai dados do PDF e salva em CSV."""
    records = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            last_boundaries = None
            log_func(f"Total de páginas: {len(pdf.pages)}", "info")
            for page in pdf.pages:
                log_func(f"Processando página {page.page_number}...", "info")
                boundaries = find_header_columns(page)
                if not boundaries and last_boundaries:
                    boundaries = last_boundaries
                last_boundaries = boundaries
                words = page.extract_words(use_text_flow=True)
                grouped = {}
                for w in words:
                    y = round(w["top"],1)
                    grouped.setdefault(y, []).append(w)
                sorted_lines = sorted(grouped.items(), key=lambda x: x[0])
                line_infos = [{"y": y, "text": " ".join(it["text"] for it in sorted(items, key=lambda k:k["x0"])), "tokens": sorted(items, key=lambda k:k["x0"])} for y,items in sorted_lines]
                i = 0
                while i < len(line_infos):
                    line = line_infos[i]
                    dates = re.findall(r"\d{2}/\d{2}/\d{4}", line["text"])
                    if len(dates) >= 2:
                        combined_tokens = list(line["tokens"])
                        j = i+1
                        while j < len(line_infos) and (j - i) <= 6:
                            if len(re.findall(r"\d{2}/\d{2}/\d{4}", line_infos[j]["text"])) >= 2: break
                            if "Tipo de aparelho" in line_infos[j]["text"]:
                                break
                            combined_tokens.extend(line_infos[j]["tokens"]); j += 1

                        tokens = []
                        if boundaries:
                            for it in combined_tokens:
                                center = (it["x0"] + it["x1"]) / 2.0
                                if any(left <= center < right for _,left,right,_ in boundaries):
                                    tokens.append({"text": it["text"], "x0": it["x0"], "x1": it["x1"]})
                        else:
                            tokens = [{"text": it["text"], "x0": it["x0"], "x1": it["x1"]} for it in combined_tokens]

                        tokens = [t for t in tokens if not re.search(r"(Usuário:|Emitido pela|Página:|Posição do Dia:)", t["text"], re.IGNORECASE)]

                        assigned = assign_tokens_using_boundaries(tokens, boundaries) if boundaries else {}
                        assigned["data_ini"] = dates[0]; assigned["data_fim"] = dates[1]
                        full_text = " ".join(t["text"] for t in tokens)
                        numbers = re.findall(r"[\d]+[.,\d]*", full_text)
                        if numbers and (not assigned.get("qtd_potencia")):
                            assigned["qtd_potencia"] = numbers[-1]
                        if numbers and (not assigned.get("dic")) and len(numbers)>=3:
                            assigned["dic"] = numbers[-2]
                        if not assigned.get("dias"):
                            try:
                                from datetime import datetime
                                d1 = datetime.strptime(assigned["data_ini"], "%d/%m/%Y")
                                d2 = datetime.strptime(assigned["data_fim"], "%d/%m/%Y")
                                assigned["dias"] = str((d2-d1).days + 1)
                            except:
                                assigned["dias"] = ""
                        for k in ["tipo","subtipo","data_ini","data_fim","dias","oper","qtd_aparelhos","horas","qtd_potencia","dic","qtd_fat"]:
                            assigned.setdefault(k,""); assigned[k] = assigned[k].strip()

                        if assigned.get("dias"):
                            m = re.search(r"\b\d+\b", assigned["dias"])
                            if m:
                                assigned["dias"] = m.group(0)

                        if assigned.get("qtd_fat"):
                            m = re.search(r"\d+[.,]?\d*", assigned["qtd_fat"])
                            if m:
                                assigned["qtd_fat"] = m.group(0)

                        if not valid_tipo(assigned.get("tipo","")):
                            i = j; continue

                        tipos = re.split(r"(?=(Lampada|Reator))", assigned.get("tipo",""), flags=re.IGNORECASE)
                        tipos = [" ".join(tipos[i:i+2]).strip() for i in range(1, len(tipos), 2)] if len(tipos)>2 else [assigned.get("tipo","")]
                        if len(tipos) > 1:
                            for t in tipos:
                                new_assigned = assigned.copy()
                                new_assigned["tipo"] = t
                                new_assigned["page"] = str(page.page_number)
                                records.append(new_assigned)
                        else:
                            assigned["page"] = str(page.page_number)
                            records.append(assigned)
                        i = j
                    else:
                        i += 1
        
        df = pd.DataFrame(records, columns=["page","tipo","subtipo","data_ini","data_fim","dias","oper","qtd_aparelhos","horas","qtd_potencia","dic","qtd_fat"])
        df.to_csv(csv_out, sep=";", index=False, encoding="utf-8-sig")
        return True
    except Exception as e:
        log_func(f"Ocorreu uma falha: {e}", "error")
        return False


# --------------------------- 
# Aplicação principal com design moderno e multi-arquivo
# --------------------------- 
class LeitorCargasApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # --- Configurações da Janela ---
        self.title("Leitor de Cargas de PDF")
        self.geometry("800x600")
        self.grid_columnconfigure(0, weight=1)
        # Ajusta para a linha do logbox expandir
        self.grid_rowconfigure(3, weight=1)
        
        self.file_paths = []

        # --- Título ---
        self.title_label = ctk.CTkLabel(self, text="Leitor de Cargas", font=ctk.CTkFont(size=28, weight="bold"))
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # --- Frame de Seleção de Arquivo ---
        self.file_frame = ctk.CTkFrame(self)
        self.file_frame.grid(row=1, column=0, padx=20, pady=20, sticky="ew")
        self.file_frame.grid_columnconfigure(0, weight=1)

        self.file_count_label = ctk.CTkLabel(self.file_frame, text="Nenhum arquivo selecionado.", text_color="#A0A0A0")
        self.file_count_label.grid(row=0, column=0, padx=(20, 10), pady=10, sticky="ew")

        self.select_button = ctk.CTkButton(self.file_frame, text="Selecionar PDF(s)", command=self.selecionar_pdfs, font=ctk.CTkFont(weight="bold"))
        self.select_button.grid(row=0, column=1, padx=20, pady=10)
        
        # --- Frame de Ação e Status ---
        self.action_frame = ctk.CTkFrame(self)
        self.action_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="ew")
        self.action_frame.grid_columnconfigure(0, weight=1)

        self.extract_button = ctk.CTkButton(self.action_frame, text="Extrair para CSV", command=self.extrair_csv, font=ctk.CTkFont(weight="bold"), state="disabled")
        self.extract_button.grid(row=0, column=0, padx=20, pady=(10, 10), sticky="ew")

        # --- Log Box ---
        self.log_textbox = ctk.CTkTextbox(self, state="disabled", fg_color="#1C2128", text_color="#C5D1DE")
        self.log_textbox.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="nsew")

        # --- Footer com o link para o LinkedIn ---
        self.footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.footer_frame.grid(row=4, column=0, padx=20, pady=(0, 10))
        
        self.footer_label = ctk.CTkLabel(self.footer_frame, text="by @pablo.bernar", text_color="#64B5F6", font=ctk.CTkFont(size=12, underline=True))
        self.footer_label.pack()
        self.footer_label.bind("<Enter>", lambda e: self.footer_label.configure(text_color="#47a1ff"))
        self.footer_label.bind("<Leave>", lambda e: self.footer_label.configure(text_color="#64B5F6"))
        self.footer_label.bind("<Button-1>", lambda e: webbrowser.open_new("https://www.linkedin.com/in/pablo-bernar/"))

        # Mensagem inicial
        self.log_message("Bem-vindo ao Leitor de Cargas de PDF. Selecione um ou mais arquivos para começar.", "info")

    def log_message(self, message, msg_type="info"):
        """Adiciona uma mensagem ao log com cor baseada no tipo."""
        # Schedule the actual GUI update to run on the main thread
        self.after(0, self._update_log_gui, message, msg_type)

    def _update_log_gui(self, message, msg_type):
        """Função interna para atualizar o logbox na thread principal."""
        self.log_textbox.configure(state="normal")
        timestamp = datetime.datetime.now().strftime("[%H:%M:%S]")
        
        color = "#C5D1DE"
        if msg_type == "error":
            color = "#DC3545"
        elif msg_type == "success":
            color = "#28A745"
        
        self.log_textbox.insert("end", f"{timestamp} {message}\n")
        
        self.log_textbox.configure(state="disabled")
        self.log_textbox.see("end") # Auto-scroll

    def selecionar_pdfs(self):
        """Abre a caixa de diálogo para seleção de múltiplos arquivos PDF."""
        caminhos = filedialog.askopenfilenames(filetypes=[("Arquivos PDF", "*.pdf")])
        if caminhos:
            self.file_paths = caminhos
            self.file_count_label.configure(text=f"{len(self.file_paths)} arquivo(s) selecionado(s).")
            self.extract_button.configure(state="normal")
            self.log_message(f"{len(self.file_paths)} arquivo(s) selecionado(s).", "info")
        else:
            self.file_paths = []
            self.file_count_label.configure(text="Nenhum arquivo selecionado.")
            self.extract_button.configure(state="disabled")
            self.log_message("Seleção de arquivo(s) cancelada.", "info")

    def extrair_csv(self):
        """Inicia a extração dos PDFs em uma thread separada."""
        if not self.file_paths:
            self.log_message("Selecione um ou mais arquivos PDF primeiro!", "error")
            return
        
        output_dir = filedialog.askdirectory(title="Selecione a pasta para salvar os arquivos CSV")
        if not output_dir:
            self.log_message("Extração cancelada. Nenhuma pasta de destino selecionada.", "info")
            return
            
        self.log_message(f"Iniciando extração de {len(self.file_paths)} arquivo(s). Os arquivos serão salvos em: {output_dir}", "info")
        self.extract_button.configure(state="disabled")
        self.select_button.configure(state="disabled")

        # Executa a extração em uma nova thread
        extraction_thread = Thread(target=self._run_extraction, args=(self.file_paths, output_dir))
        extraction_thread.start()

    def _run_extraction(self, pdf_paths, output_dir):
        """Função a ser executada na thread de extração."""
        total_success = 0
        total_failed = 0
        for pdf_path in pdf_paths:
            file_name = os.path.basename(pdf_path)
            base_name, _ = os.path.splitext(file_name)
            csv_out = os.path.join(output_dir, f"{base_name}.csv")
            self.log_message(f"Processando '{file_name}'...", "info")
            
            success = parse_pdf_to_csv(pdf_path, csv_out, self.log_message)
            
            if success:
                total_success += 1
                self.log_message(f"'{file_name}' extraído com sucesso para: {os.path.basename(csv_out)}", "success")
            else:
                total_failed += 1
                self.log_message(f"Falha ao processar '{file_name}'.", "error")
        
        self.extract_button.configure(state="normal")
        self.select_button.configure(state="normal")
        self.log_message("--- Processamento concluído. ---", "info")
        
        if total_failed == 0:
            messagebox.showinfo("Sucesso", f"Extração concluída!\n{total_success} arquivo(s) processado(s) com sucesso.")
        else:
            messagebox.showerror("Concluído com Erros", f"Extração concluída com {total_success} sucesso(s) e {total_failed} falha(s). Verifique o log para detalhes.")

if __name__ == "__main__":
    app = LeitorCargasApp()
    app.mainloop()
