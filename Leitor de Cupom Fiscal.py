import os
import re
import csv
import webbrowser
import pdfplumber
from typing import Optional, List, Dict
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, BooleanVar, Text, Scrollbar, Toplevel

class CupomReader:
    def __init__(self):
        self.agrupar_itens = True

    def set_agrupar_itens(self, valor: bool):
        self.agrupar_itens = valor

    def extract_text_with_layout(self, file_path: str) -> str:
        """Extrai texto do PDF mantendo estrutura"""
        try:
            full_text = ""
            with pdfplumber.open(file_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text(
                        layout=True,
                        x_tolerance=5,  # Aumentado para melhor captura
                        y_tolerance=3,
                        keep_blank_chars=False,
                    )
                    if text:
                        full_text += text + "\n"
            return full_text
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler PDF:\n{str(e)}")
            return ""

    def parse_items(self, text: str) -> List[Dict]:
        """Processa itens no formato exato da imagem"""
        items = []
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        i = 0
        while i < len(lines):
            line = lines[i]
            
            # Ignora cabeçalhos e linhas irrelevantes
            if any(x in line for x in ['ITEM', 'COD.', 'DESC.', 'TOTAL', 'Documento', 'Protocolo']):
                i += 1
                continue
                
            # Padrão para o formato específico da imagem
            item_match = re.match(
                r'^(\d+)\s+(\d{7,13})?\s*(.*?)\s+(\d+\.\d+|\d+\,\d+)\s+(\w+)\.?\s+(\d+\.\d+|\d+\,\d+)\s+(\d+\.\d+|\d+\,\d+)\s*$', 
                line
            )
            
            if item_match:
                try:
                    item_num = int(item_match.group(1))
                    codigo = item_match.group(2) if item_match.group(2) else ""
                    descricao = item_match.group(3).strip()
                    qtd = item_match.group(4).replace(',', '.')
                    un = item_match.group(5)
                    vl_unit = item_match.group(6).replace(',', '.')
                    vl_total = item_match.group(7).replace(',', '.')
                    
                    item = {
                        'item': item_num,
                        'codigo': codigo,
                        'descricao': descricao,
                        'quantidade': float(qtd),
                        'unidade': un,
                        'valor_unitario': float(vl_unit),
                        'valor_total': float(vl_total),
                        'desconto': 0.0
                    }
                    
                    # Verifica desconto na próxima linha
                    if i+1 < len(lines):
                        next_line = lines[i+1]
                        if f"Seq.: {item_num}" in next_line and "Desconto" in next_line:
                            desconto_match = re.search(r'Desconto\s+([\d,\.]+)', next_line)
                            if desconto_match:
                                item['desconto'] = float(desconto_match.group(1).replace(',', '.'))
                                i += 1  # Pula a linha do desconto
                    
                    items.append(item)
                except Exception as e:
                    print(f"Erro processando linha {i+1}: '{line}'\nErro: {str(e)}")
            
            i += 1
        
        return items

    def _agrupar_itens_repetidos(self, items: List[Dict]) -> List[Dict]:
        """Agrupa itens idênticos"""
        grouped = {}
        for item in items:
            key = (item['codigo'], item['descricao'], item['valor_unitario'])
            if key in grouped:
                grouped[key]['quantidade'] += item['quantidade']
                grouped[key]['valor_total'] += item['valor_total']
                grouped[key]['desconto'] += item['desconto']
            else:
                grouped[key] = item.copy()
        return sorted(grouped.values(), key=lambda x: x['item'])

    def process_cupom(self, file_path: str) -> Optional[Dict]:
        """Processa o cupom fiscal completo"""
        try:
            text = self.extract_text_with_layout(file_path)
            if not text:
                return None
                
            items = self.parse_items(text)
            if not items:
                print("DEBUG - Texto extraído:\n", text[:1000])
                messagebox.showwarning("Aviso", "Nenhum item encontrado. Verifique o console.")
                return None
                
            if self.agrupar_itens:
                items = self._agrupar_itens_repetidos(items)
                
            return {
                'total_itens': len(items),
                'total_geral': round(sum(i['valor_total'] for i in items), 2),
                'total_descontos': round(sum(i['desconto'] for i in items), 2),
                'itens': items
            }
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar:\n{str(e)}")
            return None
class CupomReaderGUI:
    def __init__(self, root):
        self.root = root
        self.reader = CupomReader()
        self.results = None
        
        # Configura o ícone
        self.set_window_icon()
        
        self.setup_ui()
        self.setup_devolucao_ui()

    def set_window_icon(self):
        """Configura o ícone do aplicativo de forma robusta"""
        try:
            # Verifica se estamos executando como script ou executável
            if getattr(sys, 'frozen', False):
                # Modo executável - usa o diretório do executável
                base_path = os.path.dirname(sys.executable)
            else:
                # Modo desenvolvimento - usa o diretório do script
                base_path = os.path.dirname(os.path.abspath(__file__))
            
            # Caminho completo para o ícone
            icon_path = os.path.join(base_path, 'icone.ico')
            
            # Verifica se o arquivo existe
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
            else:
                print(f"Aviso: Arquivo de ícone não encontrado em {icon_path}")
                
        except Exception as e:
            print(f"Erro ao carregar ícone: {str(e)}")
class CupomReaderGUI:
    def __init__(self, root):
        self.root = root
        self.reader = CupomReader()
        self.results = None
        
        self.setup_ui()
        self.setup_devolucao_ui()

    def setup_ui(self):
        """Configura a interface principal"""
        self.root.title("Processador de Cupons Muffato v7.0")
        self.root.geometry("1200x600")
        
        # Frame superior
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10, fill=tk.X)
        
        tk.Label(top_frame, text="Arquivo PDF:").pack(side=tk.LEFT)
        self.file_entry = tk.Entry(top_frame, width=80)
        self.file_entry.pack(side=tk.LEFT, padx=5, expand=True)
        
        tk.Button(top_frame, text="Procurar", command=self.browse_file).pack(side=tk.LEFT)
        tk.Button(top_frame, text="Processar", command=self.process_file).pack(side=tk.LEFT, padx=10)
        
        self.agrupar_var = BooleanVar(value=True)
        tk.Checkbutton(
            top_frame, 
            text="Agrupar itens iguais", 
            variable=self.agrupar_var,
            command=lambda: self.reader.set_agrupar_itens(self.agrupar_var.get())
        ).pack(side=tk.LEFT, padx=10)
        
        tk.Button(
            top_frame,
            text="Registrar Devolução",
            command=self.show_devolucao_window,
            bg="#4CAF50",
            fg="white"
        ).pack(side=tk.LEFT, padx=10)

        # Treeview para exibir os itens
        self.tree = ttk.Treeview(
            self.root,
            columns=('Item', 'Código', 'Descrição', 'Qtd', 'Un', 'V.Unit', 'Desconto', 'V.Total'),
            show='headings'
        )
        
        # Configuração das colunas
        for col, width in [
            ('Item', 50), ('Código', 120), ('Descrição', 300),
            ('Qtd', 60), ('Un', 40), ('V.Unit', 90), 
            ('Desconto', 90), ('V.Total', 90)
        ]:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=width, anchor='center' if width < 100 else 'w')
        
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Rodapé
        bottom_frame = tk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, pady=10)
        
        self.total_label = tk.Label(
            bottom_frame, 
            text="Total: R$ 0,00 | Itens: 0 | Descontos: R$ 0,00", 
            font=('Arial', 10, 'bold')
        )
        self.total_label.pack(side=tk.LEFT, padx=10, expand=True)
        
        tk.Button(
            bottom_frame, 
            text="Salvar CSV", 
            command=self.save_results,
            bg="#2196F3",
            fg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            bottom_frame, 
            text="Abrir PDF", 
            command=self.open_pdf,
            bg="#607D8B",
            fg="white"
        ).pack(side=tk.LEFT, padx=5)

    def setup_devolucao_ui(self):
        """Prepara a janela de devolução"""
        self.devolucao_window = None
        self.devolucao_text = None
        self.devolucao_result_tree = None

    def show_devolucao_window(self):
        """Abre a janela de registro de devolução"""
        if not self.results:
            messagebox.showwarning("Aviso", "Processe um cupom primeiro.")
            return
            
        if self.devolucao_window and self.devolucao_window.winfo_exists():
            self.devolucao_window.lift()
            return
            
        self.devolucao_window = Toplevel(self.root)
        self.devolucao_window.title("Registro de Devolução")
        self.devolucao_window.geometry("1000x700")
        
        # Frame de instruções
        instr_frame = tk.Frame(self.devolucao_window)
        instr_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(
            instr_frame,
            text="Digite os códigos dos itens para devolução (1 por linha):",
            font=('Arial', 10, 'bold')
        ).pack(anchor='w')
        
        # Área de texto para entrada
        text_frame = tk.Frame(self.devolucao_window)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.devolucao_text = Text(text_frame, wrap=tk.WORD, font=('Courier', 12), height=10)
        scrollbar = Scrollbar(text_frame, command=self.devolucao_text.yview)
        self.devolucao_text.config(yscrollcommand=scrollbar.set)
        
        self.devolucao_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Frame de botões
        button_frame = tk.Frame(self.devolucao_window)
        button_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Button(
            button_frame,
            text="Comparar Itens",
            command=self.compare_devolucao,
            bg="#4CAF50",
            fg="white",
            font=('Arial', 10, 'bold')
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="Limpar",
            command=self.clear_devolucao,
            bg="#F44336",
            fg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="Salvar Relatório",
            command=self.save_devolucao_report,
            bg="#2196F3",
            fg="white"
        ).pack(side=tk.LEFT, padx=5)
        
        # Treeview de resultados
        result_frame = tk.Frame(self.devolucao_window)
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.devolucao_result_tree = ttk.Treeview(
            result_frame,
            columns=('Item', 'Código', 'Descrição', 'Valor Unit', 'Valor Total', 'Status'),
            show='headings'
        )
        
        # Configuração das colunas
        for col, width in [
            ('Item', 50), ('Código', 120), ('Descrição', 250),
            ('Valor Unit', 100), ('Valor Total', 100), ('Status', 150)
        ]:
            self.devolucao_result_tree.heading(col, text=col)
            self.devolucao_result_tree.column(col, width=width, anchor='center')
        
        scrollbar_result = ttk.Scrollbar(result_frame, orient="vertical", command=self.devolucao_result_tree.yview)
        self.devolucao_result_tree.configure(yscrollcommand=scrollbar_result.set)
        
        self.devolucao_result_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_result.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Estilos para os resultados
        self.devolucao_result_tree.tag_configure('ok', foreground='green')
        self.devolucao_result_tree.tag_configure('erro', foreground='red')

    def compare_devolucao(self):
        """Compara itens devolvidos com o cupom"""
        if not hasattr(self, 'devolucao_result_tree'):
            return
            
        # Limpa resultados anteriores
        for item in self.devolucao_result_tree.get_children():
            self.devolucao_result_tree.delete(item)
            
        # Extrai códigos digitados
        devolucao_codigos = set()
        for line in self.devolucao_text.get('1.0', tk.END).split('\n'):
            codigo = line.strip()
            if codigo:
                devolucao_codigos.add(codigo)
        
        # Verifica cada código
        divergencias = 0
        cupom_itens = {str(item['codigo']): item for item in self.results['itens']}
        
        for codigo in devolucao_codigos:
            if codigo in cupom_itens:
                item = cupom_itens[codigo]
                self.devolucao_result_tree.insert('', 'end',
                    values=(
                        item['item'],
                        codigo,
                        item['descricao'],
                        f"R$ {item['valor_unitario']:.2f}",
                        f"R$ {item['valor_total']:.2f}",
                        "OK"
                    ),
                    tags=('ok',)
                )
            else:
                self.devolucao_result_tree.insert('', 'end',
                    values=(
                        "",
                        codigo,
                        "ITEM NÃO ENCONTRADO",
                        "",
                        "",
                        "DIVERGÊNCIA"
                    ),
                    tags=('erro',)
                )
                divergencias += 1
        
        # Mostra resumo
        messagebox.showinfo(
            "Resultado",
            f"Processamento concluído!\n\n"
            f"Itens verificados: {len(devolucao_codigos)}\n"
            f"Divergências encontradas: {divergencias}"
        )

    def clear_devolucao(self):
        """Limpa os campos de devolução"""
        if hasattr(self, 'devolucao_text'):
            self.devolucao_text.delete('1.0', tk.END)
        if hasattr(self, 'devolucao_result_tree'):
            for item in self.devolucao_result_tree.get_children():
                self.devolucao_result_tree.delete(item)

    def save_devolucao_report(self):
        """Salva o relatório de devolução em CSV"""
        if not hasattr(self, 'devolucao_result_tree'):
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivos CSV", "*.csv")],
            initialfile=f"devolucao_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        if not filename:
            return
            
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['Item', 'Código', 'Descrição', 'Valor Unitário', 'Valor Total', 'Status'])
                
                for item in self.devolucao_result_tree.get_children():
                    values = self.devolucao_result_tree.item(item, 'values')
                    writer.writerow(values)
                    
            messagebox.showinfo("Sucesso", f"Relatório salvo em:\n{filename}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{str(e)}")

    def browse_file(self):
        """Abre diálogo para selecionar arquivo PDF"""
        filename = filedialog.askopenfilename(
            title="Selecione o cupom fiscal",
            filetypes=[("Arquivos PDF", "*.pdf"), ("Todos os arquivos", "*.*")],
            initialdir=os.getcwd()
        )
        if filename:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, filename)

    def process_file(self):
        """Processa o cupom e exibe os resultados"""
        filename = self.file_entry.get()
        if not filename:
            messagebox.showerror("Erro", "Selecione um arquivo PDF.")
            return
            
        try:
            # Limpa resultados anteriores
            for i in self.tree.get_children():
                self.tree.delete(i)
                
            # Processa o arquivo
            self.results = self.reader.process_cupom(filename)
            
            if not self.results or not self.results.get('itens'):
                messagebox.showwarning(
                    "Aviso", 
                    "Nenhum item encontrado. Verifique:\n"
                    "1. Se o PDF contém texto selecionável\n"
                    "2. O formato do cupom\n"
                    "3. Consulte o console para detalhes"
                )
                return
                
            # Preenche a treeview
            for item in self.results['itens']:
                self.tree.insert('', 'end', values=(
                    item['item'],
                    item['codigo'],
                    item['descricao'],
                    f"{item['quantidade']:.3f}",
                    item['unidade'],
                    f"R$ {item['valor_unitario']:.2f}",
                    f"R$ {item['desconto']:.2f}" if item['desconto'] > 0 else "-",
                    f"R$ {item['valor_total']:.2f}"
                ))
            
            # Atualiza totais
            self.total_label.config(
                text=f"Total: R$ {self.results['total_geral']:.2f} | "
                     f"Itens: {self.results['total_itens']} | "
                     f"Descontos: R$ {self.results['total_descontos']:.2f}"
            )
            
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao processar:\n{str(e)}")

    def save_results(self):
        """Salva os resultados em CSV"""
        if not self.results:
            messagebox.showwarning("Aviso", "Nenhum resultado para salvar.")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("Arquivos CSV", "*.csv")],
            initialfile=f"cupom_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        )
        
        if not filename:
            return
            
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f, delimiter=';')
                writer.writerow(['Item', 'Código', 'Descrição', 'Qtd', 'Un', 'V.Unit', 'Desconto', 'V.Total'])
                
                for item in self.results['itens']:
                    writer.writerow([
                        item['item'],
                        item['codigo'],
                        item['descricao'],
                        item['quantidade'],
                        item['unidade'],
                        item['valor_unitario'],
                        item['desconto'],
                        item['valor_total']
                    ])
                
                writer.writerow([])
                writer.writerow(['Total itens:', '', self.results['total_itens'], '', '', '', '', ''])
                writer.writerow(['Total descontos:', '', '', '', '', '', '', f"R$ {self.results['total_descontos']:.2f}"])
                writer.writerow(['Total geral:', '', '', '', '', '', '', f"R$ {self.results['total_geral']:.2f}"])
            
            messagebox.showinfo("Sucesso", f"Resultados salvos em:\n{filename}")
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao salvar:\n{str(e)}")

    def open_pdf(self):
        """Abre o PDF no visualizador padrão"""
        if not self.file_entry.get():
            messagebox.showwarning("Aviso", "Nenhum arquivo selecionado.")
            return
            
        try:
            webbrowser.open(self.file_entry.get())
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir o PDF:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = CupomReaderGUI(root)
    root.mainloop()
