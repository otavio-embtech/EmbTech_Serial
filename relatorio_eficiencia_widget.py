import os
import re
import platform
from datetime import datetime
from collections import defaultdict, Counter

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit, QFileDialog,
    QMessageBox, QHBoxLayout, QDateEdit, QComboBox, QLineEdit
)
from PyQt6.QtCore import QDate

from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from matplotlib.ticker import MaxNLocator # Importa MaxNLocator

import pandas as pd
# xlsxwriter é necessário para formatação avançada no Excel
# Certifique-se de que está instalado: pip install XlsxWriter


class RelatorioEficienciaWidget(QWidget):
    def __init__(self, logs_dir=None, parent=None):
        super().__init__(parent)
        # Define o diretório de logs, padrão para a pasta de Documentos do usuário
        self.logs_dir = logs_dir or os.path.expanduser("~/Documents")
        # Lista para armazenar os resultados detalhados da análise de logs
        self.resultados_detalhados = []
        # Inicializa a interface do usuário
        self._init_ui()

    def _init_ui(self):
        # Layout principal da janela
        layout = QVBoxLayout(self)

        # Rótulo para os filtros de relatório
        layout.addWidget(QLabel("Filtros de Relatório:"))

        # Layout horizontal para os filtros de data e operador
        filtro_layout = QHBoxLayout()
        self.data_inicio = QDateEdit()
        self.data_fim = QDateEdit()
        hoje = QDate.currentDate()
        # Define a data de início como um mês atrás e a data de fim como hoje
        self.data_inicio.setDate(hoje.addMonths(-1))
        self.data_fim.setDate(hoje)
        filtro_layout.addWidget(QLabel("De:"))
        filtro_layout.addWidget(self.data_inicio)
        filtro_layout.addWidget(QLabel("Até:"))
        filtro_layout.addWidget(self.data_fim)

        self.operador_filtro = QComboBox()
        self.operador_filtro.addItem("Todos os Operadores")
        filtro_layout.addWidget(QLabel("Operador:"))
        filtro_layout.addWidget(self.operador_filtro)
        # Conecta a mudança de seleção do operador para filtrar e mostrar os dados
        self.operador_filtro.currentIndexChanged.connect(self._filtrar_e_mostrar)

        # Campo de filtro combinado para Lote e Série da Placa
        self.placa_filtro = QLineEdit()
        self.placa_filtro.setPlaceholderText("Filtrar por Lote (ex: 16578) ou Lote/Série (ex: 16781/32)")
        self.placa_filtro.textChanged.connect(self._filtrar_e_mostrar)
        filtro_layout.addWidget(QLabel("Lote/Série da Placa:"))
        filtro_layout.addWidget(self.placa_filtro)


        layout.addLayout(filtro_layout)

        # Layout horizontal para os botões
        btns_layout = QHBoxLayout()
        self.analisar_button = QPushButton("Analisar Logs")
        # Conecta o botão de analisar logs ao método _analisar_logs
        self.analisar_button.clicked.connect(self._analisar_logs)
        btns_layout.addWidget(self.analisar_button)

        self.exportar_button = QPushButton("Exportar Excel")
        # Conecta o botão de exportar Excel ao método _exportar_excel
        self.exportar_button.clicked.connect(self._exportar_excel)
        btns_layout.addWidget(self.exportar_button)

        layout.addLayout(btns_layout)

        # Área de texto para exibir o relatório
        self.text_area = QTextEdit()
        self.text_area.setReadOnly(True)
        layout.addWidget(self.text_area)

        # Configuração da figura e canvas para os gráficos
        self.figura = Figure(figsize=(12, 6)) # Aumentado o tamanho da figura para acomodar 4 gráficos
        self.canvas = FigureCanvas(self.figura)
        layout.addWidget(self.canvas)

    def _extrair(self, padrao, texto):
        # Extrai informações de um texto usando uma expressão regular
        match = re.search(padrao, texto)
        return match.group(1).strip() if match else None

    def _analisar_logs(self):
        # Abre uma caixa de diálogo para o usuário selecionar a pasta de logs
        pasta = QFileDialog.getExistingDirectory(self, "Selecionar Pasta de Logs", self.logs_dir)
        if not pasta:
            return

        self.resultados_detalhados.clear()
        
        operadores_encontrados = set()

        logs_por_placa = {}  # (PR, número de série) -> (data, conteúdo)

        for arquivo in os.listdir(pasta):
            if not arquivo.endswith(".txt"):
                continue
            caminho = os.path.join(pasta, arquivo)
            try:
                with open(caminho, "r", encoding="utf-8") as f:
                    conteudo = f.read()

                pr_doc = self._extrair(r"N[uú]mero do PR:\s*(.*)", conteudo) # Original PR do documento
                numero_serie = self._extrair(r"N[uú]mero de S[ée]rie da Placa:\s*(.*)", conteudo)
                data_fim = self._extrair(r"Data/Hora T[ée]rmino:\s*(.*)", conteudo)

                if not pr_doc or not numero_serie or not data_fim:
                    print(f"Aviso: Dados essenciais faltando em {arquivo}. Pulando este arquivo.")
                    continue

                dt_fim = datetime.strptime(data_fim, "%Y-%m-%d %H:%M:%S")
                chave = (pr_doc.strip(), numero_serie.strip())

                status_matches = re.findall(r"--- STATUS FINAL: (APROVADO|REPROVADO)", conteudo.upper())
                resultado_final = status_matches[-1] if status_matches else None

                if not resultado_final:
                    print(f"Aviso: Status final não encontrado em {arquivo}. Pulando este arquivo.")
                    continue  # ignora logs sem status final

                if chave not in logs_por_placa:
                    logs_por_placa[chave] = (dt_fim, conteudo)
                else:
                    dt_existente, _ = logs_por_placa[chave]
                    if dt_fim >= dt_existente: # Considera o log mais recente
                        logs_por_placa[chave] = (dt_fim, conteudo)

            except Exception as e:
                print(f"Erro ao processar {arquivo}: {e}")

        for (pr_doc, numero_serie), (dt_fim, conteudo) in logs_por_placa.items():
            operador = self._extrair(r"Operador do Teste:\s*(.*)", conteudo)
            data_inicio = self._extrair(r"Data/Hora In[ií]cio:\s*(.*)", conteudo)
            hostname = self._extrair(r"Máquina de Teste:\s*(.*)", conteudo)
            status_matches = re.findall(r"--- STATUS FINAL: (APROVADO|REPROVADO)", conteudo.upper())
            resultado = status_matches[-1] if status_matches else "DESCONHECIDO"

            if not operador or not data_inicio:
                print(f"Aviso: Operador ou data de início faltando para PR {pr_doc}, NS {numero_serie}. Pulando este registro.")
                continue

            dt_inicio = datetime.strptime(data_inicio, "%Y-%m-%d %H:%M:%S")
            duracao_segundos = (dt_fim - dt_inicio).total_seconds() # Duração em segundos

            operadores_encontrados.add(operador)

            # Expressão regular mais robusta para capturar o número do passo e a descrição do passo reprovado
            passo_regex = r"PASSO\s*(\d+):\s*(.*?)\s*-\s*Em Execução[\s\S]*?Status:\s*PASSO\s*\d+:\s*(REPROVADO)"
            passos_info = re.findall(passo_regex, conteudo, re.IGNORECASE)
            
            passos_reprovados_info = [f"PASSO {num}: {desc.strip()}" for num, desc, status in passos_info if status.upper() == "REPROVADO"]
            
            pr_formatado = pr_doc
            if pr_doc and re.match(r'^\d+$', pr_doc.strip()):
                pr_formatado = f"PR{pr_doc.strip()}"

            self.resultados_detalhados.append({
                "pr": pr_formatado,
                "numero_serie": numero_serie,
                "operador": operador,
                "inicio": dt_inicio,
                "fim": dt_fim,
                "duracao_segundos": duracao_segundos, # Corrigido nome
                "resultado": resultado,
                "maquina": hostname,
                "passos_reprovados": passos_reprovados_info
            })

        self.operador_filtro.clear()
        self.operador_filtro.addItem("Todos os Operadores")
        for op in sorted(operadores_encontrados):
            self.operador_filtro.addItem(op)

        self._filtrar_e_mostrar()
    
    def _filtrar_e_mostrar(self):
        data_ini = self.data_inicio.date().toPyDate()
        data_fim = self.data_fim.date().toPyDate()
        operador_filtro = self.operador_filtro.currentText()
        placa_filtro_txt = self.placa_filtro.text().strip()

        resultados = defaultdict(lambda: {"aprovado": 0, "reprovado": 0, "total": 0, "tempos": []})
        dias = defaultdict(list)
        
        resultados_filtrados_para_exibicao = [] 

        for r in self.resultados_detalhados:
            if not (data_ini <= r["inicio"].date() <= data_fim):
                continue
            if operador_filtro != "Todos os Operadores" and r["operador"] != operador_filtro:
                continue
            
            # Lógica do filtro de Lote/Série da Placa (versão antiga para estabilidade)
            if placa_filtro_txt:
                numero_serie_log_upper = r["numero_serie"].upper()
                placa_filtro_upper = placa_filtro_txt.upper()

                if '/' in placa_filtro_txt and not placa_filtro_txt.endswith('/'):
                    if not numero_serie_log_upper == placa_filtro_upper:
                        continue
                else:
                    search_prefix = placa_filtro_upper
                    if not search_prefix.endswith('/'):
                        search_prefix += '/'
                    if not numero_serie_log_upper.startswith(search_prefix):
                        continue
            resultados_filtrados_para_exibicao.append(r)

            resultados[r["operador"]]["total"] += 1
            resultados[r["operador"]][r["resultado"].lower()] += 1
            resultados[r["operador"]]["tempos"].append(r["duracao_segundos"]) # Corrigido nome
            dias[r["inicio"].date()].append(r["duracao_segundos"]) # Corrigido nome

        if not resultados_filtrados_para_exibicao:
            self.text_area.setPlainText("Nenhum dado encontrado para os filtros selecionados.")
            self.figura.clear() # Limpa o gráfico
            self.canvas.draw()
            return

        # Calcular o total de placas testadas após todos os filtros
        total_placas_testadas = len(resultados_filtrados_para_exibicao)

        relatorio = f"Total de Placas Testadas (Filtradas): {total_placas_testadas}\n\n"
        relatorio += "Eficiência por Operador\n" + "-"*40 + "\n"
        
        for op, dados in resultados.items():
            operador_tem_dados_filtrados = False
            for r in resultados_filtrados_para_exibicao:
                if r["operador"] == op:
                    operador_tem_dados_filtrados = True
                    break
            
            if not operador_tem_dados_filtrados:
                continue

            media = sum(dados["tempos"]) / len(dados["tempos"]) if dados["tempos"] else 0
            pct = 100 * dados["aprovado"] / dados["total"] if dados["total"] else 0

            passos_reprovados_para_contagem = []
            for r in resultados_filtrados_para_exibicao:
                if r["operador"] == op:
                    # Adiciona as descrições dos passos reprovados para contagem
                    passos_reprovados_para_contagem.extend(r.get("passos_reprovados", []))

            # Contar a frequência de cada passo reprovado (descrição completa)
            contador = Counter(passos_reprovados_para_contagem)
            top_passos = contador.most_common(3) # Top 3 passos reprovados com suas descrições
            
            top_passos_str = "Nenhum"
            if top_passos:
                top_passos_str = "\n" + "\n".join([f"    - {p} (ocorrências: {n})" for p, n in top_passos])


            relatorio += (
                f"Operador: {op}\n"
                f"  Total de Testes: {dados['total']}\n"
                f"  Aprovados: {dados['aprovado']}\n"
                f"  Reprovados: {dados['reprovado']}\n"
                f"  % Aprovação: {pct:.1f}%\n"
                f"  Tempo Médio: {media:.1f} s ({media/60:.1f} min)\n"
                f"  Passos Críticos: {top_passos_str}\n"
                f"{'-'*30}\n"
            )

        self.text_area.setPlainText(relatorio)
        self._plotar_graficos(resultados, dias)

       
    def _plotar_graficos(self, resultados, dias):
        self.figura.clear()
        
        # Ajusta o layout para 4 subplots em uma linha
        ax1 = self.figura.add_subplot(1, 4, 1) 
        ax_aprov_pct = self.figura.add_subplot(1, 4, 2) 
        ax2 = self.figura.add_subplot(1, 4, 3) # Gráfico de pizza
        ax3 = self.figura.add_subplot(1, 4, 4) 


        operadores = list(resultados.keys())
        totais = [dados["total"] for dados in resultados.values()]
        
        ax1.set_xticks(range(len(operadores)))
        ax1.bar(operadores, totais, color='skyblue')
        ax1.set_title("Testes por Operador")
        ax1.set_xticklabels(operadores, rotation=45, ha="right", fontsize=8)
        ax1.yaxis.set_major_locator(MaxNLocator(integer=True)) # Garante inteiros no eixo Y

        pct_aprovados_por_operador = []
        for op_dados in resultados.values():
            if op_dados["total"] > 0:
                pct_aprovados_por_operador.append(100 * op_dados["aprovado"] / op_dados["total"])
            else:
                pct_aprovados_por_operador.append(0) 

        ax_aprov_pct.set_xticks(range(len(operadores)))
        ax_aprov_pct.bar(operadores, pct_aprovados_por_operador, color='lightgreen')
        ax_aprov_pct.set_title("% Aprovação por Operador")
        ax_aprov_pct.set_xticklabels(operadores, rotation=45, ha="right", fontsize=8)
        ax_aprov_pct.set_ylim(0, 100) 
        for i, val in enumerate(pct_aprovados_por_operador):
            ax_aprov_pct.text(i, val - 1, f'{val:.1f}%', ha='center', va='top', fontsize=8, color='black')


        aprovados_totais_por_operador = [dados["aprovado"] for dados in resultados.values()]
        # Garante que o gráfico de pizza só seja plotado se houver dados aprovados
        if sum(aprovados_totais_por_operador) > 0: 
            ax2.pie(aprovados_totais_por_operador, labels=operadores, autopct="%1.1f%%", startangle=90)
            ax2.set_title("Contribuição nos Testes Aprovados") 
        else:
            ax2.text(0.5, 0.5, "Sem dados de aprovação", ha='center', va='center', transform=ax2.transAxes)
            ax2.set_title("Gráfico de Pizza") # Título padrão quando não há dados
        
        ax2.axis('equal') # Garante que o gráfico de pizza seja circular


        data_ini = self.data_inicio.date().toPyDate()
        data_fim = self.data_fim.date().toPyDate()

        datas_filtradas = [d for d in sorted(dias.keys()) if data_ini <= d <= data_fim]
        medias = [sum(dias[d]) / len(dias[d]) for d in datas_filtradas] # Medias em segundos

        ax3.plot(datas_filtradas, [m/60 for m in medias], marker="o") # Divide por 60 para exibir em minutos
        ax3.set_title("Tempo Médio por Dia (min)")
        ax3.tick_params(axis="x", labelrotation=45)

        self.figura.tight_layout() # Ajusta o layout para evitar sobreposição
        self.canvas.draw() # Desenha os gráficos no canvas

    def _exportar_excel(self):
        data_ini = self.data_inicio.date().toPyDate()
        data_fim = self.data_fim.date().toPyDate()
        operador_filtro = self.operador_filtro.currentText()
        placa_filtro_txt = self.placa_filtro.text().strip()

        dados_para_exportar = []
        for r in self.resultados_detalhados:
            if not (data_ini <= r["inicio"].date() <= data_fim):
                continue
            if operador_filtro != "Todos os Operadores" and r["operador"] != operador_filtro:
                continue
            
            # Aplica o mesmo filtro combinado para Lote e Série da Placa
            if placa_filtro_txt:
                numero_serie_log_upper = r["numero_serie"].upper()
                placa_filtro_upper = placa_filtro_txt.upper()

                if '/' in placa_filtro_txt and not placa_filtro_txt.endswith('/'):
                    if not numero_serie_log_upper == placa_filtro_upper:
                        continue
                else:
                    search_prefix = placa_filtro_upper
                    if not search_prefix.endswith('/'):
                        search_prefix += '/'
                    if not numero_serie_log_upper.startswith(search_prefix):
                        continue
            dados_para_exportar.append(r)

        if not dados_para_exportar:
            QMessageBox.warning(self, "Sem Dados", "Nenhum dado disponível para exportar com os filtros atuais.")
            return

        salvar_em, _ = QFileDialog.getSaveFileName(self, "Salvar Relatório Excel", "relatorio_teste.xlsx", "Excel (*.xlsx)")
        if not salvar_em:
            return

        df = pd.DataFrame(dados_para_exportar)

        column_order = [
            "pr",
            "numero_serie",
            "operador",
            "inicio",
            "fim",
            "duracao_segundos", # Corrigido nome
            "resultado",
            "maquina",
            "passos_reprovados"
        ]
        df = df[column_order]

        df.rename(columns={
            "pr": "PR",
            "numero_serie": "Número de Série",
            "operador": "Operador",
            "inicio": "Início",
            "fim": "Fim",
            "duracao_segundos": "Duração (minutos)", # Corrigido nome
            "resultado": "Resultado",
            "maquina": "Máquina",
            "passos_reprovados": "Passos Reprovados"
        }, inplace=True)

        df["Início"] = df["Início"].dt.strftime("%Y-%m-%d %H:%M:%S")
        df["Fim"] = df["Fim"].dt.strftime("%Y-%m-%d %H:%M:%S")
        df["Duração (minutos)"] = df["Duração (minutos)"].apply(lambda x: round(x / 60, 2))

        # Limita a quantidade de passos reprovados exibidos por célula
        def resumir_passos(passos):
            if isinstance(passos, list):
                max_passos = 5
                if len(passos) > max_passos:
                    return "\n".join(passos[:max_passos]) + f"\n... (+{len(passos)-max_passos} mais)"
                else:
                    return "\n".join(passos)
            return ""
        df["Passos Reprovados"] = df["Passos Reprovados"].apply(resumir_passos)

        try:
            # Usa ExcelWriter para ter mais controle sobre a formatação
            writer = pd.ExcelWriter(salvar_em, engine='xlsxwriter')
            df.to_excel(writer, sheet_name='Relatório de Eficiência', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Relatório de Eficiência']

            # Adiciona um formato para o cabeçalho com as cores da empresa
            header_format = workbook.add_format({
                'bold': True,
                'text_wrap': True,
                'valign': 'vcenter',
                'align': 'center', # Alinha o texto do cabeçalho ao centro
                'fg_color': '#2C4B7A', # Cor de fundo azul escuro (aproximado da imagem)
                'font_color': '#FFFFFF', # Cor do texto branco
                'border': 1
            })
            
            # Formato para linhas pares com quebra de texto
            light_green_row_format = workbook.add_format({'fg_color': '#D7E4BC', 'text_wrap': True})
            # Formato para linhas ímpares com quebra de texto
            default_row_format = workbook.add_format({'text_wrap': True})


            # Escreve os cabeçalhos das colunas com o formato definido
            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            # Escreve os dados e aplica formatação de linha alternada
            for row_num, r in df.iterrows():
                for col_num, value in enumerate(r):
                    # Aplica a cor de fundo verde claro para as linhas de dados, se desejado
                    if row_num % 2 == 0: # Exemplo de linhas alternadas (pares com verde claro)
                        worksheet.write(row_num + 1, col_num, value, light_green_row_format)
                    else: # Linhas ímpares com o formato padrão (fundo branco, com quebra de texto)
                        worksheet.write(row_num + 1, col_num, value, default_row_format)


            # Auto-ajusta a largura das colunas
            for i, col in enumerate(df.columns):
                # Encontra o comprimento máximo do conteúdo na coluna
                # Pega o comprimento do título da coluna e o comprimento máximo dos dados
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2 # Adiciona um pouco de espaço extra
                worksheet.set_column(i, i, max_len)

            writer.close() # Usa writer.close() para versões mais recentes do pandas
            QMessageBox.information(self, "Sucesso", f"Relatório exportado com sucesso: {salvar_em}")
        except Exception as e:
            # Exibe um erro se a exportação falhar
            QMessageBox.critical(self, "Erro", f"Falha ao exportar: {e}")

