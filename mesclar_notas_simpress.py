import os
import sys
from datetime import datetime
import fitz
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QLabel, QPushButton, QTextEdit, 
                             QProgressBar, QFileDialog, QMessageBox, QListWidget,
                             QGroupBox)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QFont

class WorkerThread(QThread):
    update_signal = pyqtSignal(str)
    progress_signal = pyqtSignal(int)
    finished_signal = pyqtSignal(dict)
    error_signal = pyqtSignal(str)

    def __init__(self, diretorio):
        super().__init__()
        self.diretorio = diretorio

    def run(self):
        try:
            self.update_signal.emit("Analisando estrutura de pastas...
            
            # Verificar se o diretório selecionado já é uma pasta Simpress
            dir_name = os.path.basename(self.diretorio)
            if dir_name.startswith("Simpress"):
                try:
                    # Tentar extrair a data do nome da pasta
                    data_str = dir_name.split()[-1]
                    datetime.strptime(data_str, "%Y.%m")
                    pasta_recente = dir_name
                    caminho_pasta_recente = self.diretorio
                    self.update_signal.emit(f"Processando pasta Simpress: {pasta_recente}")
                except ValueError:
                    # Se não for uma pasta Simpress com data válida, procurar subpastas
                    self.update_signal.emit("Procurando pastas Simpress...")
                    pastas = [pasta for pasta in os.listdir(self.diretorio) 
                             if os.path.isdir(os.path.join(self.diretorio, pasta)) and pasta.startswith("Simpress")]
                    
                    if not pastas:
                        self.error_signal.emit("Nenhuma pasta Simpress encontrada no diretório selecionado.")
                        return

                    # Ordenar pastas pela data
                    pastas_ordenadas = []
                    for pasta in pastas:
                        try:
                            data_str = pasta.split()[-1]
                            data = datetime.strptime(data_str, "%m.%Y")
                            pastas_ordenadas.append((data, pasta))
                        except ValueError:
                            continue
                    
                    if not pastas_ordenadas:
                        self.error_signal.emit("Nenhuma pasta com formato de data válido (MM.YYYY) encontrada.")
                        return
                    
                    pastas_ordenadas.sort(key=lambda x: x[0], reverse=True)
                    pasta_recente = pastas_ordenadas[0][1]
                    caminho_pasta_recente = os.path.join(self.diretorio, pasta_recente)
            else:
                # Procurar pastas Simpress no diretório selecionado
                self.update_signal.emit("Procurando pastas Simpress...")
                pastas = [pasta for pasta in os.listdir(self.diretorio) 
                         if os.path.isdir(os.path.join(self.diretorio, pasta)) and pasta.startswith("Simpress")]
                
                if not pastas:
                    self.error_signal.emit("Nenhuma pasta Simpress encontrada no diretório selecionado.")
                    return

                # Ordenar pastas pela data
                pastas_ordenadas = []
                for pasta in pastas:
                    try:
                        data_str = pasta.split()[-1]
                        data = datetime.strptime(data_str, "%m.%Y")
                        pastas_ordenadas.append((data, pasta))
                    except ValueError:
                        continue
                
                if not pastas_ordenadas:
                    self.error_signal.emit("Nenhuma pasta com formato de data válido (MM.YYYY) encontrada.")
                    return
                
                pastas_ordenadas.sort(key=lambda x: x[0], reverse=True)
                pasta_recente = pastas_ordenadas[0][1]
                caminho_pasta_recente = os.path.join(self.diretorio, pasta_recente)
            
            self.update_signal.emit(f"Pasta selecionada: {pasta_recente}")
            self.update_signal.emit(f"Caminho: {caminho_pasta_recente}")
            self.progress_signal.emit(20)

            # Criar pasta press se não existir
            pasta_press = os.path.join(caminho_pasta_recente, "press")
            
            if not os.path.exists(pasta_press):
                os.makedirs(pasta_press)
                self.update_signal.emit(f"Pasta criada: {pasta_press}")
            else:
                self.update_signal.emit(f"Pasta já existe: {pasta_press}")
            
            self.progress_signal.emit(40)

            # Verificar se existem as pastas Relatorio e Nota
            pasta_relatorio = os.path.join(caminho_pasta_recente, 'Relatorio')
            pasta_nota = os.path.join(caminho_pasta_recente, 'Nota')
            
            if not os.path.exists(pasta_relatorio):
                self.error_signal.emit(f"Pasta 'Relatorio' não encontrada em {caminho_pasta_recente}")
                return
            
            if not os.path.exists(pasta_nota):
                self.error_signal.emit(f"Pasta 'Nota' não encontrada em {caminho_pasta_recente}")
                return
            
            self.update_signal.emit("Pastas 'Relatorio' e 'Nota' encontradas. Iniciando processamento...")
            
            # Processar arquivos
            codigos = {}
            total_arquivos = len(os.listdir(pasta_relatorio)) + len(os.listdir(pasta_nota))
            processados = 0
            
            # Processar arquivos da pasta Relatorio
            for nome_arquivo in os.listdir(pasta_relatorio):
                if nome_arquivo.lower().endswith('.pdf'):
                    try:
                        partes = nome_arquivo.split('_')
                        if len(partes) >= 4:
                            codigo = partes[3]
                            if codigo not in codigos:
                                codigos[codigo] = []
                            codigos[codigo].append(os.path.join(pasta_relatorio, nome_arquivo))
                    except IndexError:
                        self.update_signal.emit(f"Formato de nome inválido: {nome_arquivo}")
                processados += 1
                self.progress_signal.emit(40 + int(20 * processados / total_arquivos))
            
            # Processar arquivos da pasta Nota
            for nome_arquivo in os.listdir(pasta_nota):
                if nome_arquivo.lower().endswith('.pdf'):
                    try:
                        partes = nome_arquivo.split('_')
                        if len(partes) >= 4:
                            codigo = partes[3]
                            if codigo not in codigos:
                                codigos[codigo] = []
                            codigos[codigo].append(os.path.join(pasta_nota, nome_arquivo))
                    except IndexError:
                        self.update_signal.emit(f"Formato de nome inválido: {nome_arquivo}")
                processados += 1
                self.progress_signal.emit(60 + int(20 * processados / total_arquivos))
            
            # Unir PDFs
            resultados = {}
            total_codigos = len(codigos)
            for i, (codigo, arquivos) in enumerate(codigos.items()):
                if len(arquivos) == 2:
                    self.update_signal.emit(f"Unindo arquivos para o código {codigo}...")
                    
                    try:
                        resultado = fitz.open()
                        for arquivo in arquivos:
                            pdf = fitz.open(arquivo)
                            resultado.insert_pdf(pdf)
                        
                        caminho_completo = os.path.join(pasta_press, f'{codigo}.pdf')
                        resultado.save(caminho_completo)
                        resultados[codigo] = caminho_completo
                        
                        self.update_signal.emit(f"Arquivo salvo: {os.path.basename(caminho_completo)}")
                    except Exception as e:
                        self.update_signal.emit(f"Erro ao processar código {codigo}: {str(e)}")
                else:
                    self.update_signal.emit(f"Código {codigo} tem {len(arquivos)} arquivos (esperados: 2)")
                
                self.progress_signal.emit(80 + int(20 * (i + 1) / total_codigos))
            
            self.progress_signal.emit(100)
            self.finished_signal.emit(resultados)
            
        except Exception as e:
            self.error_signal.emit(f"Erro durante o processamento: {str(e)}")

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Fusionador de PDFs - Simpress")
        self.setGeometry(100, 100, 800, 600)
        
        # Widget central e layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # Título
        title = QLabel("Fusionador de PDFs Simpress")
        title.setFont(QFont("Arial", 16, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("padding: 15px; background-color: #2c3e50; color: white;")
        layout.addWidget(title)
        
        # Grupo de informações
        info_group = QGroupBox("Informações do Processamento")
        info_layout = QVBoxLayout()
        
        self.info_label = QLabel("Selecione o diretório que contém as pastas Simpress ou uma pasta Simpress específica.")
        self.info_label.setWordWrap(True)
        info_layout.addWidget(self.info_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        info_layout.addWidget(self.progress_bar)
        
        info_group.setLayout(info_layout)
        layout.addWidget(info_group)
        
        # Área de log
        log_group = QGroupBox("Log de Execução")
        log_layout = QVBoxLayout()
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        log_group.setLayout(log_layout)
        layout.addWidget(log_group)
        
        # Lista de resultados
        results_group = QGroupBox("Arquivos Processados")
        results_layout = QVBoxLayout()
        
        self.results_list = QListWidget()
        results_layout.addWidget(self.results_list)
        
        results_group.setLayout(results_layout)
        layout.addWidget(results_group)
        
        # Botões
        button_layout = QHBoxLayout()
        
        self.select_button = QPushButton("Selecionar Diretório")
        self.select_button.clicked.connect(self.select_directory)
        button_layout.addWidget(self.select_button)
        
        self.process_button = QPushButton("Iniciar Processamento")
        self.process_button.clicked.connect(self.start_processing)
        self.process_button.setEnabled(False)
        button_layout.addWidget(self.process_button)
        
        self.open_button = QPushButton("Abrir Pasta de Resultados")
        self.open_button.clicked.connect(self.open_results_folder)
        self.open_button.setEnabled(False)
        button_layout.addWidget(self.open_button)
        
        layout.addLayout(button_layout)
        
        # Variáveis de estado
        self.selected_directory = None
        self.worker_thread = None
        self.result_files = {}
        
        # Aplicar estilo
        self.apply_style()
    
    def apply_style(self):
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ecf0f1;
            }
            QGroupBox {
                font-weight: bold;
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                margin-top: 1ex;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
                color: #2c3e50;
            }
            QPushButton {
                background-color: #3498db;
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
            QProgressBar {
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                text-align: center;
            }
            QProgressBar::chunk {
                background-color: #2ecc71;
                width: 10px;
            }
            QTextEdit {
                background-color: white;
            }
        """)
    
    def select_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Selecionar Diretório")
        if directory:
            self.selected_directory = directory
            self.info_label.setText(f"Diretório selecionado: {directory}")
            self.log_text.append(f"Diretório selecionado: {directory}")
            self.process_button.setEnabled(True)
    
    def start_processing(self):
        if not self.selected_directory:
            QMessageBox.warning(self, "Aviso", "Por favor, selecione um diretório primeiro.")
            return
        
        # Limpar resultados anteriores
        self.results_list.clear()
        self.result_files = {}
        self.progress_bar.setValue(0)
        self.log_text.clear()
        
        # Desabilitar botões durante o processamento
        self.select_button.setEnabled(False)
        self.process_button.setEnabled(False)
        
        # Criar e iniciar thread de processamento
        self.worker_thread = WorkerThread(self.selected_directory)
        self.worker_thread.update_signal.connect(self.update_log)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.finished_signal.connect(self.process_finished)
        self.worker_thread.error_signal.connect(self.process_error)
        self.worker_thread.start()
    
    def update_log(self, message):
        self.log_text.append(message)
        # Rolagem automática para o final
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def update_progress(self, value):
        self.progress_bar.setValue(value)
    
    def process_finished(self, results):
        self.result_files = results
        self.select_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.open_button.setEnabled(bool(results))
        
        # Adicionar resultados à lista
        for codigo, caminho in results.items():
            self.results_list.addItem(f"{codigo}: {os.path.basename(caminho)}")
        
        self.update_log("Processamento concluído com sucesso!")
        QMessageBox.information(self, "Sucesso", f"Processamento concluído. {len(results)} arquivos foram gerados.")
    
    def process_error(self, error_message):
        self.select_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.update_log(f"ERRO: {error_message}")
        QMessageBox.critical(self, "Erro", error_message)
    
    def open_results_folder(self):
        if self.result_files:
            # Abrir o diretório do primeiro arquivo de resultado
            first_file = next(iter(self.result_files.values()))
            directory = os.path.dirname(first_file)
            try:
                os.startfile(directory)  # No Windows
            except:
                # Para outros sistemas operacionais
                import subprocess
                try:
                    subprocess.Popen(['xdg-open', directory])  # Linux
                except:
                    try:
                        subprocess.Popen(['open', directory])  # macOS
                    except:
                        self.update_log(f"Não foi possível abrir a pasta: {directory}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())