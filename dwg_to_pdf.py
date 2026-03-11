import os
import shutil
import sys
import time
import comtypes.client
import pythoncom
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QFileDialog, QGroupBox, QComboBox,
    QCheckBox, QProgressBar, QTextEdit, QMessageBox, QTabWidget, QListWidget,
    QAbstractItemView, QStackedWidget, QSpinBox, QRadioButton, QButtonGroup,
    QSizePolicy, QGraphicsView, QGraphicsScene
)
from PySide6.QtCore import Qt, QThread, Signal, QSettings, QRectF
from PySide6.QtGui import QIcon, QImage, QPixmap, QPainter, QColor, QFont, QTransform
import qdarkstyle
import PyPDF2
import io
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import Color

# ====== Vistas Personalizadas ======
class ZoomGraphicsView(QGraphicsView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragMode(QGraphicsView.ScrollHandDrag)
        self.setRenderHint(QPainter.Antialiasing)
        self.setRenderHint(QPainter.SmoothPixmapTransform)
        self.custom_scene = QGraphicsScene(self)
        self.setScene(self.custom_scene)

    def wheelEvent(self, event):
        if event.modifiers() == Qt.ControlModifier:
            zoom_in_factor = 1.15
            zoom_out_factor = 1 / zoom_in_factor
            if event.angleDelta().y() > 0:
                zoom_factor = zoom_in_factor
            else:
                zoom_factor = zoom_out_factor
            self.scale(zoom_factor, zoom_factor)
        else:
            super().wheelEvent(event)

# ====== Constantes de Trazado de AutoCAD (ActiveX) ======
PLOT_AREA_MAP = {
    "Extents": 1,
    "Layout": 5,
    "Limits": 3,
    "Window": 4
}

def com_retry(func, *args, retries=15, delay=2.0, **kwargs):
    """
    Intenta ejecutar una llamada COM varias veces. 
    Previene errores 0x80010001 (RPC_E_CALL_REJECTED) cuando AutoCAD está ocupado.
    """
    for i in range(retries):
        try:
            return func(*args, **kwargs)
        except Exception as e:
            err_str = str(e).lower()
            if "-2147418111" in err_str or "rechazada" in err_str or "rejected" in err_str or "80010001" in err_str or i == retries - 1:
                if i == retries - 1:
                    raise e
                time.sleep(delay)
            else:
                raise e

class PlotWorker(QThread):
    """
    Subproceso para manejar la lógica de ActiveX en background 
    y evitar que la UI de PySide6 se congele.
    """
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished_signal = Signal()
    error_signal = Signal(str)

    def __init__(self, dwg_files, output_dir, ctb_path, plot_config):
        super().__init__()
        self.dwg_files = dwg_files
        self.output_dir = output_dir
        self.ctb_path = ctb_path
        self.plot_config = plot_config

    def run(self):
        # OBLIGATORIO inicializar COM para este hilo
        pythoncom.CoInitialize()
        
        opened_background_instance = False

        
        try:
            self.log_signal.emit("Conectando con el motor de AutoCAD vía COM...")
            try:
                acad = comtypes.client.GetActiveObject("AutoCAD.Application")
                self.log_signal.emit(" [OK] Conexión exitosa a instancia activa de AutoCAD.")
            except OSError:
                self.log_signal.emit(" Iniciando nueva instancia de AutoCAD en background...")
                acad = comtypes.client.CreateObject("AutoCAD.Application", dynamic=True)
                opened_background_instance = True
                
                # Para que NO abra la ventana gráfica de AutoCAD
                try:
                    com_retry(lambda: setattr(acad, 'Visible', False))
                except Exception:
                    pass
                
                # Respaldo: Si AutoCAD ignora el comando Visible=False por su Splash screen, 
                # lo forzamos a minimizarse al abrir (WindowState = 2: acMin).
                try:
                    com_retry(lambda: setattr(acad, 'WindowState', 2))
                except Exception:
                    pass
                
                self.log_signal.emit(" [OK] AutoCAD iniciado de forma invisible con éxito.")

            # Gestionar la ruta del CTB
            ctb_name = ""
            if self.ctb_path and os.path.exists(self.ctb_path):
                ctb_name = os.path.basename(self.ctb_path)
                try:
                    pref_files = com_retry(lambda: acad.Preferences.Files)
                    plot_style_path = com_retry(lambda: pref_files.PrinterStyleSheetPath)
                    
                    ctb_dest = os.path.join(plot_style_path, ctb_name)
                    if os.path.abspath(self.ctb_path) != os.path.abspath(ctb_dest):
                        shutil.copy2(self.ctb_path, ctb_dest)
                        self.log_signal.emit(f" CTB '{ctb_name}' añadido a AutoCAD Plot Styles.")
                except Exception as e:
                    self.log_signal.emit(f" [Aviso] No se pudo auto-copiar el CTB: {e}")

            # ====== ITERACIÓN ======
            for idx, dwg_path in enumerate(self.dwg_files):
                filename = os.path.basename(dwg_path)
                self.log_signal.emit(f"\n[{idx+1}/{len(self.dwg_files)}] >> Abriendo y procesando: {filename}")
                
                doc = None
                try:
                    doc = com_retry(lambda: acad.Documents.Open(dwg_path))
                    
                    # FUNDAMENTAL: Desactivar Ploteo en Segundo Plano (Asíncrono) 
                    # para que Python no cierre el archivo mientras AutoCAD sigue imprimiendo.
                    try:
                        com_retry(lambda: doc.SetVariable("BACKGROUNDPLOT", 0))
                    except Exception as e_bg:
                        self.log_signal.emit(f"      [Aviso] No se pudo desactivar BACKGROUNDPLOT: {e_bg}")
                    
                    layouts_found = 0
                    
                    # Convertimos la colección COM a una lista de Python de una sola vez
                    def get_layouts():
                        return list(doc.Layouts)
                    
                    layouts = com_retry(get_layouts)
                    
                    for layout in layouts:
                        layout_name = com_retry(lambda: layout.Name)
                        if layout_name.lower() == "model":
                            continue  
                            
                        layouts_found += 1
                        
                        def set_active_layout(): doc.ActiveLayout = layout
                        com_retry(set_active_layout)
                        
                        self.log_signal.emit(f"   -> Configurando Layout '{layout_name}'...")
                        
                        # --- 1. Printer / Plotter y Puntas ---
                        def set_plotter(): layout.ConfigName = self.plot_config.get("plotter", "DWG To PDF.pc3")
                        com_retry(set_plotter)
                        
                        if ctb_name:
                            def set_ctb(): layout.StyleSheet = ctb_name
                            com_retry(set_ctb)
                            
                        # --- 2. Paper Size ---
                        paper_size = self.plot_config.get("paper_size", "").strip()
                        if paper_size:
                            try:
                                def set_paper(): layout.CanonicalMediaName = paper_size
                                com_retry(set_paper)
                            except Exception as e_media:
                                self.log_signal.emit(f"      [Aviso] Tamaño de papel '{paper_size}' no soportado: {e_media}")

                        # --- 3. Plot Area ---
                        area_str = self.plot_config.get("plot_area", "Extents")
                        plot_type_const = PLOT_AREA_MAP.get(area_str, 1)
                        try:
                            def set_area(): layout.PlotType = plot_type_const
                            com_retry(set_area)
                        except Exception:
                            self.log_signal.emit(f"      [Aviso] Fallo en Area de trazado '{area_str}'. Usando Layout estándar.")
                            def set_area_std(): layout.PlotType = 5
                            com_retry(set_area_std) # acLayout
                            
                        # --- 3.5 Orientación ---
                        ori_text = self.plot_config.get("orientation", "Automático")
                        if ori_text != "Automático":
                            try:
                                rot_val = 0
                                if "90" in ori_text: rot_val = 1
                                elif "180" in ori_text: rot_val = 2
                                elif "270" in ori_text: rot_val = 3
                                
                                def set_rotation(): layout.PlotRotation = rot_val
                                com_retry(set_rotation)
                            except Exception as e_rot:
                                self.log_signal.emit(f"      [Aviso] Fallo en Orientación de trazado: {e_rot}")
                            
                        # --- 4. Plot Scale ---
                        if self.plot_config.get("fit_to_paper", True):
                            def set_fit(): layout.StandardScale = 0
                            com_retry(set_fit) # acScaleToFit
                        else:
                            try:
                                def set_custom_scale_type(): layout.StandardScale = 1
                                com_retry(set_custom_scale_type) # acScaleToCustom
                                
                                custom_num = float(self.plot_config.get("scale_num", 1.0))
                                custom_den = float(self.plot_config.get("scale_den", 1.0))
                                
                                # En AutoCAD, la asignación de escala custom se hace mediante SetCustomScale
                                def apply_custom_scale(): layout.SetCustomScale(custom_num, custom_den)
                                com_retry(apply_custom_scale)
                            except Exception as es:
                                self.log_signal.emit(f"      [Aviso] Fallo al establecer escala custom: {es}")

                        # --- 5. Offset & Lineweights ---
                        try:
                            def set_center_plot(): layout.CenterPlot = self.plot_config.get("center_plot", True)
                            com_retry(set_center_plot)
                        except Exception:
                            pass
                        
                        try:
                            def set_lineweights(): layout.ScaleLineweights = self.plot_config.get("scale_lineweights", False)
                            com_retry(set_lineweights)
                        except Exception:
                            pass
                            
                        # Refrescamos antes de imprimir
                        com_retry(lambda: layout.RefreshPlotDeviceInfo())
                        
                        # Guardado
                        base_name = os.path.splitext(filename)[0]
                        pdf_name = f"{base_name}.pdf"
                        pdf_path = os.path.join(self.output_dir, pdf_name)
                        
                        try:
                            if os.path.exists(pdf_path):
                                os.remove(pdf_path)
                        except OSError:
                            self.log_signal.emit(f"   [ERROR] El archivo {pdf_name} se encuentra abierto/bloqueado.")
                            continue
                            
                        # Plotear
                        self.log_signal.emit(f"   -> Generando PDF: {pdf_name}")
                        com_retry(lambda: doc.Plot.PlotToFile(pdf_path, self.plot_config.get("plotter", "DWG To PDF.pc3")))
                        
                    if layouts_found == 0:
                        self.log_signal.emit("   [INFO] No se encontraron Layouts disponibles en el archivo.")
                        
                except Exception as e:
                    self.log_signal.emit(f"   [ERROR] en '{filename}': {str(e)}")
                    
                finally:
                    if doc:
                        try:
                            com_retry(lambda: doc.Close(False))
                        except Exception as e_close:
                            self.log_signal.emit(f"   [Error] al cerrar documento: {str(e_close)}")
                            
                self.progress_signal.emit(idx + 1)

            self.log_signal.emit("\n====== PROCESO BATCH COMPLETADO ======")

        except Exception as e:
            self.error_signal.emit(f"Error fatal de ActiveX:\n{str(e)}")
            
        finally:
            # Si levantamos una instancia virgen de fondo, la cerramos para dejar el sistema limpio.
            if opened_background_instance:
                try:
                    self.log_signal.emit(" Cerrando instancia background de AutoCAD...")
                    com_retry(lambda: acad.Quit())
                except Exception:
                    pass
                    
            pythoncom.CoUninitialize()
            self.finished_signal.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AutoCAD Batch Plotter PRO - By: Andrés Gallo")
        self.resize(800, 700)
        
        self.dwg_files = []
        # ====== MEMORIA DE RUTAS ======
        # QSettings guarda registros en el sistema (Registro en Win, plist en Mac) 
        # asociándolo al nombre de tu app ("BatchPlotter").
        self.settings = QSettings("MiEmpresa", "BatchPlotterPRO")
        
        self.setup_ui()
        self.load_memory()
        
    def setup_ui(self):
        # ====== CONTENEDOR DE PESTAÑAS ======
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)
        
        # Pestaña 1: DWG a PDF
        self.tab_dwg = QWidget()
        self.setup_dwg_tab(self.tab_dwg)
        self.tabs.addTab(self.tab_dwg, "Conversor DWG a PDF")
        
        # Pestaña 2: Herramientas PDF
        self.tab_pdf = QWidget()
        self.setup_pdf_tab(self.tab_pdf)
        self.tabs.addTab(self.tab_pdf, "Herramientas PDF")

    def setup_dwg_tab(self, parent_widget):
        main_layout = QVBoxLayout(parent_widget)
        main_layout.setSpacing(15)
        
        # --- SECCIÓN 1: ARCHIVOS ---
        gb_files = QGroupBox("1. Selección de Archivos")
        vbox_files = QVBoxLayout(gb_files)
        
        # Selección DWG
        hbox_dwg = QHBoxLayout()
        self.btn_dwg = QPushButton("Seleccionar Archivos DWG...")
        self.btn_dwg.clicked.connect(self.select_dwg_files)
        self.lbl_dwg_count = QLabel("0 archivo(s) seleccionado(s)")
        hbox_dwg.addWidget(self.btn_dwg)
        hbox_dwg.addWidget(self.lbl_dwg_count)
        hbox_dwg.addStretch()
        
        # CTB
        hbox_ctb = QHBoxLayout()
        self.txt_ctb = QLineEdit()
        self.txt_ctb.setPlaceholderText("Seleccionar Estilo de Puntas (.ctb)...")
        self.txt_ctb.setReadOnly(True)
        btn_ctb = QPushButton("Examinar...")
        btn_ctb.clicked.connect(self.select_ctb_file)
        hbox_ctb.addWidget(self.txt_ctb)
        hbox_ctb.addWidget(btn_ctb)
        
        # Output
        hbox_out = QHBoxLayout()
        self.txt_out = QLineEdit()
        self.txt_out.setPlaceholderText("Carpeta de Destino de los PDFs...")
        self.txt_out.setReadOnly(True)
        btn_out = QPushButton("Examinar...")
        btn_out.clicked.connect(self.select_output_dir)
        hbox_out.addWidget(self.txt_out)
        hbox_out.addWidget(btn_out)
        
        vbox_files.addLayout(hbox_dwg)
        vbox_files.addLayout(hbox_ctb)
        vbox_files.addLayout(hbox_out)
        main_layout.addWidget(gb_files)
        
        # --- SECCIÓN 2: CONFIGURACIÓN DE PLOTEO (Estilo AutoCAD) ---
        gb_plot = QGroupBox("2. Configuración de Trazado (Plot Settings)")
        vbox_plot = QVBoxLayout(gb_plot)
        
        # Fila 1: Plotter y Paper Size
        hbox_p1 = QHBoxLayout()
        hbox_p1.addWidget(QLabel("Plotter (ConfigName):"))
        self.combo_plotter = QComboBox()
        self.combo_plotter.setEditable(True)
        self.combo_plotter.addItems(["DWG To PDF.pc3", "AutoCAD PDF (High Quality Print).pc3", "AutoCAD PDF (General Documentation).pc3"])
        self.combo_plotter.setMinimumWidth(250)
        
        hbox_p1.addWidget(QLabel("Paper size:"))
        self.combo_paper = QComboBox()
        self.combo_paper.setEditable(True)
        # Algunos nombres canónicos estándar según AutoCAD. A menudo varían por Plotter.
        self.combo_paper.addItems(["ISO_full_bleed_A0_(841.00_x_1189.00_MM)", "ISO_full_bleed_A1_(594.00_x_841.00_MM)", "ISO_full_bleed_A2_(420.00_x_594.00_MM)", "ISO_full_bleed_A3_(297.00_x_420.00_MM)"])
        self.combo_paper.setMinimumWidth(250)
        
        hbox_p1.addWidget(self.combo_plotter)
        hbox_p1.addWidget(self.combo_paper)
        hbox_p1.addStretch()
        
        # Fila 2: Área, Escala y Opciones
        hbox_p2 = QHBoxLayout()
        
        # Plot Area
        vbox_area = QVBoxLayout()
        vbox_area.addWidget(QLabel("Plot area:"))
        self.combo_area = QComboBox()
        self.combo_area.addItems(["Extents", "Layout", "Window", "Limits"])
        vbox_area.addWidget(self.combo_area)
        
        # Plot Scale
        vbox_scale = QVBoxLayout()
        self.chk_fit = QCheckBox("Fit to paper")
        self.chk_fit.setChecked(True)
        self.chk_fit.toggled.connect(self.toggle_scale_inputs)
        
        hbox_custom_scale = QHBoxLayout()
        self.txt_scale_num = QLineEdit("1")
        self.txt_scale_num.setMaximumWidth(40)
        self.txt_scale_num.setEnabled(False)
        self.txt_scale_den = QLineEdit("1")
        self.txt_scale_den.setMaximumWidth(40)
        self.txt_scale_den.setEnabled(False)
        
        hbox_custom_scale.addWidget(self.txt_scale_num)
        hbox_custom_scale.addWidget(QLabel("mm ="))
        hbox_custom_scale.addWidget(self.txt_scale_den)
        hbox_custom_scale.addWidget(QLabel("units"))
        hbox_custom_scale.addStretch()
        
        vbox_scale.addWidget(self.chk_fit)
        vbox_scale.addLayout(hbox_custom_scale)
        
        # Plot Offset & Options
        vbox_options = QVBoxLayout()
        vbox_options.addWidget(QLabel("Orientación:"))
        self.cmb_orientation = QComboBox()
        self.cmb_orientation.addItems(["Automático", "Horizontal (90°)", "Vertical (0°)", "Horizontal Invertido (270°)", "Vertical Invertido (180°)"])
        vbox_options.addWidget(self.cmb_orientation)
        vbox_options.addSpacing(10)
        
        self.chk_center = QCheckBox("Center the plot")
        self.chk_center.setChecked(True)
        self.chk_scale_lineweights = QCheckBox("Scale lineweights")
        vbox_options.addWidget(self.chk_center)
        vbox_options.addWidget(self.chk_scale_lineweights)
        
        hbox_p2.addLayout(vbox_area)
        hbox_p2.addSpacing(40)
        hbox_p2.addLayout(vbox_scale)
        hbox_p2.addSpacing(40)
        hbox_p2.addLayout(vbox_options)
        hbox_p2.addStretch()
        
        vbox_plot.addLayout(hbox_p1)
        vbox_plot.addSpacing(10)
        vbox_plot.addLayout(hbox_p2)
        main_layout.addWidget(gb_plot)
        
        # --- SECCIÓN 3: EJECUCIÓN Y LOG ---
        self.btn_run = QPushButton("▶ INICIAR CONVERSIÓN BATCH")
        self.btn_run.setFixedHeight(45)
        self.btn_run.setStyleSheet("background-color: #2e7d32; color: white; font-weight: bold; font-size: 14px;")
        self.btn_run.clicked.connect(self.start_conversion) # Corrected button name
        main_layout.addWidget(self.btn_run)
        
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setTextVisible(True)
        main_layout.addWidget(self.progress)
        
        # Log Header
        hbox_log_header = QHBoxLayout()
        lbl_log = QLabel("Log del Sistema:")
        lbl_log.setStyleSheet("font-weight: bold; color: #a0a0a0;")
        hbox_log_header.addWidget(lbl_log)
        hbox_log_header.addStretch()
        
        self.btn_clear_log = QPushButton("🧽 Limpiar Consola")
        self.btn_clear_log.setStyleSheet("background-color: #424242; padding: 4px; border-radius: 4px;")
        hbox_log_header.addWidget(self.btn_clear_log)
        main_layout.addLayout(hbox_log_header)
        
        self.txt_log = QTextEdit()
        self.txt_log.setReadOnly(True)
        self.txt_log.setStyleSheet("background-color: #1e1e1e; color: #d4d4d4; font-family: Consolas;")
        main_layout.addWidget(self.txt_log)
        
        # Connect signal
        self.btn_clear_log.clicked.connect(self.txt_log.clear)

    def setup_pdf_tab(self, parent_widget):
        main_layout = QHBoxLayout(parent_widget)
        main_layout.setSpacing(10)
        
        # Sidebar Menu
        self.pdf_sidebar = QListWidget()
        self.pdf_sidebar.addItems(["Unir PDFs", "Separar PDF", "Marca de Agua", "Proteger (Clave)", "Rotar PDFs"])
        self.pdf_sidebar.setMinimumWidth(180)
        self.pdf_sidebar.setMaximumWidth(220)
        self.pdf_sidebar.setStyleSheet("font-size: 13px; padding: 5px;")
        main_layout.addWidget(self.pdf_sidebar)
        
        # Stacked Widget
        self.pdf_stack = QStackedWidget()
        main_layout.addWidget(self.pdf_stack, stretch=1)
        
        # --- 1. Merge UI ---
        w_merge = QWidget()
        self.setup_pdf_merge_ui(w_merge)
        self.pdf_stack.addWidget(w_merge)
        
        # --- 2. Split UI ---
        w_split = QWidget()
        self.setup_pdf_split_ui(w_split)
        self.pdf_stack.addWidget(w_split)
        
        # --- 3. Watermark UI ---
        w_watermark = QWidget()
        self.setup_pdf_watermark_ui(w_watermark)
        self.pdf_stack.addWidget(w_watermark)
        
        # --- 4. Password UI ---
        w_password = QWidget()
        self.setup_pdf_password_ui(w_password)
        self.pdf_stack.addWidget(w_password)
        
        # --- 5. Rotate UI ---
        w_rotate = QWidget()
        self.setup_pdf_rotate_ui(w_rotate)
        self.pdf_stack.addWidget(w_rotate)
        
        # Connections
        self.pdf_sidebar.currentRowChanged.connect(self.pdf_stack.setCurrentIndex)
        self.pdf_sidebar.setCurrentRow(0)

    def setup_pdf_merge_ui(self, widget):
        lay_merge = QVBoxLayout(widget)
        grp_merge = QGroupBox("Unión de archivos PDF (Merge)")
        lay_grp = QVBoxLayout(grp_merge)
        
        lbl_inst = QLabel("Seleccione los archivos PDF que desea unir.\nArrastre para reordenarlos.")
        lay_grp.addWidget(lbl_inst)
        
        lay_list_ctrl = QHBoxLayout()
        self.btn_add_pdf = QPushButton("Añadir PDFs")
        self.btn_remove_pdf = QPushButton("Quitar Seleccionado")
        self.btn_clear_pdf = QPushButton("Limpiar Todo")
        lay_list_ctrl.addWidget(self.btn_add_pdf)
        lay_list_ctrl.addWidget(self.btn_remove_pdf)
        lay_list_ctrl.addWidget(self.btn_clear_pdf)
        lay_grp.addLayout(lay_list_ctrl)
        
        self.list_pdfs = QListWidget()
        self.list_pdfs.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_pdfs.setDragDropMode(QAbstractItemView.InternalMove)
        lay_grp.addWidget(self.list_pdfs)
        
        lay_dest = QHBoxLayout()
        lay_dest.addWidget(QLabel("Destino:"))
        self.txt_pdf_out = QLineEdit()
        self.btn_pdf_out = QPushButton("Explorar...")
        lay_dest.addWidget(self.txt_pdf_out)
        lay_dest.addWidget(self.btn_pdf_out)
        lay_grp.addLayout(lay_dest)
        
        self.btn_merge_pdf = QPushButton("UNIR ARCHIVOS PDF")
        self.btn_merge_pdf.setStyleSheet("background-color: #0277bd; color: white; font-weight: bold; font-size: 14px; padding: 10px;")
        lay_grp.addWidget(self.btn_merge_pdf)
        
        lay_merge.addWidget(grp_merge)
        
        # Connections
        self.btn_add_pdf.clicked.connect(self.pdf_add_files)
        self.btn_remove_pdf.clicked.connect(self.pdf_remove_selected)
        self.btn_clear_pdf.clicked.connect(self.list_pdfs.clear)
        self.btn_pdf_out.clicked.connect(self.pdf_select_output)
        self.btn_merge_pdf.clicked.connect(self.pdf_merge_action)

    def setup_pdf_split_ui(self, widget):
        lay = QVBoxLayout(widget)
        grp = QGroupBox("Separar / Extraer Hojas PDF")
        lay_grp = QVBoxLayout(grp)
        
        lay_src = QHBoxLayout()
        lay_src.addWidget(QLabel("PDF original:"))
        self.txt_split_in = QLineEdit()
        btn_split_in = QPushButton("Explorar...")
        btn_split_in.clicked.connect(lambda: self.select_single_pdf(self.txt_split_in))
        lay_src.addWidget(self.txt_split_in)
        lay_src.addWidget(btn_split_in)
        lay_grp.addLayout(lay_src)
        
        lay_range = QHBoxLayout()
        lay_range.addWidget(QLabel("Extraer Desde pág:"))
        self.spin_split_from = QSpinBox()
        self.spin_split_from.setMinimum(1)
        self.spin_split_from.setMaximum(9999)
        lay_range.addWidget(self.spin_split_from)
        
        lay_range.addWidget(QLabel("Hasta pág:"))
        self.spin_split_to = QSpinBox()
        self.spin_split_to.setMinimum(1)
        self.spin_split_to.setMaximum(9999)
        lay_range.addWidget(self.spin_split_to)
        
        self.chk_split_all = QCheckBox("Explotar todas las paginas a planos sueltos")
        lay_range.addWidget(self.chk_split_all)
        lay_range.addStretch()
        lay_grp.addLayout(lay_range)
        
        lay_dest = QHBoxLayout()
        lay_dest.addWidget(QLabel("Carpeta Salida:"))
        self.txt_split_out = QLineEdit()
        btn_split_out = QPushButton("Explorar...")
        btn_split_out.clicked.connect(lambda: self.select_output_dir_generic(self.txt_split_out))
        lay_dest.addWidget(self.txt_split_out)
        lay_dest.addWidget(btn_split_out)
        lay_grp.addLayout(lay_dest)
        
        btn_action = QPushButton("SEPARAR PDF")
        btn_action.setStyleSheet("background-color: #0277bd; color: white; font-weight: bold; font-size: 14px; padding: 10px;")
        btn_action.clicked.connect(self.pdf_split_action)
        lay_grp.addWidget(btn_action)
        
        lay.addWidget(grp)
        lay.addStretch()

    def setup_pdf_watermark_ui(self, widget):
        lay_main = QHBoxLayout(widget)
        
        # Panel Izquierdo: Controles
        panel_ctrl = QWidget()
        lay_ctrl = QVBoxLayout(panel_ctrl)
        lay_ctrl.setContentsMargins(0,0,10,0)
        
        grp = QGroupBox("Opciones de Marca de Agua")
        lay_grp = QVBoxLayout(grp)
        
        lay_src = QHBoxLayout()
        lay_src.addWidget(QLabel("PDFs:"))
        self.txt_wm_in = QLineEdit()
        btn_wm_in = QPushButton("Explorar...")
        self.wm_files = []
        btn_wm_in.clicked.connect(lambda: self.select_multi_pdf_wm(self.txt_wm_in, self.wm_files))
        lay_src.addWidget(self.txt_wm_in)
        lay_src.addWidget(btn_wm_in)
        lay_grp.addLayout(lay_src)
        
        lay_text = QVBoxLayout()
        lay_text.addWidget(QLabel("Texto del Sello:"))
        self.txt_wm_text = QLineEdit("APROBADO PARA CONSTRUCCIÓN")
        self.txt_wm_text.textChanged.connect(self.update_wm_preview)
        lay_text.addWidget(self.txt_wm_text)
        lay_grp.addLayout(lay_text)
        
        lay_color = QHBoxLayout()
        lay_color.addWidget(QLabel("Color:"))
        self.cmb_wm_color = QComboBox()
        self.cmb_wm_color.addItems(["Rojo", "Azul", "Negro", "Gris"])
        self.cmb_wm_color.currentTextChanged.connect(self.update_wm_preview)
        lay_color.addWidget(self.cmb_wm_color)
        lay_grp.addLayout(lay_color)
        
        lay_p1 = QHBoxLayout()
        lay_p1.addWidget(QLabel("Tamaño:"))
        self.spin_wm_size = QSpinBox()
        self.spin_wm_size.setRange(10, 200)
        self.spin_wm_size.setValue(60)
        self.spin_wm_size.valueChanged.connect(self.update_wm_preview)
        lay_p1.addWidget(self.spin_wm_size)
        
        lay_p1.addWidget(QLabel("Giro(°):"))
        self.spin_wm_rot = QSpinBox()
        self.spin_wm_rot.setRange(-180, 180)
        self.spin_wm_rot.setValue(45)
        self.spin_wm_rot.valueChanged.connect(self.update_wm_preview)
        lay_p1.addWidget(self.spin_wm_rot)
        lay_grp.addLayout(lay_p1)
        
        lay_p2 = QHBoxLayout()
        lay_p2.addWidget(QLabel("Opacidad (%):"))
        self.spin_wm_alpha = QSpinBox()
        self.spin_wm_alpha.setRange(5, 100)
        self.spin_wm_alpha.setValue(25)
        self.spin_wm_alpha.valueChanged.connect(self.update_wm_preview)
        lay_p2.addWidget(self.spin_wm_alpha)
        lay_p2.addStretch()
        lay_grp.addLayout(lay_p2)
        
        lay_dest = QHBoxLayout()
        lay_dest.addWidget(QLabel("Salida:"))
        self.txt_wm_out = QLineEdit()
        btn_wm_out = QPushButton("Explorar...")
        btn_wm_out.clicked.connect(lambda: self.select_output_dir_generic(self.txt_wm_out))
        lay_dest.addWidget(self.txt_wm_out)
        lay_dest.addWidget(btn_wm_out)
        lay_grp.addLayout(lay_dest)
        
        btn_action = QPushButton("APLICAR SELLO")
        btn_action.setStyleSheet("background-color: #0277bd; color: white; font-weight: bold; font-size: 14px; padding: 10px;")
        btn_action.clicked.connect(self.pdf_watermark_action)
        lay_grp.addWidget(btn_action)
        
        lay_ctrl.addWidget(grp)
        lay_ctrl.addStretch()
        lay_main.addWidget(panel_ctrl, stretch=1)
        
        # Panel Derecho: Vista Previa
        grp_preview = QGroupBox("Vista Previa (Sello en 1era página)")
        lay_preview = QVBoxLayout(grp_preview)
        
        # --- Navegación ---
        lay_nav = QHBoxLayout()
        self.btn_wm_prev = QPushButton("◀ Anterior")
        self.btn_wm_next = QPushButton("Siguiente ▶")
        self.lbl_wm_nav = QLabel("0 / 0")
        self.lbl_wm_nav.setAlignment(Qt.AlignCenter)
        lay_nav.addWidget(self.btn_wm_prev)
        lay_nav.addWidget(self.lbl_wm_nav)
        lay_nav.addWidget(self.btn_wm_next)
        lay_preview.addLayout(lay_nav)
        
        self.btn_wm_prev.clicked.connect(self.wm_prev_pdf)
        self.btn_wm_next.clicked.connect(self.wm_next_pdf)
        self.current_wm_idx = 0
        
        self.view_wm_preview = ZoomGraphicsView()
        self.view_wm_preview.setStyleSheet("background-color: #2b2b2b; border: 1px solid #555;")
        lay_preview.addWidget(self.view_wm_preview)
        
        # Pista para el usuario
        lbl_hint = QLabel("Use Ctrl + Rueda ratón para Zoom. Arrastre para mover.")
        lbl_hint.setStyleSheet("color: #888; font-size: 11px;")
        lay_preview.addWidget(lbl_hint)
        
        lay_main.addWidget(grp_preview, stretch=2)

    def setup_pdf_password_ui(self, widget):
        lay = QVBoxLayout(widget)
        grp = QGroupBox("Proteger Planos con Contraseña")
        lay_grp = QVBoxLayout(grp)
        
        lay_src = QHBoxLayout()
        lay_src.addWidget(QLabel("Archivo(s) origen:"))
        self.txt_pw_in = QLineEdit()
        btn_pw_in = QPushButton("Explorar...")
        self.pw_files = []
        btn_pw_in.clicked.connect(lambda: self.select_multi_pdf(self.txt_pw_in, self.pw_files))
        lay_src.addWidget(self.txt_pw_in)
        lay_src.addWidget(btn_pw_in)
        lay_grp.addLayout(lay_src)
        
        lay_pw = QHBoxLayout()
        lay_pw.addWidget(QLabel("Contraseña:"))
        self.txt_pw_pass = QLineEdit()
        self.txt_pw_pass.setEchoMode(QLineEdit.Password)
        lay_pw.addWidget(self.txt_pw_pass)
        lay_grp.addLayout(lay_pw)
        
        lay_dest = QHBoxLayout()
        lay_dest.addWidget(QLabel("Carpeta Salida:"))
        self.txt_pw_out = QLineEdit()
        btn_pw_out = QPushButton("Explorar...")
        btn_pw_out.clicked.connect(lambda: self.select_output_dir_generic(self.txt_pw_out))
        lay_dest.addWidget(self.txt_pw_out)
        lay_dest.addWidget(btn_pw_out)
        lay_grp.addLayout(lay_dest)
        
        btn_action = QPushButton("ENCRIPTAR PDFs")
        btn_action.setStyleSheet("background-color: #0277bd; color: white; font-weight: bold; font-size: 14px; padding: 10px;")
        btn_action.clicked.connect(self.pdf_password_action)
        lay_grp.addWidget(btn_action)
        
        lay.addWidget(grp)
        lay.addStretch()

    def setup_pdf_rotate_ui(self, widget):
        lay_main = QHBoxLayout(widget)
        
        # Panel Izquierdo: Controles
        panel_ctrl = QWidget()
        lay_ctrl = QVBoxLayout(panel_ctrl)
        lay_ctrl.setContentsMargins(0,0,10,0)
        
        grp = QGroupBox("Rotación Inteligente de PDFs")
        lay_grp = QVBoxLayout(grp)
        
        lay_src = QHBoxLayout()
        lay_src.addWidget(QLabel("Archivo(s) origen:"))
        self.txt_rot_in = QLineEdit()
        btn_rot_in = QPushButton("Explorar...")
        self.rot_files = []
        btn_rot_in.clicked.connect(lambda: self.select_multi_pdf_rot(self.txt_rot_in, self.rot_files))
        lay_src.addWidget(self.txt_rot_in)
        lay_src.addWidget(btn_rot_in)
        lay_grp.addLayout(lay_src)
        
        self.bg_rot = QButtonGroup()
        lay_rot = QHBoxLayout()
        rb_orig = QRadioButton("Sin Rotar (Original)")
        rb90 = QRadioButton("Rotar 90° Derecha")
        rb180 = QRadioButton("Rotar 180° Invertido")
        rb270 = QRadioButton("Rotar 90° Izquierda")
        
        # Establecer la original como activa por defecto
        rb_orig.setChecked(True)
        
        self.bg_rot.addButton(rb_orig, 0)
        self.bg_rot.addButton(rb90, 90)
        self.bg_rot.addButton(rb180, 180)
        self.bg_rot.addButton(rb270, 270)
        
        # Conectar cambio de rotación para actualizar preview
        self.bg_rot.buttonClicked.connect(self.update_rot_preview)
        
        lay_rot.addWidget(rb_orig)
        lay_rot.addWidget(rb90)
        lay_rot.addWidget(rb180)
        lay_rot.addWidget(rb270)
        lay_grp.addLayout(lay_rot)
        
        lay_dest = QHBoxLayout()
        lay_dest.addWidget(QLabel("Carpeta Salida:"))
        self.txt_rot_out = QLineEdit()
        btn_rot_out = QPushButton("Explorar...")
        btn_rot_out.clicked.connect(lambda: self.select_output_dir_generic(self.txt_rot_out))
        lay_dest.addWidget(self.txt_rot_out)
        lay_dest.addWidget(btn_rot_out)
        lay_grp.addLayout(lay_dest)
        
        btn_action = QPushButton("ROTAR PDFs")
        btn_action.setStyleSheet("background-color: #0277bd; color: white; font-weight: bold; font-size: 14px; padding: 10px;")
        btn_action.clicked.connect(self.pdf_rotate_action)
        lay_grp.addWidget(btn_action)
        
        lay_ctrl.addWidget(grp)
        lay_ctrl.addStretch()
        lay_main.addWidget(panel_ctrl, stretch=1)
        
        # Panel Derecho: Vista Previa
        grp_preview = QGroupBox("Vista Previa (1era página)")
        lay_preview = QVBoxLayout(grp_preview)
        
        # --- Navegación ---
        lay_nav = QHBoxLayout()
        self.btn_rot_prev = QPushButton("◀ Anterior")
        self.btn_rot_next = QPushButton("Siguiente ▶")
        self.lbl_rot_nav = QLabel("0 / 0")
        self.lbl_rot_nav.setAlignment(Qt.AlignCenter)
        lay_nav.addWidget(self.btn_rot_prev)
        lay_nav.addWidget(self.lbl_rot_nav)
        lay_nav.addWidget(self.btn_rot_next)
        lay_preview.addLayout(lay_nav)
        
        self.btn_rot_prev.clicked.connect(self.rot_prev_pdf)
        self.btn_rot_next.clicked.connect(self.rot_next_pdf)
        self.current_rot_idx = 0
        
        self.view_rot_preview = ZoomGraphicsView()
        self.view_rot_preview.setStyleSheet("background-color: #2b2b2b; border: 1px solid #555;")
        lay_preview.addWidget(self.view_rot_preview)
        
        # Pista para el usuario
        lbl_hint = QLabel("Use Ctrl + Rueda ratón para Zoom. Arrastre para mover.")
        lbl_hint.setStyleSheet("color: #888; font-size: 11px;")
        lay_preview.addWidget(lbl_hint)
        
        lay_main.addWidget(grp_preview, stretch=2)

    def select_single_pdf(self, line_edit):
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar PDF", self.settings.value("last_pdf_dir",""), "PDF Document (*.pdf)")
        if file:
            line_edit.setText(file)
            self.settings.setValue("last_pdf_dir", os.path.dirname(file))
            
    def select_multi_pdf(self, line_edit, target_list):
        files, _ = QFileDialog.getOpenFileNames(self, "Seleccionar PDFs", self.settings.value("last_pdf_dir",""), "PDF Document (*.pdf)")
        if files:
            target_list.clear()
            target_list.extend(files)
            if len(files) == 1:
                line_edit.setText(files[0])
            else:
                line_edit.setText(f"({len(files)} archivos seleccionados)")
            self.settings.setValue("last_pdf_dir", os.path.dirname(files[0]))
            
    def select_multi_pdf_rot(self, line_edit, target_list):
        files, _ = QFileDialog.getOpenFileNames(self, "Seleccionar PDFs", self.settings.value("last_pdf_dir",""), "PDF Document (*.pdf)")
        if files:
            target_list.clear()
            target_list.extend(files)
            if len(files) == 1:
                line_edit.setText(files[0])
            else:
                line_edit.setText(f"({len(files)} archivos seleccionados)")
            self.settings.setValue("last_pdf_dir", os.path.dirname(files[0]))
            self.current_rot_idx = 0
            # Actualizar preview de rotación
            self.update_rot_preview()

    def rot_prev_pdf(self):
        if not hasattr(self, 'rot_files') or not self.rot_files: return
        if getattr(self, 'current_rot_idx', 0) > 0:
            self.current_rot_idx -= 1
            self.update_rot_preview()

    def rot_next_pdf(self):
        if not hasattr(self, 'rot_files') or not self.rot_files: return
        if getattr(self, 'current_rot_idx', 0) < len(self.rot_files) - 1:
            self.current_rot_idx += 1
            self.update_rot_preview()

    def update_rot_preview(self):
        if not hasattr(self, 'rot_files') or not self.rot_files:
            self.view_rot_preview.custom_scene.clear()
            self.lbl_rot_nav.setText("0 / 0")
            return
            
        self.lbl_rot_nav.setText(f"{getattr(self, 'current_rot_idx', 0) + 1} / {len(self.rot_files)}")
            
        pdf_path = self.rot_files[getattr(self, 'current_rot_idx', 0)]
        if not os.path.exists(pdf_path): return
        
        try:
            doc = fitz.open(pdf_path)
            page = doc[0]
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            imgType = QImage.Format_RGB888
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, imgType)
            
            # Rotar la imagen de acuerdo a la selección del radio button
            transform = QTransform().rotate(self.bg_rot.checkedId())
            rotated_qimg = qimg.transformed(transform, Qt.SmoothTransformation)
            
            doc.close()
            
            pixmap = QPixmap.fromImage(rotated_qimg)
            self.view_rot_preview.custom_scene.clear()
            self.view_rot_preview.resetTransform()
            
            p_item = self.view_rot_preview.custom_scene.addPixmap(pixmap)
            self.view_rot_preview.setSceneRect(QRectF(pixmap.rect()))
            self.view_rot_preview.fitInView(p_item, Qt.KeepAspectRatio)
            
        except Exception as e:
            self.view_rot_preview.custom_scene.clear()
            self.lbl_rot_nav.setText("Error")
            print(f"Error preview rotación: {str(e)}")

    def select_multi_pdf_wm(self, line_edit, target_list):
        files, _ = QFileDialog.getOpenFileNames(self, "Seleccionar PDFs", self.settings.value("last_pdf_dir",""), "PDF Document (*.pdf)")
        if files:
            target_list.clear()
            target_list.extend(files)
            if len(files) == 1:
                line_edit.setText(files[0])
            else:
                line_edit.setText(f"({len(files)} archivos seleccionados)")
            self.settings.setValue("last_pdf_dir", os.path.dirname(files[0]))
            self.current_wm_idx = 0
            # Actualizar preview
            self.update_wm_preview()

    def wm_prev_pdf(self):
        if not hasattr(self, 'wm_files') or not self.wm_files: return
        if getattr(self, 'current_wm_idx', 0) > 0:
            self.current_wm_idx -= 1
            self.update_wm_preview()

    def wm_next_pdf(self):
        if not hasattr(self, 'wm_files') or not self.wm_files: return
        if getattr(self, 'current_wm_idx', 0) < len(self.wm_files) - 1:
            self.current_wm_idx += 1
            self.update_wm_preview()

    def update_wm_preview(self):
        if not hasattr(self, 'wm_files') or not self.wm_files:
            self.view_wm_preview.custom_scene.clear()
            self.lbl_wm_nav.setText("0 / 0")
            return
            
        self.lbl_wm_nav.setText(f"{getattr(self, 'current_wm_idx', 0) + 1} / {len(self.wm_files)}")
            
        pdf_path = self.wm_files[getattr(self, 'current_wm_idx', 0)]
        if not os.path.exists(pdf_path): return
        
        try:
            # Renderizar 1era hoja con PyMuPDF
            doc = fitz.open(pdf_path)
            page = doc[0]
            # Usar una matriz de zoom para que se vea legible pero rápido
            mat = fitz.Matrix(2.0, 2.0)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # Convertir a QImage
            imgType = QImage.Format_RGB888
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, imgType)
            
            # Dibujar marca de agua encima al estilo ReportLab pero con QPainter 
            # (para la aproximación visual)
            painter = QPainter(qimg)
            painter.setRenderHint(QPainter.Antialiasing)
            
            # Mapeo color
            alpha_val = int(self.spin_wm_alpha.value() * 2.55)
            c_str = self.cmb_wm_color.currentText()
            if c_str == "Rojo": col = QColor(255, 0, 0, alpha_val)
            elif c_str == "Azul": col = QColor(0, 0, 255, alpha_val)
            elif c_str == "Gris": col = QColor(128, 128, 128, alpha_val)
            else: col = QColor(0, 0, 0, alpha_val)
            
            painter.setPen(col)
            
            # El tamaño visual en PyMuPDF (escala 2.0x) vs Puntos ReportLab difiere un poco, 
            # ajustamos proporciones visuales con factor ~2.1
            v_size = int(self.spin_wm_size.value() * 2.1)
            font = QFont("Arial", v_size, QFont.Bold)
            painter.setFont(font)
            
            text = self.txt_wm_text.text()
            rect = painter.fontMetrics().boundingRect(text)
            
            # Translación y Rotación
            painter.translate(pix.width / 2, pix.height / 2)
            # En Qt la rotación es horaria positiva (alrevés que RL si RL es anti-horario, RL es anti).
            painter.rotate(-self.spin_wm_rot.value())
            
            # Dibujar texto al centro
            painter.drawText(-rect.width() / 2, rect.height() / 4, text)
            painter.end()
            
            doc.close()
            
            # Mostrar escalado al QGraphicsScene
            pixmap = QPixmap.fromImage(qimg)
            self.view_wm_preview.custom_scene.clear()
            self.view_wm_preview.resetTransform()
            
            p_item = self.view_wm_preview.custom_scene.addPixmap(pixmap)
            self.view_wm_preview.setSceneRect(QRectF(pixmap.rect()))
            self.view_wm_preview.fitInView(p_item, Qt.KeepAspectRatio)
            
        except Exception as e:
            self.view_wm_preview.custom_scene.clear()
            self.lbl_wm_nav.setText("Error")
            print(f"Error preview: {str(e)}")

    def resizeEvent(self, event):
        super().resizeEvent(event)

    def select_output_dir_generic(self, line_edit):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Destino", self.settings.value("last_output_dir", ""))
        if folder:
            line_edit.setText(folder)
            self.settings.setValue("last_output_dir", folder)

    def toggle_scale_inputs(self, checked):
        self.txt_scale_num.setEnabled(not checked)
        self.txt_scale_den.setEnabled(not checked)

    # --- LÓGICA DE MEMORIA (RUTAS RECIENTES) ---
    def load_memory(self):
        """Carga las rutas de la última sesión."""
        # CTB
        last_ctb = self.settings.value("last_ctb_path", "")
        if last_ctb and os.path.exists(last_ctb):
            self.txt_ctb.setText(last_ctb)
            
        # Directorio de Salida
        last_output = self.settings.value("last_output_dir", "")
        if last_output and os.path.exists(last_output):
            self.txt_out.setText(last_output)

    # --- DIÁLOGOS DE SELECCIÓN ---
    def select_dwg_files(self):
        last_dwg_dir = self.settings.value("last_dwg_dir", "")
        files, _ = QFileDialog.getOpenFileNames(self, "Seleccionar DWGs", last_dwg_dir, "AutoCAD Drawing (*.dwg)")
        if files:
            self.dwg_files = files
            self.lbl_dwg_count.setText(f"{len(self.dwg_files)} archivo(s) seleccionado(s)")
            # Guarda la carpeta de donde vinieron estos DWGs para abrir por defecto ahí la próxima vez
            self.settings.setValue("last_dwg_dir", os.path.dirname(self.dwg_files[0]))

    def select_ctb_file(self):
        last_ctb_dir = self.settings.value("last_ctb_path", "")
        # Si había ruta previa, la usamos como base del explorador
        start_dir = os.path.dirname(last_ctb_dir) if last_ctb_dir else ""
        file, _ = QFileDialog.getOpenFileName(self, "Seleccionar CTB", start_dir, "AutoCAD Plot Style (*.ctb)")
        if file:
            self.txt_ctb.setText(file)
            self.settings.setValue("last_ctb_path", file) # Guardar en memoria

    def select_output_dir(self):
        last_out_dir = self.settings.value("last_output_dir", "")
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta Destino", last_out_dir)
        if folder:
            self.txt_out.setText(folder)
            self.settings.setValue("last_output_dir", folder) # Guardar en memoria

    # --- EVENTOS DE INTERFAZ (PDF Tab) ---
    def pdf_add_files(self):
        last_pdf_dir = self.settings.value("last_pdf_dir", self.settings.value("last_output_dir", ""))
        files, _ = QFileDialog.getOpenFileNames(self, "Seleccionar PDFs", last_pdf_dir, "PDF Document (*.pdf)")
        if files:
            for f in files:
                # Evitar duplicados exactos visuales
                items = self.list_pdfs.findItems(f, Qt.MatchExactly)
                if not items:
                    self.list_pdfs.addItem(f)
            
            self.settings.setValue("last_pdf_dir", os.path.dirname(files[0]))
            
            # Autocompletar ruta de salida sugerida si está vacía
            if not self.txt_pdf_out.text():
                sug_path = os.path.join(os.path.dirname(files[0]), "Planos_Unidos.pdf")
                self.txt_pdf_out.setText(sug_path)

    def pdf_remove_selected(self):
        for item in self.list_pdfs.selectedItems():
            self.list_pdfs.takeItem(self.list_pdfs.row(item))
            
    def pdf_select_output(self):
        last_dir = os.path.dirname(self.txt_pdf_out.text()) if self.txt_pdf_out.text() else self.settings.value("last_pdf_dir", "")
        file, _ = QFileDialog.getSaveFileName(self, "Guardar PDF Consolidado", last_dir, "PDF Document (*.pdf)")
        if file:
            self.txt_pdf_out.setText(file)

    def pdf_merge_action(self):
        if self.list_pdfs.count() < 2:
            QMessageBox.warning(self, "Aviso", "Añada por lo menos 2 archivos PDF para unir.")
            return
            
        out_path = self.txt_pdf_out.text().strip()
        if not out_path:
            QMessageBox.warning(self, "Aviso", "Especifique el archivo destino.")
            return
            
        try:
            merger = PyPDF2.PdfMerger()
            
            # Construir la lista en el orden visual actual
            for i in range(self.list_pdfs.count()):
                pdf_path = self.list_pdfs.item(i).text()
                merger.append(pdf_path)
                
            merger.write(out_path)
            merger.close()
            
            QMessageBox.information(self, "Éxito", f"Se han unido {self.list_pdfs.count()} PDFs correctamente en:\n\n{out_path}")
            
            # Abrir carpeta del archivo o archivo en sí (Windows)
            try:
                os.startfile(os.path.dirname(out_path))
            except Exception:
                pass
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Ocurrió un problema al unir los PDFs:\n\n{str(e)}")

    def pdf_split_action(self):
        in_path = self.txt_split_in.text().strip()
        out_dir = self.txt_split_out.text().strip()
        
        if not os.path.exists(in_path) or not os.path.exists(out_dir):
            QMessageBox.warning(self, "Aviso", "Valide que el archivo de origen y la carpeta destino existan.")
            return
            
        try:
            reader = PyPDF2.PdfReader(in_path)
            total = len(reader.pages)
            
            if self.chk_split_all.isChecked():
                for i in range(total):
                    writer = PyPDF2.PdfWriter()
                    writer.add_page(reader.pages[i])
                    out_path = os.path.join(out_dir, f"{os.path.basename(in_path).replace('.pdf','')}_Hoja_{i+1}.pdf")
                    with open(out_path, "wb") as fh:
                        writer.write(fh)
                QMessageBox.information(self, "Éxito", f"Se extrajeron {total} páginas individuales con éxito.")
                
                try:
                    os.startfile(out_dir)
                except Exception:
                    pass
            else:
                f_ini = self.spin_split_from.value() - 1
                f_fin = self.spin_split_to.value() - 1
                if f_ini < 0 or f_fin >= total or f_ini > f_fin:
                    QMessageBox.warning(self, "Error", "El rango de páginas ingresado es inválido.")
                    return
                
                writer = PyPDF2.PdfWriter()
                for i in range(f_ini, f_fin + 1):
                    writer.add_page(reader.pages[i])
                
                out_path = os.path.join(out_dir, f"{os.path.basename(in_path).replace('.pdf','')}_Extracto.pdf")
                with open(out_path, "wb") as fh:
                    writer.write(fh)
                QMessageBox.information(self, "Éxito", f"Se extrajo el rango de páginas exitosamente en:\n{out_path}")
                
                try:
                    os.startfile(out_dir)
                except Exception:
                    pass
                
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def pdf_watermark_action(self):
        if not self.wm_files: return
        out_dir = self.txt_wm_out.text()
        if not os.path.exists(out_dir): return
        
        text = self.txt_wm_text.text()
        
        # Color mapping a RGB Float para ReportLab
        c_map = {"Rojo": (1,0,0), "Negro": (0,0,0), "Azul": (0,0,1), "Gris": (0.5, 0.5, 0.5)}
        rgb = c_map.get(self.cmb_wm_color.currentText(), (1,0,0))
        
        # Obtener valores personalizados
        f_size = self.spin_wm_size.value()
        f_alpha = self.spin_wm_alpha.value() / 100.0
        f_rot = self.spin_wm_rot.value()
        
        try:
            for f in self.wm_files:
                reader = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()
                
                for page in reader.pages:
                    mb = page.mediabox
                    packet = io.BytesIO()
                    can = canvas.Canvas(packet, pagesize=(mb.right, mb.top))
                    
                    can.setFillColorRGB(rgb[0], rgb[1], rgb[2], alpha=f_alpha)
                    can.setFont("Helvetica-Bold", f_size)
                    can.translate(float(mb.right)/2, float(mb.top)/2)
                    can.rotate(-f_rot)  # Invertimos rotación para igualar Qt
                    can.drawCentredString(0, 0, text)
                    can.save()
                    
                    packet.seek(0)
                    watermark = PyPDF2.PdfReader(packet)
                    page.merge_page(watermark.pages[0])
                    writer.add_page(page)
                    
                bn = os.path.basename(f)
                with open(os.path.join(out_dir, bn.replace(".pdf", "_sello.pdf")), "wb") as outF:
                    writer.write(outF)
            QMessageBox.information(self, "Éxito", "Se estamparon los sellos correctamente.")
            
            try:
                os.startfile(out_dir)
            except Exception:
                pass
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def pdf_password_action(self):
        if not self.pw_files: return
        out_dir = self.txt_pw_out.text()
        pw = self.txt_pw_pass.text()
        if not os.path.exists(out_dir) or not pw: return
        
        try:
            for f in self.pw_files:
                reader = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()
                for page in reader.pages:
                    writer.add_page(page)
                writer.encrypt(pw)
                bn = os.path.basename(f)
                with open(os.path.join(out_dir, bn.replace(".pdf", "_seguro.pdf")), "wb") as outF:
                    writer.write(outF)
            QMessageBox.information(self, "Éxito", "Se protegió el PDF con contraseña maestra (AES128).")
            
            try:
                os.startfile(out_dir)
            except Exception:
                pass
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def pdf_rotate_action(self):
        if not self.rot_files: return
        out_dir = self.txt_rot_out.text()
        if not os.path.exists(out_dir): return
        deg = self.bg_rot.checkedId()
        
        if deg == 0:
            QMessageBox.information(self, "Aviso", "Seleccionó 'Sin Rotar'. No se harán cambios.\nPor favor elija un ángulo de rotación si desea modificar el documento.")
            return
            
        try:
            for f in self.rot_files:
                reader = PyPDF2.PdfReader(f)
                writer = PyPDF2.PdfWriter()
                for page in reader.pages:
                    page.rotate(deg)
                    writer.add_page(page)
                bn = os.path.basename(f)
                with open(os.path.join(out_dir, bn.replace(".pdf", "_rotado.pdf")), "wb") as outF:
                    writer.write(outF)
            QMessageBox.information(self, "Éxito", "Se aplicó la rotación al nivel documento.")
            
            try:
                os.startfile(out_dir)
            except Exception:
                pass
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    # --- LÓGICA DE EJECUCIÓN DEL BATCH PLOTTER ---
    def append_log(self, text):
        self.txt_log.append(text)
        
    def worker_finished(self):
        self.btn_run.setEnabled(True)
        self.btn_dwg.setEnabled(True)
        QMessageBox.information(self, "Completado", "El proceso de conversión ha finalizado.")
        
    def worker_error(self, err_msg):
        QMessageBox.critical(self, "Error Fatal", err_msg)
        self.btn_run.setEnabled(True)
        self.btn_dwg.setEnabled(True)

    def update_progress(self, val):
        self.progress.setValue(val)

    def start_conversion(self):
        if not self.dwg_files:
            QMessageBox.warning(self, "Atención", "Seleccione al menos un archivo DWG.")
            return
        if not self.txt_out.text():
            QMessageBox.warning(self, "Atención", "Seleccione la carpeta de destino.")
            return

        # Recopilar la configuración elegida por el usuario
        plot_config = {
            "plotter": self.combo_plotter.currentText(),
            "paper_size": self.combo_paper.currentText(),
            "plot_area": self.combo_area.currentText(),
            "orientation": self.cmb_orientation.currentText(),
            "fit_to_paper": self.chk_fit.isChecked(),
            "scale_num": self.txt_scale_num.text(),
            "scale_den": self.txt_scale_den.text(),
            "center_plot": self.chk_center.isChecked(),
            "scale_lineweights": self.chk_scale_lineweights.isChecked()
        }

        self.btn_run.setEnabled(False)
        self.btn_dwg.setEnabled(False)
        self.progress.setMaximum(len(self.dwg_files))
        self.progress.setValue(0)
        self.txt_log.clear()
        
        # Iniciar Workers
        self.worker = PlotWorker(self.dwg_files, self.txt_out.text(), self.txt_ctb.text(), plot_config)
        self.worker.log_signal.connect(self.append_log)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.worker_finished)
        self.worker.error_signal.connect(self.worker_error)
        
        self.worker.start()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    # Aplicar Dark Theme profesional usando qdarkstyle
    app.setStyleSheet(qdarkstyle.load_stylesheet(qt_api='pyside6'))
    
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
