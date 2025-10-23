#pip install PySide6 pandas matplotlib openpyxl

import sys
import os
import platform
from pathlib import Path
from datetime import datetime
import pandas as pd

import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QTabWidget, QFileDialog,
    QMessageBox, QLabel, QScrollArea, QDialog, QLineEdit,
    QCheckBox, QFormLayout, QDialogButtonBox, QVBoxLayout,
    QHBoxLayout
)
from PySide6.QtGui import (
    QPixmap, QAction, QIcon, QDoubleValidator, QIntValidator
)
from PySide6.QtCore import Qt, Slot, QLocale


class PeakConfigDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("ピーク値計算 設定")
        
        self.config = None 
        
        self.locale = QLocale(QLocale.Language.English, QLocale.Country.UnitedStates)
        
        layout = QVBoxLayout(self)
        
        form_layout = QFormLayout()

        threshold_layout = QHBoxLayout()
        
        self.mantissa_edit = QLineEdit()
        self_mantissa_validator = QDoubleValidator()
        self_mantissa_validator.setLocale(self.locale)
        self_mantissa_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.mantissa_edit.setValidator(self_mantissa_validator)
        self.mantissa_edit.setPlaceholderText("2.5")
        
        self.exponent_edit = QLineEdit()
        self_exponent_validator = QIntValidator()
        self.exponent_edit.setValidator(self_exponent_validator)
        self.exponent_edit.setPlaceholderText("-6")
        
        threshold_layout.addWidget(QLabel("("))
        threshold_layout.addWidget(self.mantissa_edit)
        threshold_layout.addWidget(QLabel(") * 10 ^ ("))
        threshold_layout.addWidget(self.exponent_edit)
        threshold_layout.addWidget(QLabel(")"))
        
        form_layout.addRow("縦軸 (Y軸) 閾値 ( > V):", threshold_layout)

        self.time_edit = QLineEdit()
        self_time_validator = QDoubleValidator()
        self_time_validator.setLocale(self.locale)
        self_time_validator.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.time_edit.setValidator(self_time_validator)
        self.time_edit.setPlaceholderText("0.0 (空白時は0.0)")

        time_threshold_layout = QHBoxLayout()
        time_threshold_layout.addWidget(self.time_edit)
        time_threshold_layout.addWidget(QLabel("秒以降"))
        
        form_layout.addRow("横軸 (X軸) 閾値 ( > S):", time_threshold_layout)
        
        layout.addLayout(form_layout)
        
        self.calc_avg_check = QCheckBox("平均値を計算してグラフに表示する")
        self.calc_avg_check.setChecked(True) 
        layout.addWidget(self.calc_avg_check)
        
        self.save_peak_check = QCheckBox("ピークグラフを別途出力する")
        self.save_peak_check.setChecked(True) 
        layout.addWidget(self.save_peak_check)
        
        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(self.validate_and_accept)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    @Slot()
    def validate_and_accept(self):
        try:
            mantissa_str = self.mantissa_edit.text().strip()
            exponent_str = self.exponent_edit.text().strip()
            
            if not mantissa_str or not exponent_str:
                QMessageBox.warning(self, "入力エラー", "縦軸の閾値 (仮数部と指数部) は必須です。")
                return
                
            mantissa = float(mantissa_str)
            exponent = int(exponent_str)
            threshold_value = mantissa * (10 ** exponent)

            time_str = self.time_edit.text().strip()
            time_threshold = 0.0
            if time_str:
                time_threshold = float(time_str)

            self.config = {
                'threshold': threshold_value,      
                'time_threshold': time_threshold,  
                'calc_avg': self.calc_avg_check.isChecked(),
                'save_peak': self.save_peak_check.isChecked()
            }
            
            self.accept()
            
        except ValueError:
            QMessageBox.warning(self, "入力エラー", "有効な数値を入力してください。\n例: (2.5) * 10 ^ (-6)")
        except Exception as e:
            QMessageBox.critical(self, "エラー", f"予期せぬエラーが発生しました: {e}")

    def get_config(self) -> (dict | None):
        return self.config


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("松永研究室 ゾルゲル班用解析ツール")
        self.setGeometry(100, 100, 800, 600)

        home_dir = Path.home()
        
        doc_dir = home_dir / "Documents"
        self.config_path = doc_dir / "silane_ave_app.config"
        
        desktop_dir = home_dir / "Desktop"
        self.output_dir = desktop_dir / "shilane_result"

        try:
            self.output_dir.mkdir(parents=True, exist_ok=True)
        except OSError as e:
            QMessageBox.critical(self, "起動エラー", f"出力ディレクトリの作成に失敗しました: {e}")
            sys.exit(1) 

        self.plot_colors = ['red', 'blue', 'green', 'black', 'orange', 'pink', 'purple', 'cyan']
        self.color_index = 0 

        self.last_opened_file_path = None 
        self.default_open_dir = doc_dir    

        self.tabs = QTabWidget()
        self.tabs.setTabsClosable(True)
        self.tabs.tabCloseRequested.connect(self.close_tab)
        
        self.setCentralWidget(self.tabs)

        self.setup_menu()

        self.load_config()
        self.open_default_file()

    def setup_menu(self):
        menu_bar = self.menuBar()
        file_menu = menu_bar.addMenu("&ファイル (F)")

        action_open = QAction("&開く (O)...", self)
        action_open.triggered.connect(self.open_file)
        file_menu.addAction(action_open)
        
        file_menu.addSeparator()
        action_peak = QAction("ピーク値平均計算 (&P)...", self)
        action_peak.triggered.connect(self.open_peak_config)
        file_menu.addAction(action_peak)

        file_menu.addSeparator()

        action_close = QAction("&閉じる (C)", self)
        action_close.triggered.connect(self.close_current_tab)
        file_menu.addAction(action_close)

        action_close_all = QAction("全て閉じる (&A)", self)
        action_close_all.triggered.connect(self.close_all_tabs)
        file_menu.addAction(action_close_all)

    @Slot()
    def open_peak_config(self):
        current_widget = self.tabs.currentWidget()
        if not current_widget:
            QMessageBox.information(self, "情報", "対象のタブが開かれていません。")
            return

        df = current_widget.property("dataframe")
        filename = current_widget.property("filename")
        color = current_widget.property("plot_color")
        
        if df is None or filename is None or color is None:
            QMessageBox.warning(self, "エラー", "タブのデータが破損しているか、見つかりません。")
            return

        dialog = PeakConfigDialog(self)
        result = dialog.exec() 

        if result == QDialog.DialogCode.Accepted:
            peak_config = dialog.get_config()
            if not peak_config:
                return 

            print(f"ピーク設定を適用: {peak_config}")
            
            new_png_path = self.create_graph(df, filename, color, peak_config)
            
            if new_png_path:
                try:
                    image_label = current_widget.widget() 
                    if not isinstance(image_label, QLabel):
                            raise ValueError("タブの構造が予期したものではありません。")
                    
                    pixmap = QPixmap(str(new_png_path))
                    if pixmap.isNull():
                            raise ValueError(f"更新された画像ファイルの読み込みに失敗しました: {new_png_path}")

                    image_label.setPixmap(pixmap)
                
                except Exception as e:
                        QMessageBox.critical(self, "タブ更新エラー", f"グラフの表示更新に失敗しました:\n{e}")

    def load_config(self):
        try:
            if self.config_path.exists():
                path_str = self.config_path.read_text(encoding='utf-8').strip()
                if path_str:
                    config_path_obj = Path(path_str)
                    
                    if config_path_obj.is_file():
                        self.last_opened_file_path = config_path_obj
                        self.default_open_dir = config_path_obj.parent
                    elif config_path_obj.is_dir():
                        self.default_open_dir = config_path_obj
                        self.last_opened_file_path = None
                    else:
                        self.default_open_dir = Path.home() / "Documents"
                        self.last_opened_file_path = None
                else:
                    self.default_open_dir = Path.home() / "Documents"
                    self.last_opened_file_path = None
            else:
                self.default_open_dir = Path.home() / "Documents"
                self.last_opened_file_path = None
        except Exception as e:
            print(f"設定ファイルの読み込みエラー: {e}")
            self.default_open_dir = Path.home() / "Documents"
            self.last_opened_file_path = None

    def save_config(self):
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            
            path_to_save = None
            if self.last_opened_file_path and self.last_opened_file_path.is_file():
                path_to_save = self.last_opened_file_path
            elif self.default_open_dir and self.default_open_dir.is_dir():
                path_to_save = self.default_open_dir
            else:
                path_to_save = Path.home() / "Documents"

            self.config_path.write_text(str(path_to_save), encoding='utf-8')
            
        except Exception as e:
            print(f"設定ファイルの保存エラー: {e}")

    def open_default_file(self):
        if self.last_opened_file_path and self.last_opened_file_path.is_file():
            print(f"デフォルトファイルを開きます: {self.last_opened_file_path}")
            self.process_file(str(self.last_opened_file_path))
        else:
            print("デフォルトファイル、または有効なパスが設定されていません。")

    @Slot()
    def open_file(self):
        file_filter = "データファイル (*.xlsx *.csv);;全てのファイル (*.*)"
        
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "ファイルを開く",
            str(self.default_open_dir),
            file_filter
        )

        if file_path:
            file_path_obj = Path(file_path)
            self.last_opened_file_path = file_path_obj    
            self.default_open_dir = file_path_obj.parent    
            
            self.process_file(file_path)

    def process_file(self, file_path: str):
        try:
            file_path_obj = Path(file_path)
            filename = file_path_obj.name
            df = None

            if file_path_obj.suffix == '.xlsx':
                df = pd.read_excel(file_path, header=0, usecols=[1, 2])
            elif file_path_obj.suffix == '.csv':
                df = pd.read_csv(file_path, header=0, usecols=[1, 2])
            else:
                raise ValueError("サポートされていないファイル形式です (.xlsx または .csv のみ)。")

            if df.empty or len(df.columns) < 2:
                raise ValueError("ファイルから有効なデータ（2列）を読み込めませんでした。")

            color = self.plot_colors[self.color_index % len(self.plot_colors)]
            self.color_index += 1

            saved_png_path = self.create_graph(df, filename, color) 

            if saved_png_path:
                self.add_tab(saved_png_path, filename, df, color)

        except Exception as e:
            QMessageBox.critical(self, "ファイル処理エラー", f"ファイルの処理に失敗しました:\n{file_path}\n\n詳細: {e}")

    def create_graph(self, df: pd.DataFrame, filename: str, color: str, peak_config: dict = None) -> (Path | None):
        try:
            timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
            basename = Path(filename).stem
            
            col_x_name = df.columns[0] 
            col_y_name = df.columns[1] 
            
            average_current = None
            peak_df = None
            threshold_v = 0 
            threshold_t = 0.0 
            
            if peak_config:
                threshold_v = peak_config.get('threshold', 0)
                threshold_t = peak_config.get('time_threshold', 0.0)
                
                peak_df = df[
                    (df[col_y_name] > threshold_v) &    
                    (df[col_x_name] > threshold_t)
                ]
                
                if not peak_df.empty and peak_config.get('calc_avg', False):
                    average_current = peak_df[col_y_name].mean()

            fig, ax = plt.subplots()
            save_path = None 
            
            try:
                fig.set_facecolor('white')
                ax.set_facecolor('white')

                ax.plot(df[col_x_name], df[col_y_name], color=color)

                ax.set_xlabel(col_x_name)
                ax.set_ylabel(col_y_name)
                ax.set_ylim(bottom=0)
                ax.grid(False)

                if average_current is not None:
                    text_str = (
                        f"Peak Average: {average_current:.3e} A\n"
                        f"(Y > {threshold_v:.3e} A)\n"
                        f"(X > {threshold_t:.2f} s)"
                    )
                    ax.text(0.05, 0.95, text_str,    
                            transform=ax.transAxes, 
                            fontsize=9,    
                            verticalalignment='top',    
                            bbox=dict(boxstyle='round,pad=0.3', facecolor='white', alpha=0.7))

                output_filename = f"{timestamp}_shilane_{basename}.png"
                save_path = self.output_dir / output_filename
                plt.savefig(str(save_path), facecolor='white', bbox_inches='tight')
                print(f"グラフを保存しました: {save_path}")
                
            finally:
                plt.close(fig)

            if peak_df is not None and not peak_df.empty and peak_config.get('save_peak', False):
                fig_peak, ax_peak = plt.subplots()
                try:
                    fig_peak.set_facecolor('white')
                    ax_peak.set_facecolor('white')

                    ax_peak.plot(peak_df[col_x_name], peak_df[col_y_name], color='magenta')
                    
                    ax_peak.set_xlabel(col_x_name)
                    ax_peak.set_ylabel(col_y_name)
                    ax_peak.set_ylim(bottom=0.0)    
                    
                    ax_peak.set_xlim(left=0.0)    
                    
                    ax_peak.grid(False)

                    peak_output_filename = f"{timestamp}_shilane_{basename}_peak.png"
                    peak_save_path = self.output_dir / peak_output_filename
                    plt.savefig(str(peak_save_path), facecolor='white', bbox_inches='tight')
                    print(f"ピークグラフを保存しました: {peak_save_path}")

                finally:
                    plt.close(fig_peak)

            return save_path

        except Exception as e:
            QMessageBox.critical(self, "グラフ作成エラー", f"グラフの作成または保存に失敗しました:\n{e}")
            return None

    def add_tab(self, image_path: Path, tab_name: str, df: pd.DataFrame, color: str):
        try:
            scroll_area = QScrollArea()
            scroll_area.setWidgetResizable(True)
            
            scroll_area.setProperty("dataframe", df)
            scroll_area.setProperty("filename", tab_name)
            scroll_area.setProperty("plot_color", color)
            
            image_label = QLabel()
            pixmap = QPixmap(str(image_path))
            
            if pixmap.isNull():
                    raise ValueError(f"画像ファイルの読み込みに失敗しました: {image_path}")
            
            image_label.setPixmap(pixmap)
            image_label.setAlignment(Qt.AlignmentFlag.AlignCenter) 

            scroll_area.setWidget(image_label)
            new_tab_index = self.tabs.addTab(scroll_area, tab_name)
            self.tabs.setCurrentIndex(new_tab_index)

        except Exception as e:
            QMessageBox.critical(self, "タブ追加エラー", f"グラフの表示タブ追加に失敗しました:\n{e}")

    @Slot(int)
    def close_tab(self, index: int):
        widget = self.tabs.widget(index)
        if widget:
            widget.deleteLater()
        self.tabs.removeTab(index)

    @Slot()
    def close_current_tab(self):
        current_index = self.tabs.currentIndex()
        if current_index != -1: 
            self.close_tab(current_index)

    @Slot()
    def close_all_tabs(self):
        while self.tabs.count() > 0:
            self.close_tab(0)

    def closeEvent(self, event):
        print("アプリケーションを終了します。設定を保存中...")
        self.save_config()
        event.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())