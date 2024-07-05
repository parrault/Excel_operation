import sys
from PySide6.QtWidgets import QDialog, QApplication, QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QHBoxLayout, QFileDialog, QCheckBox, QGridLayout, QInputDialog
from PySide6.QtCore import Qt
from excel_operations import excel_colonne
import json
from utils import get_resource_path
import os 

class ExcelSearchApp(QWidget):
    def __init__(self):
        super().__init__()
                               
        self.dossier_choisi = ""
        self.cellule_recherchee = ""
        self.options = []
        self.default_options = [
            {'label': 'Secteur', 'checked': False, 'cell': 'E2'},
            {'label': 'Sous-secteur', 'checked': False, 'cell': 'F2'},
            {'label': 'Cours de bourse', 'checked': False, 'cell': 'B5'},
            {'label': 'WACC', 'checked': False, 'cell': 'S144'},
            {'label': 'PGR', 'checked': False, 'cell': 'S145'},
            {'label': 'DCFvalue', 'checked': False, 'cell': 'U196'},
            {'label': 'Upside/Downside', 'checked': False, 'cell': 'U198'},
            {'label': 'Organic24', 'checked': False, 'cell': 'U50'},
            {'label': 'Organic25', 'checked': False, 'cell': 'U51'},
            {'label': 'Organic26', 'checked': False, 'cell': 'U52'},
            {'label': 'Organic27', 'checked': False, 'cell': 'U53'}
        ]
        self.checkbox_vars = []

        #user_dir = os.path.expanduser("~")
        user_dir = ""
        self.config_path = os.path.join(user_dir, 'config.json')
        if not os.path.exists(self.config_path):
            with open(self.config_path, 'w') as f:
                json.dump({'options': self.default_options, 'default': self.default_options}, f, indent=4)
       
        self.init_ui()
        self.load_config()
    

    def init_ui(self):
        layout = QVBoxLayout()

        # Dossier contenant les fichiers Excel
        dossier_layout = QHBoxLayout()
        self.dossier_label = QLabel("Dossier contenant les fichiers Excel:")
        self.dossier_input = QLineEdit()
        self.dossier_button = QPushButton("Choisir")
        self.dossier_button.clicked.connect(self.choisir_dossier)
        dossier_layout.addWidget(self.dossier_label)
        dossier_layout.addWidget(self.dossier_input)
        dossier_layout.addWidget(self.dossier_button)
        layout.addLayout(dossier_layout)

        # Colonnes voulues
        self.colonnes_label = QLabel("Colonnes voulues:")
        layout.addWidget(self.colonnes_label)

        self.grid_layout = QGridLayout()
        layout.addLayout(self.grid_layout)

        # Bouton pour générer l'Excel des colonnes voulues
        self.colonnes_button = QPushButton("Excel des colonnes voulues")
        self.colonnes_button.clicked.connect(self.excel_colonne)
        layout.addWidget(self.colonnes_button)

        
        


        self.add_checkbox_button = QPushButton("Ajouter une case à cocher")
        self.remove_checkbox_button = QPushButton("Supprimer la dernière case à cocher")
        self.add_checkbox_button.clicked.connect(self.add_checkbox)
        self.remove_checkbox_button.clicked.connect(self.remove_checkbox)
        layout.addWidget(self.add_checkbox_button)
        layout.addWidget(self.remove_checkbox_button)

        # Bouton pour réinitialiser les options
        self.reset_button = QPushButton("Réinitialiser les options")
        self.reset_button.clicked.connect(self.reset_options)
        layout.addWidget(self.reset_button)

        self.setLayout(layout)
        self.setWindowTitle("Recherche dans fichiers Excel")
        self.show()
    
    def load_config(self):
        try:
            with open(self.config_path, 'r') as file:
                config = json.load(file)
                self.options = config.get('options', [])
                #self.default_options = config.get('default', [])
        except (FileNotFoundError, json.JSONDecodeError):
            print("Configuration file not found or corrupted. Using defaults.")
            self.options = self.default_options
        self.update_checkboxes()
        
    def save_config(self):
        config = {
            'options': self.options,
            'default': self.default_options
        }
        with open(self.config_path, 'w') as file:
            json.dump(config, file, indent=4)
        print("Fichier de configuration mis à jour :", self.config_path)
        print("Données enregistrées :", config)
    
    def update_checkboxes(self):
        for checkbox in self.checkbox_vars:
            checkbox.setParent(None)
        self.checkbox_vars = []

        for i, option in enumerate(self.options):
            checkbox = QCheckBox(option['label'])
            checkbox.setChecked(option['checked'])
            checkbox.setProperty("cell", option['cell'])
            checkbox.stateChanged.connect(self.checkbox_state_changed)
            self.grid_layout.addWidget(checkbox, i // 3, i % 3)
            self.checkbox_vars.append(checkbox)
    
    def checkbox_state_changed(self):
        for i, checkbox in enumerate(self.checkbox_vars):
            self.options[i]['checked'] = checkbox.isChecked()
        self.save_config()

    def choisir_dossier(self):
        dossier = QFileDialog.getExistingDirectory(self, "Choisir le dossier")
        if dossier:
            self.dossier_input.setText(dossier)
            self.dossier_choisi = dossier

    def excel_colonne(self):
        excel_colonne(self.dossier_choisi, self.options)
    
    def add_checkbox(self):
        dialog = InputDialog(self)
        if dialog.exec() == QDialog.Accepted:
            label, cell = dialog.get_inputs()
            if label and cell:
                new_checkbox = QCheckBox(label)
                self.grid_layout.addWidget(new_checkbox, len(self.checkbox_vars) // 3, len(self.checkbox_vars) % 3)
                self.checkbox_vars.append(new_checkbox)
                self.options.append({"label": label, "checked": False, "cell": cell})
                self.save_config()  # Appel à la méthode de sauvegarde
    
    def remove_checkbox(self):
        if self.checkbox_vars:
            checkbox_to_remove = self.checkbox_vars.pop()
            checkbox_to_remove.setParent(None)
            checkbox_to_remove.deleteLater()
            self.options.pop()  # Remove the last option
            self.save_config()  # Save changes

    def reset_options(self):
        self.options = self.default_options.copy()
        self.save_config()
        self.update_checkboxes()

class InputDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)

        self.setWindowTitle("Ajouter une Option")
        layout = QVBoxLayout(self)

        # Ajouter des champs pour le libellé de l'option et le numéro de cellule
        self.label_input = QLineEdit(self)
        self.cell_input = QLineEdit(self)
        self.ok_button = QPushButton("OK", self)
        self.cancel_button = QPushButton("Annuler", self)

        layout.addWidget(QLabel("Libellé de la case à cocher:"))
        layout.addWidget(self.label_input)
        layout.addWidget(QLabel("Numéro de cellule associée:"))
        layout.addWidget(self.cell_input)
        layout.addWidget(self.ok_button)
        layout.addWidget(self.cancel_button)

        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)

    def get_inputs(self):
        return self.label_input.text(), self.cell_input.text()

def main():
    app = QApplication(sys.argv)
    ex = ExcelSearchApp()
    sys.exit(app.exec_())

