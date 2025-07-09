import sys
import random
import PySide6.QtCore
from PySide6 import QtCore, QtWidgets, QtGui
from docx import Document
from autoBilan import *
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QLineEdit, QVBoxLayout, QGridLayout

print(PySide6.__version__)

class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.button = QtWidgets.QPushButton("Génerer le bilan")

        # Champs
        self.champs_nom = QtWidgets.QLineEdit()
        self.champs_prenom = QtWidgets.QLineEdit()
        self.champs_date_naiss = QtWidgets.QLineEdit()
        self.champs_age = QtWidgets.QLineEdit()
        self.champs_lat = QtWidgets.QLineEdit()
        self.champs_date = QtWidgets.QLineEdit()
        self.champs_ens_cv = QtWidgets.QLineEdit()
        self.champs_ens_v = QtWidgets.QLineEdit()
        self.champs_ens_rf = QtWidgets.QLineEdit()
        self.champs_ens_mdt = QtWidgets.QLineEdit()
        self.champs_ens_vt = QtWidgets.QLineEdit()
        self.champs_ens_et = QtWidgets.QLineEdit()
        self.champs_nc_cv = QtWidgets.QLineEdit()
        self.champs_nc_v = QtWidgets.QLineEdit()
        self.champs_nc_rf = QtWidgets.QLineEdit()
        self.champs_nc_mdt = QtWidgets.QLineEdit()
        self.champs_nc_vt = QtWidgets.QLineEdit()
        self.champs_nc_et = QtWidgets.QLineEdit()
        self.champs_rp_cv = QtWidgets.QLineEdit()
        self.champs_rp_v = QtWidgets.QLineEdit()
        self.champs_rp_rf = QtWidgets.QLineEdit()
        self.champs_rp_mdt = QtWidgets.QLineEdit()
        self.champs_rp_vt = QtWidgets.QLineEdit()
        self.champs_rp_et = QtWidgets.QLineEdit()
        self.champs_idc_cv = QtWidgets.QLineEdit()
        self.champs_idc_v = QtWidgets.QLineEdit()
        self.champs_idc_rf = QtWidgets.QLineEdit()
        self.champs_idc_mdt = QtWidgets.QLineEdit()
        self.champs_idc_vt = QtWidgets.QLineEdit()
        self.champs_idc_et = QtWidgets.QLineEdit()

        # Labels
        self.title_nom = QtWidgets.QLabel("Nom du patient")
        self.title_prenom = QtWidgets.QLabel("Prénom du patient")
        self.title_date_naiss = QtWidgets.QLabel("Date de naissance du patient")
        self.title_age = QtWidgets.QLabel("Age du patient")
        self.title_lat = QtWidgets.QLabel("Latitude du patient")
        self.title_date = QtWidgets.QLabel("Date du bilan")
        self.title_ens_cv = QtWidgets.QLabel("Ensemble des Notes Standard - Compréhension Verbale")
        self.title_ens_v = QtWidgets.QLabel("Ensemble des Notes Standard - Visuospatial")
        self.title_ens_rf = QtWidgets.QLabel("Ensemble des Notes Standard - Raisonnement Fluide")
        self.title_ens_mdt = QtWidgets.QLabel("Ensemble des Notes Standard - Mémoire de travail")
        self.title_ens_vt = QtWidgets.QLabel("Ensemble des Notes Standard - Vitesse de traitement")
        self.title_ens_et = QtWidgets.QLabel("Ensemble des Notes Standard - Echelle Totale")
        self.title_nc_cv = QtWidgets.QLabel("Note Composite - Compréhension Verbale")
        self.title_nc_v = QtWidgets.QLabel("Note Composite - Visuospatial")
        self.title_nc_rf = QtWidgets.QLabel("Note Composite - Raisonnement Fluide")
        self.title_nc_mdt = QtWidgets.QLabel("Note Composite - Mémoire de travail")
        self.title_nc_vt = QtWidgets.QLabel("Note Composite - Vitesse de traitement")
        self.title_nc_et = QtWidgets.QLabel("Note Composite - Echelle Totale")
        self.title_rp_cv = QtWidgets.QLabel("Rang Percentile - Compréhension Verbale")
        self.title_rp_v = QtWidgets.QLabel("Rang Percentile - Visuospatial")
        self.title_rp_rf = QtWidgets.QLabel("Rang Percentile - Raisonnement Fluide")
        self.title_rp_mdt = QtWidgets.QLabel("Rang Percentile - Mémoire de travail")
        self.title_rp_vt = QtWidgets.QLabel("Rang Percentile - Vitesse de traitement")
        self.title_rp_et = QtWidgets.QLabel("Rang Percentile - Echelle Totale")
        self.title_idc_cv = QtWidgets.QLabel("Intervalle de Confiance - Compréhension Verbale")
        self.title_idc_v = QtWidgets.QLabel("Intervalle de Confiance - Visuospatial")
        self.title_idc_rf = QtWidgets.QLabel("Intervalle de Confiance - Raisonnement Fluide")
        self.title_idc_mdt = QtWidgets.QLabel("Intervalle de Confiance - Mémoire de travail")
        self.title_idc_vt = QtWidgets.QLabel("Intervalle de Confiance - Vitesse de traitement")
        self.title_idc_et = QtWidgets.QLabel("Intervalle de Confiance - Echelle Totale")

        # --- Zone scrollable setup ---
        content_widget = QtWidgets.QWidget()
        content_layout = QtWidgets.QVBoxLayout(content_widget)

        # Ajout de tous les widgets à content_layout
        fields = [
            (self.title_nom, self.champs_nom),
            (self.title_prenom, self.champs_prenom),
            (self.title_date_naiss, self.champs_date_naiss),
            (self.title_age, self.champs_age),
            (self.title_lat, self.champs_lat),
            (self.title_date, self.champs_date),
            (self.title_ens_cv, self.champs_ens_cv),
            (self.title_ens_v, self.champs_ens_v),
            (self.title_ens_rf, self.champs_ens_rf),
            (self.title_ens_mdt, self.champs_ens_mdt),
            (self.title_ens_vt, self.champs_ens_vt),
            (self.title_ens_et, self.champs_ens_et),
            (self.title_nc_cv, self.champs_nc_cv),
            (self.title_nc_v, self.champs_nc_v),
            (self.title_nc_rf, self.champs_nc_rf),
            (self.title_nc_mdt, self.champs_nc_mdt),
            (self.title_nc_vt, self.champs_nc_vt),
            (self.title_nc_et, self.champs_nc_et),
            (self.title_rp_cv, self.champs_rp_cv),
            (self.title_rp_v, self.champs_rp_v),
            (self.title_rp_rf, self.champs_rp_rf),
            (self.title_rp_mdt, self.champs_rp_mdt),
            (self.title_rp_vt, self.champs_rp_vt),
            (self.title_rp_et, self.champs_rp_et),
            (self.title_idc_cv, self.champs_idc_cv),
            (self.title_idc_v, self.champs_idc_v),
            (self.title_idc_rf, self.champs_idc_rf),
            (self.title_idc_mdt, self.champs_idc_mdt),
            (self.title_idc_vt, self.champs_idc_vt),
            (self.title_idc_et, self.champs_idc_et),
        ]

        for label, champ in fields:
            content_layout.addWidget(label)
            content_layout.addWidget(champ)

        content_layout.addWidget(self.button)

        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(content_widget)

        main_layout = QtWidgets.QVBoxLayout(self)
        main_layout.addWidget(scroll_area)

        # Style du bouton
        self.button.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border-radius: 10px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:pressed {
                background-color: #1c5980;
            }
        """)

        # Style des champs de texte
        self.setStyleSheet("""
            QLineEdit {
                border-radius: 10px;
                padding: 8px 16px;
            }""")


        self.button.clicked.connect(self.generate_bilan)

    @QtCore.Slot()
    def magic(self):
        self.text.setText("Hello " + self.champs_nom.text())

    doc = Document()

    def generate_bilan(self):
        cadrePrésentation(
            self.champs_nom.text(),
            self.champs_prenom.text(),
            self.champs_date_naiss.text(),
            self.champs_age.text(),
            self.champs_lat.text(),
            self.champs_date.text(),
            doc
        )
        notes_compo_principales(self.champs_ens_cv.text(), self.champs_ens_v.text(), self.champs_ens_rf.text(), self.champs_ens_mdt.text(), self.champs_ens_vt.text(), self.champs_ens_et.text(),
                            self.champs_nc_cv.text(), self.champs_nc_v.text(), self.champs_nc_rf.text(), self.champs_nc_mdt.text(), self.champs_nc_vt.text(), self.champs_nc_et.text(),
                            self.champs_rp_cv.text(), self.champs_rp_v.text(), self.champs_rp_rf.text(), self.champs_rp_mdt.text(), self.champs_rp_vt.text(), self.champs_rp_et.text(),
                            self.champs_idc_cv.text(), self.champs_idc_v.text(), self.champs_idc_rf.text(), self.champs_idc_mdt.text(), self.champs_idc_vt.text(), self.champs_idc_et.text())
        doc.save('TestV2.docx')

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    app.setStyle("Macintosh")
    widget = MyWidget()
    widget.resize(600, 500)
    widget.show()

    sys.exit(app.exec())
