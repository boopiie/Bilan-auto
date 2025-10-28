import sys
import random
import PySide6.QtCore
from PySide6 import QtCore, QtWidgets, QtGui
from docx import Document
from autoBilan import *
from PySide6.QtCore import QPropertyAnimation, QEasingCurve, QParallelAnimationGroup, QRect


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

        self.champs_indices_IAG = QtWidgets.QLineEdit()
        self.champs_indices_rp1 = QtWidgets.QLineEdit()
        self.champs_indices_ICC = QtWidgets.QLineEdit()
        self.champs_indices_rp2 = QtWidgets.QLineEdit()
        self.champs_indices_INV = QtWidgets.QLineEdit()
        self.champs_indices_rp3 = QtWidgets.QLineEdit()

        self.champs_capacite_verbal_note_stand_simi = QtWidgets.QLineEdit()
        self.champs_capacite_verbal_note_stand_vocab = QtWidgets.QLineEdit()

        self.champs_visuo_spatial_note_stand_cubes = QtWidgets.QLineEdit()
        self.champs_visuo_spatial_note_stand_puzz = QtWidgets.QLineEdit()

        self.champs_rf_note_stand_mat = QtWidgets.QLineEdit()
        self.champs_rf_note_stand_bal = QtWidgets.QLineEdit()

        self.champs_mdt_note_stand_chiffre = QtWidgets.QLineEdit()
        self.champs_mdt_note_stand_image = QtWidgets.QLineEdit()

        self.champs_vdt_note_stand_code = QtWidgets.QLineEdit()
        self.champs_vdt_note_stand_symb = QtWidgets.QLineEdit()

        self.champs_nom_fichier = QtWidgets.QLineEdit()


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
        self.title_indices_IAG = QtWidgets.QLabel("Indice complémentaire d’aptitude générale - IAG")
        self.title_indices_rp1 = QtWidgets.QLabel("Indice complémentaire d’aptitude générale - Rang percentile")
        self.title_indices_ICC = QtWidgets.QLabel("Indice de compétence cognitive - ICC")
        self.title_indices_rp2 = QtWidgets.QLabel("Indice de compétence cognitive - Rang percentile")
        self.title_indices_INV = QtWidgets.QLabel("Indice non verbal - INV")
        self.title_indices_rp3 = QtWidgets.QLabel("Indice non verbal - Rang percentile")

        self.title_capacite_verbal_note_stand_simi = QtWidgets.QLabel("Similitudes - Notes Standards")
        self.title_capacite_verbal_note_stand_vocab = QtWidgets.QLabel("Vocabulaire - Notes Standards")

        self.title_visuo_saptial_note_stand_cubes = QtWidgets.QLabel("Cubes - Notes Standards")
        self.title_visuo_saptial_note_stand_puzz = QtWidgets.QLabel("Puzzle - Notes Standards")

        self.title_rf_note_stand_mat = QtWidgets.QLabel("Matrices - Notes Standards")
        self.title_rf_note_stand_bal = QtWidgets.QLabel("Balances - Notes Standards")

        self.title_mdt_note_stand_chiffre = QtWidgets.QLabel("Chiffre - Notes Standards")
        self.title_mdt_note_stand_image = QtWidgets.QLabel("Image - Notes Standards")

        self.title_vdt_note_stand_code = QtWidgets.QLabel("Code - Notes Standards")
        self.title_vdt_note_stand_symb = QtWidgets.QLabel("Symbole - Notes Standards")

        self.title_nom_fichier = QtWidgets.QLabel("Nom du fichier")

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
            (self.title_indices_IAG, self.champs_indices_IAG),
            (self.title_indices_rp1, self.champs_indices_rp1),
            (self.title_indices_ICC, self.champs_indices_ICC),
            (self.title_indices_rp2, self.champs_indices_rp2),
            (self.title_indices_INV, self.champs_indices_INV),
            (self.title_indices_rp3, self.champs_indices_rp3),
            (self.title_capacite_verbal_note_stand_simi, self.champs_capacite_verbal_note_stand_simi),
            (self.title_capacite_verbal_note_stand_vocab, self.champs_capacite_verbal_note_stand_vocab),
            (self.title_visuo_saptial_note_stand_cubes, self.champs_visuo_spatial_note_stand_cubes),
            (self.title_visuo_saptial_note_stand_puzz, self.champs_visuo_spatial_note_stand_puzz),
            (self.title_rf_note_stand_mat, self.champs_rf_note_stand_mat),
            (self.title_rf_note_stand_bal, self.champs_rf_note_stand_bal),
            (self.title_mdt_note_stand_chiffre, self.champs_mdt_note_stand_chiffre),
            (self.title_mdt_note_stand_image, self.champs_mdt_note_stand_image),
            (self.title_vdt_note_stand_code, self.champs_vdt_note_stand_code),
            (self.title_vdt_note_stand_symb, self.champs_vdt_note_stand_symb),
            (self.title_nom_fichier, self.champs_nom_fichier)
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

        # Style du bouton futuriste & dynamique
        self.button.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #00c6ff, stop:1 #0072ff);
                color: white;
                border: none;
                border-radius: 18px;
                padding: 14px 28px;
                font-size: 16px;
                font-weight: bold;
                letter-spacing: 1px;
                box-shadow: 0 5px 15px rgba(0, 114, 255, 0.4);
            }

            QPushButton:hover {
                background: qlineargradient(x1:1, y1:0, x2:0, y2:1,
                                            stop:0 #00c6ff, stop:1 #0072ff);
            }

            QPushButton:pressed {
                background-color: #0051a3;
                box-shadow: inset 0 4px 10px rgba(0,0,0,0.3);
            }
        """)

        # Style général de l'app
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                                            stop:0 #1e1e2f, stop:1 #2c2c3e);
                font-family: 'Segoe UI', sans-serif;
                font-size: 13px;
                color: #ecf0f1;
            }

            QLabel {
                color: #f0f0f0;
                font-weight: 600;
                padding: 6px 0;
                font-size: 13px;
                letter-spacing: 0.5px;
            }

            QLineEdit {
                background: rgba(255, 255, 255, 0.08);
                border: 1px solid rgba(255, 255, 255, 0.2);
                border-radius: 12px;
                padding: 12px 16px;
                margin-bottom: 14px;
                color: #ffffff;
                font-weight: 500;
                transition: all 0.3s ease-in-out;
            }

            QLineEdit:hover {
                background: rgba(255, 255, 255, 0.12);
            }

            QLineEdit:focus {
                border: 1px solid #00c6ff;
            }
        """)


        self.button.clicked.connect(self.generate_bilan)

    @QtCore.Slot()
    def magic(self):
        self.text.setText("Hello " + self.champs_nom.text())

    doc = Document()

    def generate_bilan(self):

        #Appel de chaque fonctions pour generer le bilan en fonction des informations renseignées

        cadrePrésentation(
            self.champs_nom.text(),
            self.champs_prenom.text(),
            self.champs_date_naiss.text(),
            self.champs_age.text(),
            self.champs_lat.text(),
            self.champs_date.text(),
            doc
        )
        """notes_compo_principales(self.champs_ens_cv.text(), self.champs_ens_v.text(), self.champs_ens_rf.text(), self.champs_ens_mdt.text(), self.champs_ens_vt.text(), self.champs_ens_et.text(),
                            self.champs_nc_cv.text(), self.champs_nc_v.text(), self.champs_nc_rf.text(), self.champs_nc_mdt.text(), self.champs_nc_vt.text(), self.champs_nc_et.text(),
                            self.champs_rp_cv.text(), self.champs_rp_v.text(), self.champs_rp_rf.text(), self.champs_rp_mdt.text(), self.champs_rp_vt.text(), self.champs_rp_et.text(),
                            self.champs_idc_cv.text(), self.champs_idc_v.text(), self.champs_idc_rf.text(), self.champs_idc_mdt.text(), self.champs_idc_vt.text(), self.champs_idc_et.text(),
                            self.champs_prenom.text())
        
        indices(self.champs_indices_IAG.text(), self.champs_indices_rp1.text(),
                self.champs_indices_ICC.text(), self.champs_indices_rp2.text(),
                self.champs_indices_INV.text(), self.champs_indices_rp3.text(),
                self.champs_prenom.text())
        
        capacite_verbal(self.champs_ens_cv.text(), self.champs_rp_cv.text(),
                        self.champs_capacite_verbal_note_stand_simi.text(),
                        self.champs_capacite_verbal_note_stand_vocab.text(),
                        self.champs_prenom.text())
        
        visuo_spatial(self.champs_ens_v.text(), self.champs_rp_v.text(),
                      self.champs_visuo_spatial_note_stand_cubes.text(),
                      self.champs_visuo_spatial_note_stand_puzz.text(),
                      self.champs_prenom.text())

        raisonnement_fluide(self.champs_ens_rf.text(), self.champs_rp_rf.text(),
                            self.champs_rf_note_stand_mat.text(),
                            self.champs_rf_note_stand_bal.text(),
                            self.champs_prenom.text())

        memoire_de_travail(self.champs_ens_mdt.text(), self.champs_rp_mdt.text(),
                           self.champs_mdt_note_stand_chiffre.text(),
                           self.champs_mdt_note_stand_image.text(),
                           self.champs_prenom.text())
        
        vitesse_de_traitement(self.champs_ens_vt.text(), self.champs_rp_vt.text(),
                              self.champs_vdt_note_stand_code.text(),
                              self.champs_vdt_note_stand_symb.text(),
                              self.champs_prenom.text()) """
        
        notes_compo_principales()
        indices()
        capacite_verbal()
        visuo_spatial()
        raisonnement_fluide
        memoire_de_travail()
        vitesse_de_traitement()

        font_reset()

        alignement_reset()
        
        doc.save(f'/Users/Arcimboldo/Desktop/logiciel/bilan/{self.champs_nom_fichier.text()}.docx')

if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    widget = MyWidget()
    widget.resize(600, 500)
    widget.show()

    sys.exit(app.exec())
