import sys
import random
import PySide6.QtCore
from PySide6 import QtCore, QtWidgets, QtGui
from docx import Document
from autoBilan import *

print(PySide6.__version__)

class MyWidget(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        # Bouton
        self.button = QtWidgets.QPushButton("Génerer le bilan")
        
        # Champs de texte à remplir par l'utilisateur
        self.champs_nom = QtWidgets.QLineEdit()
        self.champs_prenom = QtWidgets.QLineEdit()
        self.champs_date_naiss= QtWidgets.QLineEdit()
        self.champs_age = QtWidgets.QLineEdit()
        self.champs_lat = QtWidgets.QLineEdit()
        self.champs_date = QtWidgets.QLineEdit()

        # Titre des champs à remplir
        self.title_nom = QtWidgets.QLabel("Nom du patient")
        self.title_prenom = QtWidgets.QLabel("Prénom du patient")
        self.title_date_naiss = QtWidgets.QLabel("Date de naissance du patient")
        self.title_age = QtWidgets.QLabel("Age du patient")
        self.title_lat = QtWidgets.QLabel("Latitude du patient")
        self.title_date = QtWidgets.QLabel("Date du bilan")

        self.layout = QtWidgets.QVBoxLayout(self)
        self.layout.addWidget(self.title_nom)
        self.layout.addWidget(self.champs_nom)
        self.layout.addWidget(self.title_prenom)
        self.layout.addWidget(self.champs_prenom)
        self.layout.addWidget(self.title_date_naiss)
        self.layout.addWidget(self.champs_date_naiss)
        self.layout.addWidget(self.title_age)
        self.layout.addWidget(self.champs_age)
        self.layout.addWidget(self.title_lat)
        self.layout.addWidget(self.champs_lat)
        self.layout.addWidget(self.title_date)
        self.layout.addWidget(self.champs_date)
        self.layout.addWidget(self.button)

        # Modification des styles 

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

        # Style des champs de texte et du background
        self.setStyleSheet("""
            QLineEdit {
                border-radius: 10px;
                padding: 8px 16px;
            
            QApplication {
                color = black;}
            }
        """)

        

        self.button.clicked.connect(self.generate_bilan)

    @QtCore.Slot()
    def magic(self):
        self.text.setText("Hello " + self.champs_nom.text())

    def generate_bilan(self):
        cadrePrésentation(self.champs_nom, self.champs_prenom, self.champs_date_naiss, self.champs_age, self.champs_lat, self.champs_date)
        doc.save('Test.docx')

if __name__ == "__main__":
    app = QtWidgets.QApplication([])

    app.setStyle("Macintosh") 
    widget = MyWidget()
    widget.resize(400, 300)
    widget.show()

    sys.exit(app.exec())