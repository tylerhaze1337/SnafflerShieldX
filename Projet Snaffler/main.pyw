import sys
import subprocess
import os
import pandas as pd
import socket
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QTextEdit, QLineEdit
)
from PyQt6.QtGui import QTextCharFormat, QColor
from PyQt6.QtCore import Qt
import threading
import matplotlib.pyplot as plt

class SnafflerGUI(QWidget):
    def __init__(self):
        super().__init__()
        self.snaffler_process = None  # Référence au processus Snaffler
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Labels d'informations
        self.info_label = QLabel(f"📌 Machine détectée : {socket.gethostname()}")
        layout.addWidget(self.info_label)

        # Bouton Exécuter
        self.run_button = QPushButton("🚀 Lancer Snaffler")
        self.run_button.clicked.connect(self.run_snaffler)
        layout.addWidget(self.run_button)

        # Zone d'affichage des résultats
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        layout.addWidget(self.output_text)

        # Champ pour attente de l'entrée utilisateur
        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Appuyez sur 'Entrée' pour continuer après une erreur.")
        self.input_field.setReadOnly(True)  # Initialement désactivé
        self.input_field.returnPressed.connect(self.continue_scan)
        layout.addWidget(self.input_field)

        # Bouton Afficher Histogramme
        self.graph_button = QPushButton("📊 Afficher Histogramme")
        self.graph_button.clicked.connect(self.show_histogram)
        layout.addWidget(self.graph_button)

        # Bouton Ouvrir CSV
        self.open_csv_button = QPushButton("📂 Ouvrir le rapport CSV")
        self.open_csv_button.clicked.connect(self.open_csv)
        layout.addWidget(self.open_csv_button)

        self.setLayout(layout)
        self.setWindowTitle("Snaffler GUI - Analyse et Rapports")
        self.setGeometry(100, 100, 600, 500)

        # Compteurs pour les alertes
        self.green_count = 0
        self.orange_count = 0
        self.red_count = 0
        self.black_count = 0

        # Chemins des fichiers
        self.snaffler_exe = os.path.join(os.getcwd(), "Snaffler.exe")
        self.parser_exe = os.path.join(os.getcwd(), "SnafflerParser.ps1")
        self.output_json = os.path.join(os.getcwd(), "output.json")
        self.output_csv = os.path.join(os.getcwd(), "output_loot_full.csv")

    def run_snaffler(self):
        # Lancer le thread pour exécuter la commande
        threading.Thread(target=self.execute_snaffler, daemon=True).start()

    def execute_snaffler(self):
        machine = socket.gethostname()

        # Vider le fichier JSON avant chaque scan
        with open(self.output_json, 'w', encoding='utf-8') as json_file:
            json_file.truncate(0)  # Vide le contenu du fichier

        # Vider le fichier CSV avant chaque scan
        with open(self.output_csv, 'w', encoding='utf-8') as csv_file:
            csv_file.truncate(0)  # Vider le contenu du fichier CSV

        # Vérifier et créer les fichiers CSV s'ils n'existent pas
        if not os.path.exists(self.output_csv):
            open(self.output_csv, 'w').close()

        # Commande complète sous forme de chaîne
        command = f'"{self.snaffler_exe}" -s -n {machine} -o "{self.output_json}" -r 1000 -j 200 -v debug -y'
        
        print(f'"{self.snaffler_exe}" -s -n {machine} -o "{self.output_json}" -r 10000000 -j 200 -v debug -y')

        try:
            # Exécuter la commande complète dans cmd ou PowerShell avec encodage flexible
            self.snaffler_process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True, encoding='utf-8', errors='ignore')

            # Lire les flux stdout et stderr en temps réel
            for stdout_line in iter(self.snaffler_process.stdout.readline, ""):
                self.process_output(stdout_line.strip())  # Traiter la ligne stdout

            for stderr_line in iter(self.snaffler_process.stderr.readline, ""):
                self.process_output(stderr_line.strip(), is_error=True)  # Traiter la ligne stderr

            self.snaffler_process.stdout.close()
            self.snaffler_process.stderr.close()
            self.snaffler_process.wait()

            # Lancer le parser après Snaffler
            self.output_text.append("📌 Exécution du parser en cours...")
            parse_command = f"powershell.exe -ExecutionPolicy Bypass -File \"{self.parser_exe}\" -in \"{self.output_json}\" -outformat csv -output \"{self.output_csv}\""

            parse_process = subprocess.Popen(parse_command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, shell=True, encoding='utf-8', errors='ignore')
            
            for parse_stdout_line in iter(parse_process.stdout.readline, ""):
                self.process_output(parse_stdout_line.strip())  # Traiter la ligne stdout du parser

            for parse_stderr_line in iter(parse_process.stderr.readline, ""):
                self.process_output(parse_stderr_line.strip(), is_error=True)  # Traiter la ligne stderr du parser

            parse_process.stdout.close()
            parse_process.stderr.close()
            parse_process.wait()

            self.output_text.append(f"✅ Rapport généré: {self.output_csv}")

        except Exception as e:
            self.output_text.append(f"❌ Erreur d'exécution: {str(e)}")
            self.input_field.setReadOnly(False)  # Activer le champ de texte pour l'entrée de l'utilisateur
            self.output_text.append("Appuyez sur 'Entrée' pour continuer le scan...")

    def process_output(self, line, is_error=False):
        """
        Process each output line and append it to the QTextEdit with color based on content.
        """
        if is_error:
            # Erreurs (stderr) seront en vert (changer à votre demande)
            self.append_colored_text(f"⚠️ Erreur: {line}", QColor("green"))
        else:
            # Vérification du contenu de la ligne pour la couleur
            self.append_colored_text(line, QColor("green"))

    def append_colored_text(self, text, color):
        char_format = QTextCharFormat()
        char_format.setForeground(color)
        cursor = self.output_text.textCursor()
        cursor.setCharFormat(char_format)
        cursor.insertText(text + "\n")
        
        # Déplace le curseur à la fin du texte pour que la barre de défilement suive
        cursor.movePosition(cursor.MoveOperation.End)
        self.output_text.setTextCursor(cursor)

        # Assurez-vous que la barre de défilement défile vers le bas
        self.output_text.verticalScrollBar().setValue(self.output_text.verticalScrollBar().maximum())

    def show_histogram(self):
            try:
                # Charger le fichier CSV
                df = pd.read_csv(self.output_csv)

                # Vérifier si la colonne 'severity' existe
                if 'severity' in df.columns:
                    # Normaliser les valeurs de 'severity' (convertir en minuscules)
                    df['severity'] = df['severity'].str.lower()

                    # Compter le nombre d'occurrences pour chaque niveau de 'severity'
                    severity_counts = df['severity'].value_counts()

                    # Créer les labels et les valeurs pour l'histogramme
                    labels = severity_counts.index.tolist()
                    values = severity_counts.tolist()

                    # Définir une palette de couleurs dans un ordre spécifique
                    severity_order = ['yellow', 'green', 'orange', 'red', 'black']
                    color_map = {severity: color for severity, color in zip(severity_order, ['yellow', 'green', 'orange', 'red', 'black'])}
                    bar_colors = [color_map.get(label, 'gray') for label in labels]  # Utiliser 'gray' pour des labels inconnus

                    # Création de la figure avec 2 sous-graphiques côte à côte
                    fig, axs = plt.subplots(1, 2, figsize=(16, 6))  # 1 ligne, 2 colonnes

                    # Premier graphique : Histogramme
                    axs[0].bar(labels, values, color=bar_colors)  # Assure-toi que les couleurs sont bien associées
                    axs[0].set_xlabel("Niveau de Criticité")
                    axs[0].set_ylabel("Nombre d'occurrences")
                    axs[0].set_title("Histogramme des alertes par niveau de criticité")

                    # Deuxième graphique : Camembert
                    wedges, texts, autotexts = axs[1].pie(values, labels=labels, autopct='%1.1f%%', colors=bar_colors, startangle=90, wedgeprops={'edgecolor': 'black'})
                    axs[1].set_title("Répartition des alertes par niveau de criticité")

                    # Changer la couleur des textes des pourcentages pour les rendre visibles
                    for autotext in autotexts:
                        autotext.set_color('white')  # Définit la couleur des pourcentages sur blanc

                    # Afficher les graphiques
                    plt.tight_layout()  # Ajuste les espacements pour que les graphiques ne se chevauchent pas
                    plt.show()

                else:
                    self.output_text.append("❌ La colonne 'severity' n'existe pas dans le fichier CSV.")

            except Exception as e:
                self.output_text.append(f"❌ Erreur lors du traitement du fichier CSV : {str(e)}")


    def open_csv(self):
                os.startfile(self.output_csv)

    def continue_scan(self):
                # Une fois que l'utilisateur appuie sur "Entrée", réactiver le processus de scan
                self.input_field.setReadOnly(True)  # Désactiver l'entrée
                self.output_text.append("🔄 Reprise du scan...")

                # Relancer le processus d'exécution
                threading.Thread(target=self.execute_snaffler, daemon=True).start()

    def closeEvent(self, event):
                if self.snaffler_process:
                    # Tuer le processus Snaffler si il est en cours
                    self.snaffler_process.terminate()
                    self.snaffler_process.wait()  # Attendre la fin du processus avant de fermer
                    self.output_text.append("❌ Processus Snaffler arrêté en arrière-plan.")

                event.accept()
      

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = SnafflerGUI()
    window.show()
    sys.exit(app.exec())
