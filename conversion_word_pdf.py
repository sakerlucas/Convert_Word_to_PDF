import os
import win32com.client as win32

# Chemin d'accès au dossier contenant les fichiers Word
dossier_word = r"D:/0WORDenPDF"

# Créer une instance de l'application Word
word = win32.gencache.EnsureDispatch("Word.Application")

# Désactiver les alertes de Word
word.DisplayAlerts = False

# Parcourir les fichiers Word du dossier
for fichier in os.listdir(dossier_word):
    if fichier.endswith(".docx") or fichier.endswith(".doc"):
        # Chemin d'accès au fichier Word
        chemin_fichier = os.path.join(dossier_word, fichier)

        # Ouvrir le fichier Word
        doc = word.Documents.Open(chemin_fichier)

        # Convertir en PDF
        pdf_chemin = os.path.splitext(chemin_fichier)[0] + ".pdf"
        doc.SaveAs(pdf_chemin, FileFormat=17)  # 17 correspond au format PDF

        # Fermer le fichier Word
        doc.Close()

# Réactiver les alertes de Word
word.DisplayAlerts = True

# Fermer l'application Word
word.Quit()
