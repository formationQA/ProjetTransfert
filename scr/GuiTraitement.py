import tkinter as tk
from tkinter import filedialog, messagebox
from TraitementDonnees import ImputationProcessor

class ImputationInterface:
    def __init__(self, processor):
        self.processor = processor

    def choisir_fichiers(self):
        input_files = filedialog.askopenfilenames(
            title="Choisir les fichiers d'imputation",
            filetypes=[("Fichiers Excel", "*.xlsx;*.xls"), ("Tous les fichiers", "*.*")]
        )
        if input_files:
            try:
                self.processor.inserer_donnees(input_files)
                messagebox.showinfo("Succès", f"Données ajoutées et sauvegardées dans '{self.processor.output_file}'.")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du traitement : {str(e)}")
        else:
            messagebox.showinfo("Info", "Aucun fichier sélectionné.")

    def lancer_interface(self):
        self.root = tk.Tk()
        self.root.title("Ajouter des données d'imputation")
        self.root.geometry("800x600")

        btn_choisir = tk.Button(self.root, text="Choisir les fichiers d'imputation", command=self.choisir_fichiers)
        btn_choisir.pack(pady=20)

        self.root.mainloop()
