import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import datetime
import win32com.client as win32
import pythoncom
import os
import sys
import time

class ExcelManager:
    def __init__(self):
        self._clean_com_cache()
        self.excel = None

    def __enter__(self):
        pythoncom.CoInitialize()
        self._kill_excel()
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False
        self.excel.EnableEvents = False
        return self.excel

    def __exit__(self, exc_type, exc_val, exc_tb):
        self._cleanup()

    def _clean_com_cache(self):
        """Nettoyage du cache COM pour éviter les conflits"""
        cache_dir = os.path.join(sys.prefix, 'Lib', 'site-packages', 'win32com', 'gen_py')
        if os.path.exists(cache_dir):
            for f in os.listdir(cache_dir):
                try: os.remove(os.path.join(cache_dir, f))
                except: pass

    def _kill_excel(self):
        """Fermeture forcée de tous les processus Excel"""
        os.system('taskkill /f /im excel.exe > nul 2>&1')
        time.sleep(1)

    def _cleanup(self):
        """Nettoyage en profondeur des ressources COM"""
        try:
            if self.excel:
                self.excel.DisplayAlerts = False
                self.excel.Quit()
                del self.excel
        except: pass
        pythoncom.CoUninitialize()
        self._kill_excel()
        time.sleep(0.5)

def get_valid_save_path(dossier_sortie):
    """Valide et formate le chemin de sauvegarde"""
    try:
        # Nettoyage du chemin
        dossier_sortie = os.path.normpath(dossier_sortie.strip())
        dossier_sortie = dossier_sortie.replace('\xa0', '')  # Suppression caractères spéciaux
        
        # Création du dossier si besoin
        os.makedirs(dossier_sortie, exist_ok=True)
        
        # Génération nom de fichier
        date_str = datetime.datetime.now().strftime("%d%m%y")
        filename = f"Suivi_commande_{date_str}.xlsx"
        
        # Construction chemin complet
        full_path = os.path.abspath(os.path.join(dossier_sortie, filename))
        full_path = full_path.replace('/', '\\')  # Format Windows obligatoire
        
        # Vérification finale
        if not os.path.exists(dossier_sortie):
            raise ValueError("Chemin invalide ou permissions insuffisantes")
            
        if len(full_path) > 220:
            raise ValueError("Chemin trop long (>220 caractères)")
            
        return full_path
    except Exception as e:
        raise ValueError(f"Erreur de chemin : {str(e)}\nChemin utilisé : {dossier_sortie}")

def copy_comment(src_cell, dst_cell):
    """Copie un commentaire moderne (avec auteur, tâches, etc.)"""
    if not src_cell.Comment:
        return
    
    # Créer un nouveau commentaire
    new_comment = dst_cell.AddComment()
    
    # Copier le texte
    new_comment.Text(src_cell.Comment.Text())
    
    # Copier l'auteur
    try:
        new_comment.Author = src_cell.Comment.Author
    except:
        pass
    
    # Copier les tâches et balises (si disponibles)
    try:
        for task in src_cell.Comment.Tasks:
            new_task = new_comment.Tasks.Add(task.Text, task.AssignedTo, task.DueDate)
            new_task.Status = task.Status
            new_task.Priority = task.Priority
    except:
        pass
    
    # Copier le formatage
    new_comment.Shape.TextFrame.AutoSize = True
    new_comment.Shape.TextFrame.Characters().Font.Name = "Calibri"
    new_comment.Shape.TextFrame.Characters().Font.Size = 10

def traiter_fichiers(ancien_path, nouveau_path):
    with ExcelManager() as excel:
        try:
            # Ouverture fichiers
            wb_ancien = excel.Workbooks.Open(ancien_path, ReadOnly=True, CorruptLoad=1)
            ws_ancien = wb_ancien.Sheets(1)
            
            wb_nouveau = excel.Workbooks.Open(nouveau_path, CorruptLoad=1)
            ws_nouveau = wb_nouveau.Sheets(1)

            # Créer la colonne Remarques si elle n'existe pas
            header_nouveau = [cell.Value for cell in ws_nouveau.Rows(1)]
            colonne_remarques = 14  # Colonne N
            
            if "Remarques" not in header_nouveau:
                ws_nouveau.Columns(colonne_remarques).Insert()
                ws_nouveau.Cells(1, colonne_remarques).Value = "Remarques"

            # Collecte des données ancien fichier
            donnees = {}
            last_row = ws_ancien.UsedRange.Rows.Count
            
            for row in range(2, last_row + 1):
                key = (
                    str(ws_ancien.Cells(row, 2).Value).strip().upper(),  # N° Pièce
                    str(ws_ancien.Cells(row, 5).Value).strip().upper()   # Réf. Article
                )
                cell = ws_ancien.Cells(row, colonne_remarques)
                donnees[key] = {
                    'value': cell.Value,
                    'comment': cell
                }

            # Application sur nouveau fichier
            last_row_nouveau = ws_nouveau.UsedRange.Rows.Count
            
            for row in range(2, last_row_nouveau + 1):
                key = (
                    str(ws_nouveau.Cells(row, 2).Value).strip().upper(),
                    str(ws_nouveau.Cells(row, 5).Value).strip().upper()
                )
                
                if key in donnees:
                    cell = ws_nouveau.Cells(row, colonne_remarques)
                    data = donnees[key]
                    
                    # Mise à jour valeur
                    cell.Value = data['value']
                    
                    # Mise à jour commentaire
                    if data['comment'].Comment:
                        if cell.Comment:
                            cell.Comment.Delete()
                        copy_comment(data['comment'], cell)

            # Sauvegarde
            dossier_sortie = filedialog.askdirectory(title="Dossier de sauvegarde")
            if not dossier_sortie: return
            
            save_path = get_valid_save_path(dossier_sortie)
            
            if os.path.exists(save_path):
                os.remove(save_path)
                
            wb_nouveau.SaveAs(save_path)
            return save_path

        except Exception as e:
            raise Exception(f"Erreur : {str(e)}")
        finally:
            try: wb_ancien.Close(False)
            except: pass
            try: wb_nouveau.Close(False)
            except: pass


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Gestionnaire de Commandes")
        self.geometry("500x350")
        self._setup_ui()
        
    def _setup_ui(self):
        frame = ttk.Frame(self, padding=20)
        frame.pack(expand=True, fill=tk.BOTH)
        
        ttk.Button(frame, 
                 text="Sélectionner fichier ancien", 
                 command=self._select_ancien).pack(pady=5)
        self.lbl_ancien = ttk.Label(frame, text="Aucun fichier sélectionné")
        self.lbl_ancien.pack()
        
        ttk.Button(frame, 
                 text="Sélectionner fichier nouveau", 
                 command=self._select_nouveau).pack(pady=5)
        self.lbl_nouveau = ttk.Label(frame, text="Aucun fichier sélectionné")
        self.lbl_nouveau.pack()
        
        ttk.Button(frame,
                 text="Exécuter le traitement",
                 command=self._executer).pack(pady=15)
                 
        self.status = ttk.Label(frame, text="Prêt")
        self.status.pack()
        
    def _select_ancien(self):
        path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
        if path:
            self.ancien_path = path
            self.lbl_ancien.config(text=f"Ancien : {os.path.basename(path)}")
            
    def _select_nouveau(self):
        path = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx;*.xls")])
        if path:
            self.nouveau_path = path
            self.lbl_nouveau.config(text=f"Nouveau : {os.path.basename(path)}")
            
    def _executer(self):
        if not hasattr(self, 'ancien_path') or not hasattr(self, 'nouveau_path'):
            messagebox.showerror("Erreur", "Sélectionnez les deux fichiers")
            return
            
        self.status.config(text="Traitement en cours...")
        self.update()
        
        try:
            result = traiter_fichiers(self.ancien_path, self.nouveau_path)
            messagebox.showinfo("Succès", f"Fichier généré :\n{result}")
        except Exception as e:
            messagebox.showerror("Erreur", str(e))
        finally:
            self.status.config(text="Prêt")

if __name__ == "__main__":
    app = Application()
    app.mainloop()