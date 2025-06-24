import datetime
import tempfile
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
        cache_dir = os.path.join(sys.prefix, 'Lib', 'site-packages',
                                 'win32com', 'gen_py')
        if os.path.exists(cache_dir):
            for f in os.listdir(cache_dir):
                try: os.remove(os.path.join(cache_dir, f))
                except:
                    pass

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
            colonne_remarques = 14  # Colonne N (colonne 14)

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
                donnees[key] = ws_ancien.Cells(row, colonne_remarques).Value

            # Application sur nouveau fichier
            last_row_nouveau = ws_nouveau.UsedRange.Rows.Count

            for row in range(2, last_row_nouveau + 1):
                key = (
                    str(ws_nouveau.Cells(row, 2).Value).strip().upper(),
                    str(ws_nouveau.Cells(row, 5).Value).strip().upper()
                )
                if key in donnees:
                    ws_nouveau.Cells(row, colonne_remarques).Value = donnees[key]

            # Chemin temporaire pour Streamlit
            date_str = datetime.datetime.now().strftime("%d%m%y_%H%M%S")
            temp_dir = tempfile.mkdtemp()
            filename = f"Suivi_commandes_{date_str}.xlsx"
            save_path = os.path.join(temp_dir, filename)

            wb_nouveau.SaveAs(save_path)
            return save_path

        except Exception as e:
            raise Exception(f"Erreur : {str(e)}")

        finally:
            try: wb_ancien.Close(False)
            except: pass
            try: wb_nouveau.Close(False)
            except: pass
