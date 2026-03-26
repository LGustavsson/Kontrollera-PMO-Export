import FreeSimpleGUI as sg
import csv
import os
import re
import shutil

class KontrollPMOExport:
    def __init__(self):
        self.file_names = [
            "Kontroll_PMO_Export_Statistik.csv",                #0
            "Kontroll_PMO_Export_Sökvägar.csv",                 #1
            "Kontroll_PMO_Export_Fel_Personnummer.csv",         #2
            "Kontroll_PMO_Export_Lyckade_Gallringar.txt",       #3
            "Kontroll_PMO_Export_Misslyckade_Gallringar.txt",   #4
            ]

        self.variable_names = [
            "search_folder",        #0
            "search_folder_text",   #1
            "search_target",        #2
            "search_target_text",   #3
            "remove_file",          #4
            "remove_file_text",     #5
            "remove_taget",         #6
            "remove_target_text" ,  #7
            ]

    def main(self):
        window = self.gui()
        while True:
            event, self.values = window.read()
            match event:
                case sg.WIN_CLOSED:
                    break

                case "Genomför kontroll":
                    if self.search_error_check():
                        continue
                    else:
                        self.search()

                case "Genomför gallring":
                    if self.remove_error_check():
                        continue
                    else:
                        self.remove()
        window.close()

    def gui(self):
        tab_information = [
            [sg.Text("Detta programstöd erbjuder två arkivaliska funktioner gentemot exportpaket från elevhälsovårdsprogramet PMO.")],
            [sg.Text("Programmet är framarbetat utifrån den struktur som exportpaketen har enligt PMOs arkivexport standard version 1.22, 2025-09-09")],
            [sg.Text("")],
            [sg.Text("I fliken 'Genomför kontroll' genomsöks exportpaketen från PMO för att:")],
            [sg.Text("  - Identifiera elever som saknar elevhälsovårdsjournal.")],
            [sg.Text("  - Identifiera felaktiga personnummer, så som TF-nummer.")],
            [sg.Text("  - Sammanställa tre filer: En för statistik över tomma elever, en för filsökväg för felaktiga personnummer, och en för filsökväg till tomma elever")],
            [sg.Text("Mappstrukturen för PMO fungerar enligt följande:")],
            [sg.Text("1. Serveryta (det är denna sökväg som väljs i PMO-admin för var exportpaket ska placeras), 2. Exportmapp, 3. Månadsmapp, 4. Personmapp.")],
            [sg.Text("Programmet kan söka från server-(all exporterad PMO-data), export-(individuell PMO-export), eller månadsmappen(elever födda en viss månad i en export).'")],
            [sg.Text("")],
            [sg.Text("I fliken 'Genomför gallring' genomförs gallring av de tomma eleverna som skapats i filen över tomma elever.")],
            ]
        
        tab_search = [
            [sg.FolderBrowse("Välj mapp för PMO-exporter", key=self.variable_names[0], size=30, target=self.variable_names[1]), sg.Text(key=self.variable_names[1], size=60)],
            [sg.Text("Här väljer ni den mappnivå som är aktuell för sökningen.")],
            [sg.FolderBrowse("Välj mapp för resultatfiler", key=self.variable_names[2], size=30, target=self.variable_names[3]), sg.Text(key=self.variable_names[3], size=60)],
            [sg.Text("Här väljer ni den mapp där resultatsfilerna ska placeras.")],
            [sg.VPush()],
            [sg.Button("Genomför kontroll", size=30)], 
            ]
        
        tab_remove = [
            [sg.FileBrowse("Välj filen med sökvägarna", key=self.variable_names[4], size=30, target=self.variable_names[5]), sg.Text(key=self.variable_names[5], size=60)],
            [sg.Text("Här väljer ni filen 'Kontroll_PMO_Export_Sökvägar.csv' som genererades från kontroll-funktionen i förra fliken.")],
            [sg.FolderBrowse("Välj mapp för resultatfil", key=self.variable_names[6], size=30, target=self.variable_names[7]), sg.Text(key=self.variable_names[7], size=60)],
            [sg.Text("Här väljer ni den mapp där resultatsfilen ska placeras.")],
            [sg.VPush()],
            [sg.Button("Genomför gallring", size=30)], 
            ]
        
        layout = [
            [sg.TabGroup([[sg.Tab("Information", tab_information), sg.Tab("Genomför kontroll", tab_search), sg.Tab("Genomför gallring", tab_remove)]])],
            [sg.Output(key="output", expand_x=True, size=(None,10))],
            ]
       
        return sg.Window(title="Kontroll PMO Export", layout=layout, icon=os.path.join(os.path.dirname(__file__), "icon.ico"))

    def search(self):
        data = {}
        students_folder_paths = []
        wrong_folder_paths =  []
        count = 1
        size = 0

        # Identifierar rätt nivå för elevens huvudmapp.
        next_folder = os.listdir(self.values[self.variable_names[0]])[0]
        depth = int(len(str(self.values[self.variable_names[0]]).split(os.path.sep)))
        # En individuell exportmapp har valts.
        if re.match(r"[\d]{4}-[\d]{2}", next_folder):
            depth += 2
        # En månadsmapp har valts.
        elif re.match(r"[\d]{12}", next_folder):
            depth += 1
        # Servermappen har valts.
        else:
            depth += 3

        for root, dirs, _ in os.walk(self.values[self.variable_names[0]]):
            for dir in dirs:
                print(f"Genomsökta mappar: {count}")
                count += 1

                # Hoppar över alla mappar som inte är elevens huvudmapp
                if len(str(os.path.join(root, dir)).split(os.path.sep)) != depth:
                    continue

                # Identifierar ifall personnumet är fel. Ordningen på kontrollerna är viktiga.
                if not str(dir).isnumeric() or int(dir[6:8]) > 31 or len(dir) != 12:
                    wrong_folder_paths.append([os.path.join(root, dir)])
                    continue

                # Bearbetar elever.
                if not dir[:4] in data:
                    data[dir[:4]] = {"stats": {"total": 1, "tomma": 0}}
                else:
                    data[dir[:4]]["stats"]["total"] += 1 

                # Identifierar tomma elever.            
                if len([file for file in os.listdir(os.path.join(root, dir)) if os.path.isfile(os.path.join(root, dir, file))]) <= 2:
                    data[dir[:4]]["stats"]["tomma"] += 1  
                    for file in os.listdir(os.path.join(root, dir)):
                        size += os.stat(os.path.join(root, dir, file)).st_size
                    students_folder_paths.append([os.path.join(root, dir)])
        
        # Skapar statistikfil.
        with open(os.path.join(self.values[self.variable_names[2]], self.file_names[0]), mode="w", encoding="UTF-8") as file:
            writer = csv.writer(file, delimiter=",", lineterminator="\n")
            writer.writerow(["Årtal", "Antal Total", "Antal Tomma", "Procent"])
            for key in sorted(data):
                writer.writerow([key, data[key]["stats"]["total"], data[key]["stats"]["tomma"], round((data[key]["stats"]["tomma"] / data[key]["stats"]["total"]) * 100)])
            writer.writerow(["Totalt antal tomma elever", f"{len(students_folder_paths)} st"])
            writer.writerow(["Total storlek på datan", f"{round((size / 1024) / 1024, 2)} mb"])
        print("\n", f"Fil med statistik har genererats här: {file.name}")

        # Om det finns tomma elever.
        if students_folder_paths:
            with open(os.path.join(self.values[self.variable_names[2]], self.file_names[1]), mode="w", encoding="UTF-8") as file:
                writer = csv.writer(file, delimiter=",", lineterminator="\n")
                writer.writerows(students_folder_paths)
                print("\n", f"Fil med sökvägar har genererats här: {file.name}")

        # Om det finns felaktiga personnummer.
        if wrong_folder_paths:
            with open(os.path.join(self.values[self.variable_names[2]], self.file_names[2]), mode="w", encoding="UTF-8") as file:
                writer = csv.writer(file, delimiter=",", lineterminator="\n")
                writer.writerows(wrong_folder_paths)
                print("\n", f"Fil med felaktiga personnummer har genererats här: {file.name}")
        
        print("Sökningen är avslutad.")

    def remove(self):
        success = []
        failure = []
        count = 1
        
        with open(self.values[self.variable_names[4]], mode="r", encoding="UTF-8") as file:
            for row in list(file):
                path = str(row).replace("/", os.path.sep).split(os.path.sep)
                
                # Tar bort elevens log.
                try:
                    os.remove(os.path.join(os.path.sep.join(i for i in path[:-2]), "Logs", f"{str(path[-1]).strip()}_log.xml"))
                    success.append(f"{path[-1]} logfil")
                except OSError:
                    failure.append(os.path.join("/".join(i for i in path[:-2]), "Logs", f"{str(path[-1]).strip()}_log.xml"))
            
                # Tar bort elevens mapp.
                try:
                    shutil.rmtree(os.path.abspath(str(row).strip()))
                    success.append(f"{path[-1]} mapp")
                except OSError:
                    failure.append(str(row).strip())

                # Tar bort filhänvisningarna i hash-dokumentet. Verkar inte påverka inläsningen så hoppar över detta.
                
                print(f"Genomfört gallringsförsök {count}")
                count += 1

        # Skapar gallringslog över lyckade gallringar.
        if success:
            with open(os.path.join(self.values[self.variable_names[6]], self.file_names[3]), mode="w", encoding="UTF-8") as file:
                file.write("Protokoll över lyckade gallringar.\n")
                file.write(f"Gallringen har utgått från {self.values[self.variable_names[4]]}.\n")
                file.write("Följande information har gallrats:\n")
                for row in success:
                    file.write(f"{row}\n")
            print(f"Lyckade gallringar finns dokumenterade här {file.name}")
        
        # Skapar gallringslog över misslyckade gallringar.
        if failure:
            with open(os.path.join(self.values[self.variable_names[6]], self.file_names[4]), mode="w", encoding="UTF-8") as file:
                file.write("Protokoll över misslyckade gallringar.\n")
                file.write(f"Gallringen har utgått från {self.values[self.variable_names[4]]}.\n")
                file.write("Följande information har inte lyckats gallrats:\n")
                for row in failure:
                    file.write(f"{row}\n")
            print(f"Misslyckade gallringar finns dokumenterade här {file.name}")
        
    def search_error_check(self):
        # Kontrollerar att sökvägar har valts.
        print(self.values)
        if self.values[self.variable_names[0]] == "" or self.values[self.variable_names[2]] == "":
            print("Mappsökväg saknas.")
            return True
        
        # Kontrollerar att filer kan skrivas.
        try:
            os.access(self.values[self.variable_names[2]], os.W_OK)
        except OSError:
            print("Mappsökvägen för resultatfilerna saknar skrivrättigheter. Välj en annan mapp.")
            return True
        
    def remove_error_check(self):
        # Kontrollerar att sökvägar har valts.
        if self.values[self.variable_names[4]] == "" or self.values[self.variable_names[6]] == "":
            print("Sökväg saknas.")
            return True

        # Kontrollerar att rätt fil med sökvägar har valts. 
        if os.path.basename(self.values[self.variable_names[4]]) != self.file_names[1]:
            print("Fel fil har valts.")
            return True

        # Kontrollerar att filer kan skrivas.
        try:
            os.access(self.values[self.variable_names[6]], os.W_OK)
        except OSError:
            print("Mappsökvägen för resultatfilen saknar skrivrättigheter. Välj en annan mapp.")
            return True

if __name__ == "__main__":
    kontroll_pmo_export = KontrollPMOExport()
    kontroll_pmo_export.main()
