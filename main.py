import customtkinter as ctk
from tkinter import filedialog, messagebox
from logic import analyze_excel, export_to_excel

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class ExcelTool(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Excel-Auswertungs-Tool")
        self.geometry("500x520")
        self.file_path = None

        self.create_widgets()

    def create_widgets(self):
        self.label_title = ctk.CTkLabel(
            self, text="Excel-Auswertungs-Tool", font=("Arial", 20, "bold")
        )
        self.label_title.pack(pady=20)

        self.btn_select = ctk.CTkButton(
            self, text="Excel-Datei auswählen", command=self.select_file
        )
        self.btn_select.pack(pady=10)

        self.label_file = ctk.CTkLabel(self, text="Keine Datei ausgewählt")
        self.label_file.pack(pady=5)

        self.btn_analyze = ctk.CTkButton(
            self, text="Auswertung starten", command=self.run_analysis
        )
        self.btn_analyze.pack(pady=15)

        self.result_box = ctk.CTkTextbox(self, width=420, height=200)
        self.result_box.pack(pady=10)

        self.btn_export = ctk.CTkButton(
            self, text="Ergebnis als Excel speichern", command=self.export_result
        )
        self.btn_export.pack(pady=10)


    def select_file(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel Dateien", "*.xlsx *.xls")]
        )
        if path:
            self.file_path = path
            self.label_file.configure(text=path.split("/")[-1])

    def run_analysis(self):
        if not self.file_path:
            messagebox.showwarning("Hinweis", "Bitte zuerst eine Datei auswählen.")
            return

        result, error = analyze_excel(self.file_path)

        if error:
            messagebox.showerror("Fehler", error)
            return

        self.last_result = result

        self.result_box.delete("1.0", "end")

        self.result_box.insert(
            "end", f"Gesamtsumme: {result['total']:.2f} €\n\n"
        )

        self.result_box.insert("end", "Summe pro Kategorie:\n")
        for cat, val in result["by_category"].items():
            self.result_box.insert("end", f"  {cat}: {val:.2f} €\n")

        self.result_box.insert("end", "\nSumme pro Monat:\n")
        for month, val in result["by_month"].items():
            self.result_box.insert("end", f"  {month}: {val:.2f} €\n")

    def export_result(self):
        if not hasattr(self, "last_result"):
            messagebox.showwarning(
                "Hinweis", "Bitte zuerst eine Auswertung durchführen."
            )
            return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Datei", "*.xlsx")]
        )

        if not save_path:
            return

        try:
            export_to_excel(self.last_result, save_path)
            messagebox.showinfo(
                "Erfolg", "Excel-Datei wurde erfolgreich gespeichert."
            )
        except Exception as e:
            messagebox.showerror("Fehler", str(e))



if __name__ == "__main__":
    app = ExcelTool()
    app.mainloop()
