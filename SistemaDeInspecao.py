import tkinter as tk
from tkinter import messagebox
import openpyxl
from openpyxl import Workbook

class InspectionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inspeção de qualidade")

        self.label = tk.Label(root, text="Adicione as medidas")
        self.label.pack(pady=5)

        self.entries = []
        self.checkbuttons = []
        self.add_measurement()

        # Criação de um frame para fixar o botão "Finalizar" na parte inferior
        self.bottom_frame = tk.Frame(root)
        self.bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.submit_button = tk.Button(self.bottom_frame, text="Finalizar", command=self.submit)
        self.submit_button.pack(pady=10)

    def add_measurement(self):
        frame = tk.Frame(self.root)
        frame.pack(pady=5)

        entry = tk.Entry(frame, width=50)
        entry.pack(side=tk.LEFT, padx=5)
        self.entries.append(entry)

        check_var = tk.BooleanVar()
        checkbutton = tk.Checkbutton(frame, text="OK", variable=check_var)
        checkbutton.pack(side=tk.LEFT)
        self.checkbuttons.append(check_var)

        add_button = tk.Button(frame, text="Adicionar medida", command=self.add_measurement)
        add_button.pack(side=tk.LEFT, padx=5)

    def submit(self):
        measurements = []
        for entry, check_var in zip(self.entries, self.checkbuttons):
            measure = entry.get()
            if not measure:
                messagebox.showwarning("Erro", "Por favor, preencha todas as medidas.")
                return
            status = "OK" if check_var.get() else ""
            measurements.append((measure, status))
        
        self.save_to_excel(measurements)
        messagebox.showinfo("Sucesso", "Medidas salvas com sucesso.")

    def save_to_excel(self, data):
        wb = Workbook()
        ws = wb.active
        ws.title = "Inspeção de Qualidade"

        # Adicionar cabeçalhos
        ws.append(["Medida", "Status"])

        # Adicionar dados
        for measure, status in data:
            ws.append([measure, status])

        # Salvar o arquivo Excel
        wb.save("inspecao_qualidade.xlsx")

if __name__ == "__main__":
    root = tk.Tk()
    app = InspectionApp(root)
    root.mainloop()
