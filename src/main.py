import tkinter as tk
from tkinter import ttk
from views.main_window import MainWindow

def main():
    root = tk.Tk()
    root.title("Automação de Lançamentos Contábeis")
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main() 