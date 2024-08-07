import pandas as pd
from tkinter import Tk, Label, Text, Button, END, messagebox
from datetime import datetime
import os

# Função para adicionar feedback e sugestão de melhoria
def add_feedback():
    global df  # Declarar df como global
    feedback = feedback_entry.get("1.0", END).strip()
    suggestion = suggestion_entry.get("1.0", END).strip()
    if feedback:
        new_row = {
            'employee_id': len(df) + 1,
            'date': datetime.now().strftime('%Y-%m-%d'),
            'feedback': feedback,
            'suggestion': suggestion
        }
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
        # Usar um nome de arquivo temporário para evitar problemas de permissão
        temp_filename = 'feedbacks_temp.xlsx'
        df.to_excel(temp_filename, index=False)
        # Substituir o arquivo original pelo temporário
        if os.path.exists('feedbacks.xlsx'):
            os.remove('feedbacks.xlsx')
        os.rename(temp_filename, 'feedbacks.xlsx')
        feedback_entry.delete("1.0", END)
        suggestion_entry.delete("1.0", END)
        messagebox.showinfo("Feedback", "Feedback e sugestão adicionados com sucesso!")
    else:
        messagebox.showwarning("Feedback", "Por favor, insira um feedback.")

# Inicializar DataFrame
try:
    df = pd.read_excel('feedbacks.xlsx')
except FileNotFoundError:
    df = pd.DataFrame(columns=['employee_id', 'date', 'feedback', 'suggestion'])

# Interface gráfica com tkinter
root = Tk()
root.title("Análise de Feedback dos Funcionários")

Label(root, text="Digite seu feedback:").pack()
feedback_entry = Text(root, height=10, width=50)
feedback_entry.pack()

Label(root, text="Digite sua sugestão de melhoria (opcional):").pack()
suggestion_entry = Text(root, height=5, width=50)
suggestion_entry.pack()

Button(root, text="Adicionar Feedback e Sugestão", command=add_feedback).pack()

root.mainloop()
