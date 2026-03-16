from pathlib import Path
import hashlib
import sys
import uuid
import tkinter as tk
from tkinter import ttk, messagebox

SECRET = "NFSE_ANALYZER_2026_RODER"


def get_base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def formatar_machine_id(node: int) -> str:
    return f"{node:012X}"


def get_machine_id() -> str:
    return formatar_machine_id(uuid.getnode())


def gerar_hash_licenca(machine_id: str) -> str:
    texto = machine_id.strip().upper() + SECRET
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def salvar_arquivo_licenca(licenca: str, nome_arquivo: str) -> Path:
    destino = get_base_dir() / nome_arquivo
    destino.write_text(licenca, encoding="utf-8")
    return destino


class GeradorLicencaApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Gerador de Licença - NFSe Analyzer")
        self.geometry("760x380")
        self.resizable(False, False)

        self.var_machine_id = tk.StringVar(value=get_machine_id())
        self.var_licenca = tk.StringVar()
        self.var_nome_arquivo = tk.StringVar(value="licenca.key")

        self._montar_interface()

    def _montar_interface(self) -> None:
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Gerador de Licença - NFSe Analyzer",
            font=("Segoe UI", 15, "bold")
        ).pack(anchor="w", pady=(0, 12))

        ttk.Label(frame, text="ID da máquina:").pack(anchor="w")
        entry_machine = ttk.Entry(frame, textvariable=self.var_machine_id, width=70)
        entry_machine.pack(fill="x", pady=(4, 8))

        botoes_topo = ttk.Frame(frame)
        botoes_topo.pack(fill="x", pady=(0, 12))

        ttk.Button(botoes_topo, text="Usar ID desta máquina", command=self.usar_id_local).pack(side="left")
        ttk.Button(botoes_topo, text="Copiar ID", command=self.copiar_id).pack(side="left", padx=(8, 0))

        ttk.Label(frame, text="Licença gerada:").pack(anchor="w")
        entry_licenca = ttk.Entry(frame, textvariable=self.var_licenca, width=90)
        entry_licenca.pack(fill="x", pady=(4, 8))

        botoes_meio = ttk.Frame(frame)
        botoes_meio.pack(fill="x", pady=(0, 12))

        ttk.Button(botoes_meio, text="Gerar licença", command=self.gerar_licenca).pack(side="left")
        ttk.Button(botoes_meio, text="Copiar licença", command=self.copiar_licenca).pack(side="left", padx=(8, 0))

        nome_frame = ttk.LabelFrame(frame, text="Nome do arquivo ao salvar", padding=10)
        nome_frame.pack(fill="x", pady=(0, 12))

        ttk.Radiobutton(nome_frame, text="licenca.key", value="licenca.key", variable=self.var_nome_arquivo).pack(anchor="w")

        botoes_final = ttk.Frame(frame)
        botoes_final.pack(fill="x", pady=(4, 0))

        ttk.Button(botoes_final, text="Salvar arquivo de licença", command=self.salvar_licenca).pack(side="left")
        ttk.Button(botoes_final, text="Fechar", command=self.destroy).pack(side="right")

        dica = (
            "Dica: o cliente deve colocar o arquivo licenca.key na mesma pasta do .exe ou do .py."
        )
        ttk.Label(frame, text=dica, foreground="#444444", wraplength=700).pack(anchor="w", pady=(18, 0))

    def usar_id_local(self) -> None:
        self.var_machine_id.set(get_machine_id())
        messagebox.showinfo("ID carregado", "O ID desta máquina foi carregado no campo.")

    def copiar_id(self) -> None:
        valor = self.var_machine_id.get().strip().upper()
        if not valor:
            messagebox.showwarning("Atenção", "Informe um ID de máquina.")
            return
        self.clipboard_clear()
        self.clipboard_append(valor)
        self.update()
        messagebox.showinfo("Copiado", "ID copiado para a área de transferência.")

    def gerar_licenca(self) -> None:
        machine_id = self.var_machine_id.get().strip().upper()
        if not machine_id:
            messagebox.showwarning("Atenção", "Informe o ID da máquina.")
            return
        self.var_licenca.set(gerar_hash_licenca(machine_id))

    def copiar_licenca(self) -> None:
        valor = self.var_licenca.get().strip()
        if not valor:
            messagebox.showwarning("Atenção", "Gere a licença primeiro.")
            return
        self.clipboard_clear()
        self.clipboard_append(valor)
        self.update()
        messagebox.showinfo("Copiado", "Licença copiada para a área de transferência.")

    def salvar_licenca(self) -> None:
        licenca = self.var_licenca.get().strip()
        if not licenca:
            messagebox.showwarning("Atenção", "Gere a licença antes de salvar.")
            return
        nome = self.var_nome_arquivo.get().strip() or "licenca.key"
        destino = salvar_arquivo_licenca(licenca, nome)
        messagebox.showinfo("Sucesso", f"Arquivo salvo com sucesso em:\n{destino}")


if __name__ == "__main__":
    app = GeradorLicencaApp()
    app.mainloop()
