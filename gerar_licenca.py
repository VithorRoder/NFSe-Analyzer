import hashlib

SECRET = "NFSE_ANALYZER_2026_RODER"


def gerar_hash_licenca(machine_id: str) -> str:
    texto = machine_id + SECRET
    return hashlib.sha256(texto.encode("utf-8")).hexdigest()


def main():
    machine_id = input("Digite o ID da máquina: ").strip()
    licenca = gerar_hash_licenca(machine_id)

    print("\nLicença gerada:\n")
    print(licenca)

    salvar = input("\nDeseja salvar em licenca.key? (s/n): ").strip().lower()
    if salvar == "s":
        with open("licenca.key", "w", encoding="utf-8") as f:
            f.write(licenca)
        print("\nArquivo licenca.key salvo com sucesso.")


if __name__ == "__main__":
    main()