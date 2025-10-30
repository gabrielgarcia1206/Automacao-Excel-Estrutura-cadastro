import pandas as pd
print("\n")
print("🔧 GERADOR DE ESTRUTURA — CORRIGIDO (SEM LOOP DUPLICADO)\n")

# === CONFIGURAÇÕES ===
arquivo = "dados.xlsx"
saida = "estrutura_gerada.xlsx"

MAPA_MOD = {
    "CC": "MOD17001",
    "EE": "MOD17002",
    "CB": "MOD17003",
    "UG": "MOD17004",
    "SS": "MOD17005",
    "PP": "MOD17006",
    "MM": "MOD17007",
}

# === LEITURA ===
df = pd.read_excel(arquivo, dtype=str).fillna("0")
df.columns = [c.strip().upper() for c in df.columns]

NIVEL = "NIVEL"
COD_LETRA = "COD+LETRA"
LETRA = "LETRA"
QTD = "QTD"
MATERIAL = "MATERIAL"
PESO = "PESO"
JA_CAD = "JÁ CADASTRADOS" if "JÁ CADASTRADOS" in df.columns else "JA CADASTRADOS"
OP1 = "OPERAÇÃO1" if "OPERAÇÃO1" in df.columns else ("OPERAÇÃO 1" if "OPERAÇÃO 1" in df.columns else "0")
OP2 = "OPERAÇÃO 2" if "OPERAÇÃO 2" in df.columns else "0"
OP3 = "OPERAÇÃO 3" if "OPERAÇÃO 3" in df.columns else "0"

df[NIVEL] = pd.to_numeric(df[NIVEL], errors="coerce").fillna(0).astype(int)
df[QTD] = pd.to_numeric(df[QTD], errors="coerce").fillna(0)
df[PESO] = pd.to_numeric(df[PESO], errors="coerce").fillna(0)


def nao_cadastrado(v: str) -> bool:
    s = str(v).strip().upper()
    return s in ("", "0", "#N/D")


def mods_do_item(row) -> list[str]:
    letra_final = str(row[LETRA]).strip().upper()
    if letra_final == "PI":
        return []
    ops_raw = [
        str(row.get(OP1, "0")).strip().upper(),
        str(row.get(OP2, "0")).strip().upper(),
        str(row.get(OP3, "0")).strip().upper(),
    ]
    ops = [op for op in ops_raw if op in MAPA_MOD]
    if letra_final in MAPA_MOD and letra_final in ops:
        ops = ops[: ops.index(letra_final) + 1]
    return [MAPA_MOD[op] for op in ops]


def gerar_estrutura(df: pd.DataFrame) -> pd.DataFrame:
    saida = []
    linhas = set()
    pais_processados = set()

    total_linhas = len(df)
    i = 0

    while i < total_linhas:
        row = df.iloc[i]
        nivel = int(row[NIVEL])
        codigo = str(row[COD_LETRA]).strip()
        qtd = float(row[QTD])
        letra = str(row[LETRA]).strip().upper()
        material = str(row[MATERIAL]).strip().upper()
        peso = float(row[PESO])
        ja_val = str(row.get(JA_CAD, "0")).strip().upper()

        # Só processa pais não cadastrados
        if nao_cadastrado(ja_val):
            # Garante que não processe o mesmo pai novamente
            if codigo not in pais_processados:
                pais_processados.add(codigo)

                # --- filhos diretos ---
                j = i + 1
                while j < total_linhas:
                    prox = df.iloc[j]
                    nv = int(prox[NIVEL])

                    if nv <= nivel:
                        break  # encerra hierarquia do pai
                    if nv == nivel + 1:
                        filho = str(prox[COD_LETRA]).strip()
                        qtd_f = float(prox[QTD])
                        if (codigo, filho) not in linhas:
                            saida.append([codigo, filho, qtd_f, 0])
                            linhas.add((codigo, filho))
                    j += 1

                # --- Matéria-prima ---
                if material != "0" and peso > 0:
                    if (codigo, material) not in linhas:
                        saida.append([codigo, material, peso, 0])
                        linhas.add((codigo, material))

                # --- MODs ---
                for mod in mods_do_item(row):
                    if (codigo, mod) not in linhas:
                        saida.append([codigo, mod, 1, 0])
                        linhas.add((codigo, mod))

        # avança para próxima linha
        i += 1

    return pd.DataFrame(saida, columns=["Código", "Componente", "Quantidade", "%Perda"])


resultado = gerar_estrutura(df)
resultado.to_excel(saida, index=False)
print("\n")
print("\n✅ Estrutura final gerada com sucesso.")
print("Arquivo salvo como:", saida)
print("Total de linhas:", len(resultado))
print("\n")
print("\n")
print("\n Dev by:Gabriel Bueno Garcia.")
print("\n")
