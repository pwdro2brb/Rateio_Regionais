import re
import unicodedata
from pathlib import Path
from typing import Union, Tuple
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import numbers

# ==========================================
#              FUNÇÕES ÚTEIS
# ==========================================

def _strip_accents(s: str) -> str:
    if pd.isna(s) or s is None:
        return ''
    s = unicodedata.normalize('NFKD', str(s))
    return ''.join(ch for ch in s if not unicodedata.combining(ch))

def _norm_colname(s: str) -> str:
    return _strip_accents(str(s)).lower().strip()

def _clean_str(x) -> str:
    """Evita o bug do Excel de transformar diagramas numéricos em float (ex: 100000611074.0)"""
    if pd.isna(x): return ''
    if isinstance(x, float):
        if x.is_integer(): return str(int(x))
        return str(x)
    s = str(x).strip()
    if s.endswith('.0'): s = s[:-2]
    return s

def _norm_coletor(x: str) -> str:
    up = _strip_accents(_clean_str(x)).upper()
    return re.sub(r'[^A-Z0-9]', '', up)

def _is_valid_coletor(coletor: str) -> bool:
    if pd.isna(coletor) or str(coletor).strip() == '' or str(coletor).strip().lower() == 'nan':
        return False
    c = _strip_accents(_clean_str(coletor)).upper()
    c_clean = re.sub(r'[^A-Z0-9]', '', c)
    if re.fullmatch(r'\d{6,}', c_clean): 
        return True
    if 6 <= len(c_clean) <= 12 and any(x.isalpha() for x in c_clean) and any(x.isdigit() for x in c_clean): 
        return True
    return False

def _tipo_de_coletor(coletor: str) -> str:
    c = _norm_coletor(coletor)
    if c == 'SEMCENTRODECUSTO': return '-'
    if re.fullmatch(r'\d+', c): return 'N'
    if any(x.isalpha() for x in c) and any(x.isdigit() for x in c): return 'K'
    return ''

def _clean_valor_series(s: pd.Series) -> pd.Series:
    def limpa_valor(val):
        if pd.isna(val) or str(val).strip() == '': return None
        if isinstance(val, (int, float)): return float(val)
        v = str(val).upper().replace('R$', '').replace('\xa0', '').replace(' ', '').strip()
        if '.' in v and ',' in v: v = v.replace('.', '').replace(',', '.')
        elif ',' in v: v = v.replace(',', '.')
        try: return float(v)
        except ValueError: return None
    return s.apply(limpa_valor)

# ==========================================
#     LEITURA: RATEIO RECEBIDO (ISOLAMENTO)
# ==========================================

def ler_rr_bruto(caminho_rr: Union[str, Path]) -> pd.DataFrame:
    xls = pd.ExcelFile(caminho_rr, engine='openpyxl')
    frames = []
    
    for sh in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sh, header=None, engine='openpyxl')
        
        idx_ancora = -1
        max_knf = 0
        
        # 1. Procura a coluna com a maior concentração de K, N, F (Âncora)
        for j in range(df.shape[1]):
            col_str = df.iloc[:, j].astype(str).str.strip().str.upper()
            count_knf = col_str.isin(['K', 'N', 'F']).sum()
            if count_knf > max_knf and count_knf >= 2:
                max_knf = count_knf
                idx_ancora = j
                
        if idx_ancora == -1: 
            continue
            
        # 2. ISOLAMENTO: Recorta a tabela e descarta todo o lixo à esquerda
        df_regional = df.iloc[:, idx_ancora:].copy()
        df_regional.columns = range(df_regional.shape[1]) # Reseta as colunas para 0, 1, 2...
        
        col_coletor, col_subnum, col_valor = -1, -1, -1
        header_row_idx = -1
        
        # 3. Procura os cabeçalhos (da esquerda para a direita) DENTRO da zona isolada
        for i, row in df_regional.head(20).iterrows():
            row_norm = [_norm_colname(str(x)) for x in row.values]
            
            if 'coletor' in row_norm or 'subnumero' in row_norm:
                for j in range(len(row_norm)):
                    val = row_norm[j]
                    if val == 'coletor' and col_coletor == -1: col_coletor = j
                    elif 'subnumero' in val and col_subnum == -1: col_subnum = j
                    elif 'valor' in val and 'servico' not in val and col_valor == -1: col_valor = j
                
                # Fallback para a palavra valor (caso tenha nome composto)
                if col_valor == -1:
                    for j in range(len(row_norm)):
                        if 'valor' in row_norm[j] and col_valor == -1: col_valor = j
                        
                if col_coletor != -1 and col_valor != -1:
                    header_row_idx = i
                    break

        if header_row_idx != -1:
            tmp = df_regional.iloc[header_row_idx + 1:].copy()
            def safe_get(c): return tmp.iloc[:, c] if c != -1 and c < tmp.shape[1] else pd.Series(['']*len(tmp))
            
            tmp_clean = pd.DataFrame({
                'COLETOR_ORIG': safe_get(col_coletor),
                'SUBNUM_ORIG': safe_get(col_subnum),
                'VALOR': safe_get(col_valor)
            })
        else:
            # Fallback (sem cabeçalho): Procura conteúdo da esquerda para direita na zona isolada
            col_coletor_tb = -1
            max_validos = 0
            for j in range(1, df_regional.shape[1]): # Pula a coluna 0 (Âncora K/N/F)
                mask = df_regional.iloc[:, j].astype(str).apply(_is_valid_coletor)
                qtd = mask.sum()
                if qtd > max_validos and qtd >= 2:
                    max_validos = qtd
                    col_coletor_tb = j
                    
            col_valor_tb = -1
            max_nums = 0
            for j in range(1, df_regional.shape[1]):
                if j == col_coletor_tb: continue
                nums = _clean_valor_series(df_regional.iloc[:, j]).notna().sum()
                if nums > max_nums:
                    max_nums = nums
                    col_valor_tb = j
                    
            if col_coletor_tb == -1 or col_valor_tb == -1:
                continue
                
            tmp_clean = pd.DataFrame({
                'COLETOR_ORIG': df_regional.iloc[:, col_coletor_tb],
                'SUBNUM_ORIG': pd.Series(['']*len(df_regional)), 
                'VALOR': df_regional.iloc[:, col_valor_tb]
            })

        # 4. Limpeza e Prioridade de Centro de Custo
        tmp_clean['VALOR'] = _clean_valor_series(tmp_clean['VALOR'])
        tmp_clean = tmp_clean.dropna(subset=['VALOR'])
        
        def get_best_coletor(r):
            sub = _clean_str(r['SUBNUM_ORIG'])
            if _is_valid_coletor(sub): return sub
            col = _clean_str(r['COLETOR_ORIG'])
            if _is_valid_coletor(col): return col
            return None
            
        tmp_clean['COLETOR_FINAL'] = tmp_clean.apply(get_best_coletor, axis=1)
        tmp_clean = tmp_clean.dropna(subset=['COLETOR_FINAL'])
        
        if not tmp_clean.empty:
            tmp_clean['COLETOR'] = tmp_clean['COLETOR_FINAL'].apply(_norm_coletor)
            tmp_clean['TIPOCOLETOR'] = tmp_clean['COLETOR'].apply(_tipo_de_coletor)
            frames.append(tmp_clean[['TIPOCOLETOR', 'COLETOR', 'VALOR']])

    if frames:
        return pd.concat(frames, ignore_index=True)
    return pd.DataFrame(columns=['TIPOCOLETOR', 'COLETOR', 'VALOR'])

# ==========================================
#     LEITURA: CORREIOS (COLA-PEDAÇOS)
# ==========================================

def _extrair_coletor_de_titular(texto: str) -> str:
    if pd.isna(texto) or str(texto).strip() == '': return "SEM CENTRO DE CUSTO"
    t = _strip_accents(str(texto)).upper()
    
    # 1. Diagrama Exato
    m = re.search(r'(?<!\d)(\d{6,})(?!\d)', t)
    if m: return _norm_coletor(m.group(1))
    
    # 2. Centro de Custo Unificado
    palavras = t.split()
    for p in palavras:
        p_clean = re.sub(r'[^A-Z0-9]', '', p)
        if 6 <= len(p_clean) <= 12 and any(c.isalpha() for c in p_clean) and any(c.isdigit() for c in p_clean):
            return p_clean
            
    # 3. Cola-Pedaços (Ex: SU01 SP3080)
    for i in range(len(palavras) - 1):
        p1 = re.sub(r'[^A-Z0-9]', '', palavras[i])
        p2 = re.sub(r'[^A-Z0-9]', '', palavras[i+1])
        comb = p1 + p2
        if 8 <= len(comb) <= 12 and any(c.isalpha() for c in comb) and any(c.isdigit() for c in comb):
            return comb
            
    return "SEM CENTRO DE CUSTO"

def ler_correios_bruto(caminho_correios: Union[str, Path]) -> Tuple[pd.DataFrame, float]:
    df_raw = pd.read_excel(caminho_correios, header=None, engine='openpyxl')
    idx_header = -1
    col_titular, col_valor = -1, -1
    
    # 1. Encontrar o cabeçalho da tabela principal
    for i, row in df_raw.head(20).iterrows():
        row_norm = [_norm_colname(str(x)) for x in row.values]
        if any('titular do cartao' in c for c in row_norm) and any('valor do servico' in c for c in row_norm):
            idx_header = i
            for j, c in enumerate(row_norm):
                if 'titular do cartao' in c: col_titular = j
                if 'valor do servico' in c: col_valor = j
            break
            
    # 2. Encontrar o "Valor Líquido" no final do arquivo
    valor_liquido = 0.0
    idx_fim_tabela = len(df_raw)
    
    # Busca de baixo para cima para encontrar a seção de totais
    for i in range(len(df_raw) - 1, -1, -1):
        row_norm = [_norm_colname(str(x)) for x in df_raw.iloc[i].values]
        
        # Verifica se encontrou a linha de cabeçalho dos totais
        if any('valor liquido' in c for c in row_norm):
            idx_fim_tabela = i  # A tabela de dados principal termina antes dessa linha
            for j, c in enumerate(row_norm):
                if 'valor liquido' in c:
                    # O valor real está na linha de baixo (i + 1), na mesma coluna (j)
                    if i + 1 < len(df_raw):
                        val_raw = df_raw.iloc[i + 1, j]
                        # Reutilizamos a sua função de limpeza passando uma Series de 1 elemento
                        valor_liquido = _clean_valor_series(pd.Series([val_raw])).iloc[0]
                    break
            break

    if idx_header == -1: 
        return pd.DataFrame(columns=['TIPOCOLETOR', 'COLETOR', 'VALOR']), valor_liquido
    
    # 3. Processar a tabela principal (limitando até idx_fim_tabela para ignorar os resumos)
    df = df_raw.iloc[idx_header + 1 : idx_fim_tabela, [col_titular, col_valor]].copy()
    df.columns = ['TITULAR', 'VALOR']
    df['VALOR'] = _clean_valor_series(df['VALOR'])
    df = df.dropna(subset=['VALOR'])
    df['COLETOR'] = df['TITULAR'].apply(_extrair_coletor_de_titular)
    df['TIPOCOLETOR'] = df['COLETOR'].apply(_tipo_de_coletor)
    
    df_final = df[['TIPOCOLETOR', 'COLETOR', 'VALOR']]
    
    return df_final, valor_liquido

# ==========================================
#     CONSTRUÇÃO E CONCILIAÇÃO FINAL
# ==========================================

def _formatar_planilha_final(arquivo_xlsx: Union[str, Path], sheet='Planilha1'):
    wb = load_workbook(arquivo_xlsx)
    if sheet not in wb.sheetnames:
        wb.close(); return
    ws = wb[sheet]
    header = [c.value for c in next(ws.iter_rows(min_row=1, max_row=1))]
    try:
        idx_valor = header.index('VALOR') + 1
        idx_op = header.index('OPERACAO') + 1
    except ValueError:
        wb.close(); return
        
    for row in ws.iter_rows(min_row=2):
        cell_valor = row[idx_valor - 1]
        if isinstance(cell_valor.value, (int, float)):
            cell_valor.number_format = numbers.FORMAT_NUMBER_00
            
    wb.save(arquivo_xlsx)
    wb.close()

def gerar_rateio_pag(
    caminho_correios: Union[str, Path],
    caminho_rr: Union[str, Path],
    saida: Union[str, Path] = 'RATEIO PAG.xlsx',
    operacao_para_diagrama: int = 10,
    tolerancia_igual: float = 0.05, 
    debug: bool = True
) -> pd.DataFrame:

    df_rr_raw = ler_rr_bruto(caminho_rr)
    df_corr_raw, valor_liquido_correios = ler_correios_bruto(caminho_correios)

    total_rr = float(df_rr_raw['VALOR'].sum()) if not df_rr_raw.empty else 0.0
    total_corr_soma = float(df_corr_raw['VALOR'].sum()) if not df_corr_raw.empty else 0.0
    
    # O total oficial dos correios é o Valor Líquido do rodapé (se encontrado)
    total_corr = valor_liquido_correios if valor_liquido_correios > 0 else total_corr_soma

    if debug:
        print(f"[DEBUG] TOTAL RR               = R$ {total_rr:.2f}")
        print(f"[DEBUG] TOTAL CORREIOS (SOMA)  = R$ {total_corr_soma:.2f}")
        print(f"[DEBUG] TOTAL CORREIOS (LÍQ)   = R$ {valor_liquido_correios:.2f}")
        print(f"[DEBUG] DIFERENÇA (LÍQ - RR)   = R$ {round(total_corr - total_rr, 2):.2f}")

    if abs(total_corr - total_rr) <= tolerancia_igual:
        if debug: print("[DEBUG] Valores batem perfeitamente. Usando apenas Recebidos.")
        final_base = df_rr_raw.copy()
    else:
        if debug: print("[DEBUG] Valores divergem. Iniciando Conciliação Avançada...")
        
        # 1. Agrupar os dados ANTES de comparar para evitar duplicação por quebra de linhas
        df_rr_grp = df_rr_raw.groupby(['TIPOCOLETOR', 'COLETOR'], as_index=False)['VALOR'].sum()
        df_corr_grp = df_corr_raw.groupby(['TIPOCOLETOR', 'COLETOR'], as_index=False)['VALOR'].sum()

        coletores_rr = set(df_rr_grp['COLETOR'])
        valores_rr_disponiveis = df_rr_grp['VALOR'].round(2).tolist()
        
        faltantes = []

        def dar_baixa_valor_total(val):
            for i, v in enumerate(valores_rr_disponiveis):
                if abs(v - val) <= 0.02:
                    valores_rr_disponiveis.pop(i)
                    return True
            return False

        for _, row in df_corr_grp.iterrows():
            c = row['COLETOR']
            v = round(row['VALOR'], 2)
            
            if c in coletores_rr:
                # Coletor existe no RR. Pegamos o valor que o RR aprovou.
                val_no_rr = round(df_rr_grp[df_rr_grp['COLETOR'] == c]['VALOR'].sum(), 2)
                
                # Removemos esse valor da lista para não dar falso positivo em renomeações
                dar_baixa_valor_total(val_no_rr) 
                
                # Se os Correios cobraram MAIS do que o RR aprovou, a diferença ficou de fora
                if v > val_no_rr + 0.02:
                    faltantes.append({
                        'TIPOCOLETOR': row['TIPOCOLETOR'],
                        'COLETOR': c,
                        'VALOR': v - val_no_rr
                    })
            else:
                # Coletor NÃO existe no RR.
                if dar_baixa_valor_total(v):
                    # O valor bate exatamente com outro coletor do RR. 
                    # Assumimos que foi apenas renomeado/corrigido pela regional. Não adiciona.
                    pass
                else:
                    # Realmente faltante (regional excluiu ou esqueceu).
                    faltantes.append(row.to_dict())

        if faltantes:
            df_faltantes = pd.DataFrame(faltantes)
            final_base = pd.concat([df_rr_grp, df_faltantes], ignore_index=True)
            if debug: print(f"[DEBUG] Encontrados {len(faltantes)} centros de custo com divergência.")
        else:
            final_base = df_rr_grp.copy()
            
        # TRAVA DE SEGURANÇA: Impede que a conciliação gere um valor maior que a fatura
        soma_final = final_base['VALOR'].sum()
        if soma_final > total_corr + 0.50:
            if debug: print("[DEBUG] ALERTA: A conciliação detalhada excedeu o total da fatura. Aplicando Ajuste Global.")
            diferenca_global = total_corr - total_rr
            if diferenca_global > 0.02:
                linha_ajuste = pd.DataFrame([{
                    'TIPOCOLETOR': '-',
                    'COLETOR': 'AJUSTE DE CONCILIACAO',
                    'VALOR': diferenca_global
                }])
                final_base = pd.concat([df_rr_grp, linha_ajuste], ignore_index=True)
            else:
                final_base = df_rr_grp.copy()

    # Agrupa tudo novamente para garantir que a tabela final seja limpa
    final_base = final_base.groupby(['TIPOCOLETOR', 'COLETOR'], as_index=False)['VALOR'].sum()

    # Monta Tabela Final
    final = pd.DataFrame()
    final['ITEM'] = [1] * len(final_base)
    final['TIPOCOLETOR'] = final_base['TIPOCOLETOR']
    final['COLETOR'] = final_base['COLETOR']
    final['OPERACAO'] = final_base['TIPOCOLETOR'].apply(lambda t: operacao_para_diagrama if t == 'N' else '')
    final['SUBNUMERO'] = ''
    final['VALOR'] = final_base['VALOR']
    final['DESCRICAO'] = ''

    # Ordenação (K -> N -> Faltantes no Final)
    final['__ord'] = final['TIPOCOLETOR'].map({'K': 0, 'N': 1}).fillna(2)
    final = final.sort_values(['__ord', 'COLETOR']).drop(columns='__ord').reset_index(drop=True)

    saida = Path(saida)
    with pd.ExcelWriter(saida, engine='openpyxl') as writer:
        final.to_excel(writer, sheet_name='Planilha1', index=False)

    _formatar_planilha_final(saida, 'Planilha1')

    if debug: print(f"[DEBUG] Arquivo gerado com sucesso: {saida.resolve()}")

    return final

# ==========================================
#     INICIALIZAÇÃO AUTOMÁTICA
# ==========================================
if __name__ == '__main__':
    # >>> MUDE AQUI O CAMINHO DA SUA PASTA <<<
    pasta_trabalho = Path(r"C:/Users/pedro.henrsilva/OneDrive - MRV/Área de Trabalho/Rateio_Regionais/testar_edicao")
    
    caminho_rr = None
    caminho_correios = None
    
    # Vasculha a pasta procurando os arquivos dinamicamente
    for ficheiro in pasta_trabalho.glob('*.xlsx'):
        nome_ficheiro = ficheiro.name.upper()
        if nome_ficheiro.startswith('~$') or nome_ficheiro == 'RATEIO PAG.XLSX': continue
        if 'RATEIO RECEBIDO' in nome_ficheiro: caminho_rr = ficheiro
        elif re.match(r'^\d{7}\.XLSX$', nome_ficheiro): caminho_correios = ficheiro

    if not caminho_rr or not caminho_correios:
        print("ERRO: Faltam planilhas na pasta!")
        if not caminho_rr: print("- Faltou a planilha com nome 'RATEIO RECEBIDO'.")
        if not caminho_correios: print("- Faltou a planilha dos Correios (ex: 1234567.xlsx).")
    else:
        print(f"\n>> Iniciando processamento...")
        print(f"Planilha Correios: {caminho_correios.name}")
        print(f"Planilha Rateio: {caminho_rr.name}\n")
        
        caminho_saida = pasta_trabalho / "RATEIO PAG.xlsx"
        
        final = gerar_rateio_pag(caminho_correios=caminho_correios, caminho_rr=caminho_rr, saida=caminho_saida)
        print(f"\nTotal de linhas geradas no final: {len(final)}")