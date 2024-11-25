import streamlit as st
from docxtpl import DocxTemplate
import zipfile
import pandas as pd
from datetime import date
import os


st.set_page_config(
    page_title="Sales Dashboard",
    layout="wide",
    initial_sidebar_state='collapsed',
    )


def safe_strftime(date_value, date_format='%d/%m/%Y'):
    """Converte uma data para string no formato especificado ou retorna uma string vazia."""
    return date_value.strftime(date_format) if date_value else ''


# Inicialização da variável de controle de página - NAVEGACAO
if 'page' not in st.session_state:
    st.session_state.page = 'resolucao26' 

def go_to_main():
    st.session_state.page = 'main'

def go_to_viagem():
    st.session_state.page = 'viagem'

def go_to_main():
    st.session_state.page = 'main'

def go_to_termoderenuncia():
    st.session_state.page = 'termoderenuncia'

def go_to_solicitacao_diaria():
    st.session_state.page = 'solicitacao_diaria'

def avancar_para_resolucao26():
    st.session_state.page = 'resolucao26' 
    st.session_state.page = 'resolucao26' 



# Preencher Docs
def preencher_documento(data, path):
    template = DocxTemplate("anexo_da_resolucao_26-2023-ct.docx")
    context = {
        'DATA_INICIO_AFASTAMENTO': data['inicio_afastamento'],
        'DATA_FIM_AFASTAMENTO': data['fim_afastamento'],
        'NOMEDOEVENTO': data['nomedo_evento'],
        'LUGAR_EVENTO': data['lugar_evento'],
        'DATA_INICIO_EVENTO': data['data_inicio_evento'],
        'DATA_FIM_EVENTO': data['data_fim_evento'],
        'ATIVIDADE': data['atividade'],
        'opcao_a': 'X' if data['opcao'] == 'a' else '',
        'opcao_b': 'X' if data['opcao'] == 'b' else '',
        'opcao_c': 'X' if data['opcao'] == 'c' else '',
        'opcao_d': 'X' if data['opcao'] == 'd' else '',
        'declaracao_a': 'X' if data['declaracao'] == 'a' else '',
        'declaracao_b': 'X' if data['declaracao'] == 'b' else '',
        'declaracao_c': 'X' if data['declaracao'] == 'c' else '',
        'declaracao_final_a': 'X' if data['declaracao_final'] == 'a' else '',
        'declaracao_final_b': 'X' if data['declaracao_final'] == 'b' else '',
        'tabela_reposicoes': data['tabela_reposicoes'],
        'tabela_substitutos': data['tabela_substitutos']
    }
    template.render(context)
    template.save(path)
    return path


# Função para validar e converter datas
def validate_date(key):
    value = st.session_state.get(key, None)
    if value:
        try:
            # Converte para datetime.date se necessário
            return pd.to_datetime(value).date() if not isinstance(value, date) else value
        except Exception:
            return None  # Retorna None caso a conversão falhe
    return None


#Diaria e Passagem
def preencher_documento2(data, path):
    template = DocxTemplate("solicitacao_de_passagem_e_diaria_0-2.docx")
    
    # Verificação da opção "Não Desejo Diária e Passagem"
    if data['solicitacaodiaria'] == "Não Desejo Diária e Passagem":
        lugar_de_ida = ''
        trecho_ida = ''
        trecho_volta = ''
        cia_area_ida = ''
        cia_area_volta = ''
        numero_voo_ida = ''
        numero_voo_volta = ''
        data_partida_ida = ''
        data_chegada_ida = ''
        data_partida_volta = ''
        data_chegada_volta = ''
        solicitacao_diaria_x1 = ''
        solicitacao_diaria_x2 = ''
        roteiro_viagem = ''
    else:
        lugar_de_ida = data['lugar_de_ida']
        roteiro_viagem = data['roteiro_viagem'] if data['destino_nao_ter_aeroporto'] == "Não" else ''
        trecho_ida = data['roteiro_ida']
        trecho_volta = data['roteiro_volta']
        cia_area_ida = data['cia_area_ida']
        cia_area_volta = data['cia_area_volta']
        numero_voo_ida = data['numerovoo_ida']
        numero_voo_volta = data['numerovoo_volta']
        data_partida_ida = data['voo_ida_1']
        data_chegada_ida = data['voo_volta_1']
        data_partida_volta = data['voo_ida_2']
        data_chegada_volta = data['voo_volta_2']
        solicitacao_diaria_x1 = 'X' if data['solicitacaodiaria'] in ["Passagem Aérea", "Os Dois"] else ''
        solicitacao_diaria_x2 = 'X' if data['solicitacaodiaria'] in ["Diárias", "Os Dois"] else ''
        roteiro_viagem = data['roteiro_viagem']
    
    context = {
        'LUGAR_EVENTO': data['lugar_evento'],
        'SOLICITACAO_DIARIA_X1': solicitacao_diaria_x1,
        'SOLICITACAO_DIARIA_X2': solicitacao_diaria_x2,
        'SEXO_X1': 'X' if data['sexo'] == "Masculino" else '',
        'SEXO_X2': 'X' if data['sexo'] == "Feminino" else '',
        'TRANSPORTE_X1': 'X' if data['transporte'] == "Aéreo" else '',
        'TRANSPORTE_X2': 'X' if data['transporte'] == "Veículo Oficial" else '',
        'TRANSPORTE_X3': 'X' if data['transporte'] == "Veículo Próprio" else '',
        'DESTINO_SEM_AEROPORTO_X1': 'X' if data['destino_nao_ter_aeroporto'] == "Sim" else '',
        'DESTINO_SEM_AEROPORTO_X2': 'X' if data['destino_nao_ter_aeroporto'] == "Não" else '',
        'BAGAGEM_X1': 'X' if data['bagagem'] == "Sim" else '',
        'BAGAGEM_X2': 'X' if data['bagagem'] == "Não" else '',
        'NOME': data['nome'],
        'CPF': data['cpf'],
        'DATA_DE_NASCIMENTO': data['datadenascimento'],
        'EMAIL_PESSOAL': data['emailpessoal'],
        'RG': data['rg'],
        'TELEFONE_PESSOAL': data['telefonepessoal'],
        'VINCULO_SERVIDOR': data['vinculo_servidor'],
        'VINCULO_ALUNO': data['vinculo_aluno'],
        'VINCULO_CONVIDADO': data['vinculo_convidado'],
        'DATA_INICIO_EVENTO': data['data_inicio_evento'],
        'DATA_FIM_EVENTO': data['data_fim_evento'],
        'ATIVIDADE': data['atividade'],
        'DATA_INICIO_AFASTAMENTO': data['data_inicio_afastamento'],
        'DATA_FIM_AFASTAMENTO': data['data_fim_afastamento'],
        'LUGAR_DE_IDA': lugar_de_ida,
        'TRECHO_IDA': trecho_ida,
        'TRECHO_VOLTA': trecho_volta,
        'CIA_AEREA_IDA': cia_area_ida,
        'CIA_AEREA_VOLTA': cia_area_volta,
        'NUMERO_VOO_IDA': numero_voo_ida,
        'NUMERO_VOO_VOLTA': numero_voo_volta,
        'DATA_PARTIDA_IDA': data_partida_ida,
        'DATA_CHEGADA_IDA': data_chegada_ida,
        'DATA_PARTIDA_VOLTA': data_partida_volta,
        'DATA_CHEGADA_VOLTA': data_chegada_volta,
        'ROTEIRO_VIAGEM': roteiro_viagem,
    }
    
    template.render(context)
    template.save(path)
    return path

# Renuncia
def preencher_documento3(data, path):
    if data['solicitacaodiaria'] == "Não Desejo Diária e Passagem":
        template = DocxTemplate("termo_de_renuncia_de_diarias_e_passagens_0_1-2.docx")
        context = {
            'NOME': data['nome'],
            'CPF': data['cpf'],
            'SIAPE': data['siape'],
            'MOTIVO_DA_RENUNCIA': data['motivodarenuncia'],
            'DIARIAS_PARCIAL': 'X' if 'Diária Parcial' in data['renuncias'] else '',
            'DIARIAS_INTEGRAL': 'X' if 'Diária Integral' in data['renuncias'] else '',
            'PASSAGENS_IDA': 'X' if 'Passagem de Ida' in data['renuncias'] else '',
            'PASSAGENS_VOLTA': 'X' if 'Passagem de Volta' in data['renuncias'] else '',
        }
        template.render(context)
        template.save(path)
        return path  # Retorna o caminho do documento gerado
    else:
        # Não gera o documento, retorna None
        return None
    

# Bloco Principal
if st.session_state.page == 'main':
    st.title("Formulário de Preenchimento de Documento")

    nome = st.text_input("Nome", value=st.session_state.get('nome', ''))
    cpf = st.text_input("CPF", value=st.session_state.get('cpf', ''))
    rg = st.text_input("RG", value=st.session_state.get('rg', ''))
    datadenascimento = st.date_input("Data de Nascimento", value=st.session_state.get('datadenascimento', None))
    email = st.text_input("Email Pessoal", value=st.session_state.get('email', ''))
    telefone = st.text_input("Telefone", value=st.session_state.get('telefone', ''))
    SIAPE = st.text_input("SIAPE", value=st.session_state.get('SIAPE', ''))
    sexo = st.radio("Selecione o Sexo", options=["Masculino", "Feminino"], index=0 if st.session_state.get('sexo', 'Masculino') == "Masculino" else 1)
    vinculo = st.radio("Selecione o Vínculo", options=["Servidor Ufes", "Aluno", "Convidado", "Estrangeiro", "Nome da Mãe"], index=["Servidor Ufes", "Aluno", "Convidado", "Estrangeiro", "Nome da Mãe"].index(st.session_state.get('vinculo', 'Servidor Ufes')))

    # Armazenando as informações no session_state
    st.session_state.nome = nome
    st.session_state.cpf = cpf
    st.session_state.rg = rg
    st.session_state.datadenascimento = datadenascimento
    st.session_state.email = email
    st.session_state.telefone = telefone
    st.session_state.SIAPE = SIAPE
    st.session_state.sexo = sexo
    st.session_state.vinculo = vinculo

    col1, col2 = st.columns(2)
    with col1:
        st.button("Voltar", on_click=avancar_para_resolucao26)
    with col2:
         st.button("Próximo", on_click=go_to_viagem)

    

# Bloco Dados de Viagem
elif st.session_state.page == 'viagem':
    st.title("Formulário de Preenchimento de Documento - Dados Viagem")

    pedir_diaria_passagem = st.radio(
        "Deseja pedir Diárias e Passagem Aérea?", 
        options=["Não Desejo Diária e Passagem", "Diárias", "Passagem Aérea", "Os Dois"], 
        index=["Não Desejo Diária e Passagem", "Diárias", "Passagem Aérea", "Os Dois"].index(
            st.session_state.get('pedir_diaria_passagem', 'Não Desejo Diária e Passagem')
        )
    )

    data_inicio_evento = st.date_input(
        "Data Início do Evento", 
        value=validate_date('data_inicio_evento')
    )
    data_fim_evento = st.date_input(
        "Data Fim do Evento", 
        value=validate_date('data_fim_evento')
    )
    atividade = st.text_area(
        "Descreva seu compromisso", 
        value=st.session_state.get('atividade', '')
    )
    data_inicio_afastamento = st.date_input(
        "Data Início Afastamento", 
        value=validate_date('data_inicio_afastamento')
    )
    data_fim_afastamento = st.date_input(
        "Data Fim Afastamento", 
        value=validate_date('data_fim_afastamento')
    )
    lugar_de_ida = st.text_input(
        "Origem do Deslocamento (Município/Estado):", 
        value=st.session_state.get('lugar_de_ida', '')
    )
    lugar_evento = st.text_input(
        "Destino Final do Deslocamento (Município/Estado):", 
        value=st.session_state.get('lugar_evento', '')
    )
    transporte = st.radio(
        "Selecione o transporte", 
        options=["Aéreo", "Veículo Oficial", "Veículo Próprio"], 
        index=["Aéreo", "Veículo Oficial", "Veículo Próprio"].index(
            st.session_state.get('transporte', 'Aéreo')
        )
    )
    destino_nao_ter_aeroporto = st.radio(
        "O destino possui aeroporto?", 
        options=["Sim", "Não"], 
        index=["Sim", "Não"].index(
            st.session_state.get('destino_nao_ter_aeroporto', 'Sim')
        )
    )
    bagagem = st.radio(
        "Uma bagagem de 23kg a partir de 3 pernoites?", 
        options=["Sim", "Não"], 
        index=["Sim", "Não"].index(
            st.session_state.get('bagagem', 'Sim')
        )
    )
    if destino_nao_ter_aeroporto == 'Não':
        st.write("Se o destino não possui aeroporto, especifique o roteiro de viagem aqui:")
        roteiro_viagem = st.text_input(
            "Escreva o roteiro", 
            value=st.session_state.get('roteiro_viagem', '')
        )
    else:
        roteiro_viagem = ''

    if pedir_diaria_passagem in ['Passagem Aérea', 'Os Dois', 'Diárias']:
        st.write("Informações sobre os voos:")
        roteiro_ida = st.text_input(
            "Digite nesse formato: Cidade de Origem – Cidade de Destino (Ida)", 
            value=st.session_state.get('roteiro_ida', '')
        )
        roteiro_volta = st.text_input(
            "Digite nesse formato: Cidade de Origem – Cidade de Destino (Volta)", 
            value=st.session_state.get('roteiro_volta', '')
        )
        cia_area_ida = st.text_input(
            "Companhia Aérea (Ida)", 
            value=st.session_state.get('cia_area_ida', '')
        )
        cia_area_volta = st.text_input(
            "Companhia Aérea (Volta)", 
            value=st.session_state.get('cia_area_volta', '')
        )
        numerovoo_ida = st.text_input(
            "Numero voo (Ida)", 
            value=st.session_state.get('numerovoo_ida', '')
        )
        numerovoo_volta = st.text_input(
            "Numero voo (Volta)", 
            value=st.session_state.get('numerovoo_volta', '')
        )
        voo_ida_1 = st.date_input(
            "Data e Hora Partida (Ida)", 
            value=validate_date('voo_ida_1')
        )
        voo_volta_1 = st.date_input(
            "Data e Hora Chegada (Ida)", 
            value=validate_date('voo_volta_1')
        )
        voo_ida_2 = st.date_input(
            "Data e Hora Partida (Volta)", 
            value=validate_date('voo_ida_2')
        )
        voo_volta_2 = st.date_input(
            "Data e Hora Chegada (Volta)", 
            value=validate_date('voo_volta_2')
        )
    else:
        roteiro_ida = roteiro_volta = cia_area_ida = cia_area_volta = ''
        voo_ida_1 = voo_volta_1 = voo_ida_2 = voo_volta_2 = numerovoo_ida = numerovoo_volta = ''

    # Armazenando as informações no session_state
    st.session_state.numerovoo_ida = numerovoo_ida
    st.session_state.numerovoo_volta = numerovoo_volta
    st.session_state.pedir_diaria_passagem = pedir_diaria_passagem
    st.session_state.data_inicio_evento = data_inicio_evento
    st.session_state.data_fim_evento = data_fim_evento
    st.session_state.atividade = atividade
    st.session_state.data_inicio_afastamento = data_inicio_afastamento
    st.session_state.data_fim_afastamento = data_fim_afastamento
    st.session_state.lugar_de_ida = lugar_de_ida
    st.session_state.lugar_evento = lugar_evento
    st.session_state.transporte = transporte
    st.session_state.destino_nao_ter_aeroporto = destino_nao_ter_aeroporto
    st.session_state.bagagem = bagagem
    st.session_state.roteiro_viagem = roteiro_viagem
    st.session_state.roteiro_ida = roteiro_ida
    st.session_state.roteiro_volta = roteiro_volta
    st.session_state.cia_area_ida = cia_area_ida
    st.session_state.cia_area_volta = cia_area_volta
    st.session_state.voo_ida_1 = voo_ida_1
    st.session_state.voo_volta_1 = voo_volta_1
    st.session_state.voo_ida_2 = voo_ida_2
    st.session_state.voo_volta_2 = voo_volta_2
    st.session_state.roteiro = roteiro_viagem

    col1, col2 = st.columns(2)
    with col1:
        st.button("Voltar", on_click=go_to_main)
    with col2:
        def avancar():
            if pedir_diaria_passagem == 'Não Desejo Diária e Passagem':
                go_to_termoderenuncia()
            else:
                go_to_solicitacao_diaria()
        st.button("Próximo", on_click=avancar)

# Bloco Termo de Renúncia
elif st.session_state.page == 'termoderenuncia':
    st.title("Formulário de Preenchimento do Termo de Renúncia")

    SIAPE = st.session_state.get('SIAPE', '')
    nome = st.session_state.get('nome', '')
    cpf = st.session_state.get('cpf', '')

    motivodarenuncia = st.text_area(
        "Descreva seu motivo de renúncia:", 
        value=st.session_state.get('motivodarenuncia', '')
    )

    # Todas as opções de renúncia são pré-selecionadas
    renuncias_opcoes = ["Diária Parcial", "Diária Integral", "Passagem de Ida", "Passagem de Volta"]
    renuncias_selecionadas = renuncias_opcoes  # Todas as opções estão selecionadas

    st.session_state.motivodarenuncia = motivodarenuncia
    st.session_state.renuncias = renuncias_selecionadas

    col1, col2 = st.columns(2)
    with col1:
        st.button("Voltar", on_click=go_to_viagem)
    with col2:
        st.button("Gerar Documentos", on_click=go_to_solicitacao_diaria)

# Bloco Resolução 26
elif st.session_state.page == 'resolucao26':
    st.title("Formulário de Preenchimento de Documento Resolução 26")

    # Inputs gerais do evento
    nomedo_evento = st.text_input(
        "Nome do Evento",
        value=st.session_state.get('nomedo_evento', '')
    )
    lugar_evento = st.text_input(
        "Destino Final do Evento (Município/Estado):",
        value=st.session_state.get('lugar_evento', '')
    )

    # Mapeamento das opções para identificadores
    opcoes_opcao_dict = {
        'a': "com ônus Ufes (manutenção de salário + auxílio de viagem como diárias e/ou passagem pagos pela Ufes – Proap, Centro ou Pró-Reitoria)",
        'b': "com ônus Agência Financiadora (manutenção de salário + auxílio de viagem como diárias e/ou passagem pagos por agência – CAPES, CNPq, Fapes ou outra)",
        'c': "com ônus limitado (apenas com manutenção do salário e vantagens)",
        'd': "sem ônus (suspensão do salário e vantagens durante o afastamento)"
    }
    opcoes_opcao = list(opcoes_opcao_dict.values())
    opcao_atual_key = st.session_state.get('opcao', 'a')
    opcao_atual = opcoes_opcao_dict[opcao_atual_key]
    opcao_selecionada = st.radio(
        "Selecione a Opção",
        options=opcoes_opcao,
        index=opcoes_opcao.index(opcao_atual)
    )
    # Obter a chave correspondente
    opcao_key = [key for key, value in opcoes_opcao_dict.items() if value == opcao_selecionada][0]

    # Mapeamento para 'declaracao'
    opcoes_declaracao_dict = {
        'a': "Não há atividade de aula (teoria, exercícios ou laboratório) no período solicitado.",
        'b': "As aulas ministradas serão repostas por mim conforme quadro abaixo (Pode-se usar horas excedentes dentro do calendário acadêmico, se houver).",
        'c': "Os/as seguintes docentes serão meus substitutos/as nas datas previstas das aulas afetadas pelo afastamento, conforme quadro abaixo."
    }
    opcoes_declaracao = list(opcoes_declaracao_dict.values())
    declaracao_atual_key = st.session_state.get('declaracao', 'a')
    declaracao_atual = opcoes_declaracao_dict[declaracao_atual_key]
    declaracao_selecionada = st.radio(
        "Selecione a Declaração",
        options=opcoes_declaracao,
        index=opcoes_declaracao.index(declaracao_atual)
    )
    # Obter a chave correspondente
    declaracao_key = [key for key, value in opcoes_declaracao_dict.items() if value == declaracao_selecionada][0]

    # Mapeamento para 'declaracao_final'
    opcoes_declaracao_final_dict = {
        'a': "Trata-se de afastamento no país e atende às normas internas do departamento, podendo ser apresentado à Câmara Departamental para análise;",
        'b': "Trata-se de afastamento para o exterior e está instruído conforme instruções da PRPPG, contidas em seu site: https://prppg.ufes.br/afastamento-para-eventos-cientificos-e-outras-atividades-academicas-no-exterior"
    }
    opcoes_declaracao_final = list(opcoes_declaracao_final_dict.values())
    declaracao_final_atual_key = st.session_state.get('declaracao_final', 'a')
    declaracao_final_atual = opcoes_declaracao_final_dict[declaracao_final_atual_key]
    declaracao_final_selecionada = st.radio(
        "Selecione a Declaração Final",
        options=opcoes_declaracao_final,
        index=opcoes_declaracao_final.index(declaracao_final_atual)
    )
    # Obter a chave correspondente
    declaracao_final_key = [key for key, value in opcoes_declaracao_final_dict.items() if value == declaracao_final_selecionada][0]

    # Atualizando as informações gerais no session_state
    st.session_state.nomedo_evento = nomedo_evento
    st.session_state.lugar_evento = lugar_evento
    st.session_state.opcao = opcao_key
    st.session_state.declaracao = declaracao_key
    st.session_state.declaracao_final = declaracao_final_key

    # Recuperando os valores das tabelas no session_state ou definindo valores padrão
    tabela_reposicoes = st.session_state.get('tabela_reposicoes', [])
    tabela_substitutos = st.session_state.get('tabela_substitutos', [])

    if declaracao_key == 'b':
        st.write("Preencha a Tabela de Reposições:")
        num_linhas_reposicoes = st.number_input(
            "Número de Reposições", 
            min_value=1, 
            max_value=10, 
            step=1, 
            value=len(tabela_reposicoes) if tabela_reposicoes else 1
        )
        tabela_reposicoes_atualizada = []
        for i in range(int(num_linhas_reposicoes)):
            disciplina = st.text_input(
                f"Disciplina {i + 1}", 
                value=tabela_reposicoes[i]['disciplina'] if i < len(tabela_reposicoes) else ''
            )
            data_aula = st.date_input(
                f"Data Aula Afetada {i + 1}", 
                value=pd.to_datetime(tabela_reposicoes[i]['data_aula_afetada']).date() if i < len(tabela_reposicoes) and tabela_reposicoes[i]['data_aula_afetada'] else None
            )
            data_reposicao = st.date_input(
                f"Data Reposição {i + 1}", 
                value=pd.to_datetime(tabela_reposicoes[i]['data_reposicao']).date() if i < len(tabela_reposicoes) and tabela_reposicoes[i]['data_reposicao'] else None
            )
            tabela_reposicoes_atualizada.append({
                'disciplina': disciplina,
                'data_aula_afetada': data_aula.strftime('%d/%m/%Y') if data_aula else '',
                'data_reposicao': data_reposicao.strftime('%d/%m/%Y') if data_reposicao else ''
            })

        tabela_reposicoes = tabela_reposicoes_atualizada

    elif declaracao_key == 'c':
        st.write("Preencha a Tabela de Substitutos:")
        num_linhas_substitutos = st.number_input(
            "Número de Substituições", 
            min_value=1, 
            max_value=10, 
            step=1, 
            value=len(tabela_substitutos) if tabela_substitutos else 1
        )
        tabela_substitutos_atualizada = []
        for i in range(int(num_linhas_substitutos)):
            disciplina = st.text_input(
                f"Disciplina {i + 1}", 
                value=tabela_substitutos[i]['disciplina'] if i < len(tabela_substitutos) else ''
            )
            data_aula = st.date_input(
                f"Data Aula Afetada {i + 1}", 
                value=pd.to_datetime(tabela_substitutos[i]['data_aula_afetada']).date() if i < len(tabela_substitutos) and tabela_substitutos[i]['data_aula_afetada'] else None
            )
            professor = st.text_input(
                f"Professor Substituto {i + 1}", 
                value=tabela_substitutos[i]['professor_substituto'] if i < len(tabela_substitutos) else ''
            )
            tabela_substitutos_atualizada.append({
                'disciplina': disciplina,
                'data_aula_afetada': data_aula.strftime('%d/%m/%Y') if data_aula else '',
                'professor_substituto': professor
            })

        tabela_substitutos = tabela_substitutos_atualizada

    # Atualizando as tabelas no session_state
    st.session_state.tabela_reposicoes = tabela_reposicoes
    st.session_state.tabela_substitutos = tabela_substitutos

    col1, col2 = st.columns(2)
    with col1:
        st.button("Continuar", on_click=go_to_main)

elif st.session_state.page == 'solicitacao_diaria':
    # Gerar o documento
    dados_resolucao26 = {
        'inicio_afastamento': safe_strftime(st.session_state.data_inicio_afastamento),
        'fim_afastamento': safe_strftime(st.session_state.data_fim_afastamento),
        'nomedo_evento': st.session_state.nomedo_evento,
        'lugar_evento': st.session_state.lugar_evento,
        'data_inicio_evento': safe_strftime(st.session_state.data_inicio_evento),
        'data_fim_evento': safe_strftime(st.session_state.data_fim_evento),
        'atividade': st.session_state.atividade,
        'opcao': st.session_state.opcao,  # Agora contém 'a', 'b', 'c' ou 'd'
        'declaracao': st.session_state.declaracao,  # Agora contém 'a', 'b' ou 'c'
        'tabela_reposicoes': st.session_state.tabela_reposicoes,
        'tabela_substitutos': st.session_state.tabela_substitutos,
        'declaracao_final': st.session_state.declaracao_final  # Agora contém 'a' ou 'b'
    }

    caminho_docx = "Docs/resolucao26.docx"
    docx_path = preencher_documento(dados_resolucao26, caminho_docx)

    dados_solicitacaodiaria_positivo = {
        'solicitacaodiaria': st.session_state.pedir_diaria_passagem,
        'nome': st.session_state.nome,
        'cpf': st.session_state.cpf,
        'sexo': st.session_state.sexo,
        'roteiro': st.session_state.roteiro,
        'datadenascimento': safe_strftime(st.session_state.datadenascimento),
        'emailpessoal': st.session_state.email,
        'rg': st.session_state.rg,
        'telefonepessoal': st.session_state.telefone,
        'vinculo': st.session_state.vinculo,
        'transporte': st.session_state.transporte,
        'data_inicio_evento': safe_strftime(st.session_state.data_inicio_evento),
        'data_fim_evento': safe_strftime(st.session_state.data_fim_evento),
        'atividade': st.session_state.atividade,
        'data_inicio_afastamento': safe_strftime(st.session_state.data_inicio_afastamento),
        'data_fim_afastamento': safe_strftime(st.session_state.data_fim_afastamento),
        'lugar_de_ida': st.session_state.lugar_de_ida,
        'lugar_evento': st.session_state.lugar_evento,
        'destino_nao_ter_aeroporto': st.session_state.destino_nao_ter_aeroporto,
        'bagagem': st.session_state.bagagem,
        'roteiro_viagem': st.session_state.roteiro_viagem,
        'roteiro_ida': st.session_state.roteiro_ida,
        'roteiro_volta': st.session_state.roteiro_volta,
        'cia_area_ida': st.session_state.cia_area_ida,
        'cia_area_volta': st.session_state.cia_area_volta,
        'numerovoo_ida': st.session_state.numerovoo_ida,
        'numerovoo_volta': st.session_state.numerovoo_volta,
        'voo_ida_1': safe_strftime(st.session_state.voo_ida_1),
        'voo_volta_1': safe_strftime(st.session_state.voo_volta_1),
        'voo_ida_2': safe_strftime(st.session_state.voo_ida_2),
        'voo_volta_2': safe_strftime(st.session_state.voo_volta_2),
        'vinculo_servidor': 'X' if st.session_state.vinculo == 'Servidor Ufes' else '',
        'vinculo_aluno': 'X' if st.session_state.vinculo == 'Aluno' else '',
        'vinculo_convidado': 'X' if st.session_state.vinculo == 'Convidado' else '',
        'tabela_reposicoes': st.session_state.tabela_reposicoes,
        'tabela_substitutos': st.session_state.tabela_substitutos,
        'declaracao_final': st.session_state.declaracao_final
    }

    caminho_docx_solicitacao = "Docs/solicitacao_de_passagem_e_diaria.docx"
    docx_path_solicitacao = preencher_documento2(dados_solicitacaodiaria_positivo, caminho_docx_solicitacao)

    # Verificar se deve gerar o Termo de Renúncia
    if st.session_state.pedir_diaria_passagem == "Não Desejo Diária e Passagem":
        dados_termoderenuncia = {
            'solicitacaodiaria': st.session_state.pedir_diaria_passagem,
            'nome': st.session_state.nome,
            'cpf': st.session_state.cpf,
            'siape': st.session_state.SIAPE,
            'renuncias': ', '.join(st.session_state.renuncias),
            'motivodarenuncia': st.session_state.motivodarenuncia
        }
        caminho_termo_renuncia = "Docs/termoderenuncia.docx"
        docx_termo_renuncia_path = preencher_documento3(dados_termoderenuncia, caminho_termo_renuncia)
    else:
        docx_termo_renuncia_path = None

    # Criar um arquivo ZIP contendo todos os documentos
    zip_path = "documentos.zip"
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        if docx_termo_renuncia_path:
            zipf.write(docx_termo_renuncia_path)
        zipf.write(docx_path_solicitacao)
        zipf.write(docx_path)

    # Download do arquivo ZIP contendo todos os documentos
    with open(zip_path, 'rb') as f:
        if st.download_button('Baixar Todos os Documentos em ZIP', f, file_name='documentos.zip'):
            st.success("Documento gerado com sucesso!")
            st.session_state.clear() 
            os.remove(zip_path)
            os.remove(caminho_docx)
            os.remove(caminho_docx_solicitacao)
            if docx_termo_renuncia_path:
                os.remove(docx_termo_renuncia_path)