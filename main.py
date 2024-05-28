import os
import PySimpleGUI as sg

from fichacadastral import replace_text_in_fichacadastral
from utils import carregar_configuracoes, load_excel_data, exibir_propriedades
from procuracao import replace_text_in_procuracao
from honorarios import replace_text_in_honorarios



def main():
    pptx_file = "Procuracao.pptx"
    pptx_file_honorarios = "Proposta_Honorarios.pptx"
    docx_file= "Ficha_cadastral.docx"
    config = carregar_configuracoes()
    input_path, output_path = config.get('input_path', ''), config.get('output_path', '')

    if not input_path or not output_path:
        sg.popup_error('Os caminhos de entrada e saída não estão configurados corretamente. Por favor, verifique as configurações.')
        return

    if os.path.exists(os.path.join(input_path, "NLcadastro.xlsm")):
        ws, names = load_excel_data(os.path.join(input_path, "NLcadastro.xlsm"))
    else:
        ws, names = None, []

    menu_opcoes = [['&Arquivo', ['&Propriedades']], ['&Ajuda', ['&Sobre...']]]

    tab1_layout = [
        [sg.Text("Digite a finalidade:")],
        [sg.Multiline(key='finalidade', size=(40, 5), expand_x=True, expand_y=True)],
    ]

    tab2_layout = [
        [sg.Text("Objeto da Proposta:")],
        [sg.Multiline(key='objeto_proposta', size=(40, 5), expand_x=True, expand_y=True)],
        [sg.Text("Valor de Honorarios:"), sg.Multiline(key='valor_de_honorarios', size=(30, 1), expand_x=True)],
        [sg.Text("Observação honorarios:")],
        [sg.Multiline(key='obs_honorario', size=(40, 3), expand_x=True, expand_y=True)],
        [sg.Text("Parcelamento:"), sg.Multiline(key='parcelamento', size=(30, 1), expand_x=True)],
        [sg.Text("Validade da Proposta:"), sg.Multiline(key='validade_da_proposta', size=(30, 1), expand_x=True)],
        [sg.Text("Prazo de Execução:"), sg.Multiline(key='prazo_de_execução', size=(30, 1), expand_x=True, expand_y=True)],
    ]
    tab3_layout = [
        [sg.Text("Gerador de Ficha Cadastral Clientes")]

    ]

    tab_group = sg.TabGroup([
        [sg.Tab('Procuração', tab1_layout, tooltip='Procuração'),
         sg.Tab('Honorários', tab2_layout, tooltip='Honorários'),
         sg.Tab('Ficha_Cadastral', tab3_layout, tooltip='Ficha_Cadastral'),
         ]
    ], key='_TABGROUP_', expand_x=True, expand_y=True)

    layout = [
        [sg.Menu(menu_opcoes)],
        [sg.VPush()],
        [sg.Push(), sg.Text("Escolha o cliente:"), sg.Combo(values=names, key="combo", size=(30, 1), expand_x=True), sg.Push()],
        [sg.Push(), tab_group, sg.Push()],
        [sg.Push(), sg.Button("Criar", size=(15, 1)), sg.Push()],
        [sg.VPush(), sg.Sizegrip()]
    ]

    window = sg.Window("Gerador", layout, size=(600, 500), resizable=True, finalize=True)
    window['finalidade'].expand(True, True)
    window['objeto_proposta'].expand(True, True)
    window['valor_de_honorarios'].expand(True, False)
    window['obs_honorario'].expand(True, True)
    window['parcelamento'].expand(True, False)
    window['validade_da_proposta'].expand(True, False)
    window['prazo_de_execução'].expand(True, True)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED:
            break
        elif event == 'Propriedades':
            input_path, output_path = exibir_propriedades()
            if input_path and output_path:
                config.update({'input_path': input_path, 'output_path': output_path})
                if os.path.exists(os.path.join(input_path, "NLcadastro.xlsm")):
                    ws, names = load_excel_data(os.path.join(input_path, "NLcadastro.xlsm"))
                else:
                    ws, names = None, []
        elif event == "Criar":
            active_tab = values['_TABGROUP_']
            selected_name = values["combo"]
            if active_tab == 'Procuração':
                if selected_name and values['finalidade']:
                    finalidade_text = values['finalidade']
                    pptx_file_path = os.path.join(input_path, pptx_file)
                    replace_text_in_procuracao(selected_name, finalidade_text, pptx_file_path, output_path, input_path)
                else:
                    sg.popup("Por favor, complete todas as seleções e inserções necessárias.")
            elif active_tab == 'Honorários':
                objeto_proposta = values['objeto_proposta']
                valor_de_honorarios = values['valor_de_honorarios']
                obs_honorario = values['obs_honorario']
                parcelamento = values['parcelamento']
                validade_da_proposta = values['validade_da_proposta']
                prazo_de_execução = values['prazo_de_execução']
                pptx_file_path = os.path.join(input_path, pptx_file_honorarios)
                replace_text_in_honorarios(selected_name, pptx_file_path, output_path, input_path, objeto_proposta, valor_de_honorarios, obs_honorario, parcelamento, validade_da_proposta, prazo_de_execução)
            elif active_tab == 'Ficha_Cadastral':
                docx_file_path= os.path.join(input_path, docx_file)
                replace_text_in_fichacadastral(selected_name,docx_file_path, output_path, input_path)
    window.close()

if __name__ == "__main__":
    main()
