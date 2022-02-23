import email
import re
from email import policy
from email.parser import BytesParser
import pandas as pd
import os
import PySimpleGUI as sg
# variáveis
text3 = []
count = 0
loop = True
layout = [[sg.Combo(sorted(sg.user_settings_get_entry('-filenames-', [])), default_value=sg.user_settings_get_entry('-last filename-', ''), size=(50, 1), key='-FILENAME-'), sg.FolderBrowse(), sg.B('Clear History')],
          [sg.Button('Iniciar'),  sg.Button('Sair')]]
window = sg.Window('GERADOR DE SOLICITACAO DE MATRICULA', layout)
while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, 'Sair'):
        break
    if event == 'Iniciar':
        # If OK, then need to add the filename to the list of files and also set as the last used filename
        sg.user_settings_set_entry('-filenames-', list(set(sg.user_settings_get_entry('-filenames-', []) + [values['-FILENAME-'], ])))
        sg.user_settings_set_entry('-last filename-', values['-FILENAME-'])
        pathfiles = str(values.get('-FILENAME-'))
        for docfiles in os.listdir(pathfiles):
            emlmatricula = os.path.join(pathfiles, docfiles)
            with open(f'{emlmatricula}', 'rb') as fp:
                print(emlmatricula)
                msg = BytesParser(policy=policy.default).parse(fp)

            text = msg.get_body(preferencelist=('plain')).get_content()

            text = re.sub('[*]', '', text)
            text = text.replace("disciplinas:", "")
            text = text.replace("disciplina:", "")
            text2 = re.findall(r":(.*)", text)
            text2 = text2[:8]
            text3.append(text2)

        df = pd.DataFrame(text3, columns=['ALUNO', 'MATRÍCULA', 'Curso', 'RESERVA', 'CÓDIGO', 'DISCIPLINA', 'TURMA', 'PROFESSOR'])
        df['ALUNO'] = 'Aluno: ' + df['ALUNO'].astype(str)

        df['MATRÍCULA'] = 'Matrícula: ' + df['MATRÍCULA'].astype(str)

        df['Curso'] = 'Curso: ' + df['Curso'].astype(str)

        df['RESERVA'] = 'Curso ofertante: ' + df['RESERVA'].astype(str)

        df['CÓDIGO'] = 'Componente/Código: ' + df['CÓDIGO'].astype(str)

        df['DISCIPLINA'] = 'Disciplina: ' + df['DISCIPLINA'].astype(str)

        df['TURMA'] = 'Turma: ' + df['TURMA'].astype(str)

        df['PROFESSOR'] = 'Professor: ' + df['PROFESSOR'].astype(str)

        df.to_excel("output.xlsx", index=False)

    elif event == 'Clear History':
        sg.user_settings_set_entry('-filenames-', [])
        sg.user_settings_set_entry('-last filename-', '')
        window['-FILENAME-'].update(values=[], value='')

window.close()
