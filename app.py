# -*- coding: utf-8 -*-
'''PDF pages counter with GUI.'''

import os.path
import traceback

import docx2txt
import fitz
import PySimpleGUI as sg
import spacy_udpipe

try:
    sg.theme('Green Tan')

    nlp_ud = spacy_udpipe.load_from_path(lang='tr',
                                    path='./kazakh-ud-2.0-170801.udpipe',
                                    meta={"description": "Custom 'kz' model"})

    file_list_column = [
        [
            sg.Text('Open folder:'),
            sg.In(size=(25, 1), enable_events=True, key='-FOLDER-'),
            sg.FolderBrowse(),
        ],
        [sg.Text('Select file:')],
        [
            sg.Listbox(
                values=[], enable_events=True, size=(40, 24), key='-FILE LIST-'
            )
        ],
        [sg.Text('Delete symbols:')],
        [
            sg.Multiline(default_text='\\n\n\\t\n\\xa0', size=(40, 10), key='-DEL SYMBOLS-')
        ],
        [sg.Button('Open file >>', key='-OPEN-')],
    ]

    raw_viewer_column = [
        [sg.Text('Preview raw text')],
        [sg.Multiline(size=(60, 41), key='-OUT RAW-')],
        [sg.Button('Analyze >>', key='-ANALYZE-')],
    ]

    result_viewer_column = [
        [sg.Text('Result')],
        [sg.Multiline(size=(40, 41), key='-OUT RESULT-')],
        [
            sg.In(size=(25, 1), key='-OUT FILE-'),
            sg.Button('Save', key='-SAVE-'),
        ],
    ]

    layout = [
        [
            sg.Column(file_list_column),
            sg.VSeperator(),
            sg.Column(raw_viewer_column),
            sg.VSeperator(),
            sg.Column(result_viewer_column),
        ]
    ]

    window = sg.Window('Kazakh', layout)


    while True:
        event, values = window.read()

        if event == 'Exit' or event == sg.WIN_CLOSED:
            break

        if event == '-FOLDER-':
            folder = values['-FOLDER-']
            try:
                file_list = os.listdir(folder)
            except:
                file_list = []

            f_names = [
                f
                for f in file_list
                if os.path.isfile(os.path.join(folder, f))
                and f.lower().endswith(('.pdf', '.docx', '.txt'))
            ]
            window['-FILE LIST-'].update(f_names)

        if event == '-OPEN-':
            del_symbols = values['-DEL SYMBOLS-']
            del_symbols = del_symbols.split('\n')

            if len(values["-FILE LIST-"]):
                f_name = values["-FILE LIST-"][0]
                f_path = os.path.join(
                    values["-FOLDER-"], f_name
                )

                out_file_pre = 'result_'
                out_file_ext = '.csv'
                out_file = out_file_pre + f_name.replace('.', '_') + out_file_ext
                out_path = os.path.join(
                    values["-FOLDER-"], out_file
                )
                window['-OUT FILE-'].update(out_path)

                text_raw = ''
                if f_name.endswith('.pdf'):
                    fd = fitz.open(f_path)
                    for page in fd:
                        t = page.getText()
                        text_raw += '\n' + t

                elif f_name.endswith('docx'):
                    text_raw = docx2txt.process(f_path)

                elif f_name.endswith('.txt'):
                    with open(f_path, 'r') as f:
                        text_raw = f.read()

                for ds in del_symbols:
                    text_raw = text_raw.replace(ds, '')

                window['-OUT RAW-'].update(text_raw)

        if event == '-ANALYZE-':
            text_raw = values['-OUT RAW-']
            if text_raw != '':
                result = []
                doc = nlp_ud(text_raw)
                csv_header = ['token','lemma','pos']
                result.append(';'.join(csv_header))

                for token in doc:
                    r = (token.text, token.lemma_, token.pos_)
                    result.append(';'.join([f'"{t}"' for t in r]))

                text_proc = '\n'.join(result)
                window['-OUT RESULT-'].update(text_proc)

        if event == '-SAVE-':
            out_file = values['-OUT FILE-']
            text_proc = values['-OUT RESULT-']
            if text_proc != '':
                with open(out_file, 'wt', encoding='UTF-8') as f:
                    f.write(text_proc)
                    sg.popup_ok(f'File was saved: "{out_file}"')

    window.close()

except Exception as e:
    tb = traceback.format_exc()
    sg.Print(f'An error happened.  Here is the info:', e, tb)
    sg.popup_error(f'AN EXCEPTION OCCURRED!', e, tb)
