# -*- coding: utf-8 -*-
'''PDF pages counter with GUI.'''

import os.path
import sys
import traceback

import docx2txt
import fitz
import PySimpleGUI as sg
import spacy_udpipe

try:
    ROOT_PATH = sys._MEIPASS
except:
    ROOT_PATH = '.'

def main():
    """Main function."""

    result_header = ['token','lemma','pos']
    result_data = []
    result_csv_text = ''

    try:
        sg.theme('Green Tan')

        nlp_ud = spacy_udpipe.load_from_path(lang='tr',
                                        path=f'{ROOT_PATH}/kazakh-ud-2.0-170801.udpipe',
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
            [sg.Table(
                values=[],
                headings=result_header,
                def_col_width=10,
                # max_col_width=30,
                col_widths=[15, 15, 10],
                auto_size_columns=False,
                display_row_numbers=True,
                justification='right',
                num_rows=41,
                alternating_row_color='lightyellow',
                key='-RESULT TABLE-',
                selected_row_colors='red on yellow',
                enable_events=True,
                expand_x=True,
                expand_y=True,
                enable_click_events=True,
                tooltip='Result'
            )],
            [
                sg.InputText(
                    default_text='',
                    visible=False,
                    do_not_clear=False,
                    enable_events=True,
                    key='-OUT FILE-'
                ),
                sg.FileSaveAs(
                    button_text='Save as...',
                    key='-SAVE-',
                    disabled=True,
                    file_types=(('CSV', '.csv')),
                    default_extension='*.csv'
                )
            ],
        ]

        layout = [
            [
                sg.Column(file_list_column),
                sg.VSeperator(),
                sg.Column(raw_viewer_column),
                sg.VSeperator(),
                sg.Column(result_viewer_column),
            ],
        ]

        window = sg.Window(
            title='Kazakh',
            layout=layout,
            enable_close_attempted_event=True,
            location=sg.user_settings_get_entry('-location-', (None, None))
        )



        while True:
            event, values = window.read()

            if event in ('Exit', sg.WIN_CLOSED, sg.WINDOW_CLOSE_ATTEMPTED_EVENT):
                # save window position
                sg.user_settings_set_entry('-location-', window.current_location())
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
                    window['-SAVE-'].update(disabled=True)
                    result_data = []
                    window['-RESULT TABLE-'].update(values=result_data)

            if event == '-ANALYZE-':
                text_raw = values['-OUT RAW-']
                if text_raw != '':
                    doc = nlp_ud(text_raw)

                    for token in doc:
                        r = [token.text, token.lemma_, token.pos_]
                        result_data.append(r)

                    window['-RESULT TABLE-'].update(values=result_data)

                    # construct csv
                    result_csv = []
                    result_csv.append(';'.join(result_header))
                    for r in result_data:
                        result_csv.append(';'.join([f'"{t}"' for t in r]))

                    result_csv_text = '\n'.join(result_csv)
                    window['-SAVE-'].update(disabled=False)

            if event == '-OUT FILE-':
                out_file = values['-OUT FILE-']
                if result_csv_text != '' and out_file != '':
                    with open(out_file, 'wt', encoding='UTF-8') as f:
                        f.write(result_csv_text)
                        sg.popup_ok(f'File was saved: "{out_file}"')

        window.close()

    except Exception as e:
        tb = traceback.format_exc()
        sg.Print(f'An error happened.  Here is the info:', e, tb)
        sg.popup_error(f'AN EXCEPTION OCCURRED!', e, tb)



if __name__ == '__main__':
    main()
