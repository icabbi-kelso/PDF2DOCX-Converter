import os
import PySimpleGUI as sg
from pdf2docx import Converter

def open_window():
    # Layout of the GUI
    layout = [[sg.Text("Choose a folder: "), sg.Input(), sg.FolderBrowse()],
              [sg.pin(sg.Text('', key='-PROGRESS-', size=(40,1)))], 
              [sg.Button('Convert')]]

    # Create the window
    return sg.Window('PDF to DOCX Converter', layout)

window = open_window()

while True:
    event, values = window.read()
    # If user closes window or clicks cancel
    if event == sg.WINDOW_CLOSED:
        break
    if event == 'Convert':
        folder = values[0]
        if folder:
            # Create output directory if it doesn't exist
            output_dir = os.path.join(folder, "output")
            os.makedirs(output_dir, exist_ok=True)

            # Get all PDF files in the selected folder
            pdf_files = [f for f in os.listdir(folder) if f.endswith('.pdf')]
            
            # Add a progress bar after clicking convert
            window['-PROGRESS-'].update(visible=False)
            progress_bar = sg.ProgressBar(max_value=len(pdf_files), orientation='h', size=(20, 20), visible=True)
            window.extend_layout(window, [[sg.pin(progress_bar)]])

            # Store converted file names
            converted_files = []

            # Loop through all files in the selected folder
            for i, file in enumerate(pdf_files):
                # Set paths for pdf and docx files
                pdf_file = os.path.join(folder, file)
                docx_file = os.path.join(output_dir, os.path.splitext(file)[0] + '.docx')

                # Convert pdf to docx
                cv = Converter(pdf_file)
                cv.convert(docx_file)
                cv.close()

                # Update progress bar
                progress_bar.UpdateBar(i + 1)

                # Append the converted file to the list
                converted_files.append(docx_file)

            # Show popup with list of converted files
            sg.popup_ok('Conversion completed!', 'Converted files are saved in the "output" folder:\n' + '\n'.join(converted_files))

            # Close the current window and reopen it without a progress bar
            window.close()
            window = open_window()

window.close()
