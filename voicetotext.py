import speech_recognition as sr
import os
import tkinter as tk
from tkinter import filedialog
from docx import Document  
from pydub import AudioSegment

def convert_audio_to_wav(input_file, output_file):
    audio = AudioSegment.from_file(input_file)
    audio.export(output_file, format="wav")

def audio_to_text(input_file, output_folder, custom_file_name):
    recognizer = sr.Recognizer()


    if not input_file.lower().endswith(".wav"):
        temp_wav_file = os.path.join(output_folder, "temp.wav")
        convert_audio_to_wav(input_file, temp_wav_file)
        input_file = temp_wav_file

    with sr.AudioFile(input_file) as source:
        audio_data = recognizer.record(source)

        try:
            text = recognizer.recognize_google(audio_data)

            if custom_file_name:
                output_file_path = os.path.join(output_folder, "{}.docx".format(custom_file_name))
            else:
                output_file_name = "transkripsi"
                count = 1
                while os.path.exists(os.path.join(output_folder, "{}({}).docx".format(output_file_name, count))):
                    count += 1
                output_file_path = os.path.join(output_folder, "{}({}).docx".format(output_file_name, count))

            
            document = Document()
            document.add_paragraph(text)
            document.save(output_file_path)

            print("Transkripsi disimpan di: {}".format(output_file_path))

        except sr.UnknownValueError:
            print("Google Web Speech API tidak dapat mengenali audio")
        except sr.RequestError as e:
            print("Gagal menghubungi Google Web Speech API; {0}".format(e))


    if input_file.lower().endswith(".wav"):
        os.remove(input_file)


def browse_file():
    file_path = filedialog.askopenfilename()
    entry_file_path.delete(0, tk.END)
    entry_file_path.insert(0, file_path)

def browse_folder():
    folder_path = filedialog.askdirectory()
    entry_output_folder.delete(0, tk.END)
    entry_output_folder.insert(0, folder_path)

def process_audio():
    input_file = entry_file_path.get()
    output_folder = entry_output_folder.get()
    custom_file_name = entry_custom_file_name.get()

    if input_file and output_folder:
        audio_to_text(input_file, output_folder, custom_file_name)

root = tk.Tk()
root.title("Audio to Text Converter")

label_file_path = tk.Label(root, text="File Suara")
label_file_path.grid(row=0, column=0, pady=10, padx=10, sticky=tk.W)
entry_file_path = tk.Entry(root, width=50)
entry_file_path.grid(row=0, column=1, padx=10)
button_browse_file = tk.Button(root, text="Browse", command=browse_file)
button_browse_file.grid(row=0, column=2, padx=10)

label_output_folder = tk.Label(root, text="Output")
label_output_folder.grid(row=1, column=0, pady=10, padx=10, sticky=tk.W)
entry_output_folder = tk.Entry(root, width=50)
entry_output_folder.grid(row=1, column=1, padx=10)
button_browse_folder = tk.Button(root, text="Browse", command=browse_folder)
button_browse_folder.grid(row=1, column=2, padx=10)

label_custom_file_name = tk.Label(root, text="Nama")
label_custom_file_name.grid(row=2, column=0, pady=10, padx=10, sticky=tk.W)
entry_custom_file_name = tk.Entry(root, width=50)
entry_custom_file_name.grid(row=2, column=1, padx=10)

label_creator1 = tk.Label(root, text="Dibuat oleh: ")
label_creator1.grid(row=3, column=0, columnspan=3, pady=5, padx=10, sticky=tk.W)
label_creator1 = tk.Label(root, text="Yoga Maulana Febriansyah 21.83.0596")
label_creator1.grid(row=4, column=0, columnspan=3, pady=5, padx=10, sticky=tk.W)
label_creator2 = tk.Label(root, text="Rakhaezza Nabella 21.83.0598")
label_creator2.grid(row=5, column=0, columnspan=3, pady=5, padx=10, sticky=tk.W)
label_creator3 = tk.Label(root, text="Eleazar Hendro Tri Putra 21.83.0645")
label_creator3.grid(row=6, column=0, columnspan=3, pady=5, padx=10, sticky=tk.W)

button_process = tk.Button(root, text="Process Audio", command=process_audio)
button_process.grid(row=7, column=0, columnspan=3, pady=20)

root.mainloop()
