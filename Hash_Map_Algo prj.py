import tkinter as tk
from tkinter import filedialog
import re
import docx

class SpellChecker:
    def __init__(self, dictionary_file):
        self.dictionary = set()
        self.load_dictionary(dictionary_file)

    def load_dictionary(self, dictionary_file):
        try:
            with open(dictionary_file, 'r', encoding='utf-8') as file:
                for line in file:
                    word = line.strip().lower()
                    self.dictionary.add(word)
        except FileNotFoundError:
            print(f"Error: Dictionary file '{dictionary_file}' not found.")
            exit(1)
        except Exception as e:
            print(f"An error occurred while loading the dictionary: {e}")
            exit(1)

    def check_spelling(self, file_path):
        misspelled_words = set()

        try:
            if file_path.endswith(".txt"):
                with open(file_path, 'r', encoding='utf-8') as file:
                    for line_number, line in enumerate(file, start=1):
                        words = re.findall(r'\b\w+\b', line)
                        for word in words:
                            clean_word = self.clean_word(word)
                            if clean_word and clean_word not in self.dictionary:
                                misspelled_words.add((clean_word, line_number))
            elif file_path.endswith(".docx"):
                doc = docx.Document(file_path)
                for para_number, para in enumerate(doc.paragraphs, start=1):
                    words = re.findall(r'\b\w+\b', para.text)
                    for word in words:
                        clean_word = self.clean_word(word)
                        if clean_word and clean_word not in self.dictionary:
                            misspelled_words.add((clean_word, para_number))
            else:
                print("Unsupported file format. Please use .txt or .docx files.")
                return None
        except FileNotFoundError:
            print(f"Error: File '{file_path}' not found.")
            exit(1)
        except Exception as e:
            print(f"An error occurred while checking spelling: {e}")
            exit(1)

        return misspelled_words

    def clean_word(self, word):
        # Remove non-alphabetic characters from the word
        return ''.join(char.lower() for char in word if char.isalpha())

class SpellCheckerGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Spell Checker")

        default_dictionary_file = 'words.txt'

        self.dictionary_label = tk.Label(master, text=f"Dictionary File: {default_dictionary_file}")
        self.dictionary_label.grid(row=0, column=0, padx=10, pady=5)

        self.text_label = tk.Label(master, text="File to Check:")
        self.text_label.grid(row=1, column=0, padx=10, pady=5)

        self.text_entry = tk.Entry(master)
        self.text_entry.grid(row=1, column=1, padx=10, pady=5)

        self.browse_text_button = tk.Button(master, text="Browse", command=self.browse_file)
        self.browse_text_button.grid(row=1, column=2, padx=10, pady=5)

        self.check_button = tk.Button(master, text="Check Spelling", command=self.check_spelling)
        self.check_button.grid(row=2, column=0, columnspan=2, pady=10)

        self.result_text = tk.Text(master, height=10, width=50)
        self.result_text.grid(row=3, column=0, columnspan=3, pady=10)

        self.spell_checker = SpellChecker(default_dictionary_file)

    def browse_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt"), ("Word Files", "*.docx")])
        self.text_entry.delete(0, tk.END)
        self.text_entry.insert(0, file_path)

    def check_spelling(self):
        file_path = self.text_entry.get()

        misspelled_words = self.spell_checker.check_spelling(file_path)

        if misspelled_words is None:
            return

        if not misspelled_words:
            result = "No misspelled words found."
        else:
            result = "Misspelled words:\n"
            for word, line_number in sorted(misspelled_words, key=lambda x: (x[1], x[0])):
                result += f"{word} at line {line_number}\n"

        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, result)

def main():
    root = tk.Tk()
    app = SpellCheckerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
