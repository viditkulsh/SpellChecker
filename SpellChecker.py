import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox
import re
import docx
from spellchecker import SpellChecker as PySpellChecker

class SpellChecker:
    def __init__(self, dictionary_file):
        self.dictionary_file = dictionary_file
        self.dictionary = set()
        self.load_dictionary()

    def load_dictionary(self):
        try:
            with open(self.dictionary_file, 'r', encoding='utf-8') as file:
                for line in file:
                    word = line.strip().lower()
                    self.dictionary.add(word)
        except FileNotFoundError:
            print(f"Error: Dictionary file '{self.dictionary_file}' not found.")
            exit(1)
        except Exception as e:
            print(f"An error occurred while loading the dictionary: {e}")
            exit(1)

    def add_to_dictionary(self, word):
        # Add a word to the dictionary file
        with open(self.dictionary_file, 'a', encoding='utf-8') as file:
            file.write(word.lower() + '\n')

        # Update the dictionary set
        self.dictionary.add(word.lower())

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

    def get_suggestions(self, misspelled_word):
        spell = PySpellChecker()
        return spell.candidates(misspelled_word)

class SpellCheckerGUI:
    def __init__(self, master):
        self.master = master
        self.master.title("Spell Checker")

        dictionary_file = 'words.txt'

        self.dictionary_label = tk.Label(master, text=f"Dictionary File: {dictionary_file}")
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

        self.spell_checker = SpellChecker(dictionary_file)

        self.add_to_dict_button = tk.Button(master, text="Add to Dictionary", command=self.add_to_dictionary)
        self.add_to_dict_button.grid(row=4, column=0, columnspan=2, pady=10)

        # self.suggestions_button = tk.Button(master, text="Get Suggestions", command=self.get_suggestions)
        # self.suggestions_button.grid(row=5, column=0, columnspan=2, pady=10)

        # self.change_word_button = tk.Button(master, text="Change Word", command=self.change_word)
        # self.change_word_button.grid(row=6, column=0, columnspan=2, pady=10)

        self.selected_misspelled_word = None
        self.selected_suggestion = None

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

    def add_to_dictionary(self):
        misspelled_word = simpledialog.askstring("Add to Dictionary", "Enter misspelled word:")
        if misspelled_word:
            self.spell_checker.add_to_dictionary(misspelled_word)
            messagebox.showinfo("Word Added", f"The word '{misspelled_word}' has been added to the dictionary.")

    def get_suggestions(self):
        selected_word_info = self.result_text.tag_prevrange("sel", 1.0)
        if selected_word_info:
            start, end = selected_word_info
            self.selected_misspelled_word = self.result_text.get(start, end)
            suggestions = self.spell_checker.get_suggestions(self.selected_misspelled_word)
            if suggestions:
                suggestion_str = ', '.join(suggestions)
                self.selected_suggestion = simpledialog.askstring(
                    "Select Suggestion", f"Choose a suggestion for '{self.selected_misspelled_word}':\n{suggestion_str}"
                )

    def change_word(self):
        if self.selected_misspelled_word and self.selected_suggestion:
            text_content = self.result_text.get(1.0, tk.END)
            updated_content = text_content.replace(self.selected_misspelled_word, self.selected_suggestion)

            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, updated_content)
            messagebox.showinfo("Change Word", f"Word changed to: {self.selected_suggestion}")
        else:
            messagebox.showinfo("Change Word", "Please select a misspelled word and its suggestion first.")

def main():
    root = tk.Tk()
    app = SpellCheckerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

