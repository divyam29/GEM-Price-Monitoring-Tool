import tkinter as tk
from tkinter import filedialog


def browse_files():
    file_path = filedialog.askopenfilename(
        initialdir="links",
        title="Select a File",
        filetypes=(("Excel Files", "*.xlsx"), ("All files", "*.*")),
    )
    selected_file.set(file_path)
    file_path = selected_file.get()
    if file_path:
        result_label.config(text="Selected file: " + file_path)
        return file_path
    else:
        result_label.config(text="No file selected!")


def close():
    app.destroy()


def main():
    global app
    app = tk.Tk()
    app.title("Select File....")
    app.minsize(500, 500)

    global selected_file
    selected_file = tk.StringVar()

    file_select_label = tk.Label(app, text="Select a file:")
    file_select_label.pack()

    file_select_button = tk.Button(app, text="Browse Files", command=browse_files)
    file_select_button.pack()

    global result_label
    result_label = tk.Label(app, text="")
    result_label.pack()

    exit_button = tk.Button(app, text="Start Processing", command=close)
    exit_button.pack()

    app.mainloop()


if __name__ == "__main__":
    main()
    filepath = selected_file.get()
    print(filepath)
