import tinify
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import threading

# Insert Tinyfy API key below.
APIKEY = ""
tinify.key = APIKEY

def compress_image():
    input_file = input_file_entry.get()
    output_file = output_file_entry.get()

    # Create a loading screen window
    loading_window = tk.Toplevel(root)
    loading_window.title("In Progess")
    loading_label = tk.Label(loading_window, text="Compressing...")
    loading_label.pack()

    def compression_thread():
        try:
            source = tinify.from_file(input_file)
            source.to_file(output_file)
            loading_window.destroy()
            messagebox.showinfo("Success", "Image compression successful!")
        except Exception as e:
            loading_window.destroy()
            messagebox.showerror("Error", str(e))

    # Create a thread for the compression process
    compression_thread = threading.Thread(target=compression_thread)
    compression_thread.start()

def browse_input_file():
    file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff")])
    input_file_entry.delete(0, tk.END)
    input_file_entry.insert(0, file_path)

def browse_output_file():
    file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG files", "*.png")])
    output_file_entry.delete(0, tk.END)
    output_file_entry.insert(0, file_path)

# Create a GUI window
root = tk.Tk()
root.title("Image Compression Tool")

# Create and place widgets
input_label = tk.Label(root, text="Input Image:")
input_label.pack()
input_file_entry = tk.Entry(root, width=40)
input_file_entry.pack()
browse_input_button = tk.Button(root, text="Browse", command=browse_input_file)
browse_input_button.pack()

output_label = tk.Label(root, text="Output Image:")
output_label.pack()
output_file_entry = tk.Entry(root, width=40)
output_file_entry.pack()
browse_output_button = tk.Button(root, text="Browse", command=browse_output_file)
browse_output_button.pack()

compress_button = tk.Button(root, text="Compress Image", command=compress_image)
compress_button.pack()

# Start the GUI main loop
root.mainloop()
