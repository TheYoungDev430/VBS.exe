import os
import tkinter as tk
from tkinter import filedialog, messagebox

def generate_cpp_wrapper(vbs_path):
    # Read the VBS script content
    with open(vbs_path, 'r', encoding='utf-8') as f:
        vbs_content = f.read()

    # Escape double quotes and backslashes for embedding in C++ string
    escaped_vbs = vbs_content.replace('\\', '\\\\').replace('"', '\\"')

    # Generate the C++ source code
    cpp_code = f'''
#include <windows.h>
#include <fstream>
#include <string>
#include <shlobj.h>

int WINAPI WinMain(HINSTANCE, HINSTANCE, LPSTR, int) {{
    char tempPath[MAX_PATH];
    GetTempPathA(MAX_PATH, tempPath);
    std::string tempFile = std::string(tempPath) + "embedded_script.vbs";

    std::ofstream out(tempFile);
    out << "{escaped_vbs}";
    out.close();

    ShellExecuteA(NULL, "open", "wscript.exe", tempFile.c_str(), NULL, SW_HIDE);
    return 0;
}}
'''

    # Save the generated C++ source code
    cpp_output_path = os.path.join(os.path.dirname(vbs_path), 'vbs_embedded_wrapper.cpp')
    with open(cpp_output_path, 'w', encoding='utf-8') as f:
        f.write(cpp_code)

    return cpp_output_path

def compile_cpp_to_exe(cpp_path):
    exe_output_path = os.path.splitext(cpp_path)[0] + '.exe'
    compile_command = f'g++ "{cpp_path}" -o "{exe_output_path}" -mwindows'
    result = os.system(compile_command)
    return result == 0, exe_output_path

def select_vbs_and_compile():
    vbs_path = filedialog.askopenfilename(filetypes=[("VBScript files", "*.vbs")])
    if not vbs_path:
        return

    try:
        cpp_path = generate_cpp_wrapper(vbs_path)
        success, exe_path = compile_cpp_to_exe(cpp_path)
        if success:
            messagebox.showinfo("Success", f"Executable created at:\n{exe_path}")
        else:
            messagebox.showerror("Compilation Failed", "Failed to compile the generated C++ file.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# Create the GUI
root = tk.Tk()
root.title("VBS to EXE Compiler")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

label = tk.Label(frame, text="Select a .vbs file to compile into .exe")
label.pack(pady=10)

select_button = tk.Button(frame, text="Choose VBS File and Compile", command=select_vbs_and_compile)
select_button.pack(pady=10)

root.mainloop()