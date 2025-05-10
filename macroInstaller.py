import time
import os
import ctypes
import tkinter as tk
from tkinter import messagebox
from win32com.client import gencache, constants
import sys  # Import sys untuk exit program

def open_word():
    word = gencache.EnsureDispatch('Word.Application')
    word.Visible = True
    time.sleep(1)
    return word

def import_macro_and_set_shortcut(word, bas_path, key_char):
    try:
        # 1) Bersihkan Attribute VB_Name
        with open(bas_path, 'r') as f:
            lines = [l for l in f if not l.strip().startswith("Attribute VB_Name")]
        tmp = bas_path.replace('.bas', '_cleaned.bas')
        with open(tmp, 'w') as f:
            f.writelines(lines)

        # 2) Import ke Normal.dotm
        tpl = word.NormalTemplate
        proj = tpl.VBProject
        name = os.path.splitext(os.path.basename(bas_path))[0]
        try:
            comp = proj.VBComponents(name)
            proj.VBComponents.Remove(comp)
        except:
            pass
        mod = proj.VBComponents.Add(1)  # vbext_ct_StdModule
        mod.Name = name
        with open(tmp, 'r') as f:
            mod.CodeModule.AddFromString(f.read())
        os.remove(tmp)
        print(f"Imported '{name}'")

        # 3) Hapus binding lama
        for kb in word.KeyBindings:
            if kb.Command == name and kb.KeyCategory == constants.wdKeyCategoryMacro:
                kb.Clear()

        # 4) Bangun keycode: Alt + key_char
        keycode = word.BuildKeyCode(constants.wdKeyAlt, ord(key_char.upper()))

        # 5) Tambah binding baru
        word.KeyBindings.Add(
            KeyCategory=constants.wdKeyCategoryMacro,
            Command=name,
            KeyCode=keycode
        )
        print(f"Shortcut Alt+{key_char.upper()} untuk '{name}' berhasil ditambahkan")
    
    except Exception as e:
        show_error_popup(str(e))

def show_error_popup(error_message):
    # Menampilkan pesan error dengan instruksi
    root = tk.Tk()
    root.withdraw()  # Menyembunyikan jendela utama Tkinter
    messagebox.showerror("Error: Programmatic access is not trusted", 
                        f"Terjadi kesalahan: {error_message}\n\n"
                        "Untuk melanjutkan, aktifkan fitur berikut di Word:\n\n"
                        "1. Aktifkan 'Trust access to the VBA project object model': File > Options > Trust Center > Trust Center Settings > Macro Settings\n\n"
                        "Setelah itu, coba jalankan ulang program.")
    root.quit()
    sys.exit()  # Menghentikan eksekusi program setelah pop-up

def show_popup(message):
    ctypes.windll.user32.MessageBoxW(0, message, "Informasi", 0x40 | 0x1)

def main():
    bas_folder = os.path.join(os.getcwd(), 'macros')
    # Mapping macro ke satu tombol setelah Alt
    shortcuts = {
        'BuatDaftarIsi': 'D',  # Alt+D
        'FormatBab':     '1',  # Alt+1
        'FormatSubBab':  '2',  # Alt+2
    }

    word = open_word()
    for macro, key in shortcuts.items():
        bas_path = os.path.join(bas_folder, f"{macro}.bas")
        import_macro_and_set_shortcut(word, bas_path, key)

    # Pop-up notifikasi setelah semua proses selesai
    show_popup("✅ Semua macro diimpor dan shortcut diatur!")

    # Menutup Word setelah selesai
    word.Quit()
    print("Word telah ditutup!")

main()
