# MacroInstaller: Otomatisasi Pemasangan Macro di Microsoft Word

📦 **Versi**: 1.0.0  
🗓️ **Tanggal Rilis**: Mei 2025  
🔗 **Repositori**: [https://github.com/Archeoz/MacroInstaller](https://github.com/Archeoz/MacroInstaller)

---

## Deskripsi

MacroInstaller adalah aplikasi desktop berbasis Python yang memungkinkan pengguna untuk mengimpor macro ke dalam Microsoft Word dan mengonfigurasi shortcut keyboard secara otomatis. Aplikasi ini juga menyediakan file `.exe` untuk memudahkan distribusi dan penggunaan tanpa perlu instalasi Python.

---

## Fitur Utama

- **Impor Macro Otomatis**: Impor file `.bas` ke dalam template Normal.dotm Word.
- **Penetapan Shortcut Keyboard**: Menetapkan shortcut keyboard khusus untuk macro yang diimpor.

---

## Persyaratan Sistem

**_User Biasa_**

- **Sistem Operasi**: Windows 10 atau yang lebih baru
- **Microsoft Word**: Versi 2016 atau yang lebih baru
- **Aktifkan 'Trust access to the VBA project object model'**:
  - Buka Word > File > Options > Trust Center > Trust Center Settings > Macro Settings.
  - Centang opsi **"Trust access to the VBA project object model"**.

**_Developer_** 

- **Sistem Operasi**: Windows 10 atau yang lebih baru 
- **Microsoft Word**: Versi 2016 atau yang lebih baru 
- **Python**: Versi 3.10 atau yang lebih baru 
- **Pustaka Python**: 
    - *pywin32* - Untuk berinteraksi dengan Microsoft Word melalui COM. 
    - *pyautogui* - Untuk melakukan otomatisasi pengaturan dan interaksi. 
    - *time* - Untuk menambahkan jeda dalam eksekusi skrip. 
    - *os* - Untuk menangani operasi file dan direktori.

---

## Instalasi dan Penggunaan

**_User Biasa_**
-Cukup jalankan dengan menklik 2x file _MacroInstaller.exe_ lalu tunggu dan nikmati fiturnya, seperti : 1. Alt + 1 untuk membuat Bab ( Heading level 1 ) 2. Alt + 2 untuk membuat Sub Bab ( Heading level 2 ) 3. Alt + D untuk membuat Daftar Isi
