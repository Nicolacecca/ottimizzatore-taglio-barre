"""
Script per creare l'eseguibile dell'ottimizzatore di taglio barre
Esegui questo script per generare il file .exe
"""
import PyInstaller.__main__
import os

# Directory dello script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Percorsi dei file
main_script = os.path.join(script_dir, "ottimizzatore_taglio.py")
icon_file = os.path.join(script_dir, "icon.ico")

# Opzioni PyInstaller
options = [
    main_script,                          # Script principale
    '--name=OttimizzatoreTaglioBarre',    # Nome dell'eseguibile
    '--onefile',                          # Crea un singolo file .exe
    '--windowed',                         # Nasconde la console (GUI mode)
    f'--icon={icon_file}',                # Icona dell'eseguibile
    '--add-data=icon.ico;.',              # Include icona principale
    '--add-data=icon.png;.',              # Include icona principale PNG
    '--add-data=help_icon.ico;.',         # Include icona help
    '--add-data=help_icon.png;.',         # Include icona help PNG
    '--clean',                            # Pulisce cache prima di buildare
    '--noconfirm',                        # Non chiede conferma se esiste già
]

print("="*60)
print("CREAZIONE ESEGUIBILE - OTTIMIZZATORE TAGLIO BARRE")
print("="*60)
print("\nConfigurazione:")
print(f"  - Script: {main_script}")
print(f"  - Icona: {icon_file}")
print(f"  - Modalità: Finestra singola (GUI)")
print(f"  - Output: OttimizzatoreTaglioBarre.exe")
print("\nAvvio compilazione...\n")

# Esegui PyInstaller
PyInstaller.__main__.run(options)

print("\n" + "="*60)
print("COMPILAZIONE COMPLETATA!")
print("="*60)
print("\nL'eseguibile è stato creato in:")
print(f"  {os.path.join(script_dir, 'dist', 'OttimizzatoreTaglioBarre.exe')}")
print("\nPuoi copiare questo file su qualsiasi PC Windows")
print("senza bisogno di avere Python installato!")
print("="*60)
