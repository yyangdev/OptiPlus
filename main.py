from subprocess import run
import os
import sys
import ctypes
import time
from colorama import Fore, init
init()
from art import text2art
text = text2art("OptiPlus")
print(text)
# Frontend
menu_items = [
    '[1] Активация Windows 10/11',
    '[2] Установка Microsoft Office',
    '[3] Активация Microsoft Office',
    '[4] Установка драйверов на видеокарты nvidia',
    '[5] Установка драйверов на видеокарты amd',
    '[6] Установка драйверов на видеокарту intel',
    '[7] Установка драйверов на звук',
    '[8] Установка драйверов на интернет',
    '[9] Установка Steam',
    '[10] Установка Google Chrome',
    '[11] Установка 7-Zip',
    '[12] Установка CPU-Z',
    '[13] Установка VLC Media Player',
    '[14] Установка MSI Afterburner',
    ""
]

for item in menu_items:
    print(item)

try:
    choice = int(input("Введите номер 1-14: "))
except ValueError:
    print("Ошибка: введите число!")
    sys.exit()

# Backend
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    print("Запуск с правами администратора...")
    run(['powershell', 'Start-Process', 'python', f'"{sys.argv[0]}"', '-Verb', 'RunAs'])
    sys.exit()

def run_command(command, description):
    print(f"\n=== {description} ===")
    result = run(command, shell=True, capture_output=True, text=True)
    if result.returncode == 0:
        print(f"Good!: {description}")
    else:
        print(f"Error(: {description}")
        if result.stderr:
            print(f"Error(: {result.stderr}")
    return result.returncode

def activate_windows():
    commands = [
        ('slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX', 'Установка ключа Windows'),
        ('slmgr /skms kms8.msguides.com', 'Настройка KMS сервера'),
        ('slmgr /ato', 'Активация Windows')
    ]
    for cmd, desc in commands:
        run_command(cmd, desc)
        time.sleep(2)

def install_office():
    run_command('winget install -e --id Microsoft.Office --silent --accept-package-agreements', 'Установка Microsoft Office')

def activate_office():
    office_script = '''
    $paths = @("C:\\Program Files\\Microsoft Office\\Office16", "C:\\Program Files (x86)\\Microsoft Office\\Office16")
    foreach ($path in $paths) {
        if (Test-Path $path) {
            Set-Location $path
            .\\ospp.vbs /sethst:kms8.msguides.com
            .\\ospp.vbs /act
            break
        }
    }
    '''
    run(['powershell', '-Command', office_script])

def install_programs():
    programs = {
        4: ('winget install -e --id NVIDIA.GeForceExperience --silent --accept-package-agreements', 'Драйверы NVIDIA'),
        5: ('winget install -e --id AMD.Catalyst --silent --accept-package-agreements', 'Драйверы AMD'),
        6: ('winget install -e --id Intel.GraphicsDriver --silent --accept-package-agreements', 'Драйверы Intel'),
        7: ('winget install -e --id Realtek.AudioDriver --silent --accept-package-agreements', 'Драйверы звука'),
        8: ('winget install -e --id Intel.NetworkDriver --silent --accept-package-agreements', 'Драйверы сети'),
        9: ('winget install -e --id Valve.Steam --silent --accept-package-agreements', 'Установка Steam'),
        10: ('winget install -e --id Google.Chrome --silent --accept-package-agreements', 'Установка Google Chrome'),
        11: ('winget install -e --id 7zip.7zip --silent --accept-package-agreements', 'Установка 7-Zip'),
        12: ('winget install -e --id CPUID.CPU-Z --silent --accept-package-agreements', 'Установка CPU-Z'),
        13: ('winget install -e --id VideoLAN.VLC --silent --accept-package-agreements', 'Установка VLC'),
        14: ('winget install -e --id MSI.Afterburner --silent --accept-package-agreements', 'Установка MSI Afterburner')
    }
    
    if choice in programs:
        cmd, desc = programs[choice]
        run_command(cmd, desc)

# Выполнение выбранной операции
print(f"\ninput: {menu_items[choice-1]}")
print("=" * 50)

if choice == 1:
    activate_windows()
elif choice == 2:
    install_office()
elif choice == 3:
    activate_office()
elif 4 <= choice <= 14:
    install_programs()
else:
    print("Press numbers 1- 14!!!")

print("\n" + "=" * 50)
print("you PC optimized!")
print("=" * 50)
input("Нажмите Enter для выхода...")
