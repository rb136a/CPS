import sys
import Lists
from os import system, name
from colorama import init, Fore, Back, Style
init(autoreset=True)
import pyodbc
import socket

# --------------------
#   CONNECT TO DB
# --------------------
conn = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E:\PSP\Office List.accdb;')
cursor = conn.cursor()
#cursor.execute(
#    '''
#    select
#    Office_Code
#    from Offices
#    ''')

# --------------------
# START APPLICATION
# --------------------
def get_Host_name_IP():
    try:
        host_name = socket.gethostname()
        host_ip = socket.gethostbyname(host_name)
        print(Fore.RED + Style.BRIGHT + "Your Hostname :  ",Fore.LIGHTYELLOW_EX + host_name)
        print(Fore.RED + Style.BRIGHT + "Your IP : ",Fore.LIGHTYELLOW_EX + host_ip)
    except:
        print(Fore.RED + Style.BRIGHT + "Unable to get Hostname and IP")

print("")
print(Fore.MAGENTA + Style.BRIGHT + "Welcome to the Office Locater!")
print(Fore.CYAN + Style.BRIGHT + "This application was made by: " + Fore.LIGHTBLUE_EX + "Rome")
get_Host_name_IP()
print("")
print(Fore.LIGHTRED_EX + Style.BRIGHT + "At any time you want to exit the application, feel free to enter \"q\" (for quit) or \"e\" (for exit), If that doesnt work, \"CTRL + C\" will do the trick.")
print("")

time.sleep(2)

def Progress_Bar():
    toolbar_width = 43
    sys.stdout.write(("" * toolbar_width))
    sys.stdout.flush()
    sys.stdout.write("" * (toolbar_width + 1))
    for i in range(toolbar_width):
        time.sleep(0.1)
        sys.stdout.write(Fore.BLUE + ".")
        sys.stdout.flush()
    sys.stdout.write("\n")

# --------------------
# Show DB
# --------------------
#for row in cursor.fetchall():
#    print (row)

# --------------------
# Office Selection
# --------------------

def what_office():
    loc_choice = input(Fore.BLUE + Style.BRIGHT + "What Office would you like to look up? ")
    return loc_choice


def office_confirm():
    print(Fore.WHITE + Style.BRIGHT + "Your choice : " + what_office())
    confirmation = input(Fore.YELLOW + Style.BRIGHT + "Yes/No : ")
    if confirmation in Lists.yes_list:
        office_verification()
    else:
        office_confirm()

def office_verification():
    if what_office in cursor.execute("SELECT IP_Subnet FROM Offices WHERE Office_Code = ?", what_office):
        print(Fore.BLUE + Style.BRIGHT + "Great!")


office_confirm()


print("")
print("")
print("")
print("")
print("")
print(Fore.LIGHTBLUE_EX + Style.BRIGHT + "______________________")
print(
    Fore.LIGHTBLUE_EX + Style.BRIGHT + "|" + Fore.RED + Style.BRIGHT + " End of APPLICATION " + Fore.LIGHTBLUE_EX + Style.BRIGHT + "|_________________")
print(
    Fore.LIGHTBLUE_EX + Style.BRIGHT + "|" + Fore.LIGHTCYAN_EX + Style.BRIGHT + " Application will close in 5 seconds " + Fore.LIGHTBLUE_EX + Style.BRIGHT + "|")
print(Fore.LIGHTBLUE_EX + Style.BRIGHT + "|_____________________________________|")

time.sleep(5)
