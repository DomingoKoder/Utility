import win32com.client as client
import datetime
from datetime import timedelta
from colorama import Back, Fore, Style

# https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.mailitem?view=outlook-pia&redirectedfrom=MSDN#properties_
# https://stackoverflow.com/questions/22813814/clearly-documented-reading-of-emails-functionality-with-python-win32com-outlook
outlook = client.Dispatch('Outlook.Application')

date = datetime.datetime.now()
date = date.strftime("[%D]")

namespace = outlook.GetNameSpace('MAPI')
account = namespace.Folders['Outlook']   #Account name
folder = account.Folders['Shift']        #Folder name
shifty = [message for message in folder.items]

print(f'Dziś jest {date} i Folder "{folder.name}" zawiera {folder.Items.Count} wiadomosci:')
for message in folder.items:

    received_time = message.ReceivedTime
    hour = received_time.hour
    minute = received_time.minute
    formatted_time = f"{hour:02}:{minute:02}"

    if message.Sent:
        print(f' {Fore.CYAN}Wyslane{Fore.WHITE}  [{Fore.MAGENTA}{message.ReceivedTime.date()}{Fore.WHITE}] [{Fore.YELLOW}{formatted_time}{Fore.WHITE}] {message.Subject}')
    elif message.Saved:
        print(f' {Fore.CYAN}Zapisane{Fore.WHITE} [{Fore.MAGENTA}{message.ReceivedTime.date()}{Fore.WHITE}] [{Fore.YELLOW}{formatted_time}{Fore.WHITE}] {message.Subject}')    



#message.ReceivedTime()
#message.senton.date()       
#message.senton.time()    
#message.senton            


