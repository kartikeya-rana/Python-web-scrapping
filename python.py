import requests     # For URL request
from bs4 import BeautifulSoup   # For web Scrape
from xlwings import Workbook, Range     # For connection and enter data

def get_data(): 
    wb = Workbook(r'C:\Users\KARTIKEYA RANA\Desktop\Book1.xlsm') #Work Book address
    
    url = Range(1, (2,3)).value #Retrieve url from Excel Sheet
    # 1 for sheet number and (2,3) for cell from where url will be retrieved
    r =requests.get(url)
    
    soup = BeautifulSoup(r.content) #Get contents from website.

    # Find data which has h1 tag and class full-name.
    #It will include useful text as well as html syntax.
    user = soup.find_all("h1", {"class":"full-name"})

    Range(1,(6,3)).value = 'User name' # Print "User Name" at cell C6 i.e 6,3
    Range(1, (6,4)).value = user[0].text # Print only useful text using .text

    followers = soup.find_all("div", {"class":"number"})

    Range(1,(7,3)).value = 'Followers'
    Range(1, (7,4)).value = followers[0].text

    infos = soup.find_all("div", {"class": "class-info"})

    Range(1, (9,3)).value = 'Course Title'
    Range(1, (9,4)).value = 'Students enrolled'

    data_row = 10


    for info in infos:
        title = info.find_all("p" , {"class": "title-link"})[0].text
        Range(1, (data_row,3)).value = title
        students = info.find_all("span" , {"class": "num-students"})[0].text.replace(',','')
        number = int(students)
        Range(1, (data_row,4)).value = number
        data_row = data_row + 1
