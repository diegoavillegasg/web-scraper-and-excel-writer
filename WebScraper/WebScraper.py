import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
import os
import datetime

first_letter = 'A'
last_letter = 'C'
wiki_root_url = "https://de.wikipedia.org"
base_url = wiki_root_url + "/wiki/Liste_deutscher_Adelsgeschlechter/"

output = "./Template.xlsx"

# to save all the families rows to be written on a spreadsheet
full_list = []

for letter in range(ord(first_letter), ord(last_letter) + 1):
    url = base_url + chr(letter)
    print(url)
    current_page = urllib.request.urlopen(url)
    soup = BeautifulSoup(current_page, features="html.parser")

    links = soup.find_all('table', {'class': 'wikitable'})

    for elem in links:
        rows = elem.find_all('tr')
        for row in rows:
            tds = row.find('td')
            if tds is None:
                continue

            anchors = tds.find('a')
            if anchors.__sizeof__() == 0:
                continue

            if anchors is not None:
                if 'href' in anchors.attrs.keys():
                    # print("family name: " + anchors.text)
                    # print("link: " + anchors.attrs['href'])
                    data = {'family': anchors.text,
                            'link': anchors.attrs['href']}
                    full_list.append(data)

output = "Generated_" + datetime.datetime.utcnow().isoformat().replace(":", "_") + ".xlsx"
workbook = xlsxwriter.Workbook(output)
wbData = workbook.add_worksheet("Data")
cellFormatter = workbook.add_format()
cellFormatter.set_font_color('blue')
cellFormatter.set_bg_color('orange')

# write headers
wbData.write(0, 0, "NAME")
wbData.write(0, 1, "LINK")

# write data
for index in range(len(full_list)):
    wbData.write(index + 1, 0, full_list[index]['family'], cellFormatter)
    wbData.write(index + 1, 1, wiki_root_url + full_list[index]['link'])

workbook.close()

print("Bye!! You can see the output file here: " + os.path.realpath(__file__).replace(__file__, "") + output)
