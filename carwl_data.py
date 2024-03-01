!pip install fake-useragent
from bs4 import BeautifulSoup
import requests
from fake_useragent import UserAgent
import openpyxl
ua = UserAgent()
user_agent = ua.random
headers = {'User-Agent': user_agent}

#======================================FUNCTION====================================================

def should_stop(soup):
    label = soup.find('label')
    return label.text.strip() == '0'
#=====================================
def generate_url_list(base_url, result):
    page = 1
    list_tam =[]
    for i in range(result):
        url = base_url + str(page)
        print(" page {}".format(url))
        list_tam.append(url)
        page += 1

    return list_tam

def process_page(url):
    response = requests.get(url, headers=headers)
    if response.status_code == 200:
        html = response.text
        soup = BeautifulSoup(html, 'html.parser')
        next_page_class = soup.find(class_='next-page')
        if next_page_class:
            print("Đã tìm thấy class 'next-page'")
            html_content = str(next_page_class)
            result = find_max_number_in_a_tags(html_content)
            print("total page : ", result)
            return result
        else:
            print("no class 'next-page'")
            return None
    else:
        print("error :", response.status_code)
        return None

def find_max_number_in_a_tags(html_content):
    soup = BeautifulSoup(html_content, 'html.parser')

    a_tags = soup.find_all('a', href=True)

    numbers = []

    for a in a_tags:
        try:
            number = int(a.text.strip()) 
            numbers.append(number)
        except ValueError:
            pass  

    if numbers:
        max_number = max(numbers) 
        return max_number
    else:
        return None  

def get_company_info(url, headers):
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, "html.parser")
            company_info = []

            company_name = find_company_name(soup)
            if company_name:
                print("company name :", company_name)
                company_info.append(company_name)
            else:
                print("no name")
                return None

            phone_number = find_phone_number(soup)
            if phone_number:
                print("phone :", phone_number)
                company_info.append(phone_number)
            else:
                print("No phone ")
                return None

            address = find_address(soup)
            if address:
                print("add :", address)
                company_info.append(address)
            else:
                print("no add")
                return None

            return company_info
        else:
            print(f"can not acc URL '{url}'")
            return None
    except requests.exceptions.RequestException as e:
        print(f"Lỗi: {e}")
        return None

def find_company_name(soup):
    h1_element = soup.find("h1")
    if h1_element:
        return h1_element.get_text().strip()
    return None

def find_phone_number(soup):
    highlight_span = soup.find("span", class_="highlight")
    if highlight_span:
        return highlight_span.get_text().strip()
    return None

def find_address(soup):
    strong_elements = soup.find_all("strong")
    if len(strong_elements) > 1:
        return strong_elements[1].get_text().strip()
    return None

def create_workbook():
    return openpyxl.Workbook()


def write_data_to_sheet(workbook, data):
    sheet = workbook.active
    for row in data:
        sheet.append(row)

def auto_fit_columns(sheet):
    for column_cells in sheet.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            if cell.value:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
        adjusted_width = (max_length + 2) * 1.2
        sheet.column_dimensions[column].width = adjusted_width


def save_workbook(workbook, filename):
    workbook.save(filename)

def main(total_list, filename):
    workbook = create_workbook()
    write_data_to_sheet(workbook, total_list)
    sheet = workbook.active
    auto_fit_columns(sheet)
    save_workbook(workbook, filename)
#======================================FUNCTION====================================================
urls = [
    "https://hosocongty.vn/nam-2024-tien-giang/page-",
    "https://hosocongty.vn/nam-2023-tien-giang/page-"
]
url_list = []
for url in urls :
  url_list.extend(generate_url_list(url,process_page(url)))
print("n Page link : ",len(url_list))
print("Page link : ",url_list)

link_hcm = []
for url in url_list:
    try:
        response = requests.get(url, headers=headers)
        if response.status_code == 200:
            soup = BeautifulSoup(response.content, "html.parser")
            target_ul = soup.find('ul', class_='hsdn')
            target_li = target_ul.find_all('li')
            for li in target_li:
                a = li.find('a')
                #if 'Thành phố Hồ Chí Minh' in li.text or 'Bình Dương' in li.text or 'Đồng Nai' in li.text or 'Long An' in li.text or 'Tiền Giang' in li.text or 'Bến Tre' in li.text or 'Tây Ninh' in li.text or 'Vùng Tàu' in li.text or 'Bình Phước' in li.text:
                if 'Tiền Giang' in li.text :
                    href = a['href']
                    complete_link = "https://hosocongty.vn/" + href
                    link_hcm.append(complete_link)
                    print("Save : {}".format(complete_link))
    except requests.exceptions.RequestException as e:
        print(f"Lỗi: {e}")
print("Had {} Company HCM,DN,BD,LA".format(len(link_hcm)))

total_list = []
for url in link_hcm:
    company_info = get_company_info(url, headers)
    if company_info:
        total_list.append(company_info)

for row in total_list:
    for i in range(len(row)):
        row[i] = row[i].replace('"', "'")


print("Had {} Company ".format(len(total_list)))

main(total_list, 'tien_giang_2023_21_2_2024.xlsx')

