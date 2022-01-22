
import os
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

# Inicializa as variáveis
client = []
excel_file = r'challenge.xlsx'
url_site = "http://www.rpachallenge.com"

# Realiza a leitura da base
base = pd.read_excel(excel_file)

# Monta o dicionário
for index in range(len(base)):
    client.append({
        "Name": base["First Name"][index],
        "LastName": base["Last Name "][index],
        "Company": base["Company Name"][index],
        "Role": base["Role in Company"][index],
        "Address": base["Address"][index],
        "Email": base["Email"][index],
        "Phone": base["Phone Number"][index]
    })

# Desabilita todos os tipos de logs encontrados
os.environ['WDM_LOG_LEVEL'] = "0"  # Log do Download do Driver correspondente a versão do Chrome
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--log-level=OFF")
chrome_options.add_argument("--disable-logging")
chrome_options.add_argument('--disable-crash-reporter')
chrome_options.add_argument('--disable-in-process-stack-traces')
chrome_options.add_argument('--start-maximized')

# Cria instância do navegador
driver = webdriver.Chrome(ChromeDriverManager().install(), options=chrome_options, service_log_path=os.devnull)

# Acessa o site
driver.get(url_site)

# Clica em Start
driver.find_element_by_tag_name("button").click()

# Processa o Laço principal
for item in range(len(client)):

    # Preenche os campos
    driver.find_element_by_xpath("//*[text() = 'First Name']/following-sibling::input").send_keys(client[item]["Name"])
    driver.find_element_by_xpath("//*[text() = 'Last Name']/following-sibling::input").send_keys(client[item]["LastName"])
    driver.find_element_by_xpath("//*[text() = 'Company Name']/following-sibling::input").send_keys(client[item]["Company"])
    driver.find_element_by_xpath("//*[text() = 'Role in Company']/following-sibling::input").send_keys(client[item]["Role"])
    driver.find_element_by_xpath("//*[text() = 'Address']/following-sibling::input").send_keys(client[item]["Address"])
    driver.find_element_by_xpath("//*[text() = 'Email']/following-sibling::input").send_keys(client[item]["Email"])
    driver.find_element_by_xpath("//*[text() = 'Phone Number']/following-sibling::input").send_keys(str(client[item]["Phone"]))

    # Clica em submit
    submit_button = driver.find_elements_by_tag_name("input")
    for input_element in enumerate(submit_button):
        if input_element[1].get_attribute("value") == "Submit":
            input_element[1].click()

# Fecha o Chrome
driver.close()
