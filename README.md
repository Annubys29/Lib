from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import openpyxl
import re
from openpyxl.styles import PatternFill
import time
from openpyxl.styles import NamedStyle, Font
import schedule
import time
driver = webdriver.Chrome()
driver.get("https://looqbox.viavarejo.com.br/")

# Aguarda a presença dos campos de login
username_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH,"//*[@id='root']/div/div/div[2]/div/div/form/div[1]/input"))  # Substitua pelo ID correto
)
password_field = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH,"//*[@id='root']/div/div/div[2]/div/div/form/div[3]/input"))  # Substitua pelo ID correto
)

username_field.send_keys("User")
password_field.send_keys("Password")

login_button = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='root']/div/div/div[2]/div/div/form/div[4]/button"))
)
login_button.click()
#seleciona os resultados do dia

Resultados_dia = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='search-area']/div[2]/div[2]/div/div[2]/div[1]/div[3]/a/div"))
)

Resultados_dia.click()

#Copia o horario de atualização
copiar_horario = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH,"//*[@id='board0']/div[1]/span[3]")))
text = copiar_horario.text
number_match = re.search(r'\b(?:[01]\d|2[0-3]):[0-5]\d\b',text)
extract_number = number_match.group()
print(extract_number)

wb2 = openpyxl.load_workbook('C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Mercantil.xlsx')
sheet35 = wb2["Planilha1"]
sheet35 = wb2.active

sheet35["A1"] = extract_number
wb2.save('C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Mercantil.xlsx')


#Coleta os dados do mercantil
Coleta_de_dados = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[1]/div/div/div/div/div[7]/div/div/div/div/div/div/div/div/table"))
)

table_data1 = [[cell.text for cell in row.find_elements(By.TAG_NAME,"td")]
              for row in  Coleta_de_dados.find_elements(By.TAG_NAME,"tr")]
wb12 = openpyxl.load_workbook('C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Mercantil.xlsx')
sheet5 = wb12["Planilha1"]
sheet5 = wb12.active
for row_index, row in enumerate(table_data1):
    for col_index, cell_value in enumerate(row):
        sheet5.cell(row=row_index + 2, column = col_index + 1).value = cell_value

for row_index in range(4, 19) :  # Linhas de 3 a 18
    for col_index, cell in enumerate(sheet5.iter_cols(min_row=row_index, max_row=row_index, values_only=True)):
        if cell and col_index == 0:
            try:
                sheet5.cell(row=row_index, column=col_index + 1, value=float(cell[0]))
            except ValueError:
                sheet5.cell(row=row_index, column=col_index + 1, value=None)

for column_index in range(1, 12):  # Colunas de 1 a 11
    for row_index in range(3, 19):  # Linhas de 3 a 18
        cell_value = sheet5.cell(row=row_index, column=column_index).value

        # Lista de colunas que devem ter vírgulas removidas e o tipo de dado alterado para número
        colunas_alterar = [1, 2, 3, 4, 7, 8, 10, 11]  # Adapte conforme necessário

        if column_index in colunas_alterar and cell_value is not None and isinstance(cell_value, str):
            # Remova vírgulas e converta para número
            cleaned_value = cell_value.replace('.', '')

            try:
                float_value = float(cleaned_value)
                sheet5.cell(column=column_index, row=row_index, value=float_value)
            except ValueError:
                sheet5.cell(column=column_index, row=row_index, value=None)
        elif cell_value is not None and isinstance(cell_value, (int, float)):
            # Se a coluna não estiver na lista, mas o valor já for numérico, mantenha-o
            sheet5.cell(column=column_index, row=row_index, value=cell_value)
        elif cell_value is not None and not isinstance(cell_value, (int, float)):
            sheet5.cell(column=column_index, row=row_index, value=None)

 
wb12.save('C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Mercantil.xlsx')  


#Seleciona a segunda parte
Caixa_de_selecao = WebDriverWait(driver,30 ).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='rc-tabs-0-panel-b0f2e1']/div/div/div/div/div[7]/div/div/div/div/div/ul/li[3]"))
)
Caixa_de_selecao.click()



#Coleta a segunda parte
Coleta_dados2 = WebDriverWait(driver, 30).until(
     EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[1]/div/div/div/div/div[7]/div/div/div/div/div/div/div/div/table"))
)     

table_data2 = [[cell2.text for cell2 in row.find_elements(By.TAG_NAME,"td")]
              for row in  Coleta_dados2.find_elements(By.TAG_NAME,"tr")]


wb = openpyxl.load_workbook('C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Mercantil.xlsx')

sheet = wb["Planilha1"]

sheet = wb.active
for row_index, row in enumerate(table_data2):
    for col_index, cell_value in enumerate(row):
        sheet.cell(row=row_index + 18, column = col_index + 1).value = cell_value

for row_index in range(20, 31) :  # Linhas de 3 a 18
    for col_index, cell in enumerate(sheet5.iter_cols(min_row=row_index, max_row=row_index, values_only=True)):
        if cell and col_index == 0:
            try:
                sheet.cell(row=row_index, column=col_index + 1, value=float(cell[0]))
            except ValueError:
                sheet.cell(row=row_index, column=col_index + 1, value=None)

for column_index in range(1, 12):  # Colunas de 1 a 11
    for row_index in range(20, 32):  # Linhas de 3 a 18
        cell_value = sheet.cell(row=row_index, column=column_index).value

        # Lista de colunas que devem ter vírgulas removidas e o tipo de dado alterado para número
        colunas_alterar = [1, 2, 3, 4, 7, 8, 10, 11]  # Adapte conforme necessário

        if column_index in colunas_alterar and cell_value is not None and isinstance(cell_value, str):
            # Remova vírgulas e converta para número
            cleaned_value = cell_value.replace('.', '')

            try:
                float_value = float(cleaned_value)
                sheet.cell(column=column_index, row=row_index, value=float_value)
            except ValueError:
                sheet.cell(column=column_index, row=row_index, value=None)
        elif cell_value is not None and isinstance(cell_value, (int, float)):
            # Se a coluna não estiver na lista, mas o valor já for numérico, mantenha-o
            sheet.cell(column=column_index, row=row_index, value=cell_value)
        elif cell_value is not None and not isinstance(cell_value, (int, float)):
            sheet.cell(column=column_index, row=row_index, value=None)
wb.save('C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Mercantil.xlsx')

#Seleciona a parte de serviços

Servico = WebDriverWait(driver,30 ).until(
    EC.presence_of_element_located((By.XPATH, "//*[@id='board0']/div[3]/div/div/div[1]/div/div[1]/div[1]/div/div[3]"))
)
Servico.click()



#Coleta de dados de serviço

Coleta_dados_Servicos = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[1]/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[3]/div/div/div/div/div[9]/div/div/div/div/div/div/div/div/table/tbody'))
    
)     

table_data3 = [[cell3.text for cell3 in row.find_elements(By.TAG_NAME,"td")]
              for row in  Coleta_dados_Servicos.find_elements(By.TAG_NAME,"tr")]


wb1 = openpyxl.load_workbook("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Serviços.xlsx")
sheet2 = wb1["Planilha1"]

sheet2 = wb1.active
for row_index, row in enumerate(table_data3):
    for col_index, cell_value in enumerate(row):
        sheet2.cell(row=row_index + 1, column = col_index + 1).value = cell_value
        
for row_index in range(2, 16) :  # Linhas de 3 a 18
    for col_index, cell in enumerate(sheet2.iter_cols(min_row=row_index, max_row=row_index, values_only=True)):
        if cell and col_index == 0:
            try:
                sheet.cell(row=row_index, column=col_index + 1, value=float(cell[0]))
            except ValueError:
                sheet.cell(row=row_index, column=col_index + 1, value=None)

for column_index in range(1, 12):  # Colunas de 1 a 11
    for row_index in range(2, 17):  # Linhas de 3 a 18
        cell_value = sheet2.cell(row=row_index, column=column_index).value

        # Lista de colunas que devem ter vírgulas removidas e o tipo de dado alterado para número
        colunas_alterar = [1, 2, 3, 5, 6, 8, 9]  # Adapte conforme necessário

        if column_index in colunas_alterar and cell_value is not None and isinstance(cell_value, str):
            # Remova vírgulas e converta para número
            cleaned_value = cell_value.replace('.', '')

            try:
                float_value = float(cleaned_value)
                sheet2.cell(column=column_index, row=row_index, value=float_value)
            except ValueError:
                sheet2.cell(column=column_index, row=row_index, value=None)
        elif cell_value is not None and isinstance(cell_value, (int, float)):
            # Se a coluna não estiver na lista, mas o valor já for numérico, mantenha-o
            sheet2.cell(column=column_index, row=row_index, value=cell_value)
        elif cell_value is not None and not isinstance(cell_value, (int, float)):
            sheet2.cell(column=column_index, row=row_index, value=None)        

wb1.save("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Serviços.xlsx")

Caixa_de_selecao2 = WebDriverWait(driver,30 ).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div[1]/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[3]/div/div/div/div/div[9]/div/div/div/div/div/ul/li[3]"))
)
Caixa_de_selecao2.click()


Coleta_dados_Servicos2 = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[1]/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[3]/div/div/div/div/div[9]/div/div/div/div/div/div/div/div'))
    
)     

table_data3 = [[cell3.text for cell3 in row.find_elements(By.TAG_NAME,"td")]
              for row in  Coleta_dados_Servicos2.find_elements(By.TAG_NAME,"tr")]


wb1 = openpyxl.load_workbook("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Serviços.xlsx")
sheet2 = wb1["Planilha1"]

sheet2 = wb1.active
for row_index, row in enumerate(table_data3):
    for col_index, cell_value in enumerate(row):
        sheet2.cell(row=row_index + 17, column = col_index + 1).value = cell_value
        
for column_index in range(1, 12):  # Colunas de 1 a 11
    for row_index in range(19, 31):  # Linhas de 3 a 18
        cell_value = sheet2.cell(row=row_index, column=column_index).value

        # Lista de colunas que devem ter vírgulas removidas e o tipo de dado alterado para número
        colunas_alterar = [1, 2, 3, 5, 6, 8, 9]  # Adapte conforme necessário

        if column_index in colunas_alterar and cell_value is not None and isinstance(cell_value, str):
            # Remova vírgulas e converta para número
            cleaned_value = cell_value.replace('.', '')

            try:
                float_value = float(cleaned_value)
                sheet2.cell(column=column_index, row=row_index, value=float_value)
            except ValueError:
                sheet2.cell(column=column_index, row=row_index, value=None)
        elif cell_value is not None and isinstance(cell_value, (int, float)):
            # Se a coluna não estiver na lista, mas o valor já for numérico, mantenha-o
            sheet2.cell(column=column_index, row=row_index, value=cell_value)
        elif cell_value is not None and not isinstance(cell_value, (int, float)):
            sheet2.cell(column=column_index, row=row_index, value=None)        


wb1.save("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/Serviços.xlsx")

#Parte Cdc

CDC = WebDriverWait(driver,30 ).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div/div[4]/div[3]/div/div/div[1]/div/div[1]/div[1]/div/div[2]"))
)
CDC.click()

Coleta_dados_CDC = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[1]/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[2]/div/div/div/div/div[9]/div/div/div/div/div/div/div/div/table'))
    
)     

table_data4 = [[cell4.text for cell4 in row.find_elements(By.TAG_NAME,"td")]
              for row in  Coleta_dados_CDC.find_elements(By.TAG_NAME,"tr")]


wb4 = openpyxl.load_workbook("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/CDC.xlsx")
sheet3 = wb4["Planilha1"]

sheet3 = wb4.active
for row_index, row in enumerate(table_data4):
    for col_index, cell_value in enumerate(row):
        sheet3.cell(row=row_index + 1, column = col_index + 1).value = cell_value

for column_index in range(1, 12):  # Colunas de 1 a 11
    for row_index in range(2, 19):  # Linhas de 3 a 18
        cell_value = sheet3.cell(row=row_index, column=column_index).value

        # Lista de colunas que devem ter vírgulas removidas e o tipo de dado alterado para número
        colunas_alterar = [1, 2, 3, 5, 6, 8, 9]  # Adapte conforme necessário

        if column_index in colunas_alterar and cell_value is not None and isinstance(cell_value, str):
            # Remova vírgulas e converta para número
            cleaned_value = cell_value.replace('.', '')

            try:
                float_value = float(cleaned_value)
                sheet3.cell(column=column_index, row=row_index, value=float_value)
            except ValueError:
                sheet3.cell(column=column_index, row=row_index, value=None)
        elif cell_value is not None and isinstance(cell_value, (int, float)):
            # Se a coluna não estiver na lista, mas o valor já for numérico, mantenha-o
            sheet3.cell(column=column_index, row=row_index, value=cell_value)
        elif cell_value is not None and not isinstance(cell_value, (int, float)):
            sheet3.cell(column=column_index, row=row_index, value=None)        

wb4.save("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/CDC.xlsx")


Caixa_de_selecao3 = WebDriverWait(driver,30 ).until(
    EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div/div[2]/div[2]/div/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[2]/div/div/div/div/div[9]/div/div/div/div/div/ul/li[3]"))
)
Caixa_de_selecao3.click()






Coleta_dados_CDC2 = WebDriverWait(driver, 30).until(
    EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div/div[2]/div[2]/div[1]/div/div[4]/div[3]/div/div/div[1]/div/div[2]/div/div[2]/div/div/div/div/div[9]/div/div/div/div/div/div/div/div/table'))
    
)     

table_data5 = [[cell5.text for cell5 in row.find_elements(By.TAG_NAME,"td")]
              for row in  Coleta_dados_CDC2.find_elements(By.TAG_NAME,"tr")]


wb3 = openpyxl.load_workbook("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/CDC.xlsx")
sheet4 = wb3["Planilha1"]

sheet4 = wb3.active
for row_index, row in enumerate(table_data5):
    for col_index, cell_value in enumerate(row):
        sheet4.cell(row=row_index + 20, column = col_index + 1).value = cell_value



for column_index in range(1, 12):  # Colunas de 1 a 11
    for row_index in range(20, 34):  # Linhas de 3 a 18
        cell_value = sheet4.cell(row=row_index, column=column_index).value

        # Lista de colunas que devem ter vírgulas removidas e o tipo de dado alterado para número
        colunas_alterar = [1, 2, 3, 5, 6, 8, 9]  # Adapte conforme necessário

        if column_index in colunas_alterar and cell_value is not None and isinstance(cell_value, str):
            # Remova vírgulas e converta para número
            cleaned_value = cell_value.replace('.', '')

            try:
                float_value = float(cleaned_value)
                sheet4.cell(column=column_index, row=row_index, value=float_value)
            except ValueError:
                sheet4.cell(column=column_index, row=row_index, value=None)
        elif cell_value is not None and isinstance(cell_value, (int, float)):
            # Se a coluna não estiver na lista, mas o valor já for numérico, mantenha-o
            sheet4.cell(column=column_index, row=row_index, value=cell_value)
        elif cell_value is not None and not isinstance(cell_value, (int, float)):
            sheet4.cell(column=column_index, row=row_index, value=None)        

wb3.save("C:/Users/2100501557/Desktop/Atualizador de Parcial - Loogbox/CDC.xlsx")
driver.quit()
