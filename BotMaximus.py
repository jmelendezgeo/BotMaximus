from selenium import webdriver
from selenium.webdriver.support.ui import Select
import pandas as pd
import random
import time

def fill_form(row):
    # Escribir User First Name
    user_firstname_field = driver.find_element_by_id('userFirstName')
    user_firstname_field.clear()
    user_firstname_field.send_keys(row.FirstName)

    # Escribir User Last Name
    user_lastname_field = driver.find_element_by_id('userLastName')
    user_lastname_field.clear()
    user_lastname_field.send_keys(row.LastName)

    # Escribir Beneficiario
    bene_firstname_field = driver.find_element_by_id('beneFirstName')
    bene_firstname_field.clear()
    bene_firstname_field.send_keys(row.FirstName)

    # Escribir Beneficiario
    bene_lastname_field = driver.find_element_by_id('beneLastName')
    bene_lastname_field.clear()
    bene_lastname_field.send_keys(row.LastName)

    # Escribir Bene DOB
    bene_dob_field = driver.find_element_by_id('beneDOB')
    bene_dob_field.clear()
    bene_dob_field.send_keys(row.BirthDate)

    # Escribir Bene ID
    bene_id_field = driver.find_element_by_id('BeneExtId')
    bene_id_field.clear()
    bene_id_field.send_keys(row.MedicaID)

    # Select relationship
    drp_relation = Select(driver.find_element_by_id("userToBeneRelationCd"))
    drp_relation.select_by_visible_text("Self")

    # Click NextPage Buttom
    next_page_buttom = driver.find_element_by_id("login_request")

    time.sleep(random.choice(seconds_to_wait))
    next_page_buttom.click()

def extract_phonenumber():
    try:
        household_phone_number = driver.find_element_by_xpath("/html/body/div[1]/div[3]/div[3]/div[1]/form/div[1]/fieldset/div[2]/div[2]/label/span[2]")
        phone_number = driver.find_element_by_xpath("/html/body/div[1]/div[3]/div[3]/div[1]/form/div[1]/fieldset/div[2]/div[3]/label/span[2]")
        return (phone_number, household_phone_number)
    except:
        return ''

def update_dataframe(new):
    old_dataframe = pd.read_excel("data_cleaned.xlsx", sheet_name="Epaces")
    old_dataframe.update(new)
    old_dataframe.to_excel("data_cleaned.xlsx", sheet_name="Epaces", index=False)
    print("El archivo original ya esta actualizado con la nueva informacion")


if __name__ =='__main__':
    maximus_url = 'https://ssa-nyeb.maximus.com/NYSelfService/resources/portal/index.html#clients/login_public/case_information'

    driver = webdriver.Firefox(executable_path ='./geckodriver')
    driver.implicitly_wait(30)
    driver.maximize_window()
    driver.get(maximus_url)

    #Seleccionar segundos para esperar. En este caso, entre 10 y 30 segundos
    seconds_to_wait = range(10,30)

    epaces_df=pd.read_excel("data_cleaned.xlsx", sheet_name="Epaces")
    columns_to_keep =['MedicaID','Sex','BirthDate','LastName','FirstName',]
    epaces_df = epaces_df[columns_to_keep]
    epaces_df["BirthDate"] = pd.to_datetime(epaces_df["BirthDate"])
    epaces_df["BirthDate"] = epaces_df.BirthDate.dt.strftime('%m/%d/%Y')

    #Definir el rango para procesar. Recuerda guardar el documento
    epaces_df = epaces_df[2400:2500]

    print(f"Vamos a procesar {len(epaces_df)} registros")
    counter = 0
    for index,row in epaces_df.iterrows():
        fill_form(row)
        time.sleep(5)
        try:
            phone_number , household_phone_number = extract_phonenumber()
            epaces_df.at[index,"PhoneNumber"] = phone_number.text
            epaces_df.at[index,"Household_PhoneNumber"] = household_phone_number.text
        except:
            epaces_df.at[index,"PhoneNumber"] = ''
            epaces_df.at[index,"Household_PhoneNumber"] = ''

        counter += 1
        print(f"Ya he procesado {counter} de {len(epaces_df)} registros")

        time.sleep(random.choice(seconds_to_wait))
        if driver.current_url != maximus_url:
            driver.back()

    phone_found = (~(epaces_df["PhoneNumber"].values == '')).sum()
    household_phone_found = (~(epaces_df["Household_PhoneNumber"].values == '')).sum()

    driver.close()
    print(f"""Finalizamos la ejecución.\n
    Hemos conseguido {phone_found} PhoneNumbers y {household_phone_found} Household PhoneNumbers.""")
    update_dataframe(epaces_df)
