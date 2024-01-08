import os
from selenium import webdriver as webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options


def cargar_driver(sitio):
    # Inicializar el navegador
    global driver
    global espera
    
    options = webdriver.EdgeOptions()
    options.use_chromium = True
    options.add_argument('--start-maximized')
    options.add_argument('--disable-extensions')
    options.set_capability("ms:inPrivate", True)

    options.add_argument("--disable-notifications")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-features=msServices")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")

    driver_path = 'C:\\DriverEdge\\msedgedriver.exe'
    service = Service(driver_path)
    driver = webdriver.Edge(service=service, options=options)
    driver.get(sitio)
    
    espera = WebDriverWait(driver, 10)

def ok_input():
    alert = driver.switch_to.alert
    alert.accept()

def input_uno(valor):
    txt_name = driver.find_element(By.NAME, "txtName")
    txt_name.send_keys("cell_8cat")

def d_Xpatch(web_xpatch):
    ac1 = espera.until(EC.element_to_be_clickable((By.XPATH, web_xpatch)))
    if ac1 is not None:
        ac1.click()

def d_Xpatch_text(Texto, web_xpatch):
    ac1 = espera.until(EC.element_to_be_clickable((By.XPATH, web_xpatch)))
    if ac1 is not None:
        ac1.send_keys(Texto)

def d_ID(web_ID):
    ac1 = espera.until(EC.element_to_be_clickable((By.ID, web_ID)))
    if ac1 is not None:
        ac1.click()

def d_ID_text(Texto, web_ID):
    ac1 = espera.until(EC.element_to_be_clickable((By.ID, web_ID)))
    if ac1 is not None:
        ac1.send_keys(Texto)

def d_SKcss(Texto, web_CSS):
    ac1 = espera.until(EC.element_to_be_clickable((By.CSS_SELECTOR, web_CSS)))
    if ac1 is not None:
        ac1.send_keys(Texto)

def d_css_click(web_CSS):
    ac1 = espera.until(EC.element_to_be_clickable((By.CSS_SELECTOR, web_CSS)))
    if ac1 is not None:
        ac1.click()

def d_I_text(Texto, web_XPATCH):
    ac1 = espera.until(EC.element_to_be_clickable((By.LINK_TEXT, web_XPATCH)))
    if ac1 is not None:
        ac1.send_keys(Texto)

def d_t_link(web_ID):
    ac1 = espera.until(EC.element_to_be_clickable((By.LINK_TEXT, web_ID)))
    if ac1 is not None:
        ac1.click()

def d_SKcss_txt(Texto, web_CSS):
    ac1 = espera.until(EC.element_to_be_clickable((By.CSS_SELECTOR, web_CSS)))
    if ac1 is not None:
        ac1.send_keys(Texto)

def Adorisio(valor):
    txt_name = driver.find_element(By.CSS_SELECTOR, "input#txtName")
    txt_name.send_keys(valor)

#cambiar a ventana de popup
def cambiar_ventana():
    # Obtener los identificadores de las ventanas abiertas
    ventanas = driver.window_handles

# Imprimir los identificadores de las ventanas
    for ventana in ventanas:
        print("Identificador de ventana:", ventana)
    global main_window
    main_window = driver.current_window_handle
    popup_window = [window for window in driver.window_handles if window != main_window][0]
    driver.switch_to.window(popup_window)

def volver_ventana():
    driver.switch_to.window(main_window)

def Guardar_Num_Ticket(num_tk):
    valor=''

    element = espera.until(EC.element_to_be_clickable((By.ID, num_tk)))
    if element is not None:
        valor = element.text

    return valor

def Cerrar_Driver():
    driver.close()


'''
1x - Corporate Aplications
2x - Data Center
3x - Local Support
4x - Networking // 42 
'''
# CAT principales ##############################
def sel_cat(num_cat):
    categories = {
        1: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[2]/div/div[2]/div/div[1]/img",
        2: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[2]/div/div[3]/div/div[1]/img",
        3: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[2]/div/div[5]/div/div[1]/img",
        4: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[2]/div/div[6]/div/div[1]/img"
    }
    index = categories.get(num_cat, "")
    return index, num_cat

# 4 Solo para netowrking ###########################
def sel_net(num_sub):
    categories = {
        # telefonia
        1: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[4]/div/div[1]/img",
        # Networking
        2: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[3]/div/div[1]/img"
    }
    index = categories.get(num_sub, "")
    return index, num_sub



# 3 para Local Support #########################
def sel_LS(num_ls):
    categories = {
        # IT Management ok
        1: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[10]/div",
        # Commercial Applications
        2: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[4]/div/div[1]/img",
        # Computer equipment ok
        3: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[6]/div/div[1]/img",
        # Email
        5: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[8]/div/div[1]/img",
        # User Account
        6: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[14]/div/div[1]/img",
        # Computer Accesories
        7: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[7]/div/div[1]/img"
    }
    index = categories.get(num_ls, "")
    return index, num_ls

# 2 Data Center #########################
def sel_DC(num_select):
    categories = {
        # Servers
        1: "/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[3]/div/div[3]/div/div[1]/img"
    }
    index = categories.get(num_select, "")
    return index, num_select

# 1 Corporated Aplications #########################
def sel_CA(num_select):
    categories = {
        # Servers
        1: "/html/body/form/table[1]/tbody/tr[2]/td[2]/div/div/div[2]/div[1]/div/div/div/ul/div/li[1]/ul/li/div/img[1]"
    }
    index = categories.get(num_select, "")
    return index, num_select
################\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\#############################
#solo para networking 4 / 2
def sel_net_final(num_selec):
    # 1: VPN Client Tunel
    # 2: Global Protect
    # 9: Add/Remove User
    # 3: Office Network Failure
    # 4: Firewalls rules
    # 8: Network Management
    categories = {
        1: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[11]/div/div[1]/img',
        2: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[7]/div/div[1]/img',
        3: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[9]/div/div[1]/img',
        4: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[6]/div/div[1]/img',
        5: '//div[contains(text(),"DNS Management")]',
        6: '//div[contains(text(),"ADS Management")]',
        7: '//div[contains(text(),"RFC")]',
        8: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[8]/div/div[1]/img',
        9: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[2]/div/div[1]/img',
        10: '//div[contains(text(),"Debug networks")]'
    }
    index = categories.get(num_selec, "")
    return index, num_selec

#solo Data Center (es la unica categoria que tiene)
def sel_DC_final(num_selec):
    # 8: MS Server Support
    # 9: Linux Server SP

    categories = {
        1: '//span[contains(text(),"Server Request")]',    
        2: '//span[contains(text(),"Evidence collection")]',
        3: '//span[contains(text(),"Backups")]',
        4: '//span[contains(text(),"Failure")]',
        5: '//span[contains(text(),"Packages")]',
        6: '//span[contains(text(),"DC vulnerability")]',
        7: '//span[contains(text(),"Attention")]',
        8: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[13]/div/div[1]/img',
        9: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[12]/div/div[1]/img',
        10: '//span[contains(text(),"Hardware Management")]',
        11: '//span[contains(text(),"Asset Management (CMDB)")]',
        12: '//span[contains(text(),"Inventory and cloud security")]',
        13: '//span[contains(text(),"Antivirus Servers")]',
        14: '//span[contains(text(),"Patches Servers")]',
        15: '//span[contains(text(),"Request for rules in the FW")]',
        16: '//span[contains(text(),"RFC")]',
        17: '//span[contains(text(),"JDE Account Management")]'
    }
    index = categories.get(num_selec, "")
    return index, num_selec

#solo Local Support 3 / 3
def sel_LS3_final(num_selec):
    categories = {
        # 11: Coodination
        # 13: Other IT areas
        11: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[2]/div/div[1]/img',
        12: '//span[contains(text(),"Quotes/purchases")]',
        13: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[5]/div/div[1]/img',
        14: '//span[contains(text(),"Logistics")]',
        23: '/html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[4]/div/div[1]/img',
        
        
    }
    index = categories.get(num_selec, "")
    return index, num_selec


# Unica opcion de Corporated Aplication
def sel_CA_final(num_selec):
    categories = {
        1: '//span[contains(text(),"Support")]',
    }
    index = categories.get(num_selec, "")
    return index, num_selec

# //span[contains(text(),"VPN")]

    '''
        21: '//span[contains(text(),"Access")]',
        22: '//span[contains(text(),"Modify profile")]',
        23: '//span[contains(text(),"Failure")]',
        24: '//span[contains(text(),"Install")]',
        25: '//span[contains(text(),"Uninstall")]',
        26: '//span[contains(text(),"Debugging account")]',
        27: '//span[contains(text(),"License")]',
        28: '//span[contains(text(),"DB Backup")]',
        31: '//span[contains(text(),"Computer equipment installation")]',
        32: '//span[contains(text(),"Failure")]',
        /html/body/table/tbody/tr[3]/td/table/tbody/tr/td[4]/table/tbody/tr[1]/td/table/tbody/tr[3]/td/form/table/tbody/tr[4]/td/div[4]/div/div[4]/div/div[1]/img
        33: '//span[contains(text(),"Information Backups")]',
        34: '//span[contains(text(),"Recovery bitlocker")]',
        35: '//span[contains(text(),"Password BIOS")]',
        36: '//span[contains(text(),"Antivirus")]',
        37: '//span[contains(text(),"Cambio (actualización de equipo)")]',
        38: '//span[contains(text(),"Alta")]',
        39: '//span[contains(text(),"Baja")]',
        61: '//span[contains(text(),"Unlock")]',
        62: '//span[contains(text(),"Reset")]',
        63: '//span[contains(text(),"Extension")]',
        64: '//span[contains(text(),"Name change/correction")]',
        65: '//span[contains(text(),"On Board")]',
        66: '//span[contains(text(),"Termination")]'
        '''