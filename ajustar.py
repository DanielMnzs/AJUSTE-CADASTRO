from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import time
import pandas as pd


caminho = r'AJUSTE CADASTRO/PLUS.xlsx'
ler = pd.read_excel(caminho)
tabela = pd.DataFrame(ler)
print(tabela)

x = 0
y = 0





while x < len (tabela):

    plu = tabela.iloc[x, y]
    comprimento = tabela.iloc[x, y + 3]
    largura = tabela.iloc[x, y + 4]
    altura = tabela.iloc[x , y + 5]
    
    volume = tabela.iloc[x, y + 12]
    volume = volume.replace(",", ".")
    volume_2 = "{:.4f}".format(float(volume))
    volume_2 = volume_2.replace(".",",")
    embalagem = tabela.iloc[x, y + 10]
    embalagem = str(embalagem).replace(",",".")
    embalagem_2 = "{:.4f}".format(float(embalagem))
    embalagem_2 = embalagem_2.replace(".",",")
    
    pacote = tabela.iloc[x, y + 2]

    print(f"a plu = {plu}, o comprimento é {comprimento},a largura é {largura},a altura é {altura},a embalagem é {embalagem} ")


    chrome_options = Options()
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])  
    service = Service(r"C:/Program Files/Google/Chrome/Application/chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=chrome_options)

    driver.get(r'https://wmswmos1942.cbd.root.gpa/manh/index.html?i=214')

    def macro(by, caminho):
        return WebDriverWait(driver, 30).until(EC.element_to_be_clickable((by, caminho)))

    id_input = macro(By.XPATH, '//*[@id="username"]')
    id_input.send_keys("5446902")

    senha_input = macro(By.XPATH, '//*[@id="password"]')
    senha_input.send_keys("@Gpafevereiro2024")
    senha_input.send_keys(Keys.RETURN)
    
    time.sleep(3)


    actions = ActionChains(driver)
    def coord(x, y, texto):
        actions.move_by_offset(x, y).click().perform()
        time.sleep(0.5)
        actions.send_keys(texto).perform()

    tres = coord(17, 20, "itens")
    time.sleep(2)
    for i in range(4):
        actions.send_keys(Keys.ARROW_DOWN).perform()
    actions.send_keys(Keys.ENTER).perform()

    try:
        iframe = macro(By.TAG_NAME, "iframe")
        driver.switch_to.frame(iframe)
        print("iframe acessado com sucesso")

    except:
        print("iframe não encontrado")

    itens = macro(By.XPATH, '//*[@id="dataForm:ItemList_lv:ItemList_filterId:itemLookUpId"]')
    itens.send_keys(str(plu))

    button_aplicar = macro(By.XPATH, '//*[@id="dataForm:ItemList_lv:ItemList_filterId:ItemList_filterIdapply"]')
    button_aplicar.click()

    selection = macro(By.XPATH, '//*[@id="checkAll_c0_dataForm:ItemList_lv:dataTable"]')
    selection.click()

    button_see = macro(By.XPATH, '//*[@id="dataForm:ItemList_Viewbutton"]')
    button_see.click()

    button_pacotes = macro(By.XPATH, '//*[@id="Item_Package_Tab_lnk"]')
    button_pacotes.click()

    button_editar = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Edit_button"]')
    button_editar.click()

    comprimento_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_inputMask_ItemPkgLength"]')
    actions.double_click(comprimento_input).perform()
    comprimento_input.send_keys(str(comprimento))

    largura_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_inputMask_ItemPkgWidth"]')
    actions.double_click(largura_input).perform()
    largura_input.send_keys(str(largura))
    
    altura_input = macro(By.XPATH,'//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_inputMask_ItemPkgHeight"]')
    actions.double_click(altura_input).perform()
    altura_input.send_keys(str(altura))

    peso_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_inputMask_Weight"]')
    actions.double_click(peso_input).perform()
    peso_input.send_keys(embalagem_2)

    volume_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_inputMask_ItemPkgVolume"]')
    actions.double_click(volume_input).perform()
    volume_input.send_keys(volume_2)

    pacote_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_inputMask_PackageQty"]')
    actions.double_click(pacote_input).perform()
    pacote_input.send_keys(str(pacote))

    caixa = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:0:ItemPackageListEV_PackageUOM_Sel_Menu"]')
    caixa.click()
    caixa.send_keys("Caixa")


    for i in range(1):
        try: 
            pacote_kgs_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_inputMask_PackageQty"]')
            print("Produto é variável")
            actions.double_click(pacote_kgs_input).perform()
            pacote_kgs_input.send_keys("1000")

            Kg = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_PackageUOM_Sel_Menu"]')
            Kg.click()
            Kg.send_keys("Kg")

            comprimento_kgs_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_inputMask_ItemPkgLength"]')
            actions.double_click(comprimento_kgs_input).perform()
            comprimento_kgs_input.send_keys(str(comprimento))

            largura_kgs_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_inputMask_ItemPkgWidth"]')
            actions.double_click(largura_kgs_input).perform()
            largura_kgs_input.send_keys(str(largura))

            altura_kgs_input = macro(By.XPATH,'//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_inputMask_ItemPkgHeight"]')
            actions.double_click(altura_kgs_input).perform()
            altura_kgs_input.send_keys(str(altura))

            peso_kgs_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_inputMask_Weight"]')
            actions.double_click(peso_kgs_input).perform()
            peso_kgs_input.send_keys("1")

            volume_kgs_input = macro(By.XPATH, '//*[@id="dataForm:ItemPackageListEV_dataTable:1:ItemPackageListEV_inputMask_ItemPkgVolume"]')
            actions.double_click(volume_kgs_input).perform()
            volume_kgs_input.send_keys("0")

        except:
            print("OK")

    salvar = macro(By.XPATH, '//*[@id="dataForm:ItemDetailsMain_Save_Detail_button"]')
    salvar.click()

    driver.quit()

    print("PRODUTO AJUSTADO COM SUCESSO!")

    x = x + 1