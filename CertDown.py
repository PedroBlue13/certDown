import tkinter as tk
from tkinter import messagebox
import os
import sys
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from anticaptchaofficial.recaptchav2proxyless import recaptchaV2Proxyless
import time

API_KEY = "YOU-API-KEY"

def update_status(message):
    status_label.config(text=message)
    root.update_idletasks()

def run_automation(document_number):
    try:
        update_status("Configuração do WebDriver...")
        
        # Configura o WebDriver para o Chrome com opções adicionais
        chrome_options = webdriver.ChromeOptions()
        chrome_options.add_argument("--window-position=-10000,-10000")  # Modo off-screen
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")
        chrome_options.add_argument("--kiosk-printing")  # Adiciona a opção de impressão sem janela

        # Configurações para salvar automaticamente como PDF
        save_path = "C:\\certidao"
        if not os.path.exists(save_path):
            os.makedirs(save_path)

        prefs = {
            "printing.print_preview_sticky_settings.appState": '{"recentDestinations":[{"id":"Save as PDF","origin":"local","account":""}],"selectedDestinationId":"Save as PDF","version":2}',
            "savefile.default_directory": save_path,
            "download.default_directory": save_path,
            "download.prompt_for_download": False,  # Evita que a caixa "Salvar como" seja mostrada
            "download.directory_upgrade": True
        }

        chrome_options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

        update_status("Acessando o site...")
        
        # Abra o site
        driver.get("https://web.trf3.jus.br/certidao-regional/CertidaoCivelEleitoralCriminal/SolicitarDadosCertidao")

        # Simula um tempo de leitura da página antes de interagir com os elementos
        time.sleep(2)

        update_status("Selecionando o tipo de certidão...")
        
        # Seleção do tipo de certidão
        dropdowncertidao = driver.find_element(By.ID, "Tipo")
        dropdowncertidao.click()
        time.sleep(1)
        option_to_selectcetidao = driver.find_element(By.XPATH, '//*[@id="Tipo"]/option[2]')
        option_to_selectcetidao.click()
        update_status("Tipo de certidão selecionado. [OK]")

        update_status("Selecionando o tipo de documento...")
        
        # Seleção tipo de documento
        dropdowndocumento = driver.find_element(By.ID, "TipoDeDocumento")
        dropdowndocumento.click()
        time.sleep(1)
        option_to_selectdocumento = driver.find_element(By.XPATH, '//*[@id="TipoDeDocumento"]/option[3]')
        option_to_selectdocumento.click()
        update_status("Tipo de documento selecionado. [OK]")

        update_status("Preenchendo o campo documento...")
        
        # Preenchimento campo documento
        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '//*[@id="Documento"]'))
        )
        input_field.clear()
        input_field.send_keys(document_number)
        update_status("Campo documento preenchido. [OK]")

        update_status("Selecionando a abrangência...")
        
        # Seleção da abrangência
        dropdownregional = driver.find_element(By.ID, 'TipoDeAbrangencia')
        dropdownregional.click()
        time.sleep(1)
        option_to_selectregional = driver.find_element(By.XPATH, '//*[@id="TipoDeAbrangencia"]/option[2]')
        option_to_selectregional.click()
        update_status("Abrangência selecionada. [OK]")

        # Captura da URL da página atual
        page_url = driver.current_url

        update_status("Resolvendo o reCAPTCHA...")
        
        # chave da API
        SITE_KEY = "Wildest Dream - Taylor Swiftcher"

        # Usar o solver do Anti-Captcha para resolver o reCAPTCHA
        solver = recaptchaV2Proxyless()
        solver.set_verbose(1)
        solver.set_key(API_KEY)
        solver.set_website_url(page_url)
        solver.set_website_key(SITE_KEY)

        # Resolver o reCAPTCHA
        g_response = solver.solve_and_return_solution()
        if g_response != 0:
            print("g-response: " + g_response)
            
            # Tente localizar o campo de resposta do reCAPTCHA e insira o token
            driver.execute_script(f"document.getElementById('g-recaptcha-response').innerHTML = '{g_response}'")
            update_status("reCAPTCHA resolvido. [OK]")

            update_status("Submetendo o formulário...")
            
            # Simula um tempo de leitura da página antes de submeter o formulário
            time.sleep(2)

            # Submeter o formulário ou realizar a próxima ação necessária
            submit_button = driver.find_element(By.ID, "submit")
            submit_button.click()
            update_status("Formulário submetido. [OK]")

            update_status("Preparando para impressão...")
            
            # Acionar o botão de impressão
            printbutton = driver.find_element(By.ID, 'botaoImprimirCertidao')
            printbutton.click()
            time.sleep(4)
            driver.execute_script('window.print();')
            update_status("Impressão preparada. [OK]")

        else:
            print("task finished with error " + solver.error_code)
            update_status(f"Erro ao resolver reCAPTCHA: {solver.error_code}")

        update_status("Finalizando e salvando PDF...")
        
        # Simula um tempo para finalizar e visualizar o resultado
        time.sleep(5)

        # Verifica se algum arquivo PDF foi salvo no diretório
        pdf_files = [f for f in os.listdir(save_path) if f.endswith('.pdf')]
        if pdf_files:
            # Supondo que o arquivo PDF mais recente seja o que foi gerado
            latest_pdf = max(pdf_files, key=lambda x: os.path.getctime(os.path.join(save_path, x)))
            new_name = f"{document_number}.pdf"
            os.rename(os.path.join(save_path, latest_pdf), os.path.join(save_path, new_name))
            update_status("PDF salvo com sucesso. [OK]")
            messagebox.showinfo("Sucesso", f"Certidão salva em: {os.path.join(save_path, new_name)}")
        else:
            update_status("Erro ao salvar o arquivo PDF.")
            messagebox.showerror("Erro", "Falha ao salvar o arquivo PDF.")

        # Fecha o navegador após salvar
        driver.quit()
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro durante a execução: {str(e)}")
        update_status(f"Erro durante a execução: {str(e)}")

def nova_consulta():
    root.destroy()  # Fecha a aplicação atual
    os.execl(sys.executable, sys.executable, *sys.argv)  # Reinicia a aplicação

# Configuração da interface gráfica
root = tk.Tk()
root.title("Automação de Certidão")
root.geometry("400x300")
root.configure(bg='#1c1f2e')  # Tema dark-blue

# Label e Entry para o número do documento
label_document = tk.Label(root, text="Número do Documento:", bg='#1c1f2e', fg='white')
label_document.pack(pady=10)

entry_document = tk.Entry(root, width=30)
entry_document.pack(pady=5)

# Label para status
status_label = tk.Label(root, text="", bg='#1c1f2e', fg='white')
status_label.pack(pady=10)

# Botão para executar a automação
btn_run = tk.Button(root, text="Executar", command=lambda: run_automation(entry_document.get()), bg='#1c6ea4', fg='white')
btn_run.pack(pady=10)

# Botão para nova consulta
btn_nova_consulta = tk.Button(root, text="Nova Consulta", command=nova_consulta, bg='#1c6ea4', fg='white')
btn_nova_consulta.pack(pady=10)

root.mainloop()
