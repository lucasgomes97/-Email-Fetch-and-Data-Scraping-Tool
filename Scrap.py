import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image, ImageTk
from imap_tools import MailBox, AND
import csv
from datetime import datetime
import threading
import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


# Function to validate date
def validar_data(data_str):
	try:
		datetime.strptime(data_str, "%d/%m/%y")
		return True
	except ValueError:
		return False


# Function to load domains to be excluded from a CSV file
def carregar_dominios_excluir(arquivo_csv):
    try:
        with open(arquivo_csv, 'r', encoding='iso-8859-1') as file:
            reader = csv.reader(file)
            dominios = [row[0] for row in reader]
        return dominios
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo {arquivo_csv} não encontrado.")
        return []


# Load domains to be excluded
dominios_excluir = carregar_dominios_excluir("dominios_excluir.csv")


# Function to fetch emails and scrape data
def buscar_emails():
	# Validate start and end dates
	data_inicio_str = data_inicio_entry.get()
	data_fim_str = data_fim_entry.get()

	if not validar_data(data_inicio_str):
		messagebox.showerror("Erro", "Formato de data de início inválido. Por favor, insira no formato dd/mm/yy.")
		return

	if not validar_data(data_fim_str):
		messagebox.showerror("Erro", "Formato de data de fim inválido. Por favor, insira no formato dd/mm/yy.")
		return

	buscar_button.config(state=tk.DISABLED)
	carregando_label.place(relx=0.5, rely=0.5, anchor="center")
	root.update()

	def buscar_emails_thread():
		try:
			username = 'seu_email'
			senha = "sua_senha"
			server = 'imap.gmail.com'

			# Connect and login to the IMAP server
			with MailBox(server).login(username, senha, 'Leads') as meu_email:
				data_inicio = datetime.strptime(data_inicio_str, "%d/%m/%y").date()
				data_fim = datetime.strptime(data_fim_str, "%d/%m/%y").date()
				criteria = AND(from_='info@rankdone.com', date_gte=data_inicio, date_lt=data_fim)
				lista_emails = meu_email.fetch(criteria)

				dados_emails = []

				for email in lista_emails:
					if email.subject == "Novo cadastro na Rankdone":
						nome = email.text.split("Nome:")[1].split("\n")[0].strip()
						email_corporativo = email.text.split("Email:")[1].split("\n")[0].strip()
						if not any(dominio in email_corporativo for dominio in dominios_excluir):
							origem = "Formulário do site" if "Formulário de Origem" in email.text else "Experimente Grátis"
							data_formatada = email.date.date().strftime("%d/%m/%Y")
							dados_emails.append([nome, email_corporativo, data_formatada, origem])
					elif "Nome:" in email.text and "E-mail Corporativo:" in email.text:
						nome = email.text.split("Nome:")[1].split("\n")[0].strip()
						email_corporativo = email.text.split("E-mail Corporativo:")[1].split("\n")[0].strip()
						if not any(dominio in email_corporativo for dominio in dominios_excluir):
							origem = "Formulário do site" if "Formulário de Origem" in email.text else "Experimente Grátis"
							data_formatada = email.date.date().strftime("%d/%m/%Y")
							dados_emails.append([nome, email_corporativo, data_formatada, origem])

				dados_unicos = list(map(list, set(map(tuple, dados_emails))))

				# Scrape additional data using Selenium
				def scrape_rankdone_data(emails):
					usuario = "seu_email"
					senha = "sua_senha"

					service = Service(ChromeDriverManager().install())
					browser = webdriver.Chrome(service=service)
					browser.get('https://recruiter.rankdone.com/recruitment')
					time.sleep(5)

					# Perform login
					campo_usuario = browser.find_element(By.XPATH, '//*[@id="mat-input-0"]')
					campo_usuario.send_keys(usuario)
					campo_senha = browser.find_element(By.XPATH, '//*[@id="mat-input-1"]')
					campo_senha.send_keys(senha)
					botao_login = browser.find_element(By.XPATH,
					                                   '/html/body/app-root/app-root/app-public/app-login/main/section[2]/article/div/app-form-login/form/footer/button/span[1]')
					botao_login.click()
					time.sleep(12)

					email_data = []

					# Remove colons ':' and spaces from email addresses
					cleaned_emails = [email.replace(':', '').replace(' ', '') for email in emails]

					for email in cleaned_emails:
						campo_pesquisa = browser.find_element(By.XPATH,
						                                      '//*[@id="main"]/div/div[2]/div[1]/div/div/form/div[1]/div[1]/div/input')
						campo_pesquisa.clear()
						campo_pesquisa.send_keys(email)
						campo_pesquisa.send_keys(Keys.RETURN)
						time.sleep(5)

						try:
							email_element = browser.find_element(By.XPATH,
							                                     '//*[@id="contact-1"]/div[2]/div/div/ul[1]/li[2]/span')
							if email_element.text.strip() == email:
								quantidade_vagas_element = browser.find_element(By.XPATH,
								                                                '//*[@id="main"]/div/div[2]/div[1]/div/div/div[2]/app-user-list/div/table/tbody/tr/td[7]/div/strong[1]')
								quantidade_vagas = quantidade_vagas_element.text if quantidade_vagas_element else "N/A"

								convites_enviados_element = browser.find_element(By.XPATH,
								                                                 '//*[@id="main"]/div/div[2]/div[1]/div/div/div[2]/app-user-list/div/table/tbody/tr/td[7]/div/strong[4]')
								convites_enviados = convites_enviados_element.text if convites_enviados_element else "N/A"

								email_data.append([email, quantidade_vagas, convites_enviados])
							else:
								email_data.append([email, "N/A", "N/A"])
						except Exception as e:
							email_data.append([email, "Erro", "Erro"])
							print(f"Erro ao processar o email {email}: {str(e)}")

					browser.quit()
					return email_data

				# Get a list of email addresses to scrape
				email_addresses = [email[1] for email in dados_unicos]
				email_rankdone_data = scrape_rankdone_data(email_addresses)

				# Merge email data with rankdone data
				email_rankdone_dict = {data[0]: data[1:] for data in email_rankdone_data}
				for email_info in dados_unicos:
					email = email_info[1]
					if email in email_rankdone_dict:
						email_info.extend(email_rankdone_dict[email])
					else:
						email_info.extend(["N/A", "N/A"])

				# Save the combined data to a CSV file
				file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
				if file_path:
					with open(file_path, 'w', newline='', encoding='utf-8') as file:
						writer = csv.writer(file)
						writer.writerow(["Nome", "Email", "Data", "Origem", "Quantidade de Vagas", "Convites Enviados"])
						writer.writerows(dados_unicos)
					messagebox.showinfo("Sucesso", f"Os emails foram salvos com sucesso em '{file_path}'.")
				else:
					messagebox.showinfo("Aviso", "Nenhum arquivo foi selecionado para salvar.")

		except Exception as e:
			messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

		buscar_button.config(state=tk.NORMAL)
		carregando_label.place_forget()

	thread = threading.Thread(target=buscar_emails_thread)
	thread.start()


# GUI setup
root = tk.Tk()
root.title("Buscar Emails")

try:
	root.iconbitmap('rank.ico')
except Exception as e:
	print(f"Erro ao carregar ícone: {str(e)}")

try:
	background_image = Image.open("fundo meeting_3.png")
	background_photo = ImageTk.PhotoImage(background_image)
	background_label = tk.Label(root, image=background_photo)
	background_label.image = background_photo
	background_label.place(x=0, y=0, relwidth=1, relheight=1)
	root.geometry(f"{background_image.width}x{background_image.height}")
except Exception as e:
	print(f"Erro ao carregar imagem de fundo: {str(e)}")

style = ttk.Style()
style.configure("Transparent.TEntry", background="white", fieldbackground="white", borderwidth=0)

data_inicio_label = tk.Label(root, text="Data de Início (dd/mm/yy):", bg="white", font=("Arial", 10, "bold"))
data_inicio_label.place(relx=0.1, rely=0.15, relwidth=0.4, relheight=0.05)
data_inicio_entry = ttk.Entry(root, style="Transparent.TEntry", font=("Arial", 10, "bold"))
data_inicio_entry.place(relx=0.5, rely=0.15, relwidth=0.4, relheight=0.05)

data_fim_label = tk.Label(root, text="Data de Fim (dd/mm/yy):", bg="white", font=("Arial", 10, "bold"))
data_fim_label.place(relx=0.1, rely=0.25, relwidth=0.4, relheight=0.05)
data_fim_entry = ttk.Entry(root, style="Transparent.TEntry", font=("Arial", 10, "bold"))
data_fim_entry.place(relx=0.5, rely=0.25, relwidth=0.4, relheight=0.05)

buscar_button = tk.Button(root, text="Buscar Emails", command=buscar_emails, bg="lightblue", font=("Arial", 10, "bold"))
buscar_button.place(relx=0.3, rely=0.35, relwidth=0.4, relheight=0.1)

exemplo_data_inicio_label = tk.Label(root, text="Exemplo: 01/01/22", bg="white", font=("Arial", 8))
exemplo_data_inicio_label.place(relx=0.55, rely=0.21, anchor="center")

exemplo_data_fim_label = tk.Label(root, text="Exemplo: 31/12/22", bg="white", font=("Arial", 8))
exemplo_data_fim_label.place(relx=0.55, rely=0.31, anchor="center")

carregando_label = tk.Label(root, text="Carregando, aguarde...", bg="white", font=("Arial", 10, "bold"))

root.mainloop()



