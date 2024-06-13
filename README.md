# Detailed Documentation for Email Fetch and Data Scraping Tool

## Overview

This Python program provides a GUI to fetch and process emails from an IMAP server, filter them based on specific criteria, and scrape additional data from a web interface using Selenium. The results are saved to a CSV file. The tool includes functionalities for date validation, excluding specific domains, and handling multi-threaded operations for efficiency.

## Table of Contents

1. [Dependencies](#dependencies)
2. [Functions](#functions)
    - [validar_data](#validar_data)
    - [carregar_dominios_excluir](#carregar_dominios_excluir)
    - [buscar_emails](#buscar_emails)
    - [buscar_emails_thread](#buscar_emails_thread)
    - [scrape_rankdone_data](#scrape_rankdone_data)
3. [GUI Setup](#gui-setup)
4. [Execution Flow](#execution-flow)
5. [Error Handling](#error-handling)

## Dependencies

Ensure that the following libraries are installed:
- `tkinter`: For GUI creation.
- `PIL` (Pillow): For image handling in the GUI.
- `imap_tools`: For interacting with the IMAP email server.
- `csv`: For reading from and writing to CSV files.
- `datetime`: For date handling and validation.
- `threading`: For running tasks in parallel.
- `time`: For adding delays.
- `selenium`: For web scraping.
- `webdriver_manager`: For managing Selenium web drivers.

```bash
pip install tkinter pillow imap-tools selenium webdriver-manager
```

## Functions

### validar_data

```python
def validar_data(data_str):
    try:
        datetime.strptime(data_str, "%d/%m/%y")
        return True
    except ValueError:
        return False
```

This function validates if a given date string follows the `dd/mm/yy` format. It returns `True` if valid, otherwise `False`.

### carregar_dominios_excluir

```python
def carregar_dominios_excluir(arquivo_csv):
    try:
        with open(arquivo_csv, 'r', encoding='iso-8859-1') as file:
            reader = csv.reader(file)
            dominios = [row[0] for row in reader]
        return dominios
    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo {arquivo_csv} não encontrado.")
        return []
```

This function reads a CSV file to load domains that should be excluded from processing. It returns a list of domains or shows an error message if the file is not found.

### buscar_emails

```python
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
        # Function logic
        # ...

    thread = threading.Thread(target=buscar_emails_thread)
    thread.start()
```

This function initiates the email fetching process. It validates the input dates and starts a new thread to fetch and process the emails to avoid blocking the GUI.

### buscar_emails_thread

This inner function of `buscar_emails` handles the main logic for connecting to the IMAP server, fetching emails based on the criteria, and processing them. It also integrates Selenium for scraping additional data.

### scrape_rankdone_data

```python
def scrape_rankdone_data(emails):
    # Function logic for scraping data
    # ...
    return email_data
```

This function uses Selenium to log into a web interface and scrape additional data for each email. It returns a list of data for each email.

## GUI Setup

The GUI is set up using `tkinter`, with fields for date input, a button to start the process, and labels for displaying instructions and loading status. It also includes background images and styling.

```python
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
```

## Execution Flow

1. **GUI Initialization**: The main window is created with input fields, labels, and buttons.
2. **User Input**: The user enters the start and end dates.
3. **Email Fetch**: The user clicks the "Buscar Emails" button, triggering the `buscar_emails` function.
4. **Date Validation**: The input dates are validated.
5. **Thread Creation**: A new thread is started to fetch and process emails.
6. **Email Processing**: Emails are fetched from the IMAP server based on the criteria.
7. **Data Scraping**: Additional data is scraped using Selenium.
8. **Data Saving**: The processed data is saved to a CSV file selected by the user.

## Error Handling

- **File Not Found**: Shows an error message if the domains CSV file is not found.
- **Invalid Dates**: Shows an error message if the input dates are invalid.
- **Email Processing**: Catches and logs exceptions during email processing and scraping, and continues processing the next email.
- **UI Feedback**: Disables the button and shows a loading message during processing to prevent multiple triggers and inform the user.
