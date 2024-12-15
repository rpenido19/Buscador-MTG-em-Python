import os
import zipfile
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
from collections import defaultdict
import statistics
from openpyxl import Workbook
from urllib.parse import urlparse
import shutil
import logging

# Configurações básicas
BASE_URL = "https://www.mtggoldfish.com"
ARQUETYPE_URL = f"{BASE_URL}/archetype/pauper-grixis-affinity/decks"

# Configuração de logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

def setup_driver():
    """Configura o driver do Selenium."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/117.0.0.0 Safari/537.36")
    driver = webdriver.Chrome(options=options)
    return driver

def extract_archetype_name(url):
    """Extrai o nome do arquétipo da URL."""
    path = urlparse(url).path
    if "archetype" in path:
        parts = path.strip("/").split("/")
        if "archetype" in parts:
            return parts[parts.index("archetype") + 1]
    return "unknown-archetype"

def clear_directory(directory="decks"):
    """Limpa todos os arquivos e subdiretórios da pasta especificada."""
    if os.path.exists(directory):
        shutil.rmtree(directory)
    os.makedirs(directory)
    logging.info(f"A pasta '{directory}' foi limpa.")

def extract_deck_links(driver):
    """Extrai os links dos decks na página atual."""
    rows = driver.find_elements(By.CSS_SELECTOR, "body > main > div > table > tbody > tr")
    links = []
    for row in rows:
        try:
            link_element = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) > a")
            links.append(link_element.get_attribute("href"))
        except Exception as e:
            logging.warning(f"Erro ao extrair link de deck: {e}")
            continue
    return links

def extract_textarea_content(driver, deck_url):
    """Extrai o conteúdo da textarea dentro de um deck."""
    driver.get(deck_url)
    try:
        download_button = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "div.deck-tools-desktop > div > a:nth-child(9)")
            )
        )
        txt_url = download_button.get_attribute("href")
        driver.get(txt_url)
        textarea = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "body > main > div > textarea"))
        )
        return textarea.text.strip()
    except Exception as e:
        logging.error(f"Erro ao processar deck em {deck_url}: {e}")
        return None

def save_to_files(content, output_dir="decks"):
    """Salva o conteúdo separado em Commander, Deck e Sideboard."""
    os.makedirs(output_dir, exist_ok=True)

    commander = []
    companion = []
    deck = []
    sideboard = []

    lines = content.split("\n")
    current_section = None

    for line in lines:
        line = line.strip()
        if not line:
            continue
        if line.startswith("Commander"):
            current_section = commander
        elif line.startswith("Companion"):
            current_section = companion
        elif line.startswith("Deck"):
            current_section = deck
        elif line.startswith("Sideboard"):
            current_section = sideboard
        elif current_section is not None:
            if line not in current_section:  # Evita duplicação
                current_section.append(line)

    def write_section_to_file(section, file_name):
        if section:
            with open(os.path.join(output_dir, file_name), "a", encoding="utf-8") as file:
                file.write("\n".join(section) + "\n\n")

    write_section_to_file(commander, "commander.txt")
    write_section_to_file(companion, "companion.txt")
    write_section_to_file(deck, "deck.txt")
    write_section_to_file(sideboard, "sideboard.txt")

def generate_card_averages(input_dir="decks"):
    """Gera as médias, medianas e modas de cada carta nos arquivos."""
    for file_name in ["deck.txt", "commander.txt", "sideboard.txt"]:
        file_path = os.path.join(input_dir, file_name)
        if not os.path.exists(file_path):
            logging.warning(f"Arquivo {file_name} não encontrado, pulando...")
            continue

        source = "Deck" if "deck" in file_name else "Commander" if "commander" in file_name else "Sideboard"

        card_counts = defaultdict(int)
        card_occurences = defaultdict(list)

        with open(file_path, "r", encoding="utf-8") as file:
            content = file.read()
            for line in content.split("\n"):
                if line.strip() and not line.startswith("#"):
                    try:
                        count, card = line.strip().split(" ", 1)
                        count = int(count)
                        card_counts[card] += count
                        card_occurences[card].append(count)
                    except ValueError:
                        continue

        wb = Workbook()
        ws = wb.active
        ws.title = f"Card Averages - {source}"

        ws.append(["Card Name", "Total Count", "Average", "Median", "Mode"])

        sorted_cards = sorted(card_counts.items(), key=lambda x: x[1], reverse=True)
        for card, total_count in sorted_cards:
            counts = card_occurences[card]
            avg = sum(counts) / len(counts) if counts else 0
            median = statistics.median(counts) if counts else 0
            try:
                mode = statistics.mode(counts) if counts else 0
            except statistics.StatisticsError:
                mode = "N/A"

            ws.append([card, total_count, round(avg, 2), median, mode])

        file_name_excel = f"card_averages_{source.lower()}.xlsx"
        file_path_excel = os.path.join(input_dir, file_name_excel)
        wb.save(file_path_excel)

        logging.info(f"Planilha '{file_name_excel}' salva com sucesso.")

def compress_files(archetype_name, input_dir="decks", output_dir="compressed"):
    """Compacta os arquivos gerados em um .zip."""
    os.makedirs(output_dir, exist_ok=True)
    zip_name = os.path.join(output_dir, f"{archetype_name}.zip")
    
    with zipfile.ZipFile(zip_name, "w") as zipf:
        for root, _, files in os.walk(input_dir):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, start=input_dir)
                zipf.write(file_path, arcname)
    
    logging.info(f"Arquivos compactados em '{zip_name}'.")

def process_pages(driver, base_url):
    """Processa todas as páginas de decks."""
    page_number = 1
    while page_number <= 25:  # Limita o loop a 25 páginas
        url = base_url if page_number == 1 else f"{base_url}?page={page_number}"
        logging.info(f"Buscando página {page_number}: {url}")
        driver.get(url)
        time.sleep(2)
        deck_links = extract_deck_links(driver)
        if not deck_links:
            logging.info(f"Sem mais decks encontrados na página {page_number}. Encerrando.")
            break
        for deck_url in deck_links:
            process_deck(driver, deck_url)
        page_number += 1
        time.sleep(2)

def process_deck(driver, deck_url):
    """Processa um único deck."""
    try:
        logging.info(f"Processando deck: {deck_url}")
        content = extract_textarea_content(driver, deck_url)
        if content:
            save_to_files(content)
    except Exception as e:
        logging.error(f"Erro ao processar deck {deck_url}: {e}")

def main():
    archetype_name = extract_archetype_name(ARQUETYPE_URL)
    clear_directory("decks")
    driver = setup_driver()
    try:
        process_pages(driver, ARQUETYPE_URL)
        generate_card_averages()
        compress_files(archetype_name)
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
