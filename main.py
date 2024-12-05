import os
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from collections import defaultdict
import statistics
from openpyxl import Workbook

BASE_URL = "https://www.mtggoldfish.com"
ARQUETYPE_URL = f"{BASE_URL}/archetype/niv-mizzet-parun/decks"

def setup_driver():
    """Configura o driver do Selenium."""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")
    options.add_argument("--disable-gpu")
    driver = webdriver.Chrome(options=options)
    return driver

def extract_deck_links(driver):
    """Extrai os links dos decks na página atual."""
    rows = driver.find_elements(By.CSS_SELECTOR, "body > main > div > table > tbody > tr")
    links = []
    for row in rows:
        try:
            link_element = row.find_element(By.CSS_SELECTOR, "td:nth-child(2) > a")
            links.append(link_element.get_attribute("href"))
        except:
            continue
    return links

def extract_textarea_content(driver, deck_url):
    """Extrai o conteúdo da textarea dentro de um deck."""
    driver.get(deck_url)
    try:
        download_button = driver.find_element(By.CSS_SELECTOR, 
            "body > main > div > div.deck-display > div.deck-display-left-contents > div > div > div.deck-container-main > div.deck-container-sidebar > div > div.deck-tools-desktop > div > a:nth-child(9)"
        )
        txt_url = download_button.get_attribute("href")
        driver.get(txt_url)
        textarea = driver.find_element(By.CSS_SELECTOR, "body > main > div > textarea")
        return textarea.text.strip()
    except Exception as e:
        print(f"Erro ao processar deck: {e}")
        return None

def save_to_files(content, output_dir="decks"):
    """Salva o conteúdo separado em Commander, Deck e Sideboard."""
    os.makedirs(output_dir, exist_ok=True)

    commander = []
    companion = []
    deck = []
    sideboard = []

    # Dividindo o texto por linhas
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
            current_section.append(line)

    # Salvando cada seção
    if commander:
        with open(os.path.join(output_dir, "commander.txt"), "a", encoding="utf-8") as file:
            file.write("\n".join(commander) + "\n\n")
    if companion:
        with open(os.path.join(output_dir, "companion.txt"), "a", encoding="utf-8") as file:
            file.write("\n".join(companion) + "\n\n")
    if deck:
        with open(os.path.join(output_dir, "deck.txt"), "a", encoding="utf-8") as file:
            file.write("\n".join(deck) + "\n\n")
    if sideboard:
        with open(os.path.join(output_dir, "sideboard.txt"), "a", encoding="utf-8") as file:
            file.write("\n".join(sideboard) + "\n\n")

def generate_card_averages(input_dir="decks"):
    """Gera as médias, medianas e modas de cada carta nos arquivos deck.txt, commander.txt, companion.txt e sideboard.txt."""
    for file_name in ["deck.txt", "commander.txt", "companion.txt", "sideboard.txt"]:
        file_path = os.path.join(input_dir, file_name)
        if not os.path.exists(file_path):
            print(f"Arquivo {file_name} não encontrado, pulando...")
            continue

        # Determina a fonte com base no nome do arquivo
        if "deck" in file_name:
            source = "Deck"
        elif "commander" in file_name:
            source = "Commander"
        elif "companion" in file_name:
            source = "Companion"
        elif "sideboard" in file_name:
            source = "Sideboard"

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

        # Cria a planilha do Excel
        wb = Workbook()
        ws = wb.active
        ws.title = f"Card Averages - {source}"

        # Cabeçalho
        ws.append(["Card Name", "Total Count", "Average", "Median", "Mode"])

        # Dados
        sorted_cards = sorted(card_counts.items(), key=lambda x: x[1], reverse=True)
        for card, total_count in sorted_cards:
            counts = card_occurences[card]
            avg = sum(counts) / len(counts) if counts else 0
            median = statistics.median(counts) if counts else 0
            try:
                mode = statistics.mode(counts) if counts else 0
            except statistics.StatisticsError:
                mode = "N/A"  # Quando não há moda

            ws.append([card, total_count, round(avg, 2), median, mode])

        # Salvar arquivo Excel
        file_name_excel = f"card_averages_{source.lower()}.xlsx"
        file_path_excel = os.path.join(input_dir, file_name_excel)
        wb.save(file_path_excel)

        print(f"Planilha '{file_name_excel}' salva com sucesso.")

    """Gera as médias, medianas e modas de cada carta nos arquivos deck.txt, commander.txt e sideboard.txt."""
    for file_name in ["deck.txt", "commander.txt", "sideboard.txt"]:
        file_path = os.path.join(input_dir, file_name)
        if not os.path.exists(file_path):
            print(f"Arquivo {file_name} não encontrado, pulando...")
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

        # Cria a planilha do Excel
        wb = Workbook()
        ws = wb.active
        ws.title = f"Card Averages - {source}"

        # Cabeçalho
        ws.append(["Card Name", "Total Count", "Average", "Median", "Mode"])

        # Dados
        sorted_cards = sorted(card_counts.items(), key=lambda x: x[1], reverse=True)
        for card, total_count in sorted_cards:
            counts = card_occurences[card]
            avg = sum(counts) / len(counts) if counts else 0
            median = statistics.median(counts) if counts else 0
            try:
                mode = statistics.mode(counts) if counts else 0
            except statistics.StatisticsError:
                mode = 0

            ws.append([card, total_count, round(avg, 2), median, mode])

        # Salvar arquivo Excel
        file_name_excel = f"card_averages_{source.lower()}.xlsx"
        file_path_excel = os.path.join(input_dir, file_name_excel)
        wb.save(file_path_excel)

        print(f"Planilha '{file_name_excel}' salva com sucesso.")

def main():
    driver = setup_driver()
    page_number = 1
    try:
        while page_number <= 25:  # Limita o loop a 25 páginas
            print(f"Buscando página {page_number}...")
            url = ARQUETYPE_URL if page_number == 1 else f"{ARQUETYPE_URL}?page={page_number}"
            driver.get(url)
            time.sleep(2)
            deck_links = extract_deck_links(driver)
            if not deck_links:
                break
            for deck_url in deck_links:
                print(f"Processando deck: {deck_url}")
                content = extract_textarea_content(driver, deck_url)
                if content:
                    save_to_files(content)
                time.sleep(2)
            page_number += 1
            time.sleep(2)

        generate_card_averages()
    finally:
        driver.quit()

if __name__ == "__main__":
    main()
