import pandas as pd
import requests
from bs4 import BeautifulSoup
from tqdm import tqdm
import re


# Function to clean card names by removing punctuation
def clean_card_name(name):
    return re.sub(r'[^\w\s]', '', name).strip().lower()


# Read the Excel file
file_path = r'C:\Users\GrillerDriller\Downloads\Updated_MTG_Collection1 - Copy.xlsx'
df = pd.read_excel(file_path)

# Clean column names
df.columns = df.columns.str.strip()

# Initialize a DataFrame to store debug messages
debug_df = pd.DataFrame(columns=['Row', 'Debug Message'])


# Function to search eBay and collect prices
def search_ebay(card_name, series, collector_number, foil_type, art_type, row_index):
    def perform_search(query):
        url = f"https://www.ebay.com.au/sch/i.html?_from=R40&_nkw={query.replace(' ', '+')}&_sacat=0&rt=nc&LH_Sold=1&LH_Complete=1"
        response = requests.get(url)
        soup = BeautifulSoup(response.text, 'html.parser')
        return soup, url

    def extract_prices(soup, foil_type_str, art_type_str):
        prices = []
        selected_items = []
        for item in soup.select('.s-item'):
            title = item.select_one('.s-item__title')
            if title:
                title_text = title.get_text().lower()
                item_url = item.select_one('.s-item__link')['href']
                clean_title = clean_card_name(title_text)

                debug_message = f"Row {row_index}: Checking title: {title_text}"
                debug_df.loc[len(debug_df)] = [row_index, debug_message]

                # Skip if looking for foil and title mentions "regular", "non-foil", or "non foil"
                if foil_type_str != 'normal' and any(
                        term in title_text for term in ['regular', 'non-foil', 'non foil']):
                    continue
                # Skip if looking for normal and title mentions "foil"
                if foil_type_str == 'normal' and 'foil' in title_text and 'non-foil' not in title_text and 'non foil' not in title_text:
                    continue
                # Skip if looking for normal and title mentions "extended" or "borderless"
                if art_type_str == 'normal' and ('extended' in title_text or 'borderless' in title_text):
                    continue

                # Ensure card name is present in the title
                if clean_card_name(card_name) not in clean_title:
                    continue

                # Ensure foil type matches
                if foil_type_str != 'normal' and 'foil' not in title_text:
                    continue

                # Ensure art type matches
                if art_type_str != 'normal' and art_type_str not in title_text and 'borderless' not in title_text and 'extended' not in title_text:
                    continue

                price = item.select_one('.s-item__price').get_text()
                prices.append((price, item_url))
                selected_items.append(title_text)
        return prices, selected_items

    # Convert variables to strings for safe comparison
    foil_type_str = str(foil_type).lower()
    art_type_str = str(art_type).lower()

    # Create search query without "normal" foil type and art type
    search_terms = [card_name]
    if foil_type_str != 'normal':
        search_terms.append(foil_type)
    if art_type_str != 'normal':
        search_terms.append(art_type)
    search_query = ' '.join(search_terms).strip()

    soup, url = perform_search(search_query)
    prices, selected_items = extract_prices(soup, foil_type_str, art_type_str)

    # If no exact match is found, perform additional searches
    if not selected_items:
        fallback_queries = [
            f"{card_name} {collector_number} {foil_type} {art_type}".strip(),
            f"{collector_number} {series} {foil_type} {art_type}".strip(),
            f"{card_name} {series} {foil_type} {art_type}".strip()
        ]
        for query in fallback_queries:
            soup, url = perform_search(query)
            prices, selected_items = extract_prices(soup, foil_type_str, art_type_str)
            if selected_items:
                break

    # Log the selected items
    if selected_items:
        for selected_item in selected_items:
            debug_df.loc[len(debug_df)] = [row_index, f"Selected item: {selected_item}"]

    return prices[:5], url


# Function to process all rows
def process_all_rows():
    rows_to_process = range(len(df))

    # Collect prices for the specified range of rows with a progress bar
    for index in tqdm(rows_to_process, total=len(rows_to_process), desc="Processing cards"):
        row = df.iloc[index]
        series = row['Series']
        collector_number = row['Collector number']
        foil_type = row['Foil type']
        card_name = row['Name']
        art_type = row['Art type']

        # Skip rows with missing or empty names
        if pd.isna(card_name) or card_name.strip() == "":
            debug_df.loc[len(debug_df)] = [index, "Skipped due to missing or empty card name"]
            continue

        try:
            prices, url = search_ebay(card_name, series, collector_number, foil_type, art_type, index)
        except Exception as e:
            debug_df.loc[len(debug_df)] = [index, f"Error: {str(e)}"]
            continue

        if prices:
            for i, (price, item_url) in enumerate(prices):
                df.at[index, f'Price {i + 1}'] = price
                df.at[index, f'Price URL {i + 1}'] = item_url
        else:
            debug_df.loc[len(debug_df)] = [index, "No prices found"]

        df.at[index, 'Search URL'] = url  # Plain text URL

    # Save the updated DataFrame back to Excel
    output_file_path = r'C:\Users\GrillerDriller\Downloads\Updated_MTG_Collection_with_Prices_Troubleshoot9.xlsx'
    df.to_excel(output_file_path, index=False, engine='openpyxl')

    # Save the debug messages to a separate Excel file
    debug_output_file_path = r'C:\Users\GrillerDriller\Downloads\Debug_Messages.xlsx'
    debug_df.to_excel(debug_output_file_path, index=False, engine='openpyxl')

    print("Finished processing the specified entries and saved debug messages.")



# Process the specified range of rows
process_all_rows()
