import requests
from bs4 import BeautifulSoup
import json
from urllib.parse import urljoin
import time
import pandas as pd
import os

def get_company_details(url, headers):
    """Get company details from dl-horizontal"""
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'lxml')
        dl_horizontal = soup.find(class_='dl-horizontal')
        
        if not dl_horizontal:
            print(f"No dl-horizontal found on {url}")
            return {}
            
        company_info = {}
        
        # Находим нужные поля по их названиям
        dt_elements = dl_horizontal.find_all('dt')
        dd_elements = dl_horizontal.find_all('dd')
        
        for dt, dd in zip(dt_elements, dd_elements):
            key = dt.get_text(strip=True)
            
            # Сопоставляем поля с нужными нам данными
            if key == "Сайт:":
                site_link = dd.find('a')
                company_info['Сайт'] = site_link.get_text(strip=True) if site_link else ''
            elif key == "Телефон:":
                company_info['Телефон'] = dd.get_text(strip=True)
            elif key == "E-mail:":
                email_link = dd.find('a')
                company_info['E-mail'] = email_link.get_text(strip=True) if email_link else ''
            
                # Находим все span элементы с классом label label-primary
            spans = dd.find_all('span', class_='label label-primary')
            # Извлекаем только текст из каждого span
            categories = [span.get_text(strip=True) for span in spans]
            company_info['Рубрика'] = '; '.join(categories)
                
        return company_info
        
    except requests.RequestException as e:
        print(f"Error fetching company details from {url}: {e}")
        return {}

def get_table_links(url, headers):
    """Get links from the fresh-table on the /list page"""
    try:
        # Add /list to the URL if it doesn't end with it
        if not url.endswith('/list'):
            url = f"{url}/list"
            
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'lxml')
        
        # Find the fresh-table first
        fresh_table = soup.find(id='fresh-table')
        if not fresh_table:
            print(f"No fresh-table found on {url}")
            return []
            
        # Find all tr elements in the table
        table_rows = fresh_table.find_all('tr')
        
        results = []
        # for row in table_rows:

        # Limit to first 10 companies for testing
        for row in table_rows[:30]:
            # Get the first td element in the row
            first_td = row.find('td')
            if first_td:
                # Find the link in the td
                link = first_td.find('a')
                if link:
                    href = link.get('href', '')
                    text = link.get_text(strip=True)
                    full_url = urljoin(url, href) if href else ''
                    if full_url:
                        # Get company details immediately
                        print(f"Fetching details for company: {text}")
                        company_details = get_company_details(full_url, headers)
                        
                        company_data = {
                            'text': text,
                            'url': full_url,
                            'details': company_details
                        }
                        results.append(company_data)
                        time.sleep(0.5)  # Small delay between requests
        
        return results
    except requests.RequestException as e:
        print(f"Error fetching table links from {url}: {e}")
        return []

def get_exhibition_links(url, headers):
    """Get exhibition links from the main page"""
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        
        soup = BeautifulSoup(response.text, 'lxml')
        links = soup.find_all(class_='list-group-item list-group-item-action')
        
        results = []
        events_found = 0
        
        for link in links:
            href = link.get('href', '')
            text = link.get_text(strip=True)
            
            # Check if the text contains 2025
            if not any(year in text for year in ['2025']):
                continue
                
            full_url = urljoin(url, href) if href else ''
            if full_url:
                results.append({
                    'text': text,
                    'url': full_url
                })
                events_found += 1
                # Take only first two exhibitions for testing
                if events_found >= 3:
                    break
        
        return results
    except requests.RequestException as e:
        print(f"Error fetching exhibition links: {e}")
        return []

def save_to_excel(exhibition_name, companies_data):
    """Save companies data to Excel file"""
    # Create a list to store company information
    excel_data = []
    
    for company in companies_data:
        details = company.get('details', {})
        excel_data.append({
            'название': company['text'],
            'сайт': details.get('Сайт', ''),
            'телефон': details.get('Телефон', ''),
            'E-mail': details.get('E-mail', ''),
            'рубрика': details.get('Рубрика', '')
        })
    
    # Create DataFrame
    df = pd.DataFrame(excel_data)
    
    # Create 'excel' directory if it doesn't exist
    os.makedirs('excel', exist_ok=True)
    
    # Save to Excel
    filename = os.path.join('excel', f'участники выставки {exhibition_name}.xlsx')
    
    # Create Excel writer object
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        # Write DataFrame to Excel
        df.to_excel(writer, index=False, sheet_name='Участники')
        
        # Get the worksheet
        worksheet = writer.sheets['Участники']
        
        # Adjust column widths based on content
        for idx, col in enumerate(df.columns):
            # Find the maximum length in the column
            max_length = max(
                df[col].astype(str).apply(len).max(),  # max length of values
                len(str(col))  # length of column name
            ) + 2  # adding a little extra space
            
            # Convert character count to column width (approximate conversion)
            column_width = max_length * 1.2
            
            # Set column width
            worksheet.column_dimensions[chr(65 + idx)].width = column_width
    
    print(f"Excel file saved: {filename}")

def parse_expocentr():
    # URL of the website
    url = 'https://icatalog.expocentr.ru/ru'
    
    # Headers to mimic a browser request
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
    }
    
    try:
        print("Step 1: Getting exhibition links...")
        exhibition_links = get_exhibition_links(url, headers)
        
        if not exhibition_links:
            print("No exhibition links found!")
            return
            
        print(f"Found {len(exhibition_links)} exhibitions for 2025 (limited to 2 for testing)")
        
        # Create the final result structure
        result = []
        
        # Process each exhibition link
        for exhibition in exhibition_links:
            exhibition_data = {
                'exhibition_name': exhibition['text'],
                'exhibition_url': exhibition['url'],
                'companies': []
            }
            
            # Get company links from the exhibition page
            print(f"\nFetching companies from: {exhibition['text']}")
            company_links = get_table_links(exhibition['url'], headers)
            exhibition_data['companies'] = company_links
            result.append(exhibition_data)
            
            # Save to Excel
            save_to_excel(exhibition['text'], company_links)
            
            print(f"Found {len(company_links)} companies (limited to 10 for testing)")
            
            # Add a small delay to be polite to the server
            time.sleep(1)
        
        # Save to JSON file
        with open('expo_links.json', 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
            
        print(f"\nSuccessfully parsed {len(result)} exhibition for testing")
        total_companies = sum(len(expo['companies']) for expo in result)
        print(f"Total companies collected: {total_companies}")
        
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    parse_expocentr() 