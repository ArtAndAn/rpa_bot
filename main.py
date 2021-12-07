from datetime import timedelta
from random import randint
from time import sleep

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF

file_sys = FileSystem()
browser = Selenium(timeout=timedelta(10))
browser.set_download_directory(directory=file_sys.absolute_path(path='outer'))
excel_manager = Files()
pdf = PDF()


def get_agencies_amount():
    """
    Function for creating dict with agencies as keys and spending as a value
    """
    browser.open_available_browser(url="https://itdashboard.gov/")
    browser.wait_until_page_contains_element(locator='//*[@id="agency-tiles-2-widget"]/div/div/div/div/div/div/div/a')
    agencies = browser.find_elements(locator='//*[@id="agency-tiles-2-widget"]/div/div/div/div/div/div/div[1]')

    all_agencies_data = {}
    for agency in agencies:
        agency_data = agency.text.split('\n')
        all_agencies_data[agency_data[0]] = agency_data[2]
    return all_agencies_data


def fill_up_agencies():
    """
    Function for creating Excel workbook, renaming default sheet name to 'Agencies' and
    filling up this sheet with Agencies data
    """
    file = excel_manager.create_workbook(path='outer/data.xlsx')
    file.rename_worksheet('Agencies', 'Sheet')
    file.set_cell_value(row=1, column='A', value='Agency')
    file.set_cell_value(row=1, column='B', value='Spending')

    agencies_data = get_agencies_amount()
    row = 2
    for agency, spending in agencies_data.items():
        file.set_cell_value(row=row, column='A', value=agency)
        file.set_cell_value(row=row, column='B', value=spending)
        row += 1
    file.save()


def random_agency_details():
    """
    Function for choosing a random Agency, creating a new sheet with this Agency name and
    filling it with investment table data
    Downloading PDF file for all investments that have a link for investment details page
    """
    file = excel_manager.open_workbook('outer/data.xlsx')
    all_agencies_data = file.read_worksheet(name='Agencies')
    random_agency = all_agencies_data[randint(1, len(all_agencies_data) - 1)]
    agency_name = random_agency['A']
    file.create_worksheet(name=agency_name)

    browser.click_link(locator=f'partial link:{agency_name}')
    browser.wait_until_page_contains_element(locator='//*[@id="investments-table-object_length"]/label/select')
    browser.select_from_list_by_value('//*[@id="investments-table-object_length"]/label/select', '-1')
    browser.wait_until_page_does_not_contain_element(locator='class:loading')

    investments_table_rows = browser.find_elements(locator='//*[@id="investments-table-object"]/tbody/tr')
    investments_uii_links = browser.find_elements(locator='//*[@id="investments-table-object"]/tbody/tr/td[1]/a')
    investments_uii_urls = {browser.get_element_attribute(locator=link, attribute='href') for link in
                            investments_uii_links}

    for url in investments_uii_urls:
        browser.open_available_browser(url=url)
        browser.wait_until_page_contains_element(locator='//*[@id="business-case-pdf"]/a')
        browser.click_link(locator='//*[@id="business-case-pdf"]/a')
        browser.wait_until_page_does_not_contain_element(locator='//*[@id="business-case-pdf"]/span')
        sleep(1.5)
        browser.close_browser()

    table_data = []
    for row in investments_table_rows:
        cells_data = browser.find_elements(locator='tag:td', parent=row)
        row_data = {'UII': cells_data[0].text,
                    'Bureau': cells_data[1].text,
                    'Investment title': cells_data[2].text,
                    'Total FY2021 Spending($M)': cells_data[3].text,
                    'Type': cells_data[4].text,
                    'CIO Rating': cells_data[5].text,
                    '# of Projects': cells_data[6].text}
        table_data.append(row_data)

    file.append_worksheet(name=agency_name, content=table_data, header=True)
    file.save()


def compare_pdfs():
    """
    Function for comparing each pdf Investment Name and UII number with
    data in Excel detailed agency data worksheet
    If function locate any data error - it will raise Assertion Error
    """
    all_files = file_sys.list_files_in_directory(path='outer')
    all_pdfs = filter(lambda x: file_sys.get_file_extension(path=x) == '.pdf', all_files)
    for file in all_pdfs:
        pdf.open_pdf(source_path=file)
        first_page_data = pdf.get_text_from_pdf(source_path=file, pages=1, details=True, trim=False)[1]

        investment_name_str = '1. Name of this Investment: '
        uii_number_str = '2. Unique Investment Identifier (UII): '
        for textbox in first_page_data:
            if investment_name_str in textbox.text:
                investment_name = textbox.text.replace(investment_name_str, '')[:-1]
            elif uii_number_str in textbox.text:
                uii_number = textbox.text.replace(uii_number_str, '')[:-1]

        file = excel_manager.open_workbook('outer/data.xlsx')
        agency_name = excel_manager.list_worksheets()[1]
        sheet_data = file.read_worksheet(name=agency_name, header=True)

        passed_validation = False
        for row in sheet_data:
            if row['Investment title'] == investment_name and row['UII'] == uii_number:
                passed_validation = True
        if not passed_validation:
            raise AssertionError('PDF and Excel files assertion error')


def main():
    try:
        fill_up_agencies()
        random_agency_details()
        compare_pdfs()
    finally:
        browser.close_all_browsers()
        excel_manager.close_workbook()
        pdf.close_all_pdfs()


if __name__ == "__main__":
    main()
