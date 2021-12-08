from datetime import timedelta
from time import sleep

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF
from RPA.Robocorp.WorkItems import WorkItems
from robot.libraries.BuiltIn import BuiltIn

browser = Selenium(timeout=timedelta(10))
browser.set_download_directory(directory=FileSystem().absolute_path(path='outer'))
excel_manager = Files()
pdf = PDF()
logger = BuiltIn()
work_items = WorkItems()
work_items.get_input_work_item()


def fill_up_agencies():
    """
    Function for creating Excel workbook, renaming default sheet name to 'Agencies' and
    filling up this sheet with Agencies data
    """
    logger.log(message='Started fill up agencies function', console=True)
    file = excel_manager.create_workbook(path='outer/data.xlsx')
    file.rename_worksheet('Agencies', 'Sheet')
    file.set_cell_value(row=1, column='A', value='Agency')
    file.set_cell_value(row=1, column='B', value='Spending')

    site_url = work_items.get_work_item_variable(name='SITE_URL', default='https://itdashboard.gov/')
    browser.open_available_browser(url=site_url)
    browser.click_link(locator='link:DIVE IN')
    browser.wait_until_page_contains_element(locator='id:agency-tiles-widget >> class:col-sm-12')
    agencies = browser.find_elements(locator='id:agency-tiles-widget >> class:col-sm-12')
    logger.log(message='Collected all agencies data', console=True)

    row = 2
    for agency in agencies:
        agency_data = agency.text.split('\n')
        file.set_cell_value(row, column='A', value=agency_data[0])
        file.set_cell_value(row, column='B', value=agency_data[2])
        row += 1
    file.save()
    logger.log(message='Finished fill up agencies function', console=True)


def detailed_agency_investments():
    """
    Function for choosing Agency from environment variable, creating a new sheet with this Agency name and
    filling it with investment table data
    Downloading PDF file if investment has a link and comparing its data with table data
    """
    agency_name = work_items.get_work_item_variable(name='AGENCY_NAME', default='U.S. Army Corps of Engineers')
    logger.log(message=f'Started detailed agency investments function for -- {agency_name} -- agency', console=True)

    file = excel_manager.open_workbook('outer/data.xlsx')
    file.create_worksheet(name=agency_name)

    browser.click_link(locator=f'partial link:{agency_name}')
    browser.wait_until_page_contains_element(locator='name:investments-table-object_length')
    browser.select_from_list_by_value('name:investments-table-object_length', '-1')
    browser.wait_until_page_does_not_contain_element(locator='class:loading')

    investments_table_rows = browser.find_elements(locator='id:investments-table-object >> tag:tbody >> tag:tr')
    logger.log(message='Collected investments table rows data', console=True)

    for row in investments_table_rows:
        cells_data = browser.find_elements(locator='tag:td', parent=row)
        row_data = {'UII': cells_data[0].text,
                    'Bureau': cells_data[1].text,
                    'Investment title': cells_data[2].text,
                    'Total FY2021 Spending($M)': cells_data[3].text,
                    'Type': cells_data[4].text,
                    'CIO Rating': cells_data[5].text,
                    '# of Projects': cells_data[6].text}
        file.append_worksheet(name=agency_name, content=row_data, header=True)

        logger.log(message=f'Row -- {row_data["Investment title"]} -- data recorded to excel', console=True)

        individual_investment_link = browser.find_elements(locator='tag:a', parent=row)
        if not individual_investment_link:
            logger.log(message=f'Row -- {row_data["Investment title"]} -- has no link', console=True)
            continue

        logger.log(message=f'Row -- {row_data["Investment title"]} -- has a link', console=True)

        detailed_investment_data = browser.get_element_attribute(locator=individual_investment_link[0],
                                                                 attribute='href')
        browser.open_available_browser(url=detailed_investment_data)
        browser.wait_until_page_contains_element(locator='id:business-case-pdf')
        browser.click_link(locator='id:business-case-pdf >> tag:a')
        browser.wait_until_page_does_not_contain_element(locator='id:business-case-pdf >> tag:span')
        sleep(2)
        browser.close_browser()

        logger.log(message=f'Row -- {row_data["Investment title"]} -- PDF file downloaded', console=True)

        investment_title_from_table = row_data['Investment title']
        uii_from_table = row_data['UII']

        pdf.open_pdf(source_path=f'outer/{uii_from_table}.pdf')

        investment_title_search_key = '1. Name of this Investment: '
        investment_title_from_pdf = pdf.find_text(locator=f'regex:{investment_title_search_key}', pagenum=1)[0] \
            .anchor.replace(investment_title_search_key, '')

        uii_search_key = '2. Unique Investment Identifier'
        uii_from_pdf = pdf.find_text(locator=f'regex:{uii_search_key}', pagenum=1)[0] \
            .anchor.replace(f'{uii_search_key} (UII): ', '')

        if investment_title_from_table.split() != investment_title_from_pdf.split():
            logger.log(message=f'Row -- {row_data["Investment title"]} -- investment titles are not equal',
                       level='ERROR', console=True)
        elif uii_from_table.split() != uii_from_pdf.split():
            logger.log(message=f'Row -- {row_data["Investment title"]} -- UII numbers are not equal',
                       level='ERROR', console=True)
        else:
            logger.log(message=f'Row -- {row_data["Investment title"]} -- data is correct', console=True)
        pdf.close_pdf()
    file.save()
    logger.log(message='Finished detailed agency investments function', console=True)


def main():
    try:
        logger.log(message=f'Starting robot with variables -- {work_items.get_work_item_variables()}', console=True)
        fill_up_agencies()
        detailed_agency_investments()
    finally:
        browser.close_all_browsers()
        excel_manager.close_workbook()
        pdf.close_all_pdfs()


if __name__ == "__main__":
    main()
