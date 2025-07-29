from selenium import webdriver as wd
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
import time
import random
import pandas as pd
import os
from tkinter import *
from selenium.webdriver.chrome.options import Options
import tkinter as tk
from PIL import ImageTk, Image
import threading
import requests
from bs4 import BeautifulSoup
import re;

def clearText():
    text_1.delete('1.0', END)
    entry_1.delete('0', END)
    entry_2.delete('0', END)
    entry_3.delete('0', END)

def Order():
    text_1.delete('1.0', END)
    text_1.insert(tk.END, 'Scraping by Order Started...Please wait')
    df = pd.read_excel('datadotgov_main.xls')
    charity_name_list = df['Charity_Legal_Name']
    start_no = entry_1.get()
    start_no = int(start_no)
    end_no = entry_2.get()
    end_no = int(end_no)
    print(start_no)
    print(end_no)
    df = pd.read_excel('datadotgov_main.xls')
    charity_name_list = df['Charity_Legal_Name']
    charity_number = start_no-1
    for each_name in charity_name_list[start_no-2:end_no-1]:
        try:
            updated_charity_name = ''
            charity_number = charity_number + 1
            parent_dir = os.getcwd()
            charity_name = each_name
            for k in charity_name.split("\n"):
                updated_charity_name += (re.sub(r"[^a-zA-Z0-9]+", ' ', k))

            folder_name = str(charity_number) + '_' + updated_charity_name
            directory_one = 'Financial_reporting'
            directory_two = 'Documents'
            directory_three = 'Annual information statement'
            first_path = os.path.join(parent_dir, folder_name)
            os.mkdir(first_path)
            print(first_path)
            second_parent_dir = os.getcwd() + '\\%s\\' % folder_name
            path_second = os.path.join(second_parent_dir, directory_one)
            path_third = os.path.join(second_parent_dir, directory_two)
            path_fourth = os.path.join(second_parent_dir, directory_three)
            os.mkdir(path_second)
            os.mkdir(path_third)
            os.mkdir(path_fourth)
            chrome_options = Options()
            chrome_options.add_experimental_option('prefs', {
                "download.default_directory": path_second,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            }
                                                   )
            chrome_options_1 = Options()
            chrome_options_1.add_experimental_option('prefs', {
                "download.default_directory": path_third,
                "download.prompt_for_download": False,
                "download.directory_upgrade": True,
                "plugins.always_open_pdf_externally": True
            }
                                                     )


            # annual-reporting
            driver = wd.Chrome(executable_path='./chromedriver', options=chrome_options)
            driver.get(link)
            search_item = each_name
            search_input_item = driver.find_element(by=By.XPATH,
                                                    value='//*[@id="edit-name-abn"]/div/input')
            search_input_item.send_keys(search_item)
            time.sleep(random.randint(2, 3))
            driver.find_element_by_xpath('//*[@id="edit-submit-solr-charities"]').click()
            time.sleep(random.randint(3, 4))
            detail_link = driver.find_element_by_xpath(
                '//*[@id="block-views-solr-charities-index"]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
            detail_link = detail_link.get_attribute('href')
            print(detail_link)
            financila_link = detail_link + '#financials-documents'
            driver.get(financila_link)
            time.sleep(random.randint(4, 5))
            button_list = driver.find_elements_by_xpath(
                '//*[@id="financials-documents"]/div/div[1]/div/div/div/div/div/div/table/tbody/tr/td/a')
            for each_button in button_list:
                href = each_button.get_attribute('href')
                href_list = str(href).split('.')
                filter_href = href_list[-1]
                if filter_href == 'pdf' or filter_href == 'doc' or filter_href == 'xls':
                    driver.get(href)
            for each_button in button_list:
                href = each_button.get_attribute('href')
                href_list = str(href).split('.')
                filter_href = href_list[-1]
                if filter_href != 'pdf' and filter_href != 'doc' and filter_href != 'xls':
                    driver = wd.Chrome(executable_path='./chromedriver')
                    driver.get(href)

                    statement_list = []
                    statement_result_list = []
                    try:
                        title = driver.find_element_by_xpath(
                            '//*[@id="block-system-main"]/div/div/div/div/section/div[1]/div[1]')
                        title = str(title.text)
                        title = title.replace('Annual Information Statement ', '')
                        statement_one = driver.find_elements_by_xpath(
                            '//*[@id="block-system-main"]/div/div/div/div/section/div/div/div[1]')
                        statement_result__one = driver.find_elements_by_xpath(
                            '//*[@id="block-system-main"]/div/div/div/div/section/div/div/div[2]')
                        statement_second = driver.find_elements_by_xpath(
                            '//*[@id="block-system-main"]/div/div/div/div/section/div/table/tbody/tr/td[1]/div')
                        statement_result_second = driver.find_elements_by_xpath(
                            '//*[@id="block-system-main"]/div/div/div/div/section/div/table/tbody/tr/td[2]/div')
                        for each in statement_one:
                            statement_list.append(each.text)
                        for each in statement_second:
                            statement_list.append(each.text)
                        for each in statement_result__one:
                            statement_result_list.append(each.text)
                        for each in statement_result_second:
                            statement_result_list.append(each.text)
                        state_information = pd.DataFrame(
                            {
                                'State': statement_list,
                                'Detail': statement_result_list
                            }
                        )
                        state_information.to_csv(r'%s' % path_fourth + '\\' + '%s.csv' % title)
                        time.sleep(random.randint(1, 2))
                        driver.close()
                    except:
                        print('Next')
                        driver.close()
            #time.sleep(random.randint(20, 25))
            # Documents
            driver = wd.Chrome(executable_path='./chromedriver', options=chrome_options_1)
            driver.get(link)
            time.sleep(random.randint(4, 5))
            search_item = each_name
            search_input_item = driver.find_element(by=By.XPATH,
                                                    value='//*[@id="edit-name-abn"]/div/input')
            search_input_item.send_keys(search_item)
            time.sleep(random.randint(2, 3))
            driver.find_element_by_xpath('//*[@id="edit-submit-solr-charities"]').click()
            time.sleep(random.randint(3, 4))
            detail_link = driver.find_element_by_xpath(
                '//*[@id="block-views-solr-charities-index"]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
            detail_link = detail_link.get_attribute('href')
            print(detail_link)
            financila_link = detail_link + '#financials-documents'
            driver.get(financila_link)
            document_button_list = driver.find_elements_by_xpath(
                '//*[@id="financials-documents"]/div/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[4]/a')
            for each_document in document_button_list:
                docuement_href = each_document.get_attribute('href')
                docuement_href_list = str(docuement_href).split('.')
                docuement_href_filter = docuement_href_list[-1]
                if docuement_href_filter == 'pdf' or docuement_href_filter == 'doc' or docuement_href_filter == 'xls':
                    driver.get(docuement_href)
           # time.sleep(random.randint(10, 15))
           # driver.quit()
            # scraping people
            real_name_list = []
            real_postion_list = []
            people_link = detail_link + '#people'
            people_link = requests.get(people_link)
            people_soup = BeautifulSoup(people_link.text, 'html.parser')
            name_list = people_soup.find_all('div', {'class': 'views-field views-field-title'})
            position_list = people_soup.find_all('div', {'class', 'views-field views-field-field-role'})
            for each_name in name_list:
                real_name = each_name.text
                real_name_list.append(real_name)
            for each_position in position_list:
                real_position = each_position.text
                real_postion_list.append(real_position)
            people_information = pd.DataFrame(
                {
                    'Name': real_name_list,
                    'Position': real_postion_list
                }
            )
            people_information.to_csv(r'%s' % first_path + '\\' + 'people.csv', encoding='utf-8-sig')

            # scraping overview
            real_item_list = []
            real_result_list = []
            overview_link = detail_link + '#overview'
            real_overview_link = requests.get(overview_link)
            link_soup = BeautifulSoup(real_overview_link.text, 'html.parser')
            try:
                real_link = link_soup.find('div', {'class', 'group-charity-details field-group-div'})
                item_list = real_link.find_all('div', {'class', 'field-label'})
                for each in item_list:
                    real_item_list.append(each.text)
                result_list = real_link.find_all('div', {'class', 'field-item even'})
                for i in range(0, len(real_item_list)):
                    real_result_list.append(result_list[i].text)
            except :
                real_item_list.append(' ')
                real_result_list.append(' ')
            try:
                summary_content = link_soup.find('div', {'class', 'group-summary-activities field-group-div'})
                summary_content = summary_content.find('div', {'class', 'field-item even'})
                real_result_list.append(summary_content.text)
            except:
                real_result_list.append(' ')
            try:
                operate_content = link_soup.find('div', {'class', 'group-charity-operates field-group-div'})
                operate_content = operate_content.find('div', {'class', 'field-item even'})
                real_result_list.append(operate_content.text)
            except:
                real_result_list.append(' ')
            try:
                register_content = link_soup.find('div', {'class', 'group-gov-agency field-group-div'})
                register_content = register_content.find('div', {'class', 'field-item even'})
                real_result_list.append(str(register_content.text).strip())
            except:
                real_result_list.append(' ')
            try:
                total_income = driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/p[1]')
                total_income = total_income.text
                total_income = str(total_income)
                total_income = total_income.replace('Total income', '')
            except:
                total_income = ''
            try:
                total_expenses = driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/p[2]')
                total_expenses = total_expenses.text
                total_expenses = str(total_expenses)
                total_expenses = total_expenses.replace('Total expenses', '')
            except:
                total_expenses = ''
            real_item_list.append('Summary of activities')
            real_item_list.append('States')
            real_item_list.append('Using the information on the Register')
            real_item_list.append('Total Income')
            real_item_list.append('Total expenses')
            real_result_list.append(total_income)
            real_result_list.append(total_expenses)
            overview_information = pd.DataFrame(
                {
                    'Charity_Items': real_item_list,
                    'Charity_Detail': real_result_list
                }
            )
            overview_information.to_csv(r'%s' % first_path + '\\' + 'overview.csv')

            # scraping History
            real_history_item = []
            real_history_content = []
            driver = wd.Chrome(executable_path='./chromedriver')
            people_link = detail_link + '#history'
            driver.get(people_link)
            time.sleep(random.randint(4, 5))
            history_item_list = driver.find_elements_by_xpath('//*[@id="history"]/div/div/div[1]')
            for each_item in history_item_list:
                real_history_item.append(each_item.text)

            history_content_list = driver.find_elements_by_xpath('//*[@id="history"]/div/div/div[2]/div/div')
            for each_content in history_content_list:
                real_history_content.append(each_content.text)
            history_information = pd.DataFrame(
                {
                    'Name': real_history_item,
                    'Position': real_history_content
                }
            )
            history_information.to_csv(r'%s' % first_path + '\\' + 'history.csv', encoding='utf-8-sig')
            driver.close()
            driver.quit()
            # take screenshot
            options = wd.ChromeOptions()
            options.headless = True
            driver = wd.Chrome(executable_path='./chromedriver', options=options)
            driver.get(overview_link)
            time.sleep(random.randint(2, 3))
            try:
                S = lambda X: driver.execute_script('return document.body.parentNode.scroll' + X)
                driver.set_window_size(S('Width'), S('Height'))  # May need manual adjustment
                driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/div[1]/div[1]').screenshot(
                    r'%s' % first_path + '\\' + 'Financial_Overview.png')
            except :
                print('no image')
            driver.close()
            driver.quit()
            text_1.delete('1.0', END)
            text_1.insert(tk.END, '%sth ABN has been scraped' % charity_number)
        except:
            text_1.delete('1.0', END)
            text_1.insert(tk.END, 'Go to next order...')
            charity_number = charity_number + 1
            continue
    text_1.delete('1.0', END)
    text_1.insert(tk.END, 'Scraping by Order Finished')

def Group():
        text_1.delete('1.0', END)
        text_1.insert(tk.END, 'Scraping by Group Started...Please wait')
        value = my_listbox.get(my_listbox.curselection())
        print(value)
        group_start_no = entry_4.get()
        group_start_no = int(group_start_no)
        group_end_no = entry_5.get()
        group_end_no = int(group_end_no)
        df = pd.read_excel('datadotgov_main.xls')
        act_list = df[value]
        charity_name_list = df['Charity_Legal_Name']
        Y_list = []
        sample = 'Y'
        new_act_list = []
        for each in act_list:
            new_act_list.append(each)
        j = 1
        for each in new_act_list:
            j = j + 1
            if each == sample:
                Y_list.append(charity_name_list[j-2])
        group_abn_numbers = len(Y_list)
        text_1.delete('1.0', END)
        text_1.insert(tk.END, 'There are %s abns.' % group_abn_numbers)
        #m = 0
        q = group_start_no
        for each_name in Y_list[group_start_no-1:group_end_no]:
            try:
                new_each = ''
                for k in each_name.split("\n"):
                    new_each += (re.sub(r"[^a-zA-Z0-9]+", ' ', k))
                # removing punctuation
               # m = m + 1
                if q == 1:
                    parent_dir = os.getcwd()
                    group_name = value
                    zero_path = os.path.join(parent_dir, group_name)
                    os.mkdir(zero_path)
                    charity_name = str(q) + '_' + new_each
                    directory_one = 'Financial_reporting'
                    directory_two = 'Documents'
                    directory_three = 'Annual information statement'
                    first_path = os.path.join(zero_path, charity_name)
                    os.mkdir(first_path)
                    print(first_path)
                    #second_parent_dir = os.getcwd() + '\\%s\\' % charity_name
                    path_second = os.path.join(first_path, directory_one)
                    path_third = os.path.join(first_path, directory_two)
                    path_fourth = os.path.join(first_path, directory_three)
                    os.mkdir(path_second)
                    os.mkdir(path_third)
                    os.mkdir(path_fourth)
                else:
                    parent_dir = os.getcwd()
                    group_name = value
                    zero_path = os.path.join(parent_dir, group_name)
                    charity_name = str(q) + '_' + new_each
                    directory_one = 'Financial_reporting'
                    directory_two = 'Documents'
                    directory_three = 'Annual information statement'
                    first_path = os.path.join(zero_path, charity_name)
                    os.mkdir(first_path)
                    print(first_path)
                    # second_parent_dir = os.getcwd() + '\\%s\\' % charity_name
                    path_second = os.path.join(first_path, directory_one)
                    path_third = os.path.join(first_path, directory_two)
                    path_fourth = os.path.join(first_path, directory_three)
                    os.mkdir(path_second)
                    os.mkdir(path_third)
                    os.mkdir(path_fourth)

                chrome_options = Options()
                chrome_options.add_experimental_option('prefs', {
                    "download.default_directory": path_second,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "plugins.always_open_pdf_externally": True
                }
                                                       )
                chrome_options_1 = Options()
                chrome_options_1.add_experimental_option('prefs', {
                    "download.default_directory": path_third,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "plugins.always_open_pdf_externally": True
                }
                                                         )


                # annual-reporting
                driver = wd.Chrome(executable_path='./chromedriver', options=chrome_options)
                driver.get(link)
                search_item = each_name
                search_input_item = driver.find_element(by=By.XPATH,
                                                        value='//*[@id="edit-name-abn"]/div/input')
                search_input_item.send_keys(search_item)
                time.sleep(random.randint(2, 3))
                driver.find_element_by_xpath('//*[@id="edit-submit-solr-charities"]').click()
                time.sleep(random.randint(3, 4))
                detail_link = driver.find_element_by_xpath(
                    '//*[@id="block-views-solr-charities-index"]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
                detail_link = detail_link.get_attribute('href')
                print(detail_link)
                financila_link = detail_link + '#financials-documents'
                driver.get(financila_link)
                time.sleep(random.randint(4, 5))
                try:
                    button_list = driver.find_elements_by_xpath(
                        '//*[@id="financials-documents"]/div/div[1]/div/div/div/div/div/div/table/tbody/tr/td/a')
                    print(len(button_list))
                    for each_button in button_list:
                        href = each_button.get_attribute('href')
                        href_list = str(href).split('.')
                        filter_href = href_list[-1]
                        if filter_href == 'pdf' or filter_href == 'doc' or filter_href == 'xls':
                            driver.get(href)
                    for each_button in button_list:
                        href = each_button.get_attribute('href')
                        href_list = str(href).split('.')
                        filter_href = href_list[-1]
                        if filter_href != 'pdf' and filter_href != 'doc' and filter_href != 'xls':
                            driver = wd.Chrome(executable_path='./chromedriver')
                            driver.get(href)
                            statement_list = []
                            statement_result_list = []
                            try:
                                title = driver.find_element_by_xpath(
                                    '//*[@id="block-system-main"]/div/div/div/div/section/div[1]/div[1]')
                                title = str(title.text)
                                title = title.replace('Annual Information Statement ', '')
                                statement_one = driver.find_elements_by_xpath(
                                    '//*[@id="block-system-main"]/div/div/div/div/section/div/div/div[1]')
                                statement_result__one = driver.find_elements_by_xpath(
                                    '//*[@id="block-system-main"]/div/div/div/div/section/div/div/div[2]')
                                statement_second = driver.find_elements_by_xpath(
                                    '//*[@id="block-system-main"]/div/div/div/div/section/div/table/tbody/tr/td[1]/div')
                                statement_result_second = driver.find_elements_by_xpath(
                                    '//*[@id="block-system-main"]/div/div/div/div/section/div/table/tbody/tr/td[2]/div')
                                for each in statement_one:
                                    statement_list.append(each.text)
                                for each in statement_second:
                                    statement_list.append(each.text)
                                for each in statement_result__one:
                                    statement_result_list.append(each.text)
                                for each in statement_result_second:
                                    statement_result_list.append(each.text)
                                state_information = pd.DataFrame(
                                    {
                                        'State': statement_list,
                                        'Detail': statement_result_list
                                    }
                                )
                                state_information.to_csv(r'%s' % path_fourth + '\\' + '%s.csv' % title)
                                time.sleep(random.randint(1, 2))
                                driver.close()
                            except :
                                print('Next')
                                driver.close()
                except :
                    print('no data')

               # time.sleep(random.randint(20, 25))
               # driver.quit()
                # Documents
                driver = wd.Chrome(executable_path='./chromedriver', options=chrome_options_1)
                driver.get(link)
                search_input_item = driver.find_element(by=By.XPATH,
                                                        value='//*[@id="edit-name-abn"]/div/input')
                search_input_item.send_keys(search_item)
                time.sleep(random.randint(2, 3))
                driver.find_element_by_xpath('//*[@id="edit-submit-solr-charities"]').click()
                time.sleep(random.randint(3, 4))
                detail_link = driver.find_element_by_xpath(
                    '//*[@id="block-views-solr-charities-index"]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
                detail_link = detail_link.get_attribute('href')
                print(detail_link)
                financila_link = detail_link + '#financials-documents'
                driver.get(financila_link)
                document_button_list = driver.find_elements_by_xpath(
                    '//*[@id="financials-documents"]/div/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[4]/a')
                for each_document in document_button_list:
                    docuement_href = each_document.get_attribute('href')
                    docuement_href_list = str(docuement_href).split('.')
                    docuement_href_filter = docuement_href_list[-1]
                    if docuement_href_filter == 'pdf' or docuement_href_filter == 'doc' or docuement_href_filter == 'xls':
                        driver.get(docuement_href)
               # time.sleep(random.randint(10, 15))
                #driver.close()
               # driver.quit()
                # scraping people
                real_name_list = []
                real_postion_list = []
                people_link = detail_link + '#people'
                people_link = requests.get(people_link)
                people_soup = BeautifulSoup(people_link.text, 'html.parser')
                name_list = people_soup.find_all('div', {'class': 'views-field views-field-title'})
                position_list = people_soup.find_all('div', {'class', 'views-field views-field-field-role'})
                for each_name in name_list:
                    real_name = each_name.text
                    real_name_list.append(real_name)
                for each_position in position_list:
                    real_position = each_position.text
                    real_postion_list.append(real_position)
                people_information = pd.DataFrame(
                    {
                        'Name': real_name_list,
                        'Position': real_postion_list
                    }
                )
                people_information.to_csv(r'%s' % first_path + '\\' + 'people.csv', encoding='utf-8-sig')
                #driver.close()
                # scraping overview
                real_item_list = []
                real_result_list = []
                overview_link = detail_link + '#overview'
                real_overview_link = requests.get(overview_link)
                link_soup = BeautifulSoup(real_overview_link.text, 'html.parser')
                try:
                    real_link = link_soup.find('div', {'class', 'group-charity-details field-group-div'})
                    item_list = real_link.find_all('div', {'class', 'field-label'})
                    for each in item_list:
                        real_item_list.append(each.text)
                    result_list = real_link.find_all('div', {'class', 'field-item even'})
                    for i in range(0, len(real_item_list)):
                        real_result_list.append(result_list[i].text)
                except:
                    real_item_list.append(' ')
                    real_result_list.append(' ')
                try:
                    summary_content = link_soup.find('div', {'class', 'group-summary-activities field-group-div'})
                    summary_content = summary_content.find('div', {'class', 'field-item even'})
                    real_result_list.append(summary_content.text)
                except:
                    real_result_list.append(' ')
                try:
                    operate_content = link_soup.find('div', {'class', 'group-charity-operates field-group-div'})
                    operate_content = operate_content.find('div', {'class', 'field-item even'})
                    real_result_list.append(operate_content.text)
                except:
                    real_result_list.append(' ')
                try:
                    register_content = link_soup.find('div', {'class', 'group-gov-agency field-group-div'})
                    register_content = register_content.find('div', {'class', 'field-item even'})
                    real_result_list.append(str(register_content.text).strip())
                except:
                    real_result_list.append(' ')
                try:
                    total_income = driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/p[1]')
                    total_income = total_income.text
                    total_income = str(total_income)
                    total_income = total_income.replace('Total income', '')
                except:
                    total_income = ''
                try:
                    total_expenses = driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/p[2]')
                    total_expenses = total_expenses.text
                    total_expenses = str(total_expenses)
                    total_expenses = total_expenses.replace('Total expenses', '')
                except:
                    total_expenses = ''
                real_item_list.append('Summary of activities')
                real_item_list.append('States')
                real_item_list.append('Using the information on the Register')
                real_item_list.append('Total Income')
                real_item_list.append('Total expenses')
                real_result_list.append(total_income)
                real_result_list.append(total_expenses)
                overview_information = pd.DataFrame(
                    {
                        'Charity_Items': real_item_list,
                        'Charity_Detail': real_result_list
                    }
                )
                overview_information.to_csv(r'%s' % first_path + '\\' + 'overview.csv')

                # scraping History
                real_history_item = []
                real_history_content = []
                driver = wd.Chrome(executable_path='./chromedriver')
                people_link = detail_link + '#history'
                driver.get(people_link)
                time.sleep(random.randint(4, 5))
                history_item_list = driver.find_elements_by_xpath('//*[@id="history"]/div/div/div[1]')
                for each_item in history_item_list:
                    real_history_item.append(each_item.text)

                history_content_list = driver.find_elements_by_xpath('//*[@id="history"]/div/div/div[2]/div/div')
                for each_content in history_content_list:
                    real_history_content.append(each_content.text)
                history_information = pd.DataFrame(
                    {
                        'Name': real_history_item,
                        'Position': real_history_content
                    }
                )
                history_information.to_csv(r'%s' % first_path + '\\' + 'history.csv', encoding='utf-8-sig')
                driver.quit()

                # take screenshot
                options = wd.ChromeOptions()
                options.headless = True
                driver = wd.Chrome(executable_path='./chromedriver', options=options)
                driver.get(overview_link)
                time.sleep(random.randint(2, 3))
                try:
                    S = lambda X: driver.execute_script('return document.body.parentNode.scroll' + X)
                    driver.set_window_size(S('Width'), S('Height'))  # May need manual adjustment
                    driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/div[1]/div[1]').screenshot(
                        r'%s' % first_path + '\\' + 'Financial_Overview.png')
                except :
                    print('no image')
                driver.close()
                driver.quit()
                text_1.delete('1.0', END)
                if q == 1:
                   text_1.insert(tk.END, '%sst ABN has been scraped' % q)
                elif q == 2:
                   text_1.insert(tk.END, '%snd ABN has been scraped' % q)
                else:
                   text_1.insert(tk.END, '%sth ABN has been scraped' % q)
                q = q + 1
            except :
                text_1.delete('1.0', END)
                text_1.insert(tk.END, 'Go to next abns...')
                q = q + 1
                continue
        text_1.delete('1.0', END)
        text_1.insert(tk.END, 'Scraping by Group Finished...')

def Abn():
    text_1.delete('1.0', END)
    text_1.insert(tk.END, 'Scraping by ADN Started...Please wait')
    responding_charity = ''
    df = pd.read_excel('datadotgov_main.xls')
    abn_list = df['ABN']
    charity_list = df['Charity_Legal_Name']
    sample_abn = entry_3.get()
    new_abn_list = []
    for each in abn_list:
        new_abn_list.append(each)
    for each in new_abn_list:
        if each == sample_abn:
            filter_number = new_abn_list.index(each)
            responding_charity = charity_list[filter_number]
    new_name = ''
    for k in responding_charity.split("\n"):
        new_name += (re.sub(r"[^a-zA-Z0-9]+", ' ', k))
    parent_dir = os.getcwd()
    charity_name = new_name
    folder_name = sample_abn + '_' + charity_name
    directory_one = 'Financial_reporting'
    directory_two = 'Documents'
    directory_three = 'Annual information statement'
    first_path = os.path.join(parent_dir, folder_name)
    os.mkdir(first_path)
    print(first_path)
    second_parent_dir = os.getcwd() + '\\%s\\' % folder_name
    path_second = os.path.join(second_parent_dir, directory_one)
    path_third = os.path.join(second_parent_dir, directory_two)
    path_fourth = os.path.join(second_parent_dir, directory_three)
    os.mkdir(path_second)
    os.mkdir(path_third)
    os.mkdir(path_fourth)
    chrome_options = Options()
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": path_second,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
                                           )
    chrome_options_1 = Options()
    chrome_options_1.add_experimental_option('prefs', {
        "download.default_directory": path_third,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True
    }
                                             )

    # annual-reporting
    driver = wd.Chrome(executable_path='./chromedriver', options=chrome_options)
    driver.get(link)
    search_item = charity_name
    search_input_item = driver.find_element(by=By.XPATH,
                                            value='//*[@id="edit-name-abn"]/div/input')
    search_input_item.send_keys(search_item)
    time.sleep(random.randint(2, 3))
    driver.find_element_by_xpath('//*[@id="edit-submit-solr-charities"]').click()
    time.sleep(random.randint(3, 4))
    detail_link = driver.find_element_by_xpath(
        '//*[@id="block-views-solr-charities-index"]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
    detail_link = detail_link.get_attribute('href')
    print(detail_link)
    financila_link = detail_link + '#financials-documents'
    driver.get(financila_link)
    time.sleep(random.randint(4, 5))
    button_list = driver.find_elements_by_xpath(
        '//*[@id="financials-documents"]/div/div[1]/div/div/div/div/div/div/table/tbody/tr/td/a')
    for each_button in button_list:
        href = each_button.get_attribute('href')
        href_list = str(href).split('.')
        filter_href = href_list[-1]
        if filter_href == 'pdf' or filter_href == 'doc' or filter_href == 'xls':
            driver.get(href)
    for each_button in button_list:
        href = each_button.get_attribute('href')
        href_list = str(href).split('.')
        filter_href = href_list[-1]
        if filter_href != 'pdf' and filter_href != 'doc' and filter_href != 'xls':
            driver = wd.Chrome(executable_path='./chromedriver')
            driver.get(href)
            statement_list = []
            statement_result_list = []
            try:
                title = driver.find_element_by_xpath('//*[@id="block-system-main"]/div/div/div/div/section/div[1]/div[1]')
                title = str(title.text)
                title = title.replace('Annual Information Statement ', '')
                statement_one = driver.find_elements_by_xpath(
                    '//*[@id="block-system-main"]/div/div/div/div/section/div/div/div[1]')
                statement_result__one = driver.find_elements_by_xpath(
                    '//*[@id="block-system-main"]/div/div/div/div/section/div/div/div[2]')
                statement_second = driver.find_elements_by_xpath(
                    '//*[@id="block-system-main"]/div/div/div/div/section/div/table/tbody/tr/td[1]/div')
                statement_result_second = driver.find_elements_by_xpath(
                    '//*[@id="block-system-main"]/div/div/div/div/section/div/table/tbody/tr/td[2]/div')
                for each in statement_one:
                    statement_list.append(each.text)
                for each in statement_second:
                    statement_list.append(each.text)
                for each in statement_result__one:
                    statement_result_list.append(each.text)
                for each in statement_result_second:
                    statement_result_list.append(each.text)
                state_information = pd.DataFrame(
                    {
                        'State': statement_list,
                        'Detail': statement_result_list
                    }
                )
                state_information.to_csv(r'%s' % path_fourth + '\\' + '%s.csv' % title)
                time.sleep(random.randint(1, 2))
                driver.close()
            except :
                print('Next')
                driver.close()
    #time.sleep(random.randint(20, 25))
    #driver.quit()
    # Documents
    driver = wd.Chrome(executable_path='./chromedriver', options=chrome_options_1)
    driver.get(link)
    search_item = charity_name
    search_input_item = driver.find_element(by=By.XPATH,
                                            value='//*[@id="edit-name-abn"]/div/input')
    search_input_item.send_keys(search_item)
    time.sleep(random.randint(2, 3))
    driver.find_element_by_xpath('//*[@id="edit-submit-solr-charities"]').click()
    time.sleep(random.randint(3, 4))
    detail_link = driver.find_element_by_xpath(
        '//*[@id="block-views-solr-charities-index"]/div/div[2]/div/div/div[2]/div/table/tbody/tr/td[1]/a')
    detail_link = detail_link.get_attribute('href')
    print(detail_link)
    financila_link = detail_link + '#financials-documents'
    driver.get(financila_link)
    document_button_list = driver.find_elements_by_xpath(
        '//*[@id="financials-documents"]/div/div[2]/div[2]/div/div/div/div/div/table/tbody/tr/td[4]/a')

    for each_document in document_button_list:
        docuement_href = each_document.get_attribute('href')
        docuement_href_list = str(docuement_href).split('.')
        docuement_href_filter = docuement_href_list[-1]
        if docuement_href_filter == 'pdf' or docuement_href_filter == 'doc' or docuement_href_filter == 'xls':
            driver.get(docuement_href)

    #time.sleep(random.randint(10, 15))
    #driver.quit()

    # scraping people
    real_name_list = []
    real_postion_list = []
    people_link = detail_link + '#people'
    people_link = requests.get(people_link)
    people_soup = BeautifulSoup(people_link.text, 'html.parser')
    name_list = people_soup.find_all('div', {'class': 'views-field views-field-title'})
    position_list = people_soup.find_all('div', {'class', 'views-field views-field-field-role'})
    for each_name in name_list:
        real_name = each_name.text
        real_name_list.append(real_name)
    for each_position in position_list:
        real_position = each_position.text
        real_postion_list.append(real_position)
    people_information = pd.DataFrame(
        {
            'Name': real_name_list,
            'Position': real_postion_list
        }
    )
    people_information.to_csv(r'%s' % first_path + '\\' + 'people.csv')
    #driver.close()

    # scraping overview
    # scraping overview
    real_item_list = []
    real_result_list = []
    overview_link = detail_link + '#overview'
    real_overview_link = requests.get(overview_link)
    link_soup = BeautifulSoup(real_overview_link.text, 'html.parser')
    try:
        real_link = link_soup.find('div', {'class', 'group-charity-details field-group-div'})
        item_list = real_link.find_all('div', {'class', 'field-label'})
        for each in item_list:
            real_item_list.append(each.text)
        result_list = real_link.find_all('div', {'class', 'field-item even'})
        for i in range(0, len(real_item_list)):
            real_result_list.append(result_list[i].text)
    except:
        real_item_list.append(' ')
        real_result_list.append(' ')
    try:
        summary_content = link_soup.find('div', {'class', 'group-summary-activities field-group-div'})
        summary_content = summary_content.find('div', {'class', 'field-item even'})
        real_result_list.append(summary_content.text)
    except:
        real_result_list.append(' ')
    try:
        operate_content = link_soup.find('div', {'class', 'group-charity-operates field-group-div'})
        operate_content = operate_content.find('div', {'class', 'field-item even'})
        real_result_list.append(operate_content.text)
    except:
        real_result_list.append(' ')
    try:
        register_content = link_soup.find('div', {'class', 'group-gov-agency field-group-div'})
        register_content = register_content.find('div', {'class', 'field-item even'})
        real_result_list.append(str(register_content.text).strip())
    except:
        real_result_list.append(' ')
    try:
        total_income = driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/p[1]')
        total_income = total_income.text
        total_income = str(total_income)
        total_income = total_income.replace('Total income', '')
    except:
        total_income = ''
    try:
        total_expenses = driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/p[2]')
        total_expenses = total_expenses.text
        total_expenses = str(total_expenses)
        total_expenses = total_expenses.replace('Total expenses', '')
    except:
        total_expenses = ''
    real_item_list.append('Summary of activities')
    real_item_list.append('States')
    real_item_list.append('Using the information on the Register')
    real_item_list.append('Total Income')
    real_item_list.append('Total expenses')
    real_result_list.append(total_income)
    real_result_list.append(total_expenses)
    overview_information = pd.DataFrame(
        {
            'Charity_Items': real_item_list,
            'Charity_Detail': real_result_list
        }
    )
    overview_information.to_csv(r'%s' % first_path + '\\' + 'overview.csv')

    # scraping History
    real_history_item = []
    real_history_content = []
    driver = wd.Chrome(executable_path='./chromedriver')
    people_link = detail_link + '#history'
    driver.get(people_link)
    time.sleep(random.randint(4, 5))
    history_item_list = driver.find_elements_by_xpath('//*[@id="history"]/div/div/div[1]')
    for each_item in history_item_list:
        real_history_item.append(each_item.text)

    history_content_list = driver.find_elements_by_xpath('//*[@id="history"]/div/div/div[2]/div/div')
    for each_content in history_content_list:
        real_history_content.append(each_content.text)
    history_information = pd.DataFrame(
        {
            'Name': real_history_item,
            'Position': real_history_content
        }
    )
    history_information.to_csv(r'%s' % first_path + '\\' + 'history.csv')

    # take screenshot
    options = wd.ChromeOptions()
    options.headless = True
    driver = wd.Chrome(executable_path='./chromedriver', options=options)
    driver.get(overview_link)
    time.sleep(random.randint(2, 3))
    try:
        S = lambda X: driver.execute_script('return document.body.parentNode.scroll' + X)
        driver.set_window_size(S('Width'), S('Height'))  # May need manual adjustment
        driver.find_element_by_xpath('//*[@id="financial-overview"]/div/div/div/div[1]/div[1]').screenshot(r'%s' % first_path + '\\' + 'Financial_Overview.png')
    except :
        print('no image')
    driver.close()
    driver.quit()
    text_1.delete('1.0', END)
    text_1.insert(tk.END, 'Scraping by Abn Finished.....')

if __name__== "__main__":
    link = 'https://www.acnc.gov.au/charity'
    root = Tk()
    root.geometry('1000x700')
    root.title("Charity Scraper")
    my_frame = Frame(root)
    my_scrollbar = Scrollbar(my_frame, orient=VERTICAL)
    my_listbox = Listbox(my_frame, width=45)
    my_scrollbar.config(command=my_listbox.yview)
    my_listbox.config(yscrollcommand=my_scrollbar.set)
    my_scrollbar.pack(side=RIGHT, fill=Y)
    my_frame.place(x=360, y=260)
    my_listbox.pack()
    my_list = ["Operates_in_ACT", "Operates_in_NSW", "Operates_in_NT", 'Operates_in_QLD', 'Operates_in_SA', 'Operates_in_TAS',
               'Operates_in_VIC', 'Operates_in_WA', 'PBI', 'HPC', 'Preventing_or_relieving_suffering_of_animals',
               'Advancing_Culture', 'Advancing_Education', 'Advancing_Health', 'Promote_or_oppose_a_change_to_law__government_poll_or_prac',
               'Advancing_natual_environment', 'Promoting_or_protecting_human_rights', 'Purposes_beneficial_to_ther_general_public_and_other_analogous',
               'Promoting_reconciliation__mutual_respect_and_tolerance', 'Advancing_Religion', 'Advancing_social_or_public_welfare', 'Advancing_security_or_safety_of_Australia_or_Australian_public',
               'Aboriginal_or_TSI', 'Adults', 'Aged_Persons', 'Chidren', 'Communities_Overseas', 'Early_Childhood', 'Ethnic_Groups',
               'Families', 'Females', 'Financially_Disadvantaged', 'Gay_Lesbian_Bisexual', 'General_Community_in_Australia',
               'Males', 'Migrants_Refugees_or_Asylum_Seekers', 'Other_Beneficiaries', 'Other_Charities', 'People_at_risk_of_homelessness',
               'People_with_Chronic_Illness', 'People_with_Disabilities', 'Pre_Post_Release_Offenders', 'Rural_Regional_Remote_Communities',
               'Unemployed_Person', 'Veterans_or_their_families', 'Victims_of_crime', 'Victims_of_Disasters', 'Youth']
    for item in my_list:
        my_listbox.insert(END, item)
    label_1 = Label(root, text="Start No: ", width=20, font=("bold", 10))
    label_1.place(x=10, y=260)
    label_2 = Label(root, text="End No: ", width=20, font=("bold", 10))
    label_2.place(x=10, y=320)
    label_3 = Label(root, text=" (From 2): ", width=20, font=("bold", 9))
    label_3.place(x=140, y=260)
    entry_1 = Entry(root, width=10)
    entry_1.place(x=120, y=260)
    entry_1.focus_set()
    entry_2 = Entry(root, width=10)
    entry_2.place(x=120, y=320)
    label_4 = Label(root, text="ABN: ", font=("bold", 10))
    label_4.place(x=750, y=260)
    entry_3 = Entry(root, width=20)
    entry_3.place(x=800, y=260)
    var = IntVar()
    label_2 = Label(root, text="Processing.....", width=20, font=("bold", 10))
    label_2.place(x=130, y=580)
    text_1 = Text(root, width=55, height=1.5)
    text_1.place(x=270, y=570)
    #updating for group
    label_5 = Label(root, text="Start: ", width=20, font=("bold", 9))
    label_5.place(x=300, y=450)
    label_6 = Label(root, text="End: ", width=20, font=("bold", 9))
    label_6.place(x=450, y=450)
    entry_4 = Entry(root, width=10)
    entry_4.place(x=400, y=450)
    entry_5 = Entry(root, width=10)
    entry_5.place(x=550, y=450)
    Button(root, text='Search by Order', width=20, fg='black', command=lambda: threading.Thread(target=Order).start()).place(x=80, y=500)
    Button(root, text='Search by Group', width=20, fg='black', command=lambda: threading.Thread(target=Group).start()).place(x=430, y=500)
    Button(root, text='Search by ABN', width=20, fg='black', command=lambda: threading.Thread(target=Abn).start()).place(x=780, y=500)
    Button(root, text='Clear', width=15, fg='black', command=clearText).place(x=260, y=650)
    Button(root, text='Quit', width=15, fg='black', command=root.destroy).place(x=620, y=650)
    canvas = Canvas(root, width=908, height=242)
    canvas.pack()
    img = ImageTk.PhotoImage(Image.open("logo.png"))
    canvas.create_image(20, 20, anchor=NW, image=img)
    root.mainloop()








