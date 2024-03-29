from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import pandas as pd
from time import sleep
import re
from decouple import config


def wait_for_dashboard(driver):
    while True:
        print("waiting for user to log in...")
        sleep(3)
        try:
            driver.find_element_by_class_name("ic-DashboardCard__link")
            print("logged in sucessfully")
            break
        except NoSuchElementException:
            pass
    sleep(8)


def gather_class_elements(driver):
    class_link_elements = driver.find_elements_by_class_name("ic-DashboardCard__link")
    print(class_link_elements)
    class_name_elements = driver.find_elements_by_class_name("ic-DashboardCard__header-subtitle")
    print(class_name_elements)

    class_links = []  # class link/href
    class_names = []  # class name (ex: ENGE 1215)

    for elem in class_link_elements:
        class_links.append(elem.get_attribute("href"))
    print(class_links)
    for elem in class_name_elements:
        # get shortened course identifier (ex. ENGR 1045)
        print(elem.text)
        result = re.search(r'[A-Z]{2,4}[\s_]?[0-9]{4}', elem.text)
        print(result)

        if result is None:
            currPos = len(class_names)
            class_links.pop(currPos)  # get rid of corresponding href
            continue

        # standardize result format and insert into class_names
        string = result.group(0)
        major = re.search(r'[A-Z]{2,4}', string).group(0)
        class_num = re.search(r'[0-9]{4}', string).group(0)
        class_names.append(major + " " + class_num)
        # class_names.append(string[:4] + " " + string[-4:])
    return class_links, class_names


def process_grade_table(table):

    rows = table.find_all("tr", class_="student_assignment")
    table_data = {"name": [], "date": [], "score": [], "max_score": [],
                  "type": []}

    for row in rows:
        # skip elements picked up that aren't actually assignments
        if is_not_assignment(row):
            continue

        # add name of assignment
        table_data["name"].append(row.find("a").text)

        # add type of assignment
        table_data["type"].append(row.find("div", class_="context").text)

        # add due date of assignment
        date = row.find("td", class_="due").text
        formatted_date = re.search(r"[A-Za-z]{3}\s\d{1,2}", date)
        if formatted_date:
            table_data["date"].append(formatted_date.group(0))
        else:
            table_data["date"].append("N/A")

        # add grade of assignment
        score = row.find("span", class_="original_score").text
        formattedScore = re.search(r"\S+", score)
        if formattedScore:
            table_data["score"].append(formattedScore.group(0))
        else:
            table_data["score"].append("N/A")

        # add maximum score of assignment
        max_score = row.find("td", class_="points_possible").text
        formatted_max_score = re.search(r"\S+", max_score)
        table_data["max_score"].append(formatted_max_score.group(0))  # should be guaranteed to exist
    
    return table_data

def is_not_assignment(row):
    return "group_total" in row["class"] or "final_grade" in row["class"]

def create_webdriver():
    print(f"Launching with {config('PREFERRED_BROWSER')}")
    if config("PREFERRED_BROWSER") == 'chrome':
        return webdriver.Chrome(str(config("DRIVER_PATH")))
    elif config("PREFERRED_BROWSER") == 'firefox':
        return webdriver.Firefox(executable_path=str(config("DRIVER_PATH")))
    raise Exception("invalid PREFERRED_BROWSER in config")

def fetch_grades():
    with create_webdriver() as driver:
        driver.get("https://canvas.vt.edu/")
        userElem = driver.find_element_by_name("j_username")
        userElem.send_keys(config("CANVAS_USERNAME"))
        passElem = driver.find_element_by_name("j_password")
        passElem.send_keys(config("CANVAS_PASSWORD"))

        # ask user to login
        wait_for_dashboard(driver)

        # find relevant class elements on dashboard
        class_links, class_names = gather_class_elements(driver)
        print(class_links)
        print(class_names)

        for class_link, class_name in zip(class_links, class_names):
            print("gathering data for: "+class_name)

            # goto link and grab table
            driver.get(class_link + "/grades")
            html = driver.page_source
            soup = BeautifulSoup(html, features="html.parser")
            table = soup.find("table", id="grades_summary")

            table_data = process_grade_table(table)

            # store into csv
            save_table_data(table_data, class_name)

    return class_names

def save_table_data(table_data, class_name):
    table = pd.DataFrame(table_data)
    table.to_csv("ClassData\\" + class_name + ".csv")
