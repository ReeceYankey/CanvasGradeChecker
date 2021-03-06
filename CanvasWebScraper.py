from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from bs4 import BeautifulSoup
import pandas as pd
from time import sleep
import re
from decouple import config
from UpdateFromCSV import UpdateFromCSV
import os.path


if not os.path.isfile("settings.ini"):
    print("settings.ini has not been detected. Initiating first time setup...")
    CHROME_DRIVER_PATH = input("Please enter the path of your chromedriver.exe:")
    CANVAS_USERNAME = input("Please enter your canvas username (optional):")
    CANVAS_PASSWORD = input("Please enter your canvas password (optional):")

    settings = open("settings.ini", "w+")
    settings.write(("[settings]\n"
                    "CHROME_DRIVER_PATH={}\n"
                    "CANVAS_USERNAME={}\n"
                    "CANVAS_PASSWORD={}").format(CHROME_DRIVER_PATH, CANVAS_USERNAME, CANVAS_PASSWORD))
    settings.close()


with webdriver.Chrome(config("CHROME_DRIVER_PATH")) as driver:
    driver.get("https://canvas.vt.edu/")
    userElem = driver.find_element_by_name("j_username")
    userElem.send_keys(config("CANVAS_USERNAME"))
    passElem = driver.find_element_by_name("j_password")
    passElem.send_keys(config("CANVAS_PASSWORD"))

    # ask user to login
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

    # find relevant class elements on dashboard
    class_link_elements = driver.find_elements_by_class_name("ic-DashboardCard__link")
    class_name_elements = driver.find_elements_by_class_name("ic-DashboardCard__header-subtitle")

    class_links = []  # class link/href
    class_names = []  # class name (ex: ENGE 1215)

    for elem in class_link_elements:
        class_links.append(elem.get_attribute("href"))

    for elem in class_name_elements:
        # get shortened course identifier (ex. ENGR 1045)
        result = re.search(r'[A-Z]{4}[\s_]?[0-9]{4}', elem.text)

        if result is None:
            currPos = len(class_names)
            class_links.pop(currPos)  # get rid of corresponding href
            continue

        # standardize result format and insert into class_names
        string = result.group(0)
        class_names.append(string[:4] + " " + string[-4:])

    print(class_links)
    print(class_names)

    for class_link, class_name in zip(class_links, class_names):
        print("gathering data for: "+class_name)

        # goto link
        driver.get(class_link + "/grades")
        html = driver.page_source
        soup = BeautifulSoup(html, features="html.parser")
        table = soup.find("table", id="grades_summary")

        # grab all table data
        rows = table.find_all("tr", class_="student_assignment")
        table_data = {"name": [], "date": [], "score": [], "max_score": [], "type": []}
        for row in rows:
            # skip elements picked up that aren't actually assignments
            if "group_total" in row["class"] or "final_grade" in row["class"]:
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

        # store into csv
        table = pd.DataFrame(table_data)
        table.to_csv(class_name + ".csv")

    UpdateFromCSV(class_names)
