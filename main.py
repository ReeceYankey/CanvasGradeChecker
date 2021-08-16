from Setup import verify_configuration
from CanvasWebScraper import fetch_grades
from UpdateFromCSV import update_from_csv

# TODO gui to organize assignment types?
# TODO reorganize GradeSheets.py?

verify_configuration()
class_names = fetch_grades()
print(class_names)
update_from_csv(class_names)
