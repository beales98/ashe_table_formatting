import yaml
import os

from ashe_table_formatting.Create_ASHE_tables import (
    create_table,
    create_workbook
)

from ashe_table_formatting.pipeline_config import *

def run_pipeline():
    os.chdir('D:\ASHE_table_formatting')
    with open("package_config.yaml", "r") as file:
        example_config = yaml.safe_load(file)
    example_config = example_config["file_paths"][0]
    csv_path = example_config["csv_path"]
    csv_previous_year_path = example_config["csv_previous_year_path"]
    template_path = example_config["template_path"]
    output_path = example_config["output_path"]
    year = example_config["year"]
    create_table(csv_path, csv_previous_year_path, template_path, output_path, 'Table test - Test data', year)
    #create_workbook(csv_path, csv_previous_year_path, template_path, output_path, 'Table test - Test data', 'Hourly Pay', year)

if __name__ == "__main__":
    run_pipeline()
