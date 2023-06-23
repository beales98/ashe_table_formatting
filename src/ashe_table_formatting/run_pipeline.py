import yaml

from ashe_table_formatting.Create_ASHE_tables import (
    create_table,
    create_workbook
)

def run_pipeline():
    with open("package_config.yaml", "r") as file:
        example_config = yaml.safe_load(file)
    example_config = example_config["file_paths"][0]
    csv_path = example_config["csv_path"]
    csv_previous_year_path = example_config["csv_previous_year_path"]
    template_path = example_config["template_path"]
    output_path = example_config["output_path"]
    year = example_config["year"]
    create_table(csv_path, csv_previous_year_path, template_path, output_path, 'Table 2 - Occupation (2)', year)
    create_workbook(csv_path, csv_previous_year_path, template_path, output_path, 'Table 9 - Work PC', 'Hourly Pay', year)

if __name__ == "__main__":
    run_pipeline()
