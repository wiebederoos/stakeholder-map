import pandas as pd
import os
import sys
import uuid

# -------- CONFIG --------

REQUIRED_COLUMNS = ["name", "role", "department", "theme"]

THEME_COLORS = {
    "CI/CD pipelines": "#4CAF50",
    "runtime environments": "#2196F3",
    "ICT security & compliance": "#F44336",
    "Database": "#FF9800"
}


# -------- VALIDATION --------

def validate_input_file(file_path):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Input file does not exist: {file_path}")

    if not file_path.endswith((".xlsx", ".xls")):
        raise ValueError("Input file must be an Excel file (.xlsx or .xls)")


def validate_dataframe(df):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    if df.empty:
        raise ValueError("Excel file contains no data rows.")

    if df[REQUIRED_COLUMNS].isnull().any().any():
        raise ValueError("Empty values found in required columns.")

    invalid_themes = df[~df["theme"].isin(THEME_COLORS.keys())]
    if not invalid_themes.empty:
        raise ValueError(
            f"Invalid themes found:\n{invalid_themes[['name', 'theme']]}"
        )


# -------- MURAL CSV GENERATION --------

def generate_mural_csv(df, output_file):

    shapes = []
    connectors = []

    theme_positions = {}
    x_theme = 0
    y_theme = 0
    spacing_x = 600
    spacing_y = 400

    # 1️⃣ Create one sticky per theme
    for theme in df["theme"].unique():
        theme_id = str(uuid.uuid4())
        theme_positions[theme] = theme_id

        shapes.append({
            "Id": theme_id,
            "Text": f"{theme}",
            "Shape": "sticky",
            "FillColor": THEME_COLORS[theme],
            "X": x_theme,
            "Y": y_theme,
            "Width": 300,
            "Height": 200
        })

        x_theme += spacing_x

    # 2️⃣ Create stickies per row and connect to theme
    x_offset = 0
    y_offset = 300

    for _, row in df.iterrows():
        person_id = str(uuid.uuid4())
        theme_id = theme_positions[row["theme"]]

        shapes.append({
            "Id": person_id,
            "Text": f"{row['name']}\n{row['role']}\nDept: {row['department']}",
            "Shape": "sticky",
            "FillColor": THEME_COLORS[row["theme"]],
            "X": x_offset,
            "Y": y_offset,
            "Width": 250,
            "Height": 150
        })

        connectors.append({
            "FromId": person_id,
            "ToId": theme_id,
            "Type": "arrow"
        })

        x_offset += 300
        if x_offset > 2000:
            x_offset = 0
            y_offset += spacing_y

    # Export two CSV sections (Mural supports object import)
    shapes_df = pd.DataFrame(shapes)
    connectors_df = pd.DataFrame(connectors)

    shapes_df.to_csv(output_file.replace(".csv", "_shapes.csv"), index=False)
    connectors_df.to_csv(output_file.replace(".csv", "_connectors.csv"), index=False)

    print("✅ Mural import files generated:")
    print(output_file.replace(".csv", "_shapes.csv"))
    print(output_file.replace(".csv", "_connectors.csv"))


# -------- MAIN --------

def main():
    if len(sys.argv) != 3:
        print("Usage: python script.py input.xlsx output.csv")
        sys.exit(1)

    input_file = sys.argv[1]
    output_file = sys.argv[2]

    try:
        validate_input_file(input_file)
        df = pd.read_excel(input_file)  # header row automatically handled
        validate_dataframe(df)
        generate_mural_csv(df, output_file)

    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
