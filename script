import pandas as pd
import re

# Load class teachers
def load_classteachers(filename):
    df = pd.read_csv(filename, header=None)
    mapping = {}
    # Find every class row and the next two rows are subject and teacher
    for i in range(2, len(df), 2):
        class_name = str(df.iloc[i, 0]).strip()
        subject = str(df.iloc[i, 1]).strip()
        teacher = str(df.iloc[i+1, 1]).strip()
        mapping[class_name] = (subject, teacher)
    return mapping

# Parse timetable CSV to DataFrames per class
def parse_timetable(filename):
    with open(filename, encoding="utf-8") as f:
        lines = [l.strip() for l in f.readlines() if l.strip()]
    class_tables = {}
    i = 0
    while i < len(lines):
        if lines[i].startswith("Class"):
            header = lines[i+1].split(",")
            class_name = header[0].replace("Period","").strip()
            periods = header[1:]
            data = []
            j = i + 2
            while j < len(lines) and not lines[j].startswith("Class"):
                row = lines[j].split(",")
                data.append(row)
                j += 1
            df = pd.DataFrame(data, columns=["Day"] + periods)
            class_tables[class_name] = df
            i = j
        else:
            i += 1
    return class_tables

# Update timetable with class teacher for 1st period
def update_timetable(class_tables, classteacher_map):
    new_tables = {}
    for class_name, df in class_tables.items():
        canonical = class_name.strip()
        if canonical in classteacher_map:
            subj, teacher = classteacher_map[canonical]
            cell_value = f"{subj}\n{teacher}"
            df = df.copy()
            for idx in df.index:
                df.iloc[idx, 1] = cell_value  # 1st period
            # Ensure all cells are formatted as Subject\nTeacher
            for c in df.columns[1:]:
                for idx in df.index:
                    val = df.at[idx, c]
                    if val and '\n' in val:
                        subj, teacher = map(str.strip, val.split('\n', 1))
                        val = f"{subj}\n{teacher}"
                    else:
                        # Try to split with last space as fallback
                        m = re.match(r"(.+)\s+([^\s]+)$", val) if val else None
                        if m:
                            subj, teacher = m.group(1), m.group(2)
                            val = f"{subj}\n{teacher}"
                    df.at[idx, c] = val
            new_tables[class_name] = df
    return new_tables

# Write to Excel
def write_excel(class_tables, filename):
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        for cls, df in class_tables.items():
            ws_name = cls.replace("-", "_").replace(" ", "_")[:31]
            df.to_excel(writer, sheet_name=ws_name, index=False)

if __name__ == "__main__":
    ct = load_classteachers("Classteachers.csv")
    ct_tables = parse_timetable("Alright.csv")
    updated = update_timetable(ct_tables, ct)
    write_excel(updated, "New_Class_Timetable.xlsx")
