# CSV to iOS Calendar (ICS)

- Notebook: Course_CSV_to_iOS_Calendar.ipynb
- Requirements: requirements.txt

Run Course_CSV_to_iOS_Calendar.ipynb to convert your CSV to .ics.
Please modify the file path and then run it
the columns contain `name, weekday, start, end, location, start_date, end_date, count, interval, exceptions, rdates`

# Teaching Plan → iOS Calendar

- `TeachingPlan_Parse_and_Select.ipynb`: Repo-friendly notebook that reads a DOCX, lists all detected CourseCode+Class combos, lets you edit desired codes, and writes an `.ics` to output.

Quick start:
1. Place your teaching plan `.docx` in `data/` (create if missing).
2. Open `TeachingPlan_Parse_and_Select.ipynb` and run cells; edit `DESIRED_CODES` to choose courses (e.g., `ECON6001D`).
3. The generated calendar will be saved under `output/`.

Dependencies are installed from within the notebooks when first run (`python-docx`, `icalendar`)
