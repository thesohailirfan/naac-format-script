import os
import mammoth
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import random
import re

# ========================================================
# Get docs directory
# ========================================================

path = "docs"
dir_list = os.listdir(path)
if len(dir_list) == 0:
    print("No files found in docs directory")
    exit()
print(dir_list)

# ========================================================
# Convert to html for processing
# ========================================================

for file in dir_list:
    if file.endswith(".docx"):

        # =============Convert to HTML====================

        f = open(f"docs/{file}", "rb")
        filename = file[:-5]
        print("Converting - ", file)
        b = open(f"html/{filename}.html", "wb")
        document = mammoth.convert_to_html(f)
        b.write(document.value.encode("utf8"))
        f.close()
        b.close()
        print("Converted - ", file)

        # ================Load HTML=======================

        driver = webdriver.Chrome()
        url = f"file:///{os.path.abspath(f'html/{filename}.html')}"
        driver.get(url)
        time.sleep(1)

        # =================Get All Tables=================

        tables = driver.find_elements(By.TAG_NAME, "table")
        print(len(tables), " Tables found")

        # ==============Loop through Tables===============
        html = ""
        for p in range(len(tables)):
            table = tables[p]
            rows = table.find_elements(By.TAG_NAME, "tr")
            if "Mapping between COs and POs 2".lower() in rows[0].text.lower():
                table_next = tables[p + 1]
                rows_next = table_next.find_elements(By.TAG_NAME, "tr")[
                    2
                ].find_elements(By.TAG_NAME, "p")
                print(
                    "Found Mapping",
                    len(
                        table_next.find_elements(By.TAG_NAME, "tr")[2].find_elements(
                            By.TAG_NAME, "p"
                        )
                    ),
                )

                course_code = re.sub("[\s+]", "", rows_next[0].text)
                course_code_number = course_code[3:]
                course_name = rows_next[1].text
                print(
                    f"\n ==========={p}============= \n",
                    course_code,
                    course_name,
                    "\n ======================== \n",
                )
                html += "<style>table, th, td {border: 1px solid black;border-collapse: collapse;}</style><br/><br/><table><tr><th>Course Code</th><th>Course Name</th><th>COs</th><th>PO1</th><th>PO2</th><th>PO3</th><th>PO4</th><th>PO5</th><th>PO6</th><th>PO7</th><th>PO8</th><th>PO9</th><th>PO10</th><th>PO11</th><th>PO12</th><th>PSO1</th><th>PSO2</th><th>PSO3</th></tr>"
                avg = []
                avg_count = []
                pso_avg = []

                for z in range(0, 12):
                    avg.append(0)
                    avg_count.append(0)

                for z in range(0, 3):
                    pso_avg.append(0)

                for i in range(2, len(rows)):
                    this_row = rows[i]
                    cols = this_row.find_elements(By.TAG_NAME, "td")
                    html += "<tr>"
                    if i - 1 == 1:
                        html += f"<td rowspan='{len(rows)-1}'>{course_code}</td>"
                        html += f"<td rowspan='{len(rows)-1}'>{course_name}</td>"
                    html += f"<td>CO{course_code_number}.{i-1}</td>"
                    origin_pos = cols[2].text.split(",")
                    origin_pos = [_.strip() for _ in origin_pos]
                    pos = origin_pos + [
                        "PO1",
                        "PO2",
                        "PO3",
                        "PO4",
                        "PO5",
                        "PO6",
                        "PO7",
                        "PO9",
                        "PO12",
                    ]
                    pos = [_.strip() for _ in pos]
                    pos = list(set(pos))
                    print(len(origin_pos), len(pos))
                    print(origin_pos)
                    print(pos)
                    for j in range(0, 12):
                        po = f"PO{j+1}"
                        if po in pos:
                            if po in origin_pos:
                                val = random.randint(2, 3)
                            else:
                                val = random.randint(1, 3)
                            html += f"<td>{val}</td>"
                            avg[j] = avg[j] + val
                            avg_count[j] = avg_count[j] + 1
                        else:
                            html += "<td>-</td>"

                    for j in range(0, 3):
                        val = random.randint(1, 3)
                        html += f"<td>{val}</td>"
                        pso_avg[j] = pso_avg[j] + val

                    html += "</tr>"

                html += f"<tr><td>CO{course_code_number}</td>"
                avg_new = []
                for t in range(len(avg)):
                    val = 0
                    if avg_count[t] > 0:
                        val = round(avg[t] / avg_count[t], 2)
                    avg_new.append(val)
                pso_avg_new = [round(x / (len(rows) - 2), 2) for x in pso_avg]

                for j in range(0, 12):
                    if avg_new[j] > 0:
                        html += f"<td>{avg_new[j]}</td>"
                    else:
                        html += "<td>-</td>"

                for j in range(0, 3):
                    if pso_avg_new[j] > 0:
                        html += f"<td>{pso_avg_new[j]}</td>"
                    else:
                        html += "<td>-</td>"
                html += "</tr></table><br/><br/>"

        # Write HTML String to file.html
        with open(f"output/{filename}.html", "w") as file:
            file.write(html)

        driver.quit()
    else:
        print("Not docx file")
