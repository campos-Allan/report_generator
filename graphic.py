import win32com.client
import datetime
import os
import time

def get_current_year_month():
    now = datetime.datetime.now()
    year = now.year
    month = f"{now.month:02d}"
    return year, month

# Paths
query_file = r"path to excel with query \query.xlsx"
year, month = get_current_year_month()
destination_file = fr"path to the monthly spreadsheet that generates the graphic\SHEET_{year}_{month}.xlsx"

os.system("taskkill /f /im excel.exe")

excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = True

try:
    print(f"📂 Opening query workbook: {query_file}")
    wb_query = excel.Workbooks.Open(query_file)
    ws_query = wb_query.Sheets('bunker')

    print("🔄 Refreshing queries with RefreshAll...")
    wb_query.RefreshAll()

    print("⏳ Waiting for async queries to finish...")
    excel.CalculateUntilAsyncQueriesDone()
    print("✅ Queries refreshed.")

    last_row = ws_query.Cells(ws_query.Rows.Count, 1).End(-4162).Row
    print(f"📌 Last data row: {last_row}")

    ws_query.Activate()
    ws_query.Range(f"A2:H{last_row}").Select()  
    ws_query.Range(f"A2:H{last_row}").Copy() # Copying query data (A1:H[last_row])

    print(f"📂 Opening destination workbook: {destination_file}")
    wb_dest = excel.Workbooks.Open(destination_file)
    ws_dest = wb_dest.Sheets("Graphic Sheet")

    ws_dest.Range("A18").PasteSpecial(Paste=-4163) # Paste data starting here to generate graphic
    print("✅ Data pasted.")

    wb_dest.Save()

    ws_dest.Activate()
    ws_dest.Range("A1:H16").CopyPicture(Format=2) #Graphic is in these cells

    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True
    presentation = ppt.Presentations.Add()

    slide = presentation.Slides.Add(1, 12)

    print("📋 Pasting chart into PowerPoint...")
    slide.Shapes.Paste()

    for _ in range(5):
        if slide.Shapes.Count > 0:
            break
        print("⏳ Waiting for chart to paste...")
        time.sleep(1)
    else:
        raise Exception("❌ Failed to paste chart.")

    shape = slide.Shapes(1)

    output_image_path = r"path to graphic\grafico.png" #this is very important, changes to grafico.png is the trigger to activate a Power Automate flow
    # these pdf reports need to be made daily, as the graphic, but the graphic doesn't need to be stored like the pdf, so we just overwrite the file in each run

    print(f"💾 Exporting image to {output_image_path}...")
    shape.Export(output_image_path, 2)

    presentation.Close()
    ppt.Quit()
    del ppt

    print("✅ Chart exported successfully.")

except Exception as e:
    print(f"❌ Error during process: {e}")

finally:
    try:
        wb_query.Close(SaveChanges=False)
    except:
        pass
    try:
        wb_dest.Close(SaveChanges=True)
    except:
        pass
    excel.Quit()
    del excel

print("🟢 Process completed.")
