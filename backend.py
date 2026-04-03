from fastapi import FastAPI, UploadFile, File, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, HTMLResponse
import tempfile
import os
import sys
import base64
import uvicorn
import pandas as pd
from io import BytesIO

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from src.parser import TaskParser
from src.excel_exporter import ExcelExporter
from src.summarizer import TaskSummarizer
from src.google_sheets import GoogleSheetsExporter
from src.google_calendar import GoogleCalendarExporter

app = FastAPI(title="PDF Task Parser API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMP_DIR = "temp_uploads"
os.makedirs(TEMP_DIR, exist_ok=True)

summarizer = None

def get_summarizer():
    global summarizer
    if summarizer is None:
        try:
            print("🔄 Загрузка суммаризатора...")
            summarizer = TaskSummarizer()
            print("✅ Суммаризатор загружен")
        except Exception as e:
            print(f"⚠️ Суммаризатор не загружен: {e}")
            summarizer = False
    return summarizer if summarizer is not False else None


@app.get("/")
async def root():
    return {"message": "PDF Task Parser API", "status": "running"}

@app.get("/app")
async def get_app():
    html_path = os.path.join("web", "index.html")
    if os.path.exists(html_path):
        with open(html_path, "r", encoding="utf-8") as f:
            return HTMLResponse(content=f.read())
    return HTMLResponse(content="<h1>index.html not found</h1>", status_code=404)

@app.get("/style.css")
async def get_css():
    with open("web/style.css", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read(), media_type="text/css")

@app.get("/script.js")
async def get_js():
    with open("web/script.js", "r", encoding="utf-8") as f:
        return HTMLResponse(content=f.read(), media_type="application/javascript")
    
@app.get("/web/icons/{icon_name}")
async def get_icon(icon_name: str):
    icon_path = os.path.join("web", "icons", icon_name)
    if os.path.exists(icon_path):
        with open(icon_path, "rb") as f:
            return HTMLResponse(content=f.read(), media_type="image/svg+xml")
    return HTMLResponse(status_code=404)


@app.post("/parse-batch")
async def parse_batch(request: Request):
    print("\n" + "="*60)
    print("🔍 ПОЛУЧЕН ЗАПРОС НА ПАРСИНГ")
    print("="*60)
    
    form = await request.form()
    
    files = form.getlist("files")
    export_to_sheets = form.get("export_to_sheets", "false").lower() == "true"
    export_to_calendar = form.get("export_to_calendar", "false").lower() == "true"
    sheets_url = form.get("sheets_url", "")
    calendar_id = form.get("calendar_id", "")
    
    print(f"📄 Файлов: {len(files)}")
    print(f"📊 Экспорт в Sheets: {export_to_sheets}")
    print(f"📅 Экспорт в Calendar: {export_to_calendar}")
    print(f"🔗 URL Sheets: {sheets_url}")
    print(f"📆 ID Calendar: {calendar_id}")
    print("="*60 + "\n")
    
    all_results = []
    all_tasks_data = []
    sheets_export_status = None
    calendar_export_status = None
    
    all_dfs = []
    
    for file in files:
        file_ext = os.path.splitext(file.filename)[1].lower()
        with tempfile.NamedTemporaryFile(delete=False, suffix=file_ext) as tmp:
            content = await file.read()
            tmp.write(content)
            tmp_path = tmp.name
        
        try:
            parser = TaskParser(tmp_path)
            text = parser.extract_text()
            tasks = parser.parse_tasks(text)
            
            if tasks:
                summarizer = get_summarizer()
                for task in tasks:
                    task['source'] = file.filename
                    if summarizer:
                        try:
                            task['summary'] = summarizer.summarize(task['full_description'])
                        except Exception:
                            task['summary'] = task['full_description'][:100] + "..."
                    else:
                        task['summary'] = task['full_description'][:100] + "..."
                    all_tasks_data.append(task)
                
                df_data = []
                for task in tasks:
                    df_data.append({
                        '№': task['number'],
                        'Краткое описание': task.get('summary', ''),
                        'Описание': task['full_description'],
                        'Ответственный': task.get('responsible', ''),
                        'Срок': task.get('due_date_str', '')
                    })
                df = pd.DataFrame(df_data)
                all_dfs.append({
                    "df": df,
                    "filename": file.filename,
                    "tasks": tasks
                })
                
                all_results.append({
                    "filename": file.filename,
                    "tasks": tasks,
                    "count": len(tasks)
                })
            
            os.remove(tmp_path)
            
        except Exception as e:
            print(f"Ошибка в {file.filename}: {e}")
            if os.path.exists(tmp_path):
                os.remove(tmp_path)
    
    if not all_tasks_data:
        return JSONResponse({
            "success": False,
            "error": "Задачи не найдены ни в одном файле"
        })
    
    # ===== ЭКСПОРТ В GOOGLE SHEETS =====
    if export_to_sheets and sheets_url:
        try:
            print("📊 Экспорт в Google Sheets...")
            sheets_exporter = GoogleSheetsExporter()
            if sheets_exporter.use_existing_spreadsheet(sheets_url):
                for item in all_dfs:
                    sheet_name = os.path.splitext(item['filename'])[0][:30]
                    sheet_name = sheet_name.replace(' ', '_').replace('/', '_')
                    sheets_exporter.export_dataframe(item['df'], sheet_name)
                sheets_export_status = "success"
                print("✅ Экспорт в Google Sheets выполнен")
            else:
                sheets_export_status = "error: таблица не найдена"
                print("❌ Таблица не найдена")
        except Exception as e:
            sheets_export_status = f"error: {str(e)}"
            print(f"❌ Ошибка Sheets: {e}")
    
    # ===== ЭКСПОРТ В GOOGLE CALENDAR =====
    if export_to_calendar and calendar_id:
        try:
            print(f"📅 Экспорт в Google Calendar...")
            print(f"   ID календаря: {calendar_id}")
            print(f"   Количество задач: {len(all_tasks_data)}")
            
            calendar_exporter = GoogleCalendarExporter(calendar_id=calendar_id)
            calendar_exporter.create_events_from_tasks(all_tasks_data)
            calendar_export_status = "success"
            print("✅ Экспорт в Google Calendar выполнен")
        except Exception as e:
            calendar_export_status = f"error: {str(e)}"
            print(f"❌ Ошибка Calendar: {e}")
    else:
        print(f"⚠️ Экспорт в Calendar пропущен: export_to_calendar={export_to_calendar}, calendar_id={calendar_id}")
    
    # ===== СОЗДАЁМ EXCEL =====
    exporter = ExcelExporter()
    
    for item in all_dfs:
        sheet_name = os.path.splitext(item['filename'])[0][:30]
        sheet_name = sheet_name.replace(' ', '_').replace('/', '_').replace('\\', '_')
        exporter.add_sheet(item['df'], sheet_name)
    
    all_df_data = []
    for task in all_tasks_data:
        all_df_data.append({
            'Источник': task.get('source', ''),
            '№': task['number'],
            'Краткое описание': task.get('summary', ''),
            'Описание': task['full_description'],
            'Ответственный': task.get('responsible', ''),
            'Срок': task.get('due_date_str', '')
        })
    all_df = pd.DataFrame(all_df_data)
    exporter.add_sheet(all_df, "Все задачи")
    
    excel_buffer = BytesIO()
    exporter.save_to_buffer(excel_buffer)
    excel_bytes = excel_buffer.getvalue()
    excel_base64 = base64.b64encode(excel_bytes).decode('ascii')
    
    total_stats = {
        "total": len(all_tasks_data),
        "with_responsible": sum(1 for t in all_tasks_data if t.get('responsible')),
        "with_date": sum(1 for t in all_tasks_data if t.get('due_date_str')),
        "files_count": len(all_results)
    }
    
    return {
        "success": True,
        "tasks": all_tasks_data,
        "statistics": total_stats,
        "excel_base64": excel_base64,
        "files": [{"name": r["filename"], "count": r["count"]} for r in all_results],
        "sheets_export": sheets_export_status,
        "calendar_export": calendar_export_status
    }


if __name__ == "__main__":
    print("🚀 Запуск PDF Task Parser API")
    uvicorn.run(app, host="0.0.0.0", port=8000)