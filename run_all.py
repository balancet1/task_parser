import os
import glob
import pandas as pd
from datetime import datetime
from src.main_with_calendar import TaskProcessor
from src.excel_exporter import ExcelExporter

def process_all_documents():
    # Настройки
    calendar_id = "cd807be5e82c589d7a3e08ec8dbc93e55578d58603238b6754d2ce4162e12eef@group.calendar.google.com"
    sheets_url = None  # или твоя ссылка
    
    # Найти все файлы
    all_files = []
    for ext in ['*.pdf', '*.docx', '*.doc']:
        all_files.extend(glob.glob(f"data/{ext}"))
    all_files.sort()
    
    if not all_files:
        print("❌ В папке data/ нет файлов")
        return
    
    print(f"\n{'='*60}")
    print(f"📁 Найдено файлов: {len(all_files)}")
    for f in all_files:
        print(f"   📄 {os.path.basename(f)}")
    print(f"{'='*60}\n")
    
    # Единый Excel-файл
    excel_filename = "output/tasks.xlsx"
    exporter = ExcelExporter(excel_filename)
    
    # Список для сводной таблицы
    all_tasks_data = []
    
    for i, file_path in enumerate(all_files, 1):
        file_name = os.path.basename(file_path)
        print(f"\n{'='*60}")
        print(f"📄 Файл {i}/{len(all_files)}: {file_name}")
        print(f"{'='*60}")
        
        # Обрабатываем файл
        processor = TaskProcessor(file_path, use_summarizer=True)
        
        if processor.process():
            # Генерируем имя листа
            sheet_name = os.path.splitext(file_name)[0]
            sheet_name = sheet_name.replace(' ', '_')[:31]
            
            # Добавляем лист в общий Excel-файл
            exporter.add_sheet(processor.df, sheet_name)
            
            # Добавляем данные в сводную таблицу
            df_with_source = processor.df.copy()
            df_with_source['Источник'] = file_name
            all_tasks_data.append(df_with_source)
            
            # Google интеграция
            if sheets_url:
                processor.save_to_google_sheets(sheets_url, sheet_name=sheet_name)
            if calendar_id:
                processor.save_to_google_calendar(calendar_id)
    
    # Создаём сводную таблицу
    if all_tasks_data:
        all_tasks_df = pd.concat(all_tasks_data, ignore_index=True)
        cols = ['Источник'] + [c for c in all_tasks_df.columns if c != 'Источник']
        all_tasks_df = all_tasks_df[cols]
        exporter.add_sheet(all_tasks_df, "Все задачи")
        print(f"\n📊 Создан сводный лист 'Все задачи' ({len(all_tasks_df)} задач)")
    
    # Сохраняем Excel-файл
    exporter.save()
    print(f"\n✅ Все листы сохранены в {excel_filename}")
    print(f"   📑 Листы: {', '.join(exporter.wb.sheetnames)}")

if __name__ == "__main__":
    process_all_documents()