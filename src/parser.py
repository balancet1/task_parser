import pdfplumber
import re
from datetime import datetime
from typing import List, Dict, Optional

class TaskParser:
    def __init__(self, pdf_path: str):
        self.pdf_path = pdf_path
        self.tasks = []
    
    def extract_text(self) -> str:
        """Извлекает весь текст из PDF"""
        full_text = ""
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                for page in pdf.pages:
                    text = page.extract_text()
                    if text:
                        full_text += text + "\n"
            print(f"✅ Извлечено {len(full_text)} символов из PDF")
            return full_text
        except Exception as e:
            print(f"❌ Ошибка при чтении PDF: {e}")
            return ""
    
    def parse_tasks(self, text: str) -> List[Dict]:
        """Находит задачи в тексте (универсальный парсер)"""
        tasks = []
        
        # Разбиваем текст на строки
        lines = text.split('\n')
        
        current_task = None
        current_description = []
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
            
            # Проверяем, начинается ли строка с номера задачи (1., 2., 10., и т.д.)
            task_match = re.match(r'^(\d+)\.\s+(.*)', line)
            
            if task_match:
                # Если уже собирали предыдущую задачу - сохраняем её
                if current_task:
                    self._save_current_task(current_task, current_description)
                    tasks.append(current_task)
                
                # Начинаем новую задачу
                task_num = task_match.group(1)
                task_start = task_match.group(2)
                
                current_task = {
                    'number': int(task_num),
                    'raw_number': task_num,
                    'description_parts': [task_start],
                    'full_description': '',
                    'responsible': '',
                    'due_date': None,
                    'due_date_str': ''
                }
                current_description = [task_start]
            
            elif current_task:
                # Добавляем строку к описанию (в любом случае)
                current_description.append(line)
                current_task['description_parts'].append(line)
                
                # ========== ИЩЕМ ДАТУ ==========
                # Ищем дату в формате ДД.ММ.ГГГГ после слова "Срок"
                date_match = re.search(r'Срок\s*[—–-]?\s*(\d{2}\.\d{2}\.\d{4})', line)
                if date_match:
                    date_str = date_match.group(1).strip()
                    current_task['due_date_str'] = date_str
                    try:
                        current_task['due_date'] = datetime.strptime(date_str, '%d.%m.%Y').date()
                    except ValueError as e:
                        print(f"⚠️ Ошибка парсинга даты {date_str}: {e}")
                
                # ========== ИЩЕМ ОТВЕТСТВЕННОГО ==========
                # Ищем "Отв.:" и всё после до конца строки или до "Срок"
                resp_match = re.search(r'Отв\.:\s*([^С]+?)(?:\s+Срок|$)', line)
                if not resp_match:
                    # Если не нашли с "Срок", ищем просто до конца строки
                    resp_match = re.search(r'Отв\.:\s*([^\n]+)', line)
                
                if resp_match:
                    responsible = resp_match.group(1).strip()
                    # Очищаем от лишних символов
                    responsible = re.sub(r'\s+', ' ', responsible)
                    current_task['responsible'] = responsible
        
        # Сохраняем последнюю задачу
        if current_task:
            self._save_current_task(current_task, current_description)
            tasks.append(current_task)
        
        self.tasks = tasks
        return tasks
    
    def _save_current_task(self, task: Dict, description_lines: List[str]):
        """Формирует полное описание задачи"""
        # Объединяем все строки в одно описание
        full_desc = ' '.join(description_lines)
        # Очищаем от лишних пробелов
        full_desc = re.sub(r'\s+', ' ', full_desc)
        task['full_description'] = full_desc
    
    def print_tasks(self):
        """Выводит найденные задачи в читаемом виде"""
        if not self.tasks:
            print("❌ Задачи не найдены")
            return
        
        print(f"\n📋 Найдено задач: {len(self.tasks)}\n")
        print("=" * 80)
        
        for task in self.tasks:
            print(f"Задача #{task['number']}")
            print(f"📝 Описание: {task['full_description'][:100]}...")
            print(f"👤 Ответственный: {task['responsible'] or '❌ НЕ НАЙДЕН'}")
            print(f"📅 Срок: {task['due_date_str'] or '❌ НЕ НАЙДЕН'}")
            print("-" * 40)
    
    def to_dataframe(self):
        """Конвертирует задачи в pandas DataFrame"""
        import pandas as pd
        
        data = []
        for task in self.tasks:
            data.append({
                '№': task['number'],
                'Описание': task['full_description'],
                'Ответственный': task['responsible'],
                'Срок': task['due_date_str']
            })
        
        df = pd.DataFrame(data)
        return df


# Тестирование
if __name__ == "__main__":
    import sys
    import os
    
    # Путь к PDF файлу (можно передать как аргумент)
    if len(sys.argv) > 1:
        pdf_file = sys.argv[1]
    else:
        pdf_file = "data/tasks.pdf"
    
    if not os.path.exists(pdf_file):
        print(f"❌ Файл не найден: {pdf_file}")
        sys.exit(1)
    
    print(f"\n📄 Тестирование парсера на файле: {pdf_file}")
    print("=" * 60)
    
    # Создаем парсер
    parser = TaskParser(pdf_file)
    
    # Извлекаем текст
    text = parser.extract_text()
    
    if text:
        # Парсим задачи
        tasks = parser.parse_tasks(text)
        
        # Выводим результат
        parser.print_tasks()
        
        # Показываем статистику
        print("\n📊 Статистика:")
        print(f"   Всего задач: {len(tasks)}")
        print(f"   Задач с ответственным: {sum(1 for t in tasks if t['responsible'])}")
        print(f"   Задач с датой: {sum(1 for t in tasks if t['due_date_str'])}")
    else:
        print("❌ Не удалось извлечь текст из PDF")