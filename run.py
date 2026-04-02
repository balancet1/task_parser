import os
import sys
from datetime import datetime
from src.main_with_calendar import main

if __name__ == "__main__":
    # Запускаем основную программу
    main()
    
    # После завершения показываем информацию о файле
    print("\n💡 Совет: Для обработки всех файлов из папки data/ используйте:")
    print("   python3 run_all.py")