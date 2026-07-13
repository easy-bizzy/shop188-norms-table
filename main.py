from table_generator import TableGenerator


def main():
    print("🔄 Генерация таблицы норм времени цеха 188...")
    
    generator = TableGenerator()
    generator.create_table()
    generator.save()
    
    print("🎉 Готово!")


if __name__ == "__main__":
    main()
