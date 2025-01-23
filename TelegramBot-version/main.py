# 7708422693:AAHNQPcztd8B_Mj4IjOtQ8VDdC-rWj3dkd0
import telebot
from easyconvert.converter import convert_image  # Припускаємо, що ви створили функцію конвертації
import os

# Токен вашого бота
TOKEN = 'ВАШ_ТОКЕН'
bot = telebot.TeleBot(TOKEN)


# Функція обробки команд
@bot.message_handler(commands=['start', 'help'])
def send_welcome(message):
    bot.reply_to(message, "Привіт! Надішліть зображення, і я перетворю його у формат .ico або інший на ваш вибір.")


# Обробка отримання файлів
@bot.message_handler(content_types=['document'])
def handle_docs(message):
    try:
        # Отримуємо файл
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        input_path = message.document.file_name
        with open(input_path, 'wb') as new_file:
            new_file.write(downloaded_file)

        # Конвертація
        output_format = 'ICO'  # Ви можете змінювати формат на інший
        output_path = f"{os.path.splitext(input_path)[0]}.{output_format.lower()}"

        if convert_image(input_path, output_path, output_format):
            with open(output_path, 'rb') as out_file:
                bot.send_document(message.chat.id, out_file)
            os.remove(output_path)  # Видаляємо тимчасовий файл після надсилання
        else:
            bot.reply_to(message, "Не вдалося конвертувати файл.")

        os.remove(input_path)  # Видаляємо вхідний файл
    except Exception as e:
        bot.reply_to(message, f"Сталася помилка: {e}")


# Запуск бота
bot.polling()
