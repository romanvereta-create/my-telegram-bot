import datetime
import os
import shutil
from telegram import Update
from telegram.ext import Application, MessageHandler, filters
from fpdf import FPDF
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# ======================== ШРИФТЫ ========================

class PDF(FPDF):
    def __init__(self):
        super().__init__()
        self.add_font('DejaVu', '', 'DejaVuSansCondensed.ttf', uni=True)
        self.add_font('DejaVu', 'B', 'DejaVuSansCondensed-Bold.ttf', uni=True)

# ======================== НАСТРОЙКИ ИЗ EXCEL ========================

def load_settings():
    default_settings = {
        "company_name": "ИП Иванов И.И.",
        "inn": "1234567890",
        "ogrnip": "312345678901234",
        "address": "г. Москва, Волгоградский пр-т, 42, к. 9",
        "phone": "+7 (495) 123-45-67",
        "service_name": "Услуга",
        "tax_system": "ПСН",
        "email_sender": "ivanov@gmail.com",
        "thanks_text": "СПАСИБО ЗА ОПЛАТУ!",
        "website": "https://ваш-сайт.рф",
        "bank_name": "Т-Банк (АО)",
        "bik": "044525974",
        "account_number": "40802810500000012345",
        "corr_account": "30101810145250000974",
        "recipient": "ИП Иванов И.И."
    }
    
    settings = default_settings.copy()
    
    try:
        if os.path.exists('settings.xlsx'):
            wb = openpyxl.load_workbook('settings.xlsx')
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    key = str(row[0]).strip()
                    value = str(row[1]).strip()
                    settings[key] = value
            print("✅ Настройки загружены из settings.xlsx")
        else:
            # Создаём шаблон settings.xlsx
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Настройки"
            headers = ["Параметр", "Значение"]
            for col, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col, value=header)
                cell.font = Font(bold=True)
            row = 2
            for key, value in default_settings.items():
                ws.cell(row=row, column=1, value=key)
                ws.cell(row=row, column=2, value=value)
                row += 1
            ws.column_dimensions['A'].width = 25
            ws.column_dimensions['B'].width = 45
            wb.save('settings.xlsx')
            print("📄 Создан шаблон settings.xlsx (заполните его)")
    except Exception as e:
        print(f"⚠️ Ошибка загрузки settings.xlsx: {e}")
    
    return settings

# ======================== EXCEL КНИГА УЧЁТА ========================

def init_excel():
    if not os.path.exists('book.xlsx'):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Книга учёта"
        headers = ["№ п/п", "Дата и время", "№ Квитанции", "Клиент", "доходы, руб", "Статус"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 15
        ws.column_dimensions['F'].width = 12
        wb.save('book.xlsx')
        print("📄 Создан файл book.xlsx")

def add_to_excel(client_name, amount, receipt_number):
    try:
        init_excel()
        wb = openpyxl.load_workbook('book.xlsx')
        ws = wb.active
        row_num = ws.max_row + 1
        ws.cell(row=row_num, column=1, value=row_num - 1)
        ws.cell(row=row_num, column=2, value=datetime.datetime.now().strftime('%d.%m.%Y %H:%M'))
        ws.cell(row=row_num, column=3, value=receipt_number)
        ws.cell(row=row_num, column=4, value=client_name)
        ws.cell(row=row_num, column=5, value=amount)
        ws.cell(row=row_num, column=6, value="Оплачено")
        for col in range(1, 7):
            ws.cell(row=row_num, column=col).alignment = Alignment(horizontal='center')
        wb.save('book.xlsx')
        print(f"📝 Добавлена запись: {client_name} | {amount} руб. | №{receipt_number}")
    except PermissionError:
        print("⚠️ Ошибка: файл book.xlsx открыт в Excel! Закройте его и попробуйте снова.")
        if os.path.exists('book.xlsx'):
            shutil.copy('book.xlsx', 'book_backup.xlsx')
            print("📋 Создана резервная копия book_backup.xlsx")
        try:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Книга учёта"
            headers = ["№ п/п", "Дата и время", "№ Квитанции", "Клиент", "доходы, руб", "Статус"]
            for col, header in enumerate(headers, 1):
                ws.cell(row=1, column=col, value=header).font = Font(bold=True)
            row_num = 2
            ws.cell(row=row_num, column=1, value=1)
            ws.cell(row=row_num, column=2, value=datetime.datetime.now().strftime('%d.%m.%Y %H:%M'))
            ws.cell(row=row_num, column=3, value=receipt_number)
            ws.cell(row=row_num, column=4, value=client_name)
            ws.cell(row=row_num, column=5, value=amount)
            ws.cell(row=row_num, column=6, value="Оплачено")
            wb.save('book_new.xlsx')
            print("📋 Создан новый файл book_new.xlsx (закройте book.xlsx и переименуйте)")
        except:
            print("❌ Не удалось сохранить Excel-файл. Проверьте права доступа к папке.")
    except Exception as e:
        print(f"⚠️ Ошибка записи в Excel: {e}")

# ======================== ЗАГРУЗКА КЛИЕНТОВ ИЗ EXCEL ========================

def load_clients():
    clients = {}
    try:
        if os.path.exists('clients.xlsx'):
            wb = openpyxl.load_workbook('clients.xlsx')
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] and row[1]:
                    name = str(row[0]).strip()
                    try:
                        tid = int(row[1])
                        clients[name.lower()] = tid
                    except:
                        pass
            print(f"📂 Загружено {len(clients)} клиентов из clients.xlsx")
        else:
            # Создаём шаблон clients.xlsx
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Клиенты"
            headers = ["Имя", "Telegram ID", "Телефон", "Примечание"]
            for col, header in enumerate(headers, 1):
                ws.cell(row=1, column=col, value=header).font = Font(bold=True)
            ws.column_dimensions['A'].width = 20
            ws.column_dimensions['B'].width = 15
            ws.column_dimensions['C'].width = 15
            ws.column_dimensions['D'].width = 20
            wb.save('clients.xlsx')
            print("📄 Создан шаблон clients.xlsx (заполните его клиентами)")
    except Exception as e:
        print(f"⚠️ Ошибка загрузки clients.xlsx: {e}")
    return clients

# ======================== ГЕНЕРАЦИЯ PDF ========================

def generate_pdf(settings, client_name, amount, user_id):
    pdf = PDF()
    pdf.add_page()
    pdf.set_left_margin(15)
    pdf.set_right_margin(15)
    
    # Рамка вокруг чека
    pdf.rect(10, 10, 190, 277)
    
    # ЛОГОТИП (если есть)
    if os.path.exists('logo.png'):
        try:
            pdf.image('logo.png', x=80, y=15, w=50)
            pdf.ln(25)
        except:
            pdf.ln(5)
    else:
        pdf.ln(5)
    
    # ВОДЯНОЙ ЗНАК "ОПЛАЧЕНО"
    pdf.set_font('DejaVu', 'B', 50)
    pdf.set_text_color(200, 200, 200)
    pdf.rotate(45)
    pdf.set_xy(40, 100)
    pdf.cell(0, 0, "ОПЛАЧЕНО", align='C')
    pdf.rotate(0)
    pdf.set_text_color(0, 0, 0)
    
    # ВЕРХНЯЯ ЧАСТЬ
    pdf.set_font('DejaVu', 'B', 14)
    pdf.cell(0, 10, txt="═" * 52, ln=True, align='C')
    pdf.set_font('DejaVu', 'B', 16)
    pdf.cell(0, 12, txt="КАССОВЫЙ ЧЕК. ПРИХОД", ln=True, align='C')
    pdf.set_font('DejaVu', 'B', 14)
    pdf.cell(0, 10, txt="═" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # НОМЕР ЧЕКА
    now = datetime.datetime.now()
    receipt_number = f"{now.strftime('%d%m')}-{now.strftime('%H%M%S')}"
    pdf.set_font('DejaVu', '', 11)
    pdf.cell(0, 7, txt=f"Чек №: {receipt_number}", ln=True)
    pdf.cell(0, 7, txt=f"{now.strftime('%d.%m.%Y %H:%M')}", ln=True)
    pdf.ln(3)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # ДАННЫЕ ИП
    pdf.set_font('DejaVu', 'B', 12)
    pdf.cell(0, 8, txt=settings["company_name"], ln=True)
    pdf.set_font('DejaVu', '', 11)
    pdf.cell(0, 7, txt=f"ИНН: {settings['inn']}", ln=True)
    if settings.get('ogrnip'):
        pdf.cell(0, 7, txt=f"ОГРНИП: {settings['ogrnip']}", ln=True)
    pdf.cell(0, 7, txt=settings["address"], ln=True)
    if settings.get('phone'):
        pdf.cell(0, 7, txt=f"Тел.: {settings['phone']}", ln=True)
    pdf.ln(3)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # ТАБЛИЦА УСЛУГ
    pdf.set_font('DejaVu', 'B', 10)
    pdf.cell(70, 8, "Наименование", border=1, align='C')
    pdf.cell(35, 8, "Цена за ед.", border=1, align='C')
    pdf.cell(25, 8, "Кол.", border=1, align='C')
    pdf.cell(40, 8, "Сумма", border=1, align='C')
    pdf.ln()
    
    pdf.set_font('DejaVu', '', 10)
    pdf.cell(70, 8, settings['service_name'], border=1)
    pdf.cell(35, 8, f"{amount:.2f}", border=1, align='R')
    pdf.cell(25, 8, "1", border=1, align='C')
    pdf.cell(40, 8, f"{amount:.2f}", border=1, align='R')
    pdf.ln()
    pdf.ln(3)
    
    # ИТОГИ
    pdf.set_font('DejaVu', '', 10)
    pdf.cell(110, 7, txt="СУММА БЕЗ НДС", border=0)
    pdf.cell(40, 7, txt=f"{amount:.2f}", border=0, align='R')
    pdf.ln()
    pdf.set_font('DejaVu', 'B', 11)
    pdf.cell(110, 8, txt="ИТОГО:", border=0)
    pdf.cell(40, 8, txt=f"{amount:.2f}", border=0, align='R')
    pdf.ln()
    pdf.set_font('DejaVu', '', 10)
    pdf.cell(110, 7, txt="Безналичными", border=0)
    pdf.cell(40, 7, txt=f"{amount:.2f}", border=0, align='R')
    pdf.ln(6)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # НАЛОГОВАЯ ИНФОРМАЦИЯ
    pdf.set_font('DejaVu', '', 11)
    pdf.cell(100, 7, txt="Признак расчета в «Интернет»", border=0)
    pdf.cell(40, 7, txt="Да", border=0, align='R')
    pdf.ln()
    pdf.cell(100, 7, txt="Применяемая система", border=0)
    pdf.cell(40, 7, txt=settings["tax_system"], border=0, align='R')
    pdf.ln()
    pdf.cell(100, 7, txt="налогообложения", border=0)
    pdf.ln(6)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # БАНКОВСКИЕ РЕКВИЗИТЫ
    pdf.set_font('DejaVu', 'B', 10)
    pdf.cell(0, 7, txt="БАНКОВСКИЕ РЕКВИЗИТЫ:", ln=True)
    pdf.set_font('DejaVu', '', 9)
    if settings.get('bank_name'):
        pdf.cell(0, 6, txt=f"Банк: {settings['bank_name']}", ln=True)
    if settings.get('bik'):
        pdf.cell(0, 6, txt=f"БИК: {settings['bik']}", ln=True)
    if settings.get('account_number'):
        pdf.cell(0, 6, txt=f"Р/с: {settings['account_number']}", ln=True)
    if settings.get('corr_account'):
        pdf.cell(0, 6, txt=f"К/с: {settings['corr_account']}", ln=True)
    if settings.get('recipient'):
        pdf.cell(0, 6, txt=f"Получатель: {settings['recipient']}", ln=True)
    pdf.ln(3)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # КОНТАКТЫ
    pdf.set_font('DejaVu', '', 10)
    pdf.cell(0, 7, txt=f"Эл. почта: {settings['email_sender']}", ln=True)
    pdf.ln(3)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # QR-КОД (если есть)
    if os.path.exists('qrcode.png'):
        try:
            pdf.image('qrcode.png', x=80, y=240, w=40)
            pdf.ln(20)
        except:
            pdf.ln(5)
    
    if settings.get('website'):
        pdf.set_font('DejaVu', '', 10)
        pdf.cell(0, 7, txt=f"Сайт: {settings['website']}", ln=True, align='C')
    pdf.ln(3)
    
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="─" * 52, ln=True, align='C')
    pdf.ln(3)
    
    # НИЖНЯЯ ЧАСТЬ (СПАСИБО + ПОДПИСЬ)
    pdf.set_font('DejaVu', 'B', 13)
    pdf.cell(0, 10, txt=settings["thanks_text"], ln=True, align='C')
    pdf.ln(3)
    pdf.set_font('DejaVu', '', 9)
    pdf.cell(0, 6, txt="Сайт ФНС: nalog.gov.ru", ln=True, align='C')
    pdf.ln(5)
    
    # ПОДПИСЬ (из картинки или текстом)
    if os.path.exists('signature.png'):
        try:
            pdf.image('signature.png', x=45, y=255, w=120)
        except:
            pdf.set_font('DejaVu', '', 10)
            pdf.cell(0, 7, txt="____________________", ln=True, align='C')
            pdf.cell(0, 7, txt="(подпись ИП)                М.П.", ln=True, align='C')
    else:
        pdf.set_font('DejaVu', '', 10)
        pdf.cell(0, 7, txt="____________________", ln=True, align='C')
        pdf.cell(0, 7, txt="(подпись ИП)                М.П.", ln=True, align='C')
    
    # Сохраняем PDF
    filename = f"check_{user_id}_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.pdf"
    pdf.output(filename)
    return filename, receipt_number

# ======================== ОБРАБОТКА ========================

async def handle_message(update, context):
    text = update.message.text
    user_id = update.effective_user.id

    if text.startswith('/start'):
        with open('clients_auto.txt', 'a', encoding='utf-8') as f:
            f.write(f"{update.effective_user.first_name} (@{update.effective_user.username}) = {user_id}\n")
        await update.message.reply_text("✅ Вы зарегистрированы!")
        return

    if text.lower().startswith('чек'):
        rest = text[4:].strip()
        last_space = rest.rfind(' ')
        if last_space == -1:
            await update.message.reply_text("❗️ Напишите: чек Имя Сумма")
            return

        client_name = rest[:last_space].strip()
        amount_str = rest[last_space:].strip()

        try:
            amount = float(amount_str.replace(',', '.'))
        except:
            await update.message.reply_text("❌ Сумма должна быть числом")
            return

        settings = load_settings()
        pdf_file = None
        try:
            pdf_file, receipt_number = generate_pdf(settings, client_name, amount, user_id)
            add_to_excel(client_name, amount, receipt_number)

            # Отправка чека вам
            with open(pdf_file, 'rb') as f:
                await update.message.reply_document(document=f, filename=pdf_file)
            await update.message.reply_text(f"✅ Чек для {client_name} на {amount} руб. №{receipt_number}")

            # Отправка клиенту
            clients = load_clients()
            client_id = clients.get(client_name.lower())

            if client_id:
                try:
                    with open(pdf_file, 'rb') as f:
                        await context.bot.send_document(chat_id=client_id, document=f, filename=pdf_file)
                    await update.message.reply_text(f"📤 Чек отправлен клиенту {client_name}")
                except Exception as e:
                    await update.message.reply_text(f"⚠️ Ошибка отправки клиенту: {e}")
            else:
                await update.message.reply_text(f"⚠️ Клиент '{client_name}' не найден в списке. Чек отправлен только вам. Перешлите его клиенту вручную.")

        except Exception as e:
            await update.message.reply_text(f"❌ Ошибка: {e}")
        finally:
            if pdf_file and os.path.exists(pdf_file):
                try:
                    os.remove(pdf_file)
                except:
                    pass

# ======================== ЗАПУСК ========================

def main():
    try:
        with open('token.txt', 'r') as f:
            TOKEN = f.read().strip()
    except:
        print("❌ token.txt не найден!")
        input("Нажмите Enter...")
        return

    if not TOKEN:
        print("❌ Токен пуст!")
        input("Нажмите Enter...")
        return

    init_excel()
    load_settings()  # Создаст settings.xlsx если нет
    load_clients()   # Создаст clients.xlsx если нет

    app = Application.builder().token(TOKEN).build()
    app.add_handler(MessageHandler(filters.TEXT, handle_message))

    print("✅ БОТ ЗАПУЩЕН! Ожидаю сообщения...")
    print("📋 Книга учёта: book.xlsx")
    print("👥 Список клиентов: clients.xlsx")
    print("⚙️ Настройки: settings.xlsx")
    app.run_polling()

if __name__ == '__main__':
    main()