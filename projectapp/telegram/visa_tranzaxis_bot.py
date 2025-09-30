

#
# import re
# import pdfplumber
# from openpyxl import Workbook
# from aiogram import Bot, Dispatcher, types, F
# from aiogram.filters import Command
# from aiogram.types import FSInputFile
# import asyncio
# import os
#
# # 🔑 Token
# TOKEN = "8307217215:AAE65y0HoGbVcb5CUShi9D_Di1j3vaRBHd8"
#
# bot = Bot(token=TOKEN)
# dp = Dispatcher()
#
#
# # 📥 PDF’dan text olish
# def extract_text_from_pdf(pdf_path):
#     text = ""
#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:
#             page_text = page.extract_text()
#             if page_text:
#                 text += page_text + "\n"
#     return text
#
#
# # 📊 Parsing (valyuta bo‘yicha ajratib olish)
# def parse_settlement_report(text):
#     parsed_data = {}  # { "USD": [...], "UZS": [...] }
#
#     currency_blocks = re.split(r"Settlement Currency Code:", text)
#
#     for block in currency_blocks[1:]:
#         # Valyutani aniqlash
#         currency_match = re.match(r"\s*(\w+)", block)
#         currency = currency_match.group(1) if currency_match else "UNKNOWN"
#
#         parsed_data.setdefault(currency, [])
#
#         # Har bir Settl. Date bo‘yicha ma’lumot yig‘amiz
#         settl_blocks = re.split(r"Settl\. Date:", block)
#
#         for sb in settl_blocks[1:]:
#             # Sana
#             date_match = re.match(r"\s*(\d{4}\.\d{2}\.\d{2})", sb)
#             date = date_match.group(1).replace(".", "/") if date_match else ""
#
#             count, interch, reimb = ("", "", "")
#
#             # Issuer Total
#             issuer_key_match = re.search(
#                 r"Issuer Total\s+Count\s+Interch\. value\s+Reimb\. Fees\s+Net value",
#                 sb,
#             )
#             if issuer_key_match:
#                 after_key = sb[issuer_key_match.end():].strip().splitlines()
#                 if after_key:
#                     numbers = after_key[0].split()
#                     if len(numbers) >= 3:
#                         count = numbers[0].replace(",", ".")
#                         interch = numbers[1].replace(",", ".")
#                         reimb = numbers[2].replace(",", ".")
#
#             # Total for VISA charges
#             visa_match = re.search(r"Total for VISA charges\s+([-\d.,]+)", sb)
#             visa_charges = visa_match.group(1).replace(",", ".") if visa_match else ""
#
#             # Net Settlement Amount (Issuer)
#             net_match = re.search(r"Net Settlement Amount \(Issuer\)\s+([-\d.,]+)", sb)
#             net_settlement = net_match.group(1).replace(",", ".") if net_match else ""
#
#             if date:
#                 parsed_data[currency].append([
#                     date, count, interch, reimb, visa_charges, net_settlement
#                 ])
#
#     return parsed_data
#
#
# # 📤 Excel’ga yozish
# def save_to_excel(parsed_data, filename="report.xlsx"):
#     wb = Workbook()
#     ws = wb.active
#
#     headers = [
#         "Sana",
#         "Count",
#         "Interch. value",
#         "Reimb. Fees",
#         "Total for VISA charges",
#         "Net Settlement Amount (Issuer)",
#         "Total value"  # ➕ Yangi ustun
#     ]
#     ws.append(headers)
#
#     # Umumiy total (faqat Interch, Reimb, Visa, Net, Total value uchun)
#     grand_totals = [0, 0, 0, 0, 0]
#
#     for currency, rows in parsed_data.items():
#         ws.append([])  # Bo‘sh qator
#         ws.append([f"Settlement Currency Code: {currency}"])
#
#         subtotal = [0, 0, 0, 0, 0]  # Interch, Reimb, Visa, Net, Total
#
#         for row in rows:
#             # Qiymatlarni floatga aylantiramiz
#             interch = float(row[2]) if row[2] else 0
#             reimb = float(row[3]) if row[3] else 0
#             visa = float(row[4]) if row[4] else 0
#             net = float(row[5]) if row[5] else 0
#
#             # Total value = Interch + Reimb + Visa + Net
#             total_value = interch + reimb + visa + net
#
#             ws.append([row[0], row[1], interch, reimb, visa, net, total_value])
#
#             # Hisoblash
#             subtotal[0] += interch
#             subtotal[1] += reimb
#             subtotal[2] += visa
#             subtotal[3] += net
#             subtotal[4] += total_value
#
#         # Valyuta bo‘yicha subtotal
#         ws.append([
#             "Subtotal",
#             "",
#             subtotal[0],  # Interch
#             subtotal[1],  # Reimb
#             subtotal[2],  # Visa
#             subtotal[3],  # Net
#             subtotal[4],  # Total value
#         ])
#
#         # Grand totalga qo‘shamiz
#         for i in range(5):
#             grand_totals[i] += subtotal[i]
#
#     # Umumiy total
#     ws.append([])
#     ws.append([
#         "Total Value (All Currencies)",
#         "",
#         grand_totals[0],  # Interch
#         grand_totals[1],  # Reimb
#         grand_totals[2],  # Visa
#         grand_totals[3],  # Net
#         grand_totals[4],  # Total value
#     ])
#
#     wb.save(filename)
#
#
# # 🚀 Start komandasi
# @dp.message(Command("start"))
# async def start(message: types.Message):
#     await message.reply(
#         "Menga PDF fayl yuboring 📂\n"
#         "Men undan Issuer Total, VISA charges va Net Settlement Amount ma’lumotlarini Excel’ga chiqarib beraman ✅"
#     )
#
#
# # 📥 PDF qabul qilish
# @dp.message(F.document)
# async def handle_file(message: types.Message):
#     file = await bot.get_file(message.document.file_id)
#     file_path = file.file_path
#     pdf_file = f"downloads/{message.document.file_name}"
#
#     os.makedirs("downloads", exist_ok=True)
#
#     # 📂 PDF yuklab olish
#     await bot.download_file(file_path, pdf_file)
#
#     # 📂 PDF o‘qish va Excel saqlash
#     text = extract_text_from_pdf(pdf_file)
#     parsed_data = parse_settlement_report(text)
#     save_to_excel(parsed_data, "output.xlsx")
#
#     # 📤 Excel qaytarish
#     excel_file = FSInputFile("output.xlsx")
#     await message.reply_document(excel_file)
#
#
# # 🔄 Botni ishga tushirish
# async def main():
#     await dp.start_polling(bot)
#
#
# if __name__ == "__main__":
#     asyncio.run(main())


import re
import pdfplumber
from openpyxl import Workbook
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import FSInputFile
import asyncio
import os

# 🔑 Token
TOKEN = "8307217215:AAE65y0HoGbVcb5CUShi9D_Di1j3vaRBHd8"

bot = Bot(token=TOKEN)
dp = Dispatcher()


# 📥 PDF’dan text olish
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    return text


# 📊 Parsing (Visa settlement report)
def parse_settlement_report(text):
    if "Settlement Currency Code:" not in text or "Settl. Date:" not in text:
        return None  # ⚠️ Bu Visa report emas

    parsed_data = {}
    currency_blocks = re.split(r"Settlement Currency Code:", text)

    for block in currency_blocks[1:]:
        currency_match = re.match(r"\s*(\w+)", block)
        currency = currency_match.group(1) if currency_match else "UNKNOWN"

        parsed_data.setdefault(currency, [])
        settl_blocks = re.split(r"Settl\. Date:", block)

        for sb in settl_blocks[1:]:
            date_match = re.match(r"\s*(\d{4}\.\d{2}\.\d{2})", sb)
            date = date_match.group(1).replace(".", "/") if date_match else ""

            count, interch, reimb = ("", "", "")
            issuer_key_match = re.search(
                r"Issuer Total\s+Count\s+Interch\. value\s+Reimb\. Fees\s+Net value",
                sb,
            )
            if issuer_key_match:
                after_key = sb[issuer_key_match.end():].strip().splitlines()
                if after_key:
                    numbers = after_key[0].split()
                    if len(numbers) >= 3:
                        count = numbers[0].replace(",", ".")
                        interch = numbers[1].replace(",", ".")
                        reimb = numbers[2].replace(",", ".")

            visa_match = re.search(r"Total for VISA charges\s+([-\d.,]+)", sb)
            visa_charges = visa_match.group(1).replace(",", ".") if visa_match else ""

            net_match = re.search(r"Net Settlement Amount \(Issuer\)\s+([-\d.,]+)", sb)
            net_settlement = net_match.group(1).replace(",", ".") if net_match else ""

            if date:
                parsed_data[currency].append([
                    date, count, interch, reimb, visa_charges, net_settlement
                ])

    return parsed_data


# 📤 Excel’ga yozish
def save_to_excel(parsed_data, filename="report.xlsx"):
    wb = Workbook()
    ws = wb.active

    headers = [
        "Sana",
        "Count",
        "Interch. value",
        "Reimb. Fees",
        "Total for VISA charges",
        "Net Settlement Amount (Issuer)",
        "Total value"
    ]
    ws.append(headers)

    grand_totals = [0, 0, 0, 0, 0]

    for currency, rows in parsed_data.items():
        ws.append([])
        ws.append([f"Settlement Currency Code: {currency}"])

        subtotal = [0, 0, 0, 0, 0]

        for row in rows:
            interch = float(row[2]) if row[2] else 0
            reimb = float(row[3]) if row[3] else 0
            visa = float(row[4]) if row[4] else 0
            net = float(row[5]) if row[5] else 0
            total_value = interch + reimb + visa + net

            ws.append([row[0], row[1], interch, reimb, visa, net, total_value])

            subtotal[0] += interch
            subtotal[1] += reimb
            subtotal[2] += visa
            subtotal[3] += net
            subtotal[4] += total_value

        ws.append([
            "Subtotal",
            "",
            subtotal[0],
            subtotal[1],
            subtotal[2],
            subtotal[3],
            subtotal[4],
        ])

        for i in range(5):
            grand_totals[i] += subtotal[i]

    ws.append([])
    ws.append([
        "Total Value (All Currencies)",
        "",
        grand_totals[0],
        grand_totals[1],
        grand_totals[2],
        grand_totals[3],
        grand_totals[4],
    ])

    wb.save(filename)


# 🚀 Start komandasi
@dp.message(Command("start"))
async def start(message: types.Message):
    await message.reply(
        "Menga faqat VISA settlement report PDF yuboring 📂\n"
        "Men undan ma’lumotlarni Excel’ga chiqarib beraman ✅"
    )


# 📥 Faqat PDF fayllarni tekshirish
@dp.message(F.document)
async def handle_file(message: types.Message):
    if not message.document.file_name.lower().endswith(".pdf"):
        await message.reply("❌ Faqat PDF fayl yuborishingiz mumkin!")
        return

    file = await bot.get_file(message.document.file_id)
    file_path = file.file_path
    pdf_file = f"downloads/{message.document.file_name}"

    os.makedirs("downloads", exist_ok=True)
    await bot.download_file(file_path, pdf_file)

    text = extract_text_from_pdf(pdf_file)
    parsed_data = parse_settlement_report(text)

    if not parsed_data:
        await message.reply("❌ Bu VISA settlement report PDF emas. Iltimos, to‘g‘ri fayl yuboring.")
        return

    save_to_excel(parsed_data, "output.xlsx")
    excel_file = FSInputFile("output.xlsx")
    await message.reply_document(excel_file)


# 📥 Oddiy text / rasm / ovoz uchun filter
@dp.message()
async def handle_other(message: types.Message):
    await message.reply("❌ Bu turdagi faylni qabul qilmayman. Faqat VISA settlement report PDF yuboring.")


# 🔄 Botni ishga tushirish
async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())
