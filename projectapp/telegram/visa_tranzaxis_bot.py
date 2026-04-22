import re
import pdfplumber
from openpyxl import Workbook
from aiogram import Bot, Dispatcher, types, F
from aiogram.filters import Command
from aiogram.types import FSInputFile
import asyncio
import os
# jhjhjh
# 🔑 TOKEN (BU YERGA O'Z TOKENINGNI QO'Y)
TOKEN = "8307217215:AAE65y0HoGbVcb5CUShi9D_Di1j3vaRBHd8"

bot = Bot(token=TOKEN)
dp = Dispatcher()


# 📥 PDF → TEXT
def extract_text_from_pdf(pdf_path):
    text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            t = page.extract_text()
            if t:
                text += t + "\n"
    return text


# 📊 PARSER (VISA REPORT)
def parse_settlement_report(text):
    if "Settlement Currency Code:" not in text:
        return None

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

            issuer_total = re.search(
                r"Issuer Total\s+Count\s+Interch\. value\s+Reimb\. Fees\s+Net value",
                sb
            )

            if issuer_total:
                lines = sb[issuer_total.end():].strip().splitlines()
                if lines:
                    nums = lines[0].split()
                    if len(nums) >= 3:
                        count = nums[0]
                        interch = nums[1]
                        reimb = nums[2]

            visa_match = re.search(r"Total for VISA charges\s+([-\d.,]+)", sb)
            visa = visa_match.group(1) if visa_match else "0"

            net_match = re.search(r"Net Settlement Amount \(Issuer\)\s+([-\d.,]+)", sb)
            net = net_match.group(1) if net_match else "0"

            if date:
                parsed_data[currency].append([
                    date,
                    count,
                    float(interch.replace(",", ".")) if interch else 0,
                    float(reimb.replace(",", ".")) if reimb else 0,
                    float(visa.replace(",", ".")),
                    float(net.replace(",", "."))
                ])

    return parsed_data


# 📤 EXCEL
def save_to_excel(data, filename="report.xlsx"):
    wb = Workbook()
    ws = wb.active

    headers = [
        "Date", "Count", "Interch", "Reimb",
        "Visa Charges", "Net Settlement", "Total"
    ]
    ws.append(headers)

    for currency, rows in data.items():
        ws.append([])
        ws.append([f"CURRENCY: {currency}"])

        for r in rows:
            total = r[2] + r[3] + r[4] + r[5]
            ws.append([r[0], r[1], r[2], r[3], r[4], r[5], total])

    wb.save(filename)


# 🚀 START
@dp.message(Command("start"))
async def start(message: types.Message):
    await message.answer(
        "📂 VISA settlement PDF yubor\n"
        "Men Excel qilib beraman ✅"
    )


# 📥 PDF HANDLE
@dp.message(F.document)
async def handle_pdf(message: types.Message):
    if not message.document.file_name.endswith(".pdf"):
        await message.answer("❌ Faqat PDF yubor")
        return

    os.makedirs("downloads", exist_ok=True)

    file = await bot.get_file(message.document.file_id)
    path = f"downloads/{message.document.file_name}"

    await bot.download_file(file.file_path, path)

    text = extract_text_from_pdf(path)
    data = parse_settlement_report(text)

    if not data:
        await message.answer("❌ Bu VISA report emas")
        return

    save_to_excel(data)

    await message.answer_document(FSInputFile("report.xlsx"))
    try:
        os.remove(path)
    except:
        pass

        # 🔥 Excelni ham o‘chirish
    try:
        os.remove("output.xlsx")
    except:
        pass


# ❌ BOSHQA NARSA
@dp.message()
async def other(message: types.Message):
    await message.answer("❌ Faqat PDF yubor")


# ▶️ RUN
async def main():
    await dp.start_polling(bot)


if __name__ == "__main__":
    asyncio.run(main())