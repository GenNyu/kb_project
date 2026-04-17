from pdf2image import convert_from_path
import pytesseract
import os

input_file = "input/PCI_44_64.pdf"
output_file = "output/PCI_44_64.txt"

images = convert_from_path(input_file, dpi=300)

all_text = ""
for img in images:
    text = pytesseract.image_to_string(img, lang="eng")
    all_text += text + "\n"

with open(output_file, "w", encoding="utf-8") as f:
    f.write(all_text)

print("Saved:", output_file)