
from nltk.probability import FreqDist
from docx.shared import RGBColor
import docx

doc = docx.Document('lion.docx')
text = []
for paragraph in doc.paragraphs:
    text.append(paragraph.text)
print('\n'.join(text))



