from docx import Document

doc = Document('./sample.docx')

for i, p in enumerate(doc.paragraphs):
    # print(i, p.text)
    for run in p.runs:
        if run.font.name != "Meiryo UI":
            print(i, p.text)
            print(run.font.name)
            print(run.font.size / 12700)
