from pathlib import Path

from openpyxl import load_workbook
import docx


def parse_sheets(xlsx):
    wb = load_workbook(xlsx)
    columns = [col for col in wb.active.iter_cols()]
    header, data = [c.value for c in columns[1]], [[c.value for c in col] for col in columns[2:]]
    organized = {}
    for col in data:
        question = col[0]
        answers = {}
        for i in range(len(col)):
            if i == 0:
                continue
            answers[header[i]] = col[i]
        organized[question] = answers

    return organized


def export_docx(data, out_file):
    doc = docx.Document()
    for question, answers in data.items():
        doc.add_heading(question, level=1)
        for answerer, answer in answers.items():
            if not answer and not answerer:
                continue
            par = doc.add_paragraph('')
            par.add_run(answerer, 'Emphasis').font.bold = True
            par.add_run(' — ')
            par.add_run(answer)
    doc.save(out_file)


if __name__ == '__main__':
    in_path = 'input'
    out_path = 'output'
    for f in Path(in_path).glob('*.xlsx'):
        print(f'parsing {f}')
        parsed = parse_sheets(f)
        out_file = Path(out_path) / (f.stem + '.docx')
        export_docx(parsed, out_file)
