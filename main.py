from docx import Document

wordDoc = Document('doc.docx')


def only_bold_if_zero(text, bold):
  if bold:
    return '\\textbf{' + text + '}'
  else:
    return text

def convert_to_latex_text(text):
  return text.replace('\n', '\\newline ').replace('&', '\&')

for t_index, table in enumerate(wordDoc.tables):
  if t_index == 0:
    continue
  latex = '\\begin{table}\n' 
  latex += '\\begin{tabular}{|l|p{10cm}|}\n'
  latex += '\\hline\n'
  looking_for_title = False
  title = 'Your caption'
  for r_index, row in enumerate(table.rows):
    for c_index, cell in enumerate(row.cells):
      if cell.text == 'Title':
        looking_for_title = True
      elif looking_for_title:
        title = cell.text
        looking_for_title = False
      latex += only_bold_if_zero(convert_to_latex_text(cell.text), c_index == 0) + \
          (' & ' if c_index < len(row.cells) - 1 else ' \\\\ \n')
    latex += '\\hline\n'
  latex += '\\end{tabular}\n'
  latex += '\\caption{' + title + '}'
  latex += '\\end{table}\n'
  open('table_' + str(t_index) + '.tex', 'w').write(latex)