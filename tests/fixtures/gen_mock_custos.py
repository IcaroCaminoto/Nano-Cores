"""Gera tests/fixtures/mock_Cerulean_DB_Custos.xlsx: base SINTETICA para testar a spec-009.
Uso: python3 tests/fixtures/gen_mock_custos.py   (o .xlsx e ignorado pelo git; este script o recria)
Nada aqui vem de dados reais. Casos de borda: custo em branco, negativo, texto, SKU duplicado,
SKU com apostrofo, abas que nao sao pacotes e codigo numerico x texto."""
import sys, zipfile, os
from xml.sax.saxutils import escape

out = sys.argv[1] if len(sys.argv) > 1 else os.path.join(os.path.dirname(os.path.abspath(__file__)), "mock_Cerulean_DB_Custos.xlsx")
os.makedirs(os.path.dirname(out), exist_ok=True)

def col(n):
    s = ''
    while n: n, r = divmod(n - 1, 26); s = chr(65 + r) + s
    return s

HDR_MM = ['Cod_Produto', 'Descricao_Produto', 'Custo_Medio_SKU', 'Custo_Medio_Familia',
          'Qtd_Transacoes_SKU', 'Janela_Dias_Base', 'Data_Processamento']

def linha_mm(sku, familia, janela, data, decoy=999.99):
    return [str(sku), f'SINTETICO {sku}', decoy, familia, 3, janela, data]

mm90 = [HDR_MM,
        linha_mm(90001, 12.50, '90D', '2026-11-01 12:00:00'),
        linha_mm(90002, 7.35, '90D', '2026-11-01 12:00:00'),
        linha_mm(90003, None, '90D', '2026-11-01 12:00:00'),     # custo em branco
        linha_mm(90004, -3.20, '90D', '2026-11-01 12:00:00'),    # negativo
        linha_mm(90005, 'abc', '90D', '2026-11-01 12:00:00'),    # texto
        linha_mm(90006, 100, '90D', '2026-11-01 12:00:00'),
        linha_mm(90007, 20.00, '90D', '2026-11-01 12:00:00'),    # duplicado: vale a 1a
        linha_mm(90007, 99.00, '90D', '2026-11-01 12:00:00'),
        linha_mm("A'B-01", 5.55, '90D', '2026-11-01 12:00:00')]  # apostrofo
for i in range(1, 46):                                           # 45 SKUs p/ testar lotes de 40
    mm90.append(linha_mm(91000 + i, round(10 + ((i - 1) % 7) * 1.5, 2), '90D', '2026-11-01 12:00:00'))

mmult = [HDR_MM,
         linha_mm(90001, 13.75, 'ULT', '2026-11-05 09:00:00'),
         linha_mm(90002, 7.35, 'ULT', '2026-11-05 09:00:00')]

ex = [['Cod_Produto', 'Descricao_Produto', 'Custo_Medio_SKU', 'Custo_Final_Engenharia', 'Data_Processamento'],
      [90001, 'SINTETICO 90001', 1.00, 11.11, '2026-11-10 15:30:00'],   # codigo NUMERICO na base
      [90002, 'SINTETICO 90002', 1.00, 22.22, '2026-11-10 15:30:00'],
      [90009, 'SINTETICO 90009', 1.00, 33.33, '2026-11-10 15:30:00']]

sheets = [
    ('Resumo Geral', [['Item', 'Valor'], ['Nao e pacote', 1]]),
    ('Custo_MM_90D_20261101_1200', mm90),
    ('Custo_MM_ULT_20261105_0900', mmult),
    ('Custo_EX_20261110_1530', ex),
    ('Custo_XX_20261101_1200', [HDR_MM, linha_mm(90001, 1.00, '90D', '2026-11-01 12:00:00')]),   # prefixo desconhecido
    ('Notas_Custo_MM_90D', [['Obs'], ['prefixo no meio do nome']]),
]

sst, sidx = [], {}
def s(v):
    if v not in sidx: sidx[v] = len(sst); sst.append(v)
    return sidx[v]
def cell(ref, v):
    if v is None: return ''
    if isinstance(v, bool): return f'<c r="{ref}" t="b"><v>{int(v)}</v></c>'
    if isinstance(v, (int, float)):
        return f'<c r="{ref}"><v>{repr(float(v)) if not float(v).is_integer() else int(v)}</v></c>'
    return f'<c r="{ref}" t="s"><v>{s(str(v))}</v></c>'

sheet_xml = []
for name, rows in sheets:
    xr = []
    for ri, vals in enumerate(rows, start=1):
        xr.append(f'<row r="{ri}">' + ''.join(cell(f'{col(ci)}{ri}', v) for ci, v in enumerate(vals, start=1)) + '</row>')
    ncols = max(len(r) for r in rows)
    sheet_xml.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f'<dimension ref="A1:{col(ncols)}{len(rows)}"/><sheetData>' + ''.join(xr) + '</sheetData></worksheet>')

ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      '<Default Extension="xml" ContentType="application/xml"/>'
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      + ''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1, len(sheets) + 1)) +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>')
wb = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      + ''.join(f'<sheet name="{escape(n)}" sheetId="{i}" r:id="rId{i}"/>' for i, (n, _) in enumerate(sheets, start=1)) + '</sheets></workbook>')
rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + ''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
        + f'<Relationship Id="rId{len(sheets)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        + f'<Relationship Id="rId{len(sheets)+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/></Relationships>')
styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
          '<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill>'
          '<fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
          '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
          '<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs></styleSheet>')
sst_xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           f'count="{len(sst)}" uniqueCount="{len(sst)}">' + ''.join(f'<si><t xml:space="preserve">{escape(t)}</t></si>' for t in sst) + '</sst>')
root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')

with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    z.writestr('[Content_Types].xml', ct)
    z.writestr('_rels/.rels', root_rels)
    z.writestr('xl/workbook.xml', wb)
    z.writestr('xl/_rels/workbook.xml.rels', rels)
    z.writestr('xl/styles.xml', styles)
    z.writestr('xl/sharedStrings.xml', sst_xml)
    for i, x in enumerate(sheet_xml, start=1):
        z.writestr(f'xl/worksheets/sheet{i}.xml', x)
print('ok', out, [n for n, _ in sheets])
