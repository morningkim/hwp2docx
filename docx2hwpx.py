# -*- coding: utf-8 -*-
"""
docx2hwpx.py — DOCX를 HWPX(한글 OWPML)로 변환

오피스 프로그램(한컴/MS/LibreOffice) 없이 순수 Python으로 동작.
의존성: python-docx  (zipfile, xml 은 표준 라이브러리)

HWPX는 한컴의 XML+ZIP 패키지 포맷이다. 본 변환기는 본문 문단과 표(셀 병합 포함)를
section0.xml 로 생성하고, 서식 정의(header.xml)는 한컴이 생성한 표준 템플릿
(hwpx_assets/header.xml, 본문/개인정보 없는 순수 스타일 정의)을 재사용한다.

처리 범위:
  - 본문 문단(문단 내 줄바꿈은 여러 문단으로 분리)
  - 표: 행/열, 셀 병합(colSpan/rowSpan), 셀 텍스트
한계:
  - 글꼴/글자 크기/색상/이미지/도형은 재현하지 않음(텍스트·표 구조 위주)
  - 생성 결과는 한글에서 열어 확인 권장(레이아웃은 열 때 재계산됨)

사용법:
  python docx2hwpx.py                  # 현재 폴더의 모든 *.docx 변환
  python docx2hwpx.py a.docx           # a.docx -> a.hwpx
  python docx2hwpx.py a.docx out.hwpx  # 출력 경로 지정
"""

import os
import sys
import glob
import zipfile

from docx import Document
from docx.document import Document as _Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table

ASSET_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'hwpx_assets')

# 표/페이지 치수(HWPUNIT). 한글이 열 때 재계산하므로 대략값이면 충분.
TABLE_WIDTH = 42520
ROW_HEIGHT = 2697

LINESEG = ('<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" '
           'textheight="1000" baseline="850" spacing="600" horzpos="0" '
           'horzsize="%d" flags="393216"/></hp:linesegarray>')

# 첫 문단에 들어가는 구역 정의(페이지 크기/여백). 표준 템플릿에서 추출한 값.
SECPR = (
    '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000"'
    ' tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0"'
    ' textVerticalWidthHead="0" masterPageCnt="0">'
    '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
    '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
    '<hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0"'
    ' border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0"'
    ' showLineNumber="0"/>'
    '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
    '<hp:pagePr landscape="WIDELY" width="59528" height="84188" gutterType="LEFT_ONLY">'
    '<hp:margin header="2834" footer="2834" gutter="0" left="5669" right="5669"'
    ' top="2834" bottom="2834"/></hp:pagePr>'
    '<hp:footNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar=""'
    ' suffixChar=")" supscript="0"/><hp:noteLine length="-1" type="SOLID"'
    ' width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="283"'
    ' belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/>'
    '<hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr>'
    '<hp:endNotePr><hp:autoNumFormat type="DIGIT" userChar="" prefixChar=""'
    ' suffixChar=")" supscript="0"/><hp:noteLine length="14692344" type="SOLID"'
    ' width="0.12 mm" color="#000000"/><hp:noteSpacing betweenNotes="0"'
    ' belowLine="567" aboveLine="850"/><hp:numbering type="CONTINUOUS" newNum="1"/>'
    '<hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr>'
    '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER"'
    ' headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417"'
    ' right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
    '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER"'
    ' headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417"'
    ' right="1417" top="1417" bottom="1417"/></hp:pageBorderFill>'
    '<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER"'
    ' headerInside="0" footerInside="0" fillArea="PAPER"><hp:offset left="1417"'
    ' right="1417" top="1417" bottom="1417"/></hp:pageBorderFill></hp:secPr>'
)

SEC_NS = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)


def esc(s):
    """XML 텍스트 이스케이프 + 제어문자 제거"""
    s = s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    return ''.join(ch for ch in s if ch >= ' ' or ch in '\t')


class IdGen:
    def __init__(self):
        self.n = 0

    def next(self):
        self.n += 1
        return self.n


def iter_blocks(doc):
    """문서 본문의 문단/표를 순서대로 산출"""
    body = doc.element.body
    for child in body.iterchildren():
        if isinstance(child, CT_P):
            yield ('p', Paragraph(child, doc))
        elif isinstance(child, CT_Tbl):
            yield ('tbl', Table(child, doc))


def make_para(text, ids, horzsize=TABLE_WIDTH, secpr=False):
    """텍스트 한 줄 -> <hp:p> (줄바꿈은 호출부에서 분리)"""
    run = '<hp:run charPrIDRef="0">'
    if secpr:
        run += SECPR
    run += '<hp:t>%s</hp:t></hp:run>' % esc(text)
    return ('<hp:p id="%d" paraPrIDRef="0" styleIDRef="0" pageBreak="0" '
            'columnBreak="0" merged="0">%s%s</hp:p>'
            % (ids.next(), run, LINESEG % horzsize))


def split_lines(text):
    lines = text.replace('\r\n', '\n').replace('\r', '\n').split('\n')
    return lines if lines else ['']


def make_cell(td_text, col, row, cspan, rspan, colw, ids):
    width = colw * cspan
    height = ROW_HEIGHT * rspan
    inner = ''.join(make_para(ln, ids, horzsize=max(1000, width - 1020))
                    for ln in split_lines(td_text))
    sub = ('<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
           'vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" '
           'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">%s</hp:subList>'
           % inner)
    return ('<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" '
            'dirty="0" borderFillIDRef="4">%s'
            '<hp:cellAddr colAddr="%d" rowAddr="%d"/>'
            '<hp:cellSpan colSpan="%d" rowSpan="%d"/>'
            '<hp:cellSz width="%d" height="%d"/>'
            '<hp:cellMargin left="510" right="510" top="141" bottom="141"/></hp:tc>'
            % (sub, col, row, cspan, rspan, width, height))


def make_table(table, ids):
    nr = len(table.rows)
    nc = len(table.columns)
    if nr == 0 or nc == 0:
        return ''
    colw = max(1000, TABLE_WIDTH // nc)

    # 병합 셀 인식: grid 상 동일 tc 요소의 경계상자로 colSpan/rowSpan 산출
    info = {}      # tc -> [minr, minc, maxr, maxc, text]
    order = []
    for r in range(nr):
        for c in range(nc):
            try:
                cell = table.cell(r, c)
            except IndexError:
                continue
            tc = cell._tc
            if tc not in info:
                info[tc] = [r, c, r, c, cell.text]
                order.append(tc)
            else:
                v = info[tc]
                v[2] = max(v[2], r)
                v[3] = max(v[3], c)

    # 시작 행별로 셀 묶기
    rows = {}
    for tc in order:
        minr, minc, maxr, maxc, text = info[tc]
        cell_xml = make_cell(text, minc, minr,
                             maxc - minc + 1, maxr - minr + 1, colw, ids)
        rows.setdefault(minr, []).append((minc, cell_xml))

    tr_xml = []
    for r in range(nr):
        cells = sorted(rows.get(r, []), key=lambda x: x[0])
        tr_xml.append('<hp:tr>%s</hp:tr>' % ''.join(c for _, c in cells))

    tbl = ('<hp:tbl id="%d" zOrder="0" numberingType="TABLE" '
           'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" '
           'dropcapstyle="None" pageBreak="CELL" repeatHeader="1" '
           'rowCnt="%d" colCnt="%d" cellSpacing="0" borderFillIDRef="3" noAdjust="0">'
           '<hp:sz width="%d" widthRelTo="ABSOLUTE" height="%d" '
           'heightRelTo="ABSOLUTE" protect="0"/>'
           '<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" '
           'allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" '
           'vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
           '<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
           '<hp:inMargin left="510" right="510" top="141" bottom="141"/>%s</hp:tbl>'
           % (ids.next(), nr, nc, colw * nc, ROW_HEIGHT * nr, ''.join(tr_xml)))

    # 표는 문단의 run 안에 inline 으로 배치
    return ('<hp:p id="%d" paraPrIDRef="0" styleIDRef="0" pageBreak="0" '
            'columnBreak="0" merged="0"><hp:run charPrIDRef="0">%s</hp:run>%s</hp:p>'
            % (ids.next(), tbl, LINESEG % TABLE_WIDTH))


def build_section(doc):
    ids = IdGen()
    parts = []
    plain = []  # 미리보기 텍스트용
    first_done = False

    for kind, obj in iter_blocks(doc):
        if kind == 'p':
            lines = split_lines(obj.text)
            for ln in lines:
                parts.append(make_para(ln, ids, secpr=(not first_done)))
                first_done = True
                if ln.strip():
                    plain.append(ln)
        else:  # table
            if not first_done:
                # secPr 를 담을 빈 첫 문단 선행
                parts.append(make_para('', ids, secpr=True))
                first_done = True
            parts.append(make_table(obj, ids))
            for row in obj.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        plain.append(cell.text.strip())

    if not first_done:
        parts.append(make_para('', ids, secpr=True))

    sec = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
           '<hs:sec %s>%s</hs:sec>' % (SEC_NS, ''.join(parts)))
    return sec, '\n'.join(plain)


def build_hpf(title):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
        '<opf:package xmlns:opf="http://www.idpf.org/2007/opf/" '
        'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" version="" '
        'unique-identifier="" id="">'
        '<opf:metadata><opf:title>%s</opf:title><opf:language>ko</opf:language>'
        '</opf:metadata>'
        '<opf:manifest>'
        '<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
        '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
        '<opf:item id="settings" href="settings.xml" media-type="application/xml"/>'
        '</opf:manifest>'
        '<opf:spine><opf:itemref idref="header"/>'
        '<opf:itemref idref="section0" linear="yes"/></opf:spine>'
        '</opf:package>' % esc(title)
    )


VERSION_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
               '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" '
               'tagetApplication="WORDPROCESSOR" major="5" minor="0" micro="5" '
               'buildNumber="0" os="1" xmlVersion="1.4" application="docx2hwpx" '
               'appVersion="1.0"/>')

SETTINGS_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
                '<ha:HWPApplicationSetting '
                'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
                'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
                '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
                '</ha:HWPApplicationSetting>')

CONTAINER_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
                 '<ocf:container '
                 'xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" '
                 'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
                 '<ocf:rootfiles>'
                 '<ocf:rootfile full-path="Contents/content.hpf" '
                 'media-type="application/hwpml-package+xml"/>'
                 '<ocf:rootfile full-path="Preview/PrvText.txt" media-type="text/plain"/>'
                 '</ocf:rootfiles></ocf:container>')

MANIFEST_XML = ('<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
                '<odf:manifest '
                'xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>')


def load_header():
    path = os.path.join(ASSET_DIR, 'header.xml')
    if not os.path.exists(path):
        raise FileNotFoundError(
            'header 템플릿이 없습니다: %s\n(hwpx_assets/header.xml 필요)' % path)
    with open(path, 'rb') as f:
        return f.read()


def convert(src, dst=None):
    if dst is None:
        dst = os.path.splitext(src)[0] + '.hwpx'
    doc = Document(src)
    section, preview = build_section(doc)
    title = os.path.splitext(os.path.basename(src))[0]
    header = load_header()

    with zipfile.ZipFile(dst, 'w') as z:
        # mimetype 은 반드시 첫 항목 + 무압축
        z.writestr('mimetype', 'application/hwp+zip', zipfile.ZIP_STORED)
        z.writestr('version.xml', VERSION_XML, zipfile.ZIP_DEFLATED)
        z.writestr('settings.xml', SETTINGS_XML, zipfile.ZIP_DEFLATED)
        z.writestr('Contents/content.hpf', build_hpf(title), zipfile.ZIP_DEFLATED)
        z.writestr('Contents/header.xml', header, zipfile.ZIP_DEFLATED)
        z.writestr('Contents/section0.xml', section, zipfile.ZIP_DEFLATED)
        z.writestr('Preview/PrvText.txt', preview, zipfile.ZIP_DEFLATED)
        z.writestr('META-INF/container.xml', CONTAINER_XML, zipfile.ZIP_DEFLATED)
        z.writestr('META-INF/manifest.xml', MANIFEST_XML, zipfile.ZIP_DEFLATED)

    n_par = section.count('<hp:p ')
    n_tbl = section.count('<hp:tbl ')
    print('  -> %s  (문단 %d, 표 %d)' % (dst, n_par, n_tbl))
    return dst


def main(argv):
    if len(argv) >= 2:
        src = argv[1]
        dst = argv[2] if len(argv) >= 3 else None
        print('변환:', src)
        convert(src, dst)
    else:
        files = sorted(glob.glob('*.docx'))
        if not files:
            print('현재 폴더에 *.docx 파일이 없습니다.')
            return
        for f in files:
            print('변환:', f)
            try:
                convert(f)
            except Exception as e:
                print('  [실패]', e)


if __name__ == '__main__':
    main(sys.argv)
