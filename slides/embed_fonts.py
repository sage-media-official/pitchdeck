#!/usr/bin/env python3
"""Embed Glacial Indifference (Regular+Bold) into the generated .pptx so that
PowerPoint renders the real typeface even on machines where it isn't installed.
Follows the OOXML embedded-font spec (ppt/fonts/*.fntdata + embeddedFontLst)."""
import zipfile, shutil, os, re, sys

SRC = sys.argv[1] if len(sys.argv) > 1 else '/home/user/Sage-portfolio/Echon-Annual-Marketing-Plan-SageMedia.pptx'
REG = '/home/user/Sage-portfolio/fonts/GlacialIndifference-Regular.otf'
BOLD = '/home/user/Sage-portfolio/fonts/GlacialIndifference-Bold.otf'
TMP = SRC + '.tmp'

reg = open(REG,'rb').read()
bold = open(BOLD,'rb').read()

zin = zipfile.ZipFile(SRC,'r')
names = zin.namelist()
data = {n: zin.read(n) for n in names}
zin.close()

# 1) content types: add Default for fntdata
ct = data['[Content_Types].xml'].decode('utf-8')
if 'fntdata' not in ct:
    ct = ct.replace('</Types>',
        '<Default Extension="fntdata" ContentType="application/x-fontdata"/></Types>')
data['[Content_Types].xml'] = ct.encode('utf-8')

# 2) presentation.xml.rels: add font relationships
rels = data['ppt/_rels/presentation.xml.rels'].decode('utf-8')
FONT_T = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'
add_rels = (
    f'<Relationship Id="rIdFontReg" Type="{FONT_T}" Target="fonts/font1.fntdata"/>'
    f'<Relationship Id="rIdFontBold" Type="{FONT_T}" Target="fonts/font2.fntdata"/>'
)
rels = rels.replace('</Relationships>', add_rels + '</Relationships>')
data['ppt/_rels/presentation.xml.rels'] = rels.encode('utf-8')

# 3) presentation.xml: embedTrueTypeFonts + embeddedFontLst (after notesSz)
pres = data['ppt/presentation.xml'].decode('utf-8')
# add attribute on <p:presentation ...>
pres = re.sub(r'(<p:presentation\b)([^>]*?)(>)',
              lambda m: m.group(1)+m.group(2)+(' embedTrueTypeFonts="1"' if 'embedTrueTypeFonts' not in m.group(2) else '')+m.group(3),
              pres, count=1)
font_lst = ('<p:embeddedFontLst><p:embeddedFont>'
            '<p:font typeface="Glacial Indifference"/>'
            '<p:regular r:id="rIdFontReg"/>'
            '<p:bold r:id="rIdFontBold"/>'
            '</p:embeddedFont></p:embeddedFontLst>')
# insert right after the notesSz element (correct schema order)
m = re.search(r'<p:notesSz[^>]*/>', pres)
if m:
    pres = pres[:m.end()] + font_lst + pres[m.end():]
else:
    pres = pres.replace('</p:presentation>', font_lst + '</p:presentation>')
data['ppt/presentation.xml'] = pres.encode('utf-8')

# 4) add the font binaries
data['ppt/fonts/font1.fntdata'] = reg
data['ppt/fonts/font2.fntdata'] = bold

# write out
zout = zipfile.ZipFile(TMP,'w',zipfile.ZIP_DEFLATED)
for n, b in data.items():
    zout.writestr(n, b)
zout.close()
shutil.move(TMP, SRC)
print('embedded fonts ->', SRC, os.path.getsize(SRC),'bytes')
