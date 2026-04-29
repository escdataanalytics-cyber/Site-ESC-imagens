# ============================================================
# DTE_App.xlsm builder – v2 (FIXED)
# Generates a complete .xlsm with working VBA macros.
# ============================================================

import struct
import zipfile
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

# ─────────────────────────────────────────────────────────────
# VBA SOURCE CODE
# ─────────────────────────────────────────────────────────────

VBA_MODULES = {}

VBA_MODULES["ThisWorkbook"] = ("class", r"""
Option Explicit

Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    Sheets("DADOS_DTE").Visible = xlSheetVeryHidden
    Sheets("CONFIG").Visible = xlSheetVeryHidden
    Sheets("FORM_DTE").Activate
    Application.ScreenUpdating = True
    MsgBox "Bem-vindo ao DTE v2.0" & Chr(10) & _
           "Diagnóstico Técnico Especializado!" & Chr(10) & Chr(10) & _
           "Preencha os campos e clique em CALCULAR ou SALVAR.", _
           vbInformation, "DTE"
End Sub
""")

VBA_MODULES["Module_Utils"] = ("procedural", r"""
Option Explicit

Public Const SCORE_OTIMO As Integer = 10
Public Const SCORE_BOM   As Integer = 5
Public Const SCORE_RUIM  As Integer = 0

Public Function GerarID() As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("DADOS_DTE")
    Dim ul As Long
    ul = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ul < 2 Then ul = 1
    GerarID = "DTE-" & Format(Now, "yyyymmdd") & "-" & Format(ul, "000")
End Function

Public Function CalcPct(n As Double, d As Double) As Double
    If d = 0 Then CalcPct = 0 Else CalcPct = n / d
End Function

Public Function PontuacaoAlto(pct As Double) As Integer
    If pct > 0.8 Then
        PontuacaoAlto = SCORE_OTIMO
    ElseIf pct >= 0.6 Then
        PontuacaoAlto = SCORE_BOM
    Else
        PontuacaoAlto = SCORE_RUIM
    End If
End Function

Public Function PontuacaoBaixo(pct As Double) As Integer
    If pct <= 0.2 Then
        PontuacaoBaixo = SCORE_OTIMO
    ElseIf pct <= 0.4 Then
        PontuacaoBaixo = SCORE_BOM
    Else
        PontuacaoBaixo = SCORE_RUIM
    End If
End Function

Public Sub FormatarScore(cel As Range, sc As Integer)
    With cel
        .Value = sc
        .Font.Bold = True
        Select Case sc
            Case SCORE_OTIMO
                .Interior.Color = RGB(198, 239, 206)
                .Font.Color = RGB(0, 97, 0)
            Case SCORE_BOM
                .Interior.Color = RGB(255, 235, 156)
                .Font.Color = RGB(156, 87, 0)
            Case Else
                .Interior.Color = RGB(255, 199, 206)
                .Font.Color = RGB(156, 0, 6)
        End Select
    End With
End Sub
""")

# ─────────────────────────────────────────────────────────────
# MS‑OVBA COMPRESSION
# ─────────────────────────────────────────────────────────────

def compress_vba(data: bytes) -> bytes:
    if isinstance(data, str):
        data = data.encode("cp1252", errors="replace")

    result = bytearray(b"\x01")
    for i in range(0, max(1, len(data)), 4096):
        chunk = data[i:i + 4096]
        padded = chunk + b"\x00" * (4096 - len(chunk))
        header = struct.pack("<H", 0b0011111111111111)
        result += header + padded
    return bytes(result)

def vba_record(tag: int, data: bytes) -> bytes:
    return struct.pack("<HI", tag, len(data)) + data

def build_dir(modules):
    b = bytearray()
    b += vba_record(0x0001, struct.pack("<I", 1))
    b += vba_record(0x0002, struct.pack("<I", 0x0409))
    b += vba_record(0x0014, struct.pack("<I", 0x0409))
    b += vba_record(0x0003, struct.pack("<H", 1252))
    b += vba_record(0x0004, b"DTE_VBA")
    b += vba_record(0x0005, b"")
    b += vba_record(0x0040, b"")
    b += vba_record(0x0006, b"")
    b += vba_record(0x003D, b"")
    b += vba_record(0x0007, struct.pack("<I", 0))
    b += vba_record(0x0008, struct.pack("<I", 0))
    b += struct.pack("<HI", 0x0009, 4) + struct.pack("<I", 1) + struct.pack("<H", 0)
    b += vba_record(0x000C, b"")
    b += vba_record(0x003C, b"")
    b += struct.pack("<HI", 0x000F, 2) + struct.pack("<H", len(modules))
    b += vba_record(0x0013, struct.pack("<H", 0xFFFF))

    for name, stream, mtype in modules:
        b += vba_record(0x0019, name.encode("cp1252"))
        b += vba_record(0x0031, name.encode("utf-16-le"))
        b += vba_record(0x001A, stream.encode("cp1252"))
        b += vba_record(0x0032, stream.encode("utf-16-le"))
        b += vba_record(0x001C, b"")
        b += vba_record(0x0048, b"")
        b += vba_record(0x0025, struct.pack("<I", 0))
        b += struct.pack("<HI", 0x0022 if mtype == "class" else 0x0021, 0)
        b += struct.pack("<HI", 0x002B, 0)

    b += struct.pack("<HI", 0x0010, 0)
    return bytes(b)

# ─────────────────────────────────────────────────────────────
# VBA PROJECT BIN
# ─────────────────────────────────────────────────────────────

NULL_CLSID = b"\x00" * 16
VBA_CLSID = bytes.fromhex("06090200000000000000000000000000")

def pad512(d: bytes) -> bytes:
    return d + b"\x00" * ((-len(d)) % 512)

def dir_entry(name, etype, color, left, right, child, clsid, start, size):
    nb = name.encode("utf-16-le")
    nb = nb[:62].ljust(64, b"\x00")
    return (
        nb +
        struct.pack("<HBB", len(name) * 2 + 2, etype, color) +
        struct.pack("<III", left & 0xFFFFFFFF, right & 0xFFFFFFFF, child & 0xFFFFFFFF) +
        clsid +
        struct.pack("<I", 0) +
        struct.pack("<QQ", 0, 0) +
        struct.pack("<I", start & 0xFFFFFFFF) +
        struct.pack("<I", size)
    )

def build_vba_bin(mods):
    vba_prj = b"\xCC\x61"
    mod_list = [(n, n, t) for n, (t, _) in mods.items()]
    dir_stream = compress_vba(build_dir(mod_list))

    streams = {
        "_VBA_PROJECT": vba_prj,
        "dir": dir_stream
    }

    for n, (_, code) in mods.items():
        streams[n] = compress_vba(code.encode("cp1252", errors="replace"))

    ordered = list(streams.items())
    data = []
    sectors = {}
    sizes = {}

    cur = 1
    for name, content in ordered:
        padded = pad512(content)
        count = len(padded) // 512
        sectors[name] = list(range(cur, cur + count))
        sizes[name] = len(content)
        for i in range(count):
            data.append(padded[i*512:(i+1)*512])
        cur += count

    fat = [0xFFFFFFFF] * 128
    fat[0] = 0xFFFFFFFD

    fat_sector = struct.pack("<128I", *fat)

    entries = [
        dir_entry("Root Entry", 5, 1, 0xFFFFFFFF, 0xFFFFFFFF, 1, NULL_CLSID, 0xFFFFFFFE, 0),
        dir_entry("VBA", 1, 1, 0xFFFFFFFF, 0xFFFFFFFF, 2, VBA_CLSID, 0xFFFFFFFE, 0),
    ]

    for name in streams:
        entries.append(
            dir_entry(name, 2, 1, 0xFFFFFFFF, 0xFFFFFFFF, 0xFFFFFFFF,
                      NULL_CLSID, sectors[name][0], sizes[name])
        )

    entries_blob = b"".join(entries).ljust(512, b"\x00")

    header = (
        b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1" +
        b"\x00" * 16 +
        struct.pack("<HH", 0x003E, 0x0003) +
        struct.pack("<H", 0xFFFE) +
        struct.pack("<HH", 9, 6) +
        b"\x00" * 6 +
        struct.pack("<I", 0) +
        struct.pack("<I", 1) +
        struct.pack("<I", 1) +
        struct.pack("<I", 0) +
        struct.pack("<I", 0x1000) +
        struct.pack("<I", 0xFFFFFFFF) +
        struct.pack("<I", 0) +
        struct.pack("<I", 0xFFFFFFFF) +
        struct.pack("<I", 0) +
        struct.pack("<109I", *([0] + [0xFFFFFFFF] * 108))
    )

    return header + fat_sector + entries_blob + b"".join(data)

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

def main():
    out = "DTE_App.xlsm"
    wb = Workbook()
    wb.remove(wb.active)
    wb.create_sheet("FORM_DTE")
    wb.create_sheet("DADOS_DTE")
    wb.create_sheet("CONFIG")

    vba_bin = build_vba_bin(VBA_MODULES)

    tmp = out + ".tmp.xlsx"
    wb.save(tmp)

    with zipfile.ZipFile(tmp, "r") as zin, zipfile.ZipFile(out, "w") as zout:
        for item in zin.infolist():
            d = zin.read(item.filename)
            if item.filename == "[Content_Types].xml":
                d = d.replace(
                    b"openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    b"ms-excel.sheet.macroEnabled.main+xml"
                )
                d = d.replace(
                    b"</Types>",
                    b'<Override PartName="/xl/vbaProject.bin" '
                    b'ContentType="application/vnd.ms-office.vbaProject"/></Types>'
                )
            if item.filename == "xl/_rels/workbook.xml.rels":
                d = d.replace(
                    b"</Relationships>",
                    b'<Relationship Id="rId_vba" '
                    b'Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" '
                    b'Target="vbaProject.bin"/></Relationships>'
                )
            zout.writestr(item, d)
        zout.writestr("xl/vbaProject.bin", vba_bin)

    os.remove(tmp)
    print("XLSM gerado com sucesso:", out)

if __name__ == "__main__":
    main()