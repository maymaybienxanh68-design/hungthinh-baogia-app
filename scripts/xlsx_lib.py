#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
xlsx_lib.py — Sửa file Excel báo giá bằng cách thao tác trực tiếp XML trong file .xlsx
(KHÔNG dùng openpyxl để lưu, vì openpyxl làm RỚT ảnh .emf/.wmf — thường là con dấu,
chữ ký, đường viền trang trí).

Mục tiêu: đổi tên sản phẩm / số lượng / đơn giá / mô tả, chèn ảnh sản phẩm vào ô,
bỏ ẩn cột "Hình ảnh", chỉnh chiều cao dòng & độ rộng cột — mà GIỮ NGUYÊN mọi ảnh,
con dấu, chữ ký, khung kẻ của file mẫu gốc.
"""

import os, re, shutil, zipfile, subprocess, tempfile, unicodedata

PX_TO_EMU = 9525  # 1 pixel = 9525 EMU


def _nfc(s):
    return unicodedata.normalize("NFC", s) if isinstance(s, str) else s


def _xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
             .replace('"', "&quot;"))


class XlsxQuote:
    def __init__(self, src_xlsx, sheet="xl/worksheets/sheet1.xml"):
        self.src = src_xlsx
        self.tmp = tempfile.mkdtemp(prefix="xlsxq_")
        with zipfile.ZipFile(src_xlsx) as z:
            z.extractall(self.tmp)
        self.sheet_path = sheet
        self._next_img_id = self._max_drawing_id() + 1
        self._next_rid = self._max_drawing_rid() + 1

    # ---------- low level file helpers ----------
    def _p(self, rel):
        return os.path.join(self.tmp, rel)

    def _read(self, rel):
        with open(self._p(rel), encoding="utf-8") as f:
            return f.read()

    def _write(self, rel, s):
        with open(self._p(rel), "w", encoding="utf-8") as f:
            f.write(s)

    def _exists(self, rel):
        return os.path.exists(self._p(rel))

    # ---------- sheet cell editing (raw XML, self-contained per cell) ----------
    def _edit_cell(self, ref, new_inner, force_type=None):
        """Thay nội dung 1 ô, GIỮ nguyên thuộc tính style s="..". new_inner là phần
        bên trong <c ...>NEW</c>. force_type: giá trị thuộc tính t (vd 'inlineStr')."""
        s = self._read(self.sheet_path)
        pat = re.compile(r'<c r="%s"((?:\s+[a-zA-Z0-9:]+="[^"]*")*)\s*(?:/>|>.*?</c>)'
                         % re.escape(ref), re.S)
        m = pat.search(s)
        if not m:
            raise ValueError("Không tìm thấy ô %s trong %s" % (ref, self.sheet_path))
        attrs = m.group(1)
        sm = re.search(r'\s+s="([^"]*)"', attrs)
        style = ' s="%s"' % sm.group(1) if sm else ""
        t_attr = ' t="%s"' % force_type if force_type else ""
        new_cell = '<c r="%s"%s%s>%s</c>' % (ref, style, t_attr, new_inner)
        self._write(self.sheet_path, s[:m.start()] + new_cell + s[m.end():])

    def set_text(self, ref, value):
        """Đặt CHỮ cho ô bằng inlineStr (không đụng sharedStrings — an toàn nhất)."""
        v = _xml_escape(_nfc(str(value)))
        self._edit_cell(ref, '<is><t xml:space="preserve">%s</t></is>' % v, force_type="inlineStr")

    def set_number(self, ref, value):
        """Đặt SỐ cho ô (đơn giá, số lượng, hoặc cache thành tiền)."""
        self._edit_cell(ref, "<v>%s</v>" % value)

    def replace_in_strings(self, old, new):
        """Thay chuỗi con trong sharedStrings.xml — tiện cho đổi NGÀY trong rich-text."""
        rel = "xl/sharedStrings.xml"
        if not self._exists(rel):
            return False
        s = self._read(rel)
        if old not in s:
            return False
        self._write(rel, s.replace(old, _nfc(new)))
        return True

    # ---------- columns & rows ----------
    @staticmethod
    def _col_to_idx(letter):
        letter = letter.upper(); n = 0
        for ch in letter:
            n = n * 26 + (ord(ch) - 64)
        return n  # 1-based

    def unhide_col(self, letter):
        idx = self._col_to_idx(letter)
        s = self._read(self.sheet_path)
        pat = re.compile(r'(<col min="%d" max="%d"[^>]*?)\s*hidden="1"([^>]*/>)' % (idx, idx))
        self._write(self.sheet_path, pat.sub(r"\1\2", s))

    def set_col_width(self, letter, width):
        idx = self._col_to_idx(letter)
        s = self._read(self.sheet_path)
        def repl(m):
            tag = m.group(0)
            if 'width="' in tag:
                tag = re.sub(r'width="[^"]*"', 'width="%s"' % width, tag)
            else:
                tag = tag[:-2] + ' width="%s"/>' % width
            if "customWidth" not in tag:
                tag = tag[:-2] + ' customWidth="1"/>'
            return tag
        self._write(self.sheet_path,
                    re.sub(r'<col min="%d" max="%d"[^>]*/>' % (idx, idx), repl, s, count=1))

    def set_row_height(self, row, height):
        s = self._read(self.sheet_path)
        def repl(m):
            tag = m.group(0)
            if 'ht="' in tag:
                tag = re.sub(r'ht="[^"]*"', 'ht="%s"' % height, tag)
            else:
                tag = tag[:-1] + ' ht="%s"' % height + ">"
            if "customHeight" not in tag:
                tag = tag[:-1] + ' customHeight="1">'
            return tag
        self._write(self.sheet_path,
                    re.sub(r'<row r="%d"[^>]*>' % int(row), repl, s, count=1))

    # ---------- images / drawing ----------
    def _drawing_path(self):
        rel = "xl/worksheets/_rels/sheet1.xml.rels"
        if self._exists(rel):
            m = re.search(r'Target="\.\./(drawings/[^"]+)"', self._read(rel))
            if m:
                return "xl/" + m.group(1)
        return "xl/drawings/drawing1.xml"

    def _max_drawing_id(self):
        dp = self._drawing_path()
        if not self._exists(dp):
            return 0
        ids = [int(x) for x in re.findall(r'id="(\d+)"', self._read(dp))]
        return max(ids) if ids else 0

    def _max_drawing_rid(self):
        rel = "xl/drawings/_rels/" + os.path.basename(self._drawing_path()) + ".rels"
        if not self._exists(rel):
            return 0
        ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', self._read(rel))]
        return max(ids) if ids else 0

    def _ensure_content_type(self, ext):
        ext = ext.lower().lstrip(".")
        ct = {"png": "image/png", "jpeg": "image/jpeg", "jpg": "image/jpeg",
              "gif": "image/gif", "emf": "image/x-emf"}.get(ext, "application/octet-stream")
        rel = "[Content_Types].xml"
        s = self._read(rel)
        if 'Extension="%s"' % ext not in s:
            s = s.replace("</Types>",
                '<Default Extension="%s" ContentType="%s"/></Types>' % (ext, ct))
            self._write(rel, s)

    def embed_image(self, anchor_cell, img_path, width_px=110, height_px=None,
                    col_off_px=8, row_off_px=4):
        """Chèn ảnh vào ô (vd 'F12'), GIỮ tỷ lệ. Dùng oneCellAnchor để không méo ảnh."""
        if height_px is None:
            try:
                from PIL import Image
                w, h = Image.open(img_path).size
                height_px = int(round(width_px * h / w))
            except Exception:
                height_px = width_px
        m = re.match(r"([A-Za-z]+)(\d+)", anchor_cell)
        col = self._col_to_idx(m.group(1)) - 1
        row = int(m.group(2)) - 1
        ext = os.path.splitext(img_path)[1].lower().lstrip(".")
        if ext == "jpg":
            ext = "jpeg"
        os.makedirs(self._p("xl/media"), exist_ok=True)
        img_name = "image_q%d.%s" % (self._next_img_id, ext)
        shutil.copy(img_path, self._p("xl/media/" + img_name))
        self._ensure_content_type(ext)

        dp = self._drawing_path()
        if not self._exists(dp):
            raise RuntimeError("File mẫu không có drawing — cần dựng lại bằng template (xem SKILL.md)")
        rels = "xl/drawings/_rels/" + os.path.basename(dp) + ".rels"
        rid = "rId%d" % self._next_rid
        if self._exists(rels):
            s = self._read(rels).replace("</Relationships>",
                '<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/></Relationships>' % (rid, img_name))
            self._write(rels, s)
        else:
            os.makedirs(os.path.dirname(self._p(rels)), exist_ok=True)
            self._write(rels, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/%s"/></Relationships>' % (rid, img_name))

        cx = width_px * PX_TO_EMU
        cy = height_px * PX_TO_EMU
        anchor = (
            '<xdr:oneCellAnchor>'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>%d</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>%d</xdr:rowOff></xdr:from>'
            '<xdr:ext cx="%d" cy="%d"/>'
            '<xdr:pic>'
            '<xdr:nvPicPr><xdr:cNvPr id="%d" name="ProductImage%d"/>'
            '<xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
            '<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="%s"/>'
            '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
            '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>'
            '</xdr:pic><xdr:clientData/></xdr:oneCellAnchor>'
        ) % (col, col_off_px * PX_TO_EMU, row, row_off_px * PX_TO_EMU, cx, cy,
             self._next_img_id, self._next_img_id, rid, cx, cy)
        self._write(dp, self._read(dp).replace("</xdr:wsDr>", anchor + "</xdr:wsDr>"))
        self._next_img_id += 1
        self._next_rid += 1

    # ---------- save (an toàn cho mount Cowork) ----------
    def save(self, out_path):
        """Nén ra /tmp rồi cat sang đích (tên MỚI). Giữ nguyên mọi media (kể cả emf)."""
        tmp_zip = tempfile.mktemp(suffix=".xlsx")
        cwd = os.getcwd()
        try:
            os.chdir(self.tmp)
            names = []
            for root, _, files in os.walk("."):
                for fn in files:
                    names.append(os.path.relpath(os.path.join(root, fn), "."))
            names.sort(key=lambda n: (n != "[Content_Types].xml", n))
            with zipfile.ZipFile(tmp_zip, "w", zipfile.ZIP_DEFLATED) as z:
                for n in names:
                    z.write(n, n)
        finally:
            os.chdir(cwd)
        with open(tmp_zip, "rb") as fi, open(out_path, "wb") as fo:
            fo.write(fi.read())
        os.remove(tmp_zip)
        return out_path

    def cleanup(self):
        shutil.rmtree(self.tmp, ignore_errors=True)


def to_pdf(xlsx_path, out_dir):
    """Xuất PDF bằng LibreOffice (headless), nếu server có cài."""
    subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                    "--outdir", out_dir, xlsx_path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    return os.path.join(out_dir, os.path.splitext(os.path.basename(xlsx_path))[0] + ".pdf")
