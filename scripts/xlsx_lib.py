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

    def set_formula_cache(self, ref, value):
        """Giữ nguyên công thức <f> đang có sẵn trong ô (vd =G15*H15, =SUM(...)),
        chỉ ghi lại giá trị <v> đã tính sẵn để các phần mềm/trình xem KHÔNG tự
        tính lại công thức (Zalo preview, xem nhanh online...) vẫn hiển thị đúng
        số ngay lập tức thay vì để trống."""
        s = self._read(self.sheet_path)
        pat = re.compile(r'(<c r="%s"(?:\s+[a-zA-Z0-9:]+="[^"]*")*\s*>)(.*?)</c>'
                         % re.escape(ref), re.S)
        m = pat.search(s)
        if not m:
            raise ValueError("Không tìm thấy ô %s trong %s" % (ref, self.sheet_path))
        inner = m.group(2)
        fm = re.search(r'<f[^>]*>.*?</f>|<f[^>]*/>', inner, re.S)
        f_part = fm.group(0) if fm else ""
        new_inner = f_part + "<v>%s</v>" % value
        self._write(self.sheet_path, s[:m.start()] + m.group(1) + new_inner + "</c>" + s[m.end():])

    def force_recalc(self):
        """Bật fullCalcOnLoad để Excel/LibreOffice luôn tính lại toàn bộ công thức
        ngay khi mở file (phòng khi có phần mềm không nhận cache <v>)."""
        rel = "xl/workbook.xml"
        if not self._exists(rel):
            return
        s = self._read(rel)
        if "<calcPr" in s:
            if "fullCalcOnLoad" not in s:
                s = re.sub(r'<calcPr ', '<calcPr fullCalcOnLoad="1" ', s, count=1)
            else:
                s = re.sub(r'fullCalcOnLoad="[^"]*"', 'fullCalcOnLoad="1"', s, count=1)
        else:
            s = s.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>')
        self._write(rel, s)

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

    def set_row_autoheight(self, row):
        """Bỏ chiều cao cố định của dòng, để Excel tự tính chiều cao theo nội dung
        (dùng cho các dòng có xuống dòng tự động / wrap text)."""
        s = self._read(self.sheet_path)
        def repl(m):
            tag = m.group(0)
            tag = re.sub(r'\s+ht="[^"]*"', '', tag)
            tag = re.sub(r'\s+customHeight="[^"]*"', '', tag)
            return tag
        self._write(self.sheet_path,
                    re.sub(r'<row r="%d"[^>]*>' % int(row), repl, s, count=1))

    def set_row_hidden(self, row):
        """Ẩn dòng (hidden="1") — dòng ẩn KHÔNG hiển thị khi in / xuất PDF.
        Dùng để giấu các dòng sản phẩm trống trong bảng báo giá."""
        s = self._read(self.sheet_path)
        def repl(m):
            tag = m.group(0)
            if 'hidden=' in tag:
                return re.sub(r'hidden="[^"]*"', 'hidden="1"', tag)
            if tag.endswith('/>'):
                return tag[:-2] + ' hidden="1"/>'
            return tag[:-1] + ' hidden="1">'
        self._write(self.sheet_path,
                    re.sub(r'<row r="%d"[^>]*?/?>' % int(row), repl, s, count=1))

    # ---------- styles (number format / shrink-to-fit) ----------
    def _styles(self):
        return self._read("xl/styles.xml")

    def _cell_style_id(self, ref):
        s = self._read(self.sheet_path)
        m = re.search(r'<c r="%s"((?:\s+[a-zA-Z0-9:]+="[^"]*")*)\s*(?:/>|>)' % re.escape(ref), s)
        if not m:
            raise ValueError("Không tìm thấy ô %s" % ref)
        sm = re.search(r'\s+s="([^"]*)"', m.group(1))
        return int(sm.group(1)) if sm else 0

    def _set_cell_style_id(self, ref, style_id):
        s = self._read(self.sheet_path)
        pat = re.compile(r'(<c r="%s")((?:\s+[a-zA-Z0-9:]+="[^"]*")*)(\s*(?:/>|>))' % re.escape(ref))
        def repl(m):
            attrs = m.group(2)
            if re.search(r'\s+s="[^"]*"', attrs):
                attrs = re.sub(r'\s+s="[^"]*"', ' s="%d"' % style_id, attrs)
            else:
                attrs = ' s="%d"' % style_id + attrs
            return m.group(1) + attrs + m.group(3)
        self._write(self.sheet_path, pat.sub(repl, s, count=1))

    def _ensure_numfmt(self, fmt_code):
        """Đăng ký (hoặc lấy lại) 1 numFmt tuỳ chỉnh trong styles.xml, trả về numFmtId."""
        s = self._styles()
        ids = [int(x) for x in re.findall(r'<numFmt numFmtId="(\d+)"', s)]
        existing = re.search(r'<numFmt numFmtId="(\d+)" formatCode="%s"/>' % re.escape(fmt_code), s)
        if existing:
            return int(existing.group(1))
        new_id = max(ids + [163]) + 1
        new_numfmt = '<numFmt numFmtId="%d" formatCode="%s"/>' % (new_id, fmt_code)
        m_self_close = re.search(r'<numFmts count="(\d+)"\s*/>', s)
        m_open = re.search(r'<numFmts count="(\d+)">', s)
        if m_open:
            cnt = int(m_open.group(1)) + 1
            s2 = s[:m_open.start()] + '<numFmts count="%d">' % cnt + new_numfmt + s[m_open.end():]
        elif m_self_close:
            s2 = s[:m_self_close.start()] + '<numFmts count="1">' + new_numfmt + '</numFmts>' + s[m_self_close.end():]
        else:
            # chưa có khối numFmts nào — chèn mới ngay sau thẻ mở <styleSheet ...>
            m0 = re.search(r'(<styleSheet[^>]*>)', s)
            s2 = s[:m0.end()] + '<numFmts count="1">' + new_numfmt + '</numFmts>' + s[m0.end():]
        self._write("xl/styles.xml", s2)
        return new_id

    def _clone_xf_with_numfmt(self, orig_style_id, numfmt_id):
        s = self._styles()
        m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', s, re.S)
        count = int(m.group(1))
        body = m.group(2)
        xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', body, re.S)
        orig = xfs[orig_style_id]
        if re.search(r'numFmtId="\d+"', orig):
            new_xf = re.sub(r'numFmtId="\d+"', 'numFmtId="%d"' % numfmt_id, orig)
        else:
            new_xf = orig.replace("<xf ", '<xf numFmtId="%d" ' % numfmt_id, 1)
        if 'applyNumberFormat' not in new_xf:
            new_xf = new_xf.replace("<xf ", '<xf applyNumberFormat="1" ', 1)
        new_body = body + new_xf
        new_count = count + 1
        s2 = s[:m.start()] + '<cellXfs count="%d">' % new_count + new_body + '</cellXfs>' + s[m.end():]
        self._write("xl/styles.xml", s2)
        return new_count - 1  # index of appended xf

    def set_number_format(self, ref, fmt_code):
        """Đổi định dạng hiển thị số của 1 ô (vd '[$-42A]#,##0' để luôn hiện dấu CHẤM
        ngăn cách hàng nghìn, không phụ thuộc locale máy mở file)."""
        numfmt_id = self._ensure_numfmt(fmt_code)
        orig_style = self._cell_style_id(ref)
        new_style = self._clone_xf_with_numfmt(orig_style, numfmt_id)
        self._set_cell_style_id(ref, new_style)

    def set_shrink_to_fit(self, ref):
        """Bật shrinkToFit cho 1 ô để chữ tự co lại vừa 1 dòng, không bị xuống dòng/cắt."""
        s = self._styles()
        style_id = self._cell_style_id(ref)
        m = re.search(r'<cellXfs count="(\d+)">(.*?)</cellXfs>', s, re.S)
        body = m.group(2)
        xfs = re.findall(r'<xf\b[^>]*?(?:/>|>.*?</xf>)', body, re.S)
        orig = xfs[style_id]
        if "<alignment" in orig:
            if "shrinkToFit" not in orig:
                new_xf = re.sub(r'<alignment ', '<alignment shrinkToFit="1" wrapText="0" ', orig, count=1)
            else:
                new_xf = orig
        else:
            if orig.endswith("/>"):
                new_xf = orig[:-2] + '><alignment shrinkToFit="1" wrapText="0"/></xf>'
            else:
                new_xf = orig.replace("</xf>", '<alignment shrinkToFit="1" wrapText="0"/></xf>')
        new_body = body + new_xf
        new_count = int(m.group(1)) + 1
        s2 = s[:m.start()] + '<cellXfs count="%d">' % new_count + new_body + '</cellXfs>' + s[m.end():]
        self._write("xl/styles.xml", s2)
        self._set_cell_style_id(ref, new_count - 1)

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

        # Template có thể khai báo namespace vẽ hình theo 2 kiểu khác nhau:
        #   <xdr:wsDr xmlns:xdr="...">  (có tiền tố xdr:)
        #   <wsDr xmlns="...">          (namespace mặc định, KHÔNG tiền tố)
        # Phải tự phát hiện đúng kiểu template đang dùng, nếu không thẻ mới
        # chèn vào sẽ không khớp namespace gốc -> Excel không hiện ảnh
        # (hoặc tệ hơn, chèn không thành công mà không báo lỗi).
        drawing_src = self._read(dp)
        m_root = re.search(r'<(\w+:)?wsDr\b', drawing_src)
        prefix = (m_root.group(1) or "") if m_root else "xdr:"
        p = prefix  # "" hoặc "xdr:"

        # Dùng oneCellAnchor: kích thước ảnh hiển thị do <ext cx,cy> quyết định
        # trực tiếp (không phụ thuộc độ rộng/cao ô), tránh bị co nhỏ khi ô có
        # autoheight. Lưu ý bắt buộc: mỗi thẻ <a:...> phải tự khai báo
        # xmlns:a riêng (không kế thừa được từ thẻ tự đóng liền trước) —
        # đây chính là lỗi khiến ảnh từng không hiện ra dù XML "trông có vẻ đúng".
        anchor = (
            '<{p}oneCellAnchor>'
            '<{p}from><{p}col>{col}</{p}col><{p}colOff>{coloff}</{p}colOff>'
            '<{p}row>{row}</{p}row><{p}rowOff>{rowoff}</{p}rowOff></{p}from>'
            '<{p}ext cx="{cx}" cy="{cy}"/>'
            '<{p}pic>'
            '<{p}nvPicPr><{p}cNvPr id="{picid}" name="ProductImage{picid}" descr="Picture"/>'
            '<{p}cNvPicPr/></{p}nvPicPr>'
            '<{p}blipFill><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="{rid}"/>'
            '<a:stretch xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:fillRect/></a:stretch></{p}blipFill>'
            '<{p}spPr><a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
            '<a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
            '<a:prstGeom xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" prst="rect">'
            '<avLst/></a:prstGeom><a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'
            '<a:ln xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" w="0">'
            '<a:noFill/><a:prstDash val="solid"/></a:ln></{p}spPr>'
            '</{p}pic><{p}clientData/></{p}oneCellAnchor>'
        ).format(p=p, col=col, coloff=col_off_px * PX_TO_EMU, row=row,
                 rowoff=row_off_px * PX_TO_EMU, cx=cx, cy=cy,
                 picid=self._next_img_id, rid=rid)

        close_tag = "</{p}wsDr>".format(p=p)
        if close_tag not in drawing_src:
            raise RuntimeError("Không tìm thấy thẻ đóng %s trong %s — cấu trúc drawing khác thường" % (close_tag, dp))
        self._write(dp, drawing_src.replace(close_tag, anchor + close_tag))
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
        # copy file nén sang đích (an toàn cho mount Cowork), giữ nguyên mọi media
        shutil.copyfile(tmp_zip, out_path)
        try:
            os.remove(tmp_zip)
        except Exception:
            pass

    def cleanup(self):
        """Xóa thư mục tạm giải nén."""
        shutil.rmtree(self.tmp, ignore_errors=True)
