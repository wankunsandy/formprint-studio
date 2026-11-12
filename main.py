import io
import json
import os
import re
import shutil
import subprocess
import sys
import time
import uuid
import string
from typing import Any, Dict, List, Optional

import requests
from docx.shared import Mm
from docxtpl import DocxTemplate, InlineImage
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

app = FastAPI(title="FormPrint Studio - Google Form to Print Document")

STORAGE_DIR = os.path.abspath("storage")
os.makedirs(STORAGE_DIR, exist_ok=True)

ALPHABETS = list(string.ascii_uppercase)


# ==============================================================================
# 1. PEMBERSIH FILE OTOMATIS & UTILITAS SISTEM
# ==============================================================================

def cleanup_old_files(directory: str, max_age_seconds: int = 1800) -> None:
    """Menghapus berkas DOCX/PDF di folder storage yang berusia lebih dari 30 menit."""
    now = time.time()
    try:
        for filename in os.listdir(directory):
            filepath = os.path.join(directory, filename)
            if os.path.isfile(filepath):
                if now - os.path.getmtime(filepath) > max_age_seconds:
                    os.remove(filepath)
    except Exception as e:
        print(f"[Peringatan] Gagal membersihkan file lama: {e}")


def get_libreoffice_command() -> List[str]:
    """Mendeteksi lokasi program LibreOffice di Linux/Docker, Windows, atau macOS."""
    if shutil.which("libreoffice"):
        return ["libreoffice"]
    if shutil.which("soffice"):
        return ["soffice"]

    if sys.platform.startswith("win"):
        standard_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for path in standard_paths:
            if os.path.exists(path):
                return [path]
    elif sys.platform == "darwin":
        mac_path = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(mac_path):
            return [mac_path]

    return []


def convert_docx_to_pdf(docx_path: str, output_dir: str) -> Optional[str]:
    """Mengonversi berkas DOCX ke PDF menggunakan LibreOffice mode headless."""
    cmd_base = get_libreoffice_command()
    if not cmd_base:
        print("[Info] LibreOffice tidak terpasang. Konversi PDF dilewati.")
        return None

    cmd = cmd_base + ["--headless", "--convert-to", "pdf", docx_path, "--outdir", output_dir]
    try:
        # Set HOME ke /tmp untuk mencegah kendala permission di container Linux/Docker
        env = os.environ.copy()
        if not sys.platform.startswith("win"):
            env["HOME"] = "/tmp"

        subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, env=env)
        pdf_path = os.path.splitext(docx_path)[0] + ".pdf"
        if os.path.exists(pdf_path):
            return pdf_path
        return None
    except Exception as e:
        print(f"[Peringatan] Gagal mengonversi ke PDF: {e}")
        return None


def find_image_url(data: Any) -> Optional[str]:
    """Mencari URL gambar Google secara rekursif di dalam struktur JSON."""
    if isinstance(data, str):
        if "googleusercontent.com" in data or (
            data.startswith("http") and any(ext in data.lower() for ext in [".png", ".jpg", ".jpeg", ".webp"])
        ):
            return data
    elif isinstance(data, list):
        for item in data:
            found = find_image_url(item)
            if found:
                return found
    elif isinstance(data, dict):
        for value in data.values():
            found = find_image_url(value)
            if found:
                return found
    return None


def download_image_to_stream(img_url: str) -> Optional[io.BytesIO]:
    """Mengunduh gambar dari internet langsung ke memory buffer."""
    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(img_url, headers=headers, timeout=5)
        if resp.status_code == 200:
            return io.BytesIO(resp.content)
    except Exception as e:
        print(f"[Peringatan] Gagal mengunduh gambar ({img_url}): {e}")
    return None


# ==============================================================================
# 2. PARSER OPSI & SOAL KISI (GRID)
# ==============================================================================

def extract_options(choice_data: list, doc: DocxTemplate) -> List[Dict[str, Any]]:
    """Mengekstrak daftar pilihan jawaban beserta gambarnya (jika ada)."""
    options = []
    if len(choice_data) > 1 and choice_data[1]:
        for idx, opt_item in enumerate(choice_data[1]):
            if not opt_item:
                continue
            opt_text = str(opt_item[0] or "").strip()

            opt_img_url = find_image_url(opt_item)
            opt_img_obj = None
            if opt_img_url:
                opt_stream = download_image_to_stream(opt_img_url)
                if opt_stream:
                    opt_img_obj = InlineImage(doc, opt_stream, width=Mm(32))

            options.append({
                "huruf": ALPHABETS[idx] if idx < len(ALPHABETS) else f"({idx+1})",
                "teks": opt_text,
                "gambar": opt_img_obj,
                "ada_gambar": opt_img_obj is not None
            })
    return options


def parse_grid_question(item: list) -> Dict[str, Any]:
    """
    Mendeteksi soal kisi (Tipe 7) dan mengklasifikasikannya menjadi:
    - pg_kompleks (jika 1 kolom centang)
    - benar_salah (jika 2 kolom Benar/Salah)
    - kisi_pilihan_ganda (jika kolom pendek A/B/C/D atau skala ringkas)
    - menjodohkan (jika kolom berisi teks pasangan yang panjang)
    """
    q_title = item[1] or "Soal Kisi"
    rows_data = item[4] or []

    # 1. Ekstrak header kolom
    columns = []
    if rows_data and len(rows_data[0]) > 1 and rows_data[0][1]:
        for col_item in rows_data[0][1]:
            if col_item and col_item[0]:
                columns.append(str(col_item[0]).strip())

    # 2. Ekstrak baris pernyataan
    rows = []
    for idx, r in enumerate(rows_data):
        row_text = ""
        if len(r) > 3 and r[3]:
            row_text = r[3][0] if isinstance(r[3], list) else str(r[3])

        rows.append({
            "idx": idx,
            "sub_no": chr(97 + idx) if idx < 26 else str(idx + 1),  # a, b, c...
            "no_angka": str(idx + 1),                              # 1, 2, 3...
            "teks": row_text.strip()
        })

    # Kasus A: Kisi 1 kolom -> Konversi otomatis menjadi Pilihan Ganda Kompleks
    if len(columns) == 1:
        pgk_options = [
            {
                "huruf": ALPHABETS[idx] if idx < len(ALPHABETS) else f"({idx+1})",
                "teks": b["teks"],
                "gambar": None,
                "ada_gambar": False
            }
            for idx, b in enumerate(rows)
        ]
        return {
            "tipe": "pg_kompleks",
            "pertanyaan": q_title,
            "opsi": pgk_options,
            "ada_opsi": True,
            "baris": [],
            "kolom_header": [],
            "kiri": [],
            "kanan": []
        }

    # Kasus B: Kisi 2 kolom Benar-Salah
    cols_lower = [c.lower() for c in columns]
    if len(columns) == 2 and any(k in cols_lower for k in ["benar", "salah", "b", "s", "true", "false", "ya", "tidak"]):
        return {
            "tipe": "benar_salah",
            "pertanyaan": q_title,
            "kolom_header": columns,
            "baris": rows,
            "opsi": [],
            "ada_opsi": False,
            "kiri": [],
            "kanan": []
        }

    # Kasus C: Kisi Pilihan Ganda Matriks 
    # (Hanya aktif jika judul kolom di Google Form murni huruf A/B/C/D atau angka 1/2/3/4 maks 4 kolom)
    opsi_matriks_standar = {"a", "b", "c", "d", "1", "2", "3", "4"}
    if len(columns) <= 4 and all(c.lower() in opsi_matriks_standar for c in columns):
        return {
            "tipe": "kisi_pilihan_ganda",
            "pertanyaan": q_title,
            "kolom_header": columns,
            "baris": rows,
            "opsi": [],
            "ada_opsi": False,
            "kiri": [],
            "kanan": []
        }

    # Kasus D: Soal Menjodohkan (Mendukung berapa pun jumlah opsi pasangan: 4, 5, 8, dst.)
    pilihan_kanan = [
        {
            "huruf": ALPHABETS[i] if i < len(ALPHABETS) else f"({i+1})",
            "teks": col_text
        }
        for i, col_text in enumerate(columns)
    ]
    return {
        "tipe": "menjodohkan",
        "pertanyaan": q_title,
        "kiri": rows,
        "kanan": pilihan_kanan,
        "baris": [],
        "kolom_header": columns,
        "opsi": [],
        "ada_opsi": False
    }


# ==============================================================================
# 3. EKSTRAKSI STRUKTUR GOOGLE FORM PER BAGIAN (SECTIONS)
# ==============================================================================

def extract_and_parse_form(form_url: str, doc: DocxTemplate) -> Dict[str, Any]:
    clean_url = form_url.strip().split("?")[0]
    if not clean_url.endswith("/viewform"):
        clean_url = clean_url.rstrip("/") + "/viewform"

    try:
        headers = {"User-Agent": "Mozilla/5.0"}
        resp = requests.get(clean_url, headers=headers, timeout=10)
        resp.raise_for_status()
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Gagal mengakses tautan Google Form: {str(e)}")

    pattern = r"FB_PUBLIC_LOAD_DATA_\s*=\s*(.*?);</script>"
    match = re.search(pattern, resp.text, re.DOTALL)
    if not match:
        raise HTTPException(status_code=400, detail="Formulir tidak ditemukan atau belum disetel ke mode Publik.")

    try:
        raw_data = json.loads(match.group(1))
        form_title = raw_data[1][8] or "Naskah Soal Ujian"
        form_desc = raw_data[1][0] or ""
        raw_items = raw_data[1][1] or []

        daftar_bagian = []
        current_section = {
            "nama_bagian": "Bagian I",
            "petunjuk": "",
            "soal": []
        }
        global_counter = 1

        for item in raw_items:
            if not item or len(item) < 4:
                continue

            item_type = item[3]

            # 1. Deteksi Pembatas Bagian / Section Break (Tipe 8)
            if item_type == 8:
                if current_section["soal"]:
                    daftar_bagian.append(current_section)
                current_section = {
                    "nama_bagian": item[1] or f"Bagian {len(daftar_bagian) + 1}",
                    "petunjuk": item[2] or "",
                    "soal": []
                }
                continue

            # 2. Deteksi Soal Kisi / Grid (Tipe 7)
            if item_type == 7 and len(item) >= 5 and item[4]:
                grid_res = parse_grid_question(item)
                nomor_di_bagian = len(current_section["soal"]) + 1

                # Cek apakah ada gambar utama pada soal kisi
                q_img_url = find_image_url(item[1:4])
                q_img_obj = None
                if q_img_url:
                    stream = download_image_to_stream(q_img_url)
                    if stream:
                        q_img_obj = InlineImage(doc, stream, width=Mm(65))

                current_section["soal"].append({
                    "no": nomor_di_bagian,
                    "no_global": global_counter,
                    "pertanyaan": grid_res["pertanyaan"],
                    "tipe": grid_res["tipe"],
                    "ada_gambar": q_img_obj is not None,
                    "gambar": q_img_obj,
                    "ada_opsi": grid_res["ada_opsi"],
                    "opsi_bergambar": False,
                    "opsi": grid_res["opsi"],
                    "baris": grid_res["baris"],
                    "kolom_header": grid_res["kolom_header"],
                    "kiri": grid_res["kiri"],
                    "kanan": grid_res["kanan"]
                })
                global_counter += 1
                continue

            # 3. Deteksi Soal Standar (Pilihan Ganda, Kotak Centang, Dropdown, Isian/Esai)
            if len(item) >= 5 and item[4]:
                q_text = item[1] or ""

                # Gambar pada pertanyaan utama
                q_img_url = find_image_url(item[1:4])
                q_img_obj = None
                if q_img_url:
                    stream = download_image_to_stream(q_img_url)
                    if stream:
                        q_img_obj = InlineImage(doc, stream, width=Mm(65))

                options = extract_options(item[4][0], doc)
                ada_opsi = len(options) > 0
                ada_opsi_bergambar = any(opt["ada_gambar"] for opt in options)

                # Tentukan tipe soal (Tipe 4 adalah Checkboxes / Pilihan Ganda Kompleks)
                if item_type == 4 and ada_opsi:
                    tipe_soal = "pg_kompleks"
                elif ada_opsi:
                    tipe_soal = "pilihan_ganda"
                else:
                    tipe_soal = "isian"

                nomor_di_bagian = len(current_section["soal"]) + 1

                current_section["soal"].append({
                    "no": nomor_di_bagian,
                    "no_global": global_counter,
                    "pertanyaan": q_text,
                    "tipe": tipe_soal,
                    "ada_gambar": q_img_obj is not None,
                    "gambar": q_img_obj,
                    "ada_opsi": ada_opsi,
                    "opsi_bergambar": ada_opsi_bergambar,
                    "opsi": options,
                    "baris": [],
                    "kolom_header": [],
                    "kiri": [],
                    "kanan": []
                })
                global_counter += 1

        if current_section["soal"]:
            daftar_bagian.append(current_section)

        flat_questions = [q for sec in daftar_bagian for q in sec["soal"]]

        return {
            "judul_form": form_title,
            "deskripsi_form": form_desc,
            "daftar_bagian": daftar_bagian,
            "daftar_pertanyaan": flat_questions
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Gagal memproses struktur formulir: {str(e)}")


# ==============================================================================
# 4. ENDPOINT API & PELAYANAN FILE
# ==============================================================================

@app.get("/", response_class=HTMLResponse)
async def serve_index():
    with open("templates/index.html", "r", encoding="utf-8") as f:
        return f.read()


@app.post("/api/generate")
async def generate_document(
    form_url: str = Form(...),
    template_file: UploadFile = File(...)
):
    # Bersihkan file lama (> 30 menit) agar ruang server tetap lega
    cleanup_old_files(STORAGE_DIR, max_age_seconds=1800)

    if not template_file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Berkas template harus berformat .docx (Microsoft Word).")

    session_id = str(uuid.uuid4())
    temp_template_path = os.path.join(STORAGE_DIR, f"temp_{session_id}.docx")
    output_docx_path = os.path.join(STORAGE_DIR, f"{session_id}.docx")

    try:
        with open(temp_template_path, "wb") as buffer:
            shutil.copyfileobj(template_file.file, buffer)

        doc = DocxTemplate(temp_template_path)
        context = extract_and_parse_form(form_url, doc)

        doc.render(context)
        # Kunci otomatis agar soal, opsi, dan tabel tidak terpotong antarhalaman
        lock_layout_prevent_page_split(doc.docx)
        doc.save(output_docx_path)

        pdf_path = convert_docx_to_pdf(output_docx_path, STORAGE_DIR)
        has_pdf = pdf_path is not None

        return {
            "status": "success",
            "message": "Dokumen siap cetak sudah selesai di-generate!" + (
                "" if has_pdf else " (Pratinjau PDF tidak aktif karena LibreOffice belum terpasang)."
            ),
            "file_id": session_id,
            "has_pdf": has_pdf,
            "preview_url": f"/api/files/{session_id}/pdf" if has_pdf else None,
            "download_docx": f"/api/files/{session_id}/docx",
            "download_pdf": f"/api/files/{session_id}/pdf" if has_pdf else None
        }

    finally:
        if os.path.exists(temp_template_path):
            os.remove(temp_template_path)


@app.get("/api/files/{file_id}/{ext}")
async def serve_file(file_id: str, ext: str):
    target_path = os.path.join(STORAGE_DIR, f"{file_id}.{ext}")
    if not os.path.exists(target_path):
        raise HTTPException(status_code=404, detail="Berkas sudah kedaluwarsa atau tidak ditemukan.")

    media_types = {
        "pdf": "application/pdf",
        "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }

    headers = (
        {"Content-Disposition": f"inline; filename=naskah_cetak.{ext}"}
        if ext == "pdf"
        else {"Content-Disposition": f"attachment; filename=naskah_cetak.{ext}"}
    )

    return FileResponse(target_path, media_type=media_types.get(ext, "application/octet-stream"), headers=headers)

# ==============================================================================
# 5. Penguncian Otomatis
# ==============================================================================

def lock_layout_prevent_page_split(docx_obj) -> None:
    """
    Mengunci paragraf soal dan baris tabel secara otomatis agar tidak 
    terpotong di batas pergantian kolom atau halaman.
    """
    # 1. Kunci paragraf pertanyaan agar selalu menempel dengan opsi/tabel di bawahnya
    for p in docx_obj.paragraphs:
        text = p.text.strip()
        p.paragraph_format.keep_together = True

        # Jika paragraf adalah judul bagian, nomor soal (misal "12. ..."), atau petunjuk:
        # aktifkan keep_with_next agar tidak menggantung sendirian di dasar halaman
        if (
            re.match(r"^\d+\.\s+", text)
            or text.startswith("Bagian ")
            or text.startswith("*(Pilihlah")
            or "w:drawing" in p._p.xml  # Jika paragraf memuat gambar soal
        ):
            p.paragraph_format.keep_with_next = True

    # 2. Kunci seluruh baris tabel agar tidak terbelah di tengah baris (cantSplit)
    for table in docx_obj.tables:
        for row_idx, row in enumerate(table.rows):
            trPr = row._tr.get_or_add_trPr()
            
            # Cegah baris tabel terpotong lintas halaman
            if trPr.find(qn("w:cantSplit")) is None:
                trPr.append(OxmlElement("w:cantSplit"))

            # Jika baris sebelum baris terakhir pada tabel kecil (< 4 baris),
            # tempelkan antar-baris agar opsi bergambar tidak terpisah
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_together = True
                    if len(table.rows) <= 3 and row_idx < len(table.rows) - 1:
                        p.paragraph_format.keep_with_next = True