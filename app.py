import io
import os
import re
import time
import zipfile
import mimetypes
from urllib.parse import urlparse, unquote

import pandas as pd
import requests
import streamlit as st
from PyPDF2 import PdfMerger, PdfReader


APP_TITLE = "Olympiad Smart Bulk Downloader"

URL_PATTERN = re.compile(r"https?://[^\s\]\)\}\<\"']+", re.IGNORECASE)


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def clean_url(url):
    return str(url).strip().rstrip(".,);]}'\"")


def extract_urls_from_text(text):
    if not text:
        return []

    urls = URL_PATTERN.findall(str(text))
    cleaned = [clean_url(url) for url in urls]
    return list(dict.fromkeys(cleaned))


def normalize_column_name(value):
    return re.sub(r"\s+", " ", str(value).strip())


def make_safe_filename(value):
    value = unquote(str(value))
    value = re.sub(r"[\\/*?:\"<>|]", "_", value)
    value = re.sub(r"\s+", "_", value)
    value = value.strip("._ ")

    if not value:
        value = "file"

    return value[:150]


def find_header_row(raw_df):
    """
    Finds the row where actual headers exist.
    Your Excel has this row:
    # | Exam / Source | Year | Paper Type | Question Paper PDF URL | Solutions PDF URL | Notes / Status
    """
    for row_index in range(min(len(raw_df), 20)):
        row_values = [
            normalize_column_name(x).lower()
            for x in raw_df.iloc[row_index].tolist()
        ]

        row_text = " | ".join(row_values)

        if "question paper pdf url" in row_text and "solutions pdf url" in row_text:
            return row_index

        if "question" in row_text and "url" in row_text and "solution" in row_text:
            return row_index

    return None


def read_excel_as_bytes(uploaded_file):
    """
    Important fix:
    Streamlit uploaded files are file-like objects.
    Reading them multiple times can fail unless we save bytes first.
    """
    return uploaded_file.getvalue()


def read_sheet_with_detected_header(excel_bytes, sheet_name):
    raw_df = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=sheet_name,
        header=None,
        dtype=str
    )

    header_row = find_header_row(raw_df)

    if header_row is None:
        return pd.DataFrame(), None

    df = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=sheet_name,
        header=header_row,
        dtype=str
    )

    df.columns = [normalize_column_name(c) for c in df.columns]
    df = df.dropna(how="all")

    return df, header_row


def find_column(df, possible_names):
    """
    Finds matching column by exact or partial matching.
    """
    existing_cols = list(df.columns)

    lower_map = {
        normalize_column_name(col).lower(): col
        for col in existing_cols
    }

    for name in possible_names:
        key = normalize_column_name(name).lower()
        if key in lower_map:
            return lower_map[key]

    for col in existing_cols:
        col_lower = normalize_column_name(col).lower()

        for name in possible_names:
            name_lower = normalize_column_name(name).lower()

            if name_lower in col_lower:
                return col

    return None


def fallback_extract_all_urls(excel_bytes, sheet_name):
    """
    If headers are not detected, this still scans the whole sheet and extracts URLs.
    It will not have perfect exam/year names, but it prevents 'no URLs found'.
    """
    raw_df = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=sheet_name,
        header=None,
        dtype=str
    )

    rows = []

    for row_index in range(len(raw_df)):
        row_values = raw_df.iloc[row_index].tolist()
        row_text = " ".join([clean_text(x) for x in row_values])
        urls = extract_urls_from_text(row_text)

        if not urls:
            continue

        for url in urls:
            url_lower = url.lower()

            if "sol" in url_lower or "solution" in url_lower:
                file_type = "Solution"
            else:
                file_type = "Question"

            rows.append({
                "Subject": sheet_name.title(),
                "Exam": "Unknown Exam",
                "Year": "Unknown Year",
                "File Type": file_type,
                "URL": url,
                "Notes": "Extracted by fallback scanner"
            })

    return rows


def build_download_table(excel_bytes, debug=False):
    excel = pd.ExcelFile(io.BytesIO(excel_bytes))
    all_rows = []
    debug_rows = []

    for sheet_name in excel.sheet_names:
        if sheet_name.strip().upper() == "SUMMARY":
            continue

        df, header_row = read_sheet_with_detected_header(excel_bytes, sheet_name)

        debug_rows.append({
            "Sheet": sheet_name,
            "Header Row Found": header_row + 1 if header_row is not None else "Not found",
            "Rows Read": len(df) if not df.empty else 0,
            "Columns": ", ".join(df.columns.astype(str).tolist()) if not df.empty else ""
        })

        if df.empty:
            fallback_rows = fallback_extract_all_urls(excel_bytes, sheet_name)
            all_rows.extend(fallback_rows)
            continue

        exam_col = find_column(df, ["Exam / Source", "Exam", "Source"])
        year_col = find_column(df, ["Year"])
        paper_type_col = find_column(df, ["Paper Type"])
        question_col = find_column(df, [
            "Question Paper PDF URL",
            "Question Paper URL",
            "Question PDF URL",
            "Question URL",
            "Paper URL"
        ])
        solution_col = find_column(df, [
            "Solutions PDF URL",
            "Solution PDF URL",
            "Solutions URL",
            "Solution URL",
            "Answer URL"
        ])
        notes_col = find_column(df, [
            "Notes / Status",
            "Notes",
            "Status"
        ])

        for _, row in df.iterrows():
            subject = sheet_name.title()

            exam = clean_text(row.get(exam_col, "")) if exam_col else "Unknown Exam"
            year = clean_text(row.get(year_col, "")) if year_col else "Unknown Year"
            notes = clean_text(row.get(notes_col, "")) if notes_col else ""

            if not exam:
                exam = "Unknown Exam"

            if not year:
                year = "Unknown Year"

            if question_col:
                question_text = clean_text(row.get(question_col, ""))
                question_urls = extract_urls_from_text(question_text)

                for url in question_urls:
                    all_rows.append({
                        "Subject": subject,
                        "Exam": exam,
                        "Year": year,
                        "File Type": "Question",
                        "URL": url,
                        "Notes": notes
                    })

            if solution_col:
                solution_text = clean_text(row.get(solution_col, ""))
                solution_urls = extract_urls_from_text(solution_text)

                for url in solution_urls:
                    all_rows.append({
                        "Subject": subject,
                        "Exam": exam,
                        "Year": year,
                        "File Type": "Solution",
                        "URL": url,
                        "Notes": notes
                    })

            # Extra safety: scan whole row for missed URLs
            whole_row_text = " ".join([clean_text(x) for x in row.tolist()])
            row_urls = extract_urls_from_text(whole_row_text)

            known_urls = set()
            if question_col:
                known_urls.update(extract_urls_from_text(clean_text(row.get(question_col, ""))))
            if solution_col:
                known_urls.update(extract_urls_from_text(clean_text(row.get(solution_col, ""))))

            for url in row_urls:
                if url in known_urls:
                    continue

                url_lower = url.lower()
                if "sol" in url_lower or "solution" in url_lower:
                    file_type = "Solution"
                else:
                    file_type = "Question"

                all_rows.append({
                    "Subject": subject,
                    "Exam": exam,
                    "Year": year,
                    "File Type": file_type,
                    "URL": url,
                    "Notes": notes
                })

    result_df = pd.DataFrame(all_rows)

    if not result_df.empty:
        result_df = result_df.drop_duplicates(
            subset=["Subject", "Exam", "Year", "File Type", "URL"]
        )
        result_df = result_df.sort_values(
            ["Subject", "Exam", "Year", "File Type"]
        ).reset_index(drop=True)

    debug_df = pd.DataFrame(debug_rows)

    return result_df, debug_df


def get_extension_from_url_or_content_type(url, content_type):
    parsed = urlparse(url)
    path = parsed.path
    filename = os.path.basename(path)

    if "." in filename:
        ext = os.path.splitext(filename)[1]
        if ext:
            return ext

    if content_type:
        guessed = mimetypes.guess_extension(content_type.split(";")[0].strip())
        if guessed:
            return guessed

    return ".pdf"


def build_filename(item, content_type=""):
    subject = make_safe_filename(item.get("Subject", "Subject"))
    exam = make_safe_filename(item.get("Exam", "Exam"))
    year = make_safe_filename(item.get("Year", "Year"))
    file_type = make_safe_filename(item.get("File Type", "File"))

    ext = get_extension_from_url_or_content_type(item.get("URL", ""), content_type)

    filename = f"{subject}_{exam}_{year}_{file_type}{ext}"
    return make_safe_filename(filename.replace(ext, "")) + ext


def download_file(item):
    url = item["URL"]

    headers = {
        "User-Agent": "Mozilla/5.0 OlympiadBulkDownloader/1.0",
        "Accept": "application/pdf,application/octet-stream,text/html,*/*"
    }

    try:
        response = requests.get(
            url,
            headers=headers,
            timeout=45,
            allow_redirects=True
        )

        if response.status_code != 200:
            return {
                **item,
                "Status": "Failed",
                "Filename": "",
                "Error": f"HTTP {response.status_code}",
                "Content": None
            }

        content_type = response.headers.get("Content-Type", "")
        filename = build_filename(item, content_type)

        return {
            **item,
            "Status": "Success",
            "Filename": filename,
            "Error": "",
            "Content": response.content
        }

    except requests.exceptions.Timeout:
        return {
            **item,
            "Status": "Failed",
            "Filename": "",
            "Error": "Timeout",
            "Content": None
        }

    except requests.exceptions.RequestException as e:
        return {
            **item,
            "Status": "Failed",
            "Filename": "",
            "Error": str(e),
            "Content": None
        }


def add_file_to_zip(zip_file, used_paths, zip_path, content):
    base, ext = os.path.splitext(zip_path)
    final_path = zip_path
    counter = 1

    while final_path in used_paths:
        final_path = f"{base}_{counter}{ext}"
        counter += 1

    used_paths.add(final_path)
    zip_file.writestr(final_path, content)
    return final_path


def create_report_excel(results):
    rows = []

    for item in results:
        row = {k: v for k, v in item.items() if k != "Content"}
        rows.append(row)

    report_df = pd.DataFrame(rows)

    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Download Report")

    output.seek(0)
    return output.read()


def merge_pdfs(pdf_items):
    merger = PdfMerger()
    added_count = 0

    for item in pdf_items:
        content = item.get("Content")

        if not content:
            continue

        try:
            PdfReader(io.BytesIO(content))
            merger.append(io.BytesIO(content))
            added_count += 1
        except Exception:
            continue

    if added_count == 0:
        return None

    output = io.BytesIO()
    merger.write(output)
    merger.close()
    output.seek(0)

    return output.read()


def extract_text_from_pdf(content):
    try:
        reader = PdfReader(io.BytesIO(content))
        text_parts = []

        for page_number, page in enumerate(reader.pages, start=1):
            page_text = page.extract_text() or ""
            text_parts.append(f"\n\n--- Page {page_number} ---\n{page_text}")

        return "".join(text_parts).strip()

    except Exception as e:
        return f"Text extraction failed: {e}"


def create_zip(results, include_originals=True, include_merged=False, include_text=False):
    zip_buffer = io.BytesIO()
    used_paths = set()

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
        report_excel = create_report_excel(results)
        zip_file.writestr("download_report.xlsx", report_excel)

        successful_files = [
            item for item in results
            if item.get("Status") == "Success" and item.get("Content")
        ]

        if include_originals:
            for item in successful_files:
                subject = make_safe_filename(item.get("Subject", "Subject"))
                exam = make_safe_filename(item.get("Exam", "Exam"))
                year = make_safe_filename(item.get("Year", "Year"))
                file_type = make_safe_filename(item.get("File Type", "File"))
                filename = item.get("Filename", "file.pdf")

                zip_path = f"{subject}/{exam}/{year}_{file_type}_{filename}"

                final_path = add_file_to_zip(
                    zip_file,
                    used_paths,
                    zip_path,
                    item["Content"]
                )

                item["Zip Path"] = final_path

        if include_text:
            for item in successful_files:
                filename = item.get("Filename", "").lower()

                if not filename.endswith(".pdf"):
                    continue

                text = extract_text_from_pdf(item["Content"])

                subject = make_safe_filename(item.get("Subject", "Subject"))
                exam = make_safe_filename(item.get("Exam", "Exam"))
                year = make_safe_filename(item.get("Year", "Year"))
                file_type = make_safe_filename(item.get("File Type", "File"))

                text_path = f"Text_Extracted/{subject}/{exam}/{year}_{file_type}.txt"

                add_file_to_zip(
                    zip_file,
                    used_paths,
                    text_path,
                    text.encode("utf-8", errors="ignore")
                )

        if include_merged:
            for subject in sorted(set(item["Subject"] for item in successful_files)):
                subject_pdf_items = [
                    item for item in successful_files
                    if item.get("Subject") == subject
                    and item.get("Filename", "").lower().endswith(".pdf")
                ]

                merged_pdf = merge_pdfs(subject_pdf_items)

                if merged_pdf:
                    merged_path = f"Merged_PDFs/{make_safe_filename(subject)}_merged.pdf"
                    zip_file.writestr(merged_path, merged_pdf)

            all_pdf_items = [
                item for item in successful_files
                if item.get("Filename", "").lower().endswith(".pdf")
            ]

            all_merged_pdf = merge_pdfs(all_pdf_items)

            if all_merged_pdf:
                zip_file.writestr("Merged_PDFs/ALL_SUBJECTS_merged.pdf", all_merged_pdf)

    zip_buffer.seek(0)
    return zip_buffer.read()


def show_metrics(df):
    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Total URLs", len(df))
    col2.metric("Subjects", df["Subject"].nunique())
    col3.metric("Exams", df["Exam"].nunique())
    col4.metric("Years", df["Year"].nunique())


def main():
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="📚",
        layout="wide"
    )

    st.title("📚 Olympiad Smart Bulk Downloader")

    st.write(
        "Upload your Excel file with question paper and solution URLs. "
        "This app will download all selected files in one ZIP."
    )

    with st.sidebar:
        st.header("Output Options")

        include_originals = st.checkbox(
            "Original PDFs/files in folders",
            value=True
        )

        include_merged = st.checkbox(
            "Also create merged PDF by subject",
            value=False
        )

        include_text = st.checkbox(
            "Also extract text from PDFs",
            value=False
        )

        st.divider()

        show_debug = st.checkbox(
            "Show debug details",
            value=False
        )

        st.caption(
            "Best default: keep only 'Original PDFs/files in folders' selected."
        )

    uploaded_file = st.file_uploader(
        "Upload Excel file",
        type=["xlsx", "xls"]
    )

    if uploaded_file is None:
        st.info("Upload your Excel file to start.")
        return

    excel_bytes = read_excel_as_bytes(uploaded_file)

    try:
        download_df, debug_df = build_download_table(
            excel_bytes,
            debug=show_debug
        )
    except Exception as e:
        st.error(f"Could not read Excel file: {e}")
        return

    if show_debug:
        st.subheader("Debug Details")
        st.dataframe(debug_df, use_container_width=True)

    if download_df.empty:
        st.error(
            "No downloadable URLs were found. "
            "The app scanned the workbook but could not find links."
        )

        st.write("Please check that your Excel has links starting with `http://` or `https://`.")

        if show_debug:
            st.write("Debug table above shows which sheets were scanned.")

        return

    st.success(f"Detected {len(download_df)} downloadable URLs.")

    show_metrics(download_df)

    st.divider()

    subjects = sorted(download_df["Subject"].dropna().unique().tolist())
    file_types = sorted(download_df["File Type"].dropna().unique().tolist())
    exams = sorted(download_df["Exam"].dropna().unique().tolist())

    col1, col2, col3 = st.columns(3)

    with col1:
        selected_subjects = st.multiselect(
            "Subjects",
            subjects,
            default=subjects
        )

    with col2:
        selected_file_types = st.multiselect(
            "File Types",
            file_types,
            default=file_types
        )

    with col3:
        selected_exams = st.multiselect(
            "Exams",
            exams,
            default=exams
        )

    filtered_df = download_df[
        download_df["Subject"].isin(selected_subjects)
        & download_df["File Type"].isin(selected_file_types)
        & download_df["Exam"].isin(selected_exams)
    ].reset_index(drop=True)

    st.subheader("Preview Selected Files")
    st.write(f"Selected files: **{len(filtered_df)}**")

    st.dataframe(
        filtered_df,
        use_container_width=True,
        hide_index=True
    )

    if filtered_df.empty:
        st.warning("No files match your selected filters.")
        return

    if not include_originals and not include_merged and not include_text:
        st.warning("Please select at least one output option from the sidebar.")
        return

    st.divider()

    if st.button("⬇️ Download Selected Files as ZIP", type="primary"):
        records = filtered_df.to_dict("records")
        total = len(records)

        progress = st.progress(0)
        status_text = st.empty()

        results = []

        for index, item in enumerate(records, start=1):
            status_text.write(
                f"Downloading {index}/{total}: "
                f"{item['Subject']} | {item['Exam']} | {item['Year']} | {item['File Type']}"
            )

            result = download_file(item)
            results.append(result)

            progress.progress(index / total)
            time.sleep(0.05)

        success_count = sum(1 for item in results if item.get("Status") == "Success")
        failed_count = total - success_count

        status_text.write("Creating ZIP file...")

        zip_bytes = create_zip(
            results,
            include_originals=include_originals,
            include_merged=include_merged,
            include_text=include_text
        )

        st.success(
            f"Download process completed. Successful: {success_count}. Failed: {failed_count}."
        )

        report_df = pd.DataFrame([
            {k: v for k, v in item.items() if k != "Content"}
            for item in results
        ])

        st.subheader("Download Report")
        st.dataframe(report_df, use_container_width=True, hide_index=True)

        st.download_button(
            label="⬇️ Download ZIP File",
            data=zip_bytes,
            file_name="olympiad_downloads.zip",
            mime="application/zip"
        )


if __name__ == "__main__":
    main()