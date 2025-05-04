import streamlit as st
import os
from formatter import parse_docx_transcript, parse_vtt_transcript, format_transcript

st.set_page_config(page_title="MS Teams Transcript Formatter")

st.title("🎧 MS Teams Transcript Formatter for ICF Evaluation")

uploaded_file = st.file_uploader("Upload MS Teams .vtt or .docx file", type=["vtt", "docx"])

if uploaded_file:
    # Save uploaded file temporarily
    file_path = os.path.join("temp", uploaded_file.name)
    os.makedirs("temp", exist_ok=True)
    with open(file_path, "wb") as f:
        f.write(uploaded_file.read())

    # Determine file type and parse
    if uploaded_file.name.endswith(".vtt"):
        entries = parse_vtt_transcript(file_path)
    elif uploaded_file.name.endswith(".docx"):
        entries = parse_docx_transcript(file_path)
    else:
        st.error("Unsupported file format.")
        st.stop()

    if not entries:
        st.error("No transcript entries found. Please check the file formatting.")
        st.stop()

    # Coach selection
    speakers = sorted({e[0] for e in entries})
    coach_name = st.selectbox("Select the Coach's Name", speakers)

    # Enter output filename
    default_name = os.path.splitext(uploaded_file.name)[0] + "_Formatted.docx"
    output_filename = st.text_input("Output Filename", value=default_name)

    if st.button("Format and Download"):
        output_path = os.path.join("temp", output_filename)
        format_transcript(entries, coach_name=coach_name, output_path=output_path)

        with open(output_path, "rb") as f:
            st.download_button("Download Formatted Transcript", f, file_name=output_filename)

        st.success("Transcript formatted successfully.")