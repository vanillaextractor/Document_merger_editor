#!/usr/bin/env python3
import os
import re
import subprocess
import pypandoc
from pathlib import Path
from docx import Document
import PyPDF2
from tqdm import tqdm
import click

class DocumentMerger:
    def __init__(self, input_dir="input", output_dir="output", temp_dir="converted"):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.temp_dir = Path(temp_dir)
        self.temp_dir.mkdir(exist_ok=True)
        self.output_dir.mkdir(exist_ok=True)
        self.chapters = []
    
    def scan_documents(self):
        """Scan input directory for all document files"""
        supported_formats = {'.pdf', '.docx', '.tex', '.md', '.doc'}
        files = list(self.input_dir.glob('*'))
        
        for file_path in sorted(files):
            if file_path.suffix.lower() in supported_formats:
                self.chapters.append({
                    'original': file_path,
                    'name': file_path.stem.replace('_', ' ').title(),
                    'format': file_path.suffix.lower()
                })
        print(f"📁 Found {len(self.chapters)} documents to merge")
    
    def create_header_file(self):
        """Create a LaTeX header file for emoji support and layout"""
        header_content = r"""
\usepackage{newunicodechar}
% \newfontfamily\emojifont{Apple Color Emoji}[Renderer=Harfbuzz]
% \newunicodechar{🔍}{{\emojifont 🔍}}
% \newunicodechar{🔎}{{\emojifont 🔎}}
% \newunicodechar{✏}{{\emojifont ✏}}
% \newunicodechar{️}{{\emojifont ️}}

% Center all figures
\makeatletter
\g@addto@macro\@floatboxreset\centering
\makeatother

% Increase spacing around figures
\setlength{\intextsep}{25pt plus 5pt minus 5pt}
\setlength{\textfloatsep}{25pt plus 5pt minus 5pt}
"""
        header_path = self.temp_dir / "header.tex"
        with open(header_path, 'w') as f:
            f.write(header_content)
        return header_path

    def convert_to_pdf(self, chapter):
        """Convert each document to PDF"""
        output_pdf = self.temp_dir / f"{chapter['name']}.pdf"
        header_path = self.create_header_file()
        
        if chapter['format'] == '.pdf':
            output_pdf = chapter['original']
        else:
            try:
                if chapter['format'] == '.docx':
                    pypandoc.convert_file(chapter['original'], 'pdf', 
                                        outputfile=str(output_pdf),
                                        extra_args=['--pdf-engine=xelatex', 
                                                  '-V', 'mainfont=Arial',
                                                  '-V', 'monofont=Courier New',
                                                  '-V', 'fontsize=12pt',
                                                  '-V', 'linestretch=1.5',
                                                  '-V', 'pagestyle=empty',  # Suppress chapter numbering
                                                  '-H', str(header_path),
                                                  '-V', 'geometry:margin=1in'])
                elif chapter['format'] == '.tex':
                    # For tex files, we'd need to insert the header inclusion manually or assume it's standalone
                    # defaulting to simple compilation for now as user input is docx
                    subprocess.run(['xelatex', '-interaction=nonstopmode', 
                                  '-output-directory', str(self.temp_dir),
                                  str(chapter['original'])], 
                                  capture_output=True)
                    output_pdf = self.temp_dir / f"{chapter['original'].stem}.pdf"
                else:
                    pypandoc.convert_file(chapter['original'], 'pdf', 
                                        outputfile=str(output_pdf),
                                        extra_args=['--pdf-engine=xelatex',
                                                  '-V', 'mainfont=Arial',
                                                  '-V', 'monofont=Courier New',
                                                  '-V', 'fontsize=12pt',
                                                  '-V', 'linestretch=1.5',
                                                  '-V', 'pagestyle=empty',  # Suppress chapter numbering
                                                  '-H', str(header_path),
                                                  '-V', 'geometry:margin=1in'])
                chapter['pdf'] = output_pdf
                print(f"✅ Converted: {chapter['name']}")
            except Exception as e:
                print(f"❌ Failed to convert {chapter['name']}: {e}")
        
        return chapter
    
    def get_pdf_bookmarks(self, pdf_path):
        """Extract bookmarks and try to find the full Chapter title"""
        bookmarks = []
        found_chapter_title = None
        
        try:
            reader = PyPDF2.PdfReader(pdf_path)
            
            def recurse_outlines(outlines):
                nonlocal found_chapter_title
                for item in outlines:
                    if isinstance(item, list):
                        recurse_outlines(item)
                    else:
                        title = item.title.replace(',', ' ').strip()
                        
                        # Check for Chapter title (Chapter X: ... or CH-X ...)
                        # Supports: "Chapter 1", "CH-6", "Ch 6"
                        match = re.match(r'^(Chapter|CH)[-:\s]+\d+[:\.\s-]*', title, re.IGNORECASE)
                        if match:
                            # Strip the "Chapter X" prefix to get clean title: "NPCYF Introduction"
                            # This avoids TOC looking like "Chapter 1 Chapter 1: NPCYF..."
                            found_chapter_title = re.sub(r'^(Chapter|CH)[-:\s]+\d+[:\.\s-]*', '', title, flags=re.IGNORECASE).strip()
                            continue # Don't add chapter title as a section bookmark
                        
                        # Fallback: Check for Title starting with Number (e.g. "2. Introduction")
                        # Must NOT be a section (i.e., no internal dots like 1.1)
                        elif re.match(r'^\d+\.?\s+', title):
                             first_part = title.split()[0].rstrip('.')
                             if '.' not in first_part:
                                 # Strip leading "2. " or "2 " to avoid "Chapter 2 2. Introduction"
                                 found_chapter_title = re.sub(r'^\d+\.?\s*', '', title).strip()
                                 continue
                            
                        # Strict Regex: Match "X.Y" or "X.Y.Z" followed by optional dot, then space or end
                        # ^\d+\.\d+ matches 1.1
                        # (\.\d+)? matches optional .1
                        # \.? matches optional trailing dot (e.g. 1.3.)
                        # (\s|$) matches space or end
                        if re.match(r'^\d+\.\d+(\.\d+)?\.?(\s|$)', title):
                            # PyPDF2 pages are 0-indexed, addtotoc needs 1-indexed relative to the included doc
                            page_num = reader.get_destination_page_number(item) + 1
                            
                            # Determine level based on dots in the NUMBER part only
                            number_part = title.split(' ')[0].rstrip('.')
                            dot_count = number_part.count('.')
                            
                            if dot_count == 1:
                                level = "section"
                                depth = 1
                            else:
                                level = "subsection"
                                depth = 2
                            
                            
                            # Page, section/subsection, depth, Title, label
                            # Strip the leading number from the title for addtotoc to avoid double numbering
                            # e.g. "1.1 Introduction" -> "Introduction"
                            # The regex matched "X.Y" or "X.Y.Z" at the start
                            clean_title = re.sub(r'^\d+\.\d+(\.\d+)?\.?\s*', '', title).strip()
                            bookmarks.append(f"{page_num},{level},{depth},{clean_title},label{len(bookmarks)}")
            
            if reader.outline:
                recurse_outlines(reader.outline)
                
            # Fallback: If no chapter title found in bookmarks, check Page 1 text
            if not found_chapter_title and len(reader.pages) > 0:
                try:
                    first_page_text = reader.pages[0].extract_text()
                    lines = first_page_text.splitlines()
                    # Check first 5 lines
                    for line in lines[:5]:
                        line = line.strip()
                        # Match "Chapter 2: Title" or "Chapter 2"
                        match = re.match(r'^(Chapter|CH)[-:\s]+\d+[:\.\s-]*', line, re.IGNORECASE)
                        if match:
                             # Found it! Strip prefix
                             found_chapter_title = re.sub(r'^(Chapter|CH)[-:\s]+\d+[:\.\s-]*', '', line, flags=re.IGNORECASE).strip()
                             # If stripping resulted in empty string (e.g. title was just "Chapter 2"), keep original or handle?
                             # Usually there is a title. If empty, maybe look at next line?
                             if not found_chapter_title:
                                 # Edge case: "Chapter 2" on one line, "Title" on next?
                                 # For now, let's just use "Chapter X" if title is empty, or try next line logic later.
                                 # But user wants clean titles. If just "Chapter 2", we can't do much.
                                 found_chapter_title = line
                             print(f"  [Fallback] Found title in text: '{found_chapter_title}'")
                             break
                except Exception as ex:
                    print(f"  [Fallback] Text extraction failed: {ex}")

        except Exception as e:
            print(f"⚠️ Could not read bookmarks from {pdf_path.name}: {e}")
            
        return ", ".join(bookmarks), found_chapter_title

    def extract_keywords(self, keywords_file="keywords.txt"):
        """Extract keywords from PDFs and track their page numbers"""
        if not os.path.exists(keywords_file):
            print(f"ℹ️  No {keywords_file} found. Skipping index generation.")
            return None

        print(f"🔍 Reading keywords from {keywords_file}...")
        with open(keywords_file, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Split by comma or newline to handle both formats
        keywords = [k.strip() for k in re.split(r'[,\n]', content) if k.strip()]
        
        if not keywords:
            print("⚠️  Keywords file is empty.")
            return None

        # distinct keywords result found
        # Map: keyword -> set of page numbers
        keyword_map = {k: set() for k in keywords}
        
        global_page = 1
        
        print("Wait... scanning for keywords in document text (this might take a moment)")
        
        for chapter in tqdm(self.chapters, desc="Indexing Chapters"):
            if 'pdf' not in chapter:
                continue
                
            try:
                reader = PyPDF2.PdfReader(chapter['pdf'])
                num_pages = len(reader.pages)
                
                for i in range(num_pages):
                    page = reader.pages[i]
                    text = page.extract_text().lower()
                    
                    # Search for each keyword
                    current_page_num = global_page + i
                    
                    for kw in keywords:
                        if kw.lower() in text:
                            keyword_map[kw].add(current_page_num)
                
                global_page += num_pages
                
            except Exception as e:
                print(f"⚠️  Error reading {chapter['name']} for indexing: {e}")
        
        # Filter out keywords that weren't found
        found_keywords = {k: sorted(list(v)) for k, v in keyword_map.items() if v}
        print(f"✅ Found {len(found_keywords)} unique keywords in the text.")
        return found_keywords

    def generate_index_pdf(self, keyword_map):
        """Generate a PDF page for the Index"""
        if not keyword_map:
            return None
            
        index_tex = self.temp_dir / "index.tex"
        
        # Sort keywords alphabetically
        sorted_keys = sorted(keyword_map.keys(), key=lambda x: x.lower())
        
        latex_content = r"""
\documentclass[12pt,a4paper]{article}
\usepackage[margin=1in]{geometry}
\usepackage{multicol}
\usepackage{hyperref}
\usepackage{fancyhdr}

\pagestyle{fancy}
\fancyhf{}
\lhead{\textbf{Index}}
\rhead{\thepage}
\cfoot{\textbf{Institute of Data Engineering Analytics and Science}}

\title{\textbf{Index}}
\date{}

\begin{document}
\section*{Index}
\begin{multicols}{2}
\begin{description}
"""
        
        for kw in sorted_keys:
            pages = ", ".join(map(str, keyword_map[kw]))
            # Escape LaTeX special chars in keyword if necessary (basic ones)
            safe_kw = kw.replace('_', '\\_').replace('%', '\\%').replace('&', '\\&')
            latex_content += f"    \\item[{safe_kw}] {pages}\n"
            
        latex_content += r"""
\end{description}
\end{multicols}
\end{document}
"""
        
        with open(index_tex, 'w') as f:
            f.write(latex_content)
            
        # Compile index PDF
        print("⚙️  Compiling Index...")
        subprocess.run(['xelatex', '-interaction=nonstopmode', 
                      '-output-directory', str(self.temp_dir),
                      str(index_tex)], capture_output=True)
                      
        index_pdf = self.temp_dir / "index.pdf"
        if index_pdf.exists():
            return index_pdf
        else:
            print("❌ Failed to compile Index PDF")
            return None

    def create_master_latex(self, index_pdf=None, title="MERGED DOCUMENT"):
        """Create LaTeX master document with TOC"""
        latex_content = r"""
\documentclass[12pt,a4paper]{book}
\usepackage{fontspec}
\usepackage[margin=1in]{geometry}
\usepackage{pdfpages}
\usepackage{hyperref}
\usepackage{bookmark}
\usepackage{fancyhdr}
\usepackage{graphicx}
\usepackage{newunicodechar}

% Emoji support
% \newfontfamily\emojifont{Apple Color Emoji}[Renderer=Harfbuzz]
% \newunicodechar{🔍}{{\emojifont 🔍}}
% \newunicodechar{🔎}{{\emojifont 🔎}}
% \newunicodechar{✏}{{\emojifont ✏}}
% \newunicodechar{️}{{\emojifont ️}}

\hypersetup{
    colorlinks=true,
    linkcolor=blue,
}

\pagestyle{fancy}
\fancyhf{}
\lhead{\textbf{NPCYF Documentation}}
\rhead{\thepage}
\cfoot{\includegraphics[height=0.8cm]{ideas_logo.png} \hspace{10pt} \textbf{Institute of Data Engineering Analytics and Science}}

\cfoot{\includegraphics[height=0.8cm]{ideas_logo.png} \hspace{10pt} \textbf{Institute of Data Engineering Analytics and Science}}

\title{\textbf{""" + title + r"""}}
\date{\today}

\begin{document}
\maketitle
\tableofcontents
\newpage
"""
        
        for i, chapter in enumerate(self.chapters, 1):
            if 'pdf' in chapter:
                pdf_path = chapter['pdf'].resolve().as_posix()
                
                # generate addtotoc for sections/subsections
                addtotoc_str, found_title = self.get_pdf_bookmarks(chapter['pdf'])
                
                # Use found extracted title if available, else filename
                chapter_title = found_title if found_title else chapter['name']
                
                # Always add the Chapter entry first
                # Page 1, chapter, 0, Name, label
                chapter_entry = f"1,chapter,0,{chapter_title},chap{i}"
                
                if addtotoc_str:
                    full_addtotoc = f"{chapter_entry}, {addtotoc_str}"
                else:
                    full_addtotoc = chapter_entry
                
                latex_content += f"""
\\includepdf[pages=-, pagecommand={{}}, addtotoc={{{full_addtotoc}}}]{{{pdf_path}}}
"""
        
        # Append Index if available
        if index_pdf:
            index_path = index_pdf.resolve().as_posix()
            latex_content += f"""
\\includepdf[pages=-, pagecommand={{}}, addtotoc={{1,chapter,0,Index,idx}}]{{{index_path}}}
"""

        latex_content += r"""\end{document}"""
        
        master_tex = self.temp_dir / "master.tex"
        with open(master_tex, 'w') as f:
            f.write(latex_content)
        return master_tex
    
    def compile_final_pdf(self, master_tex, final_name="FINAL_MERGED_DOCUMENT"):
        """Compile LaTeX master to final PDF"""
        # Ensure extension is handled
        if not final_name.lower().endswith('.pdf'):
            final_name += ".pdf"
            
        output_pdf = self.output_dir / final_name
        
        # Compile twice for TOC
        for i in range(2):
            subprocess.run([
                'xelatex', '-interaction=nonstopmode', 
                '-output-directory', str(self.output_dir),
                str(master_tex)
            ], capture_output=True)
            
        # Rename the output file (xelatex outputs as master.pdf)
        generated_pdf = self.output_dir / "master.pdf"
        if generated_pdf.exists():
            if output_pdf.exists():
                output_pdf.unlink()
            generated_pdf.rename(output_pdf)
            print(f"🎉 Final document: {output_pdf}")
        else:
            print(f"❌ Failed to generate PDF. Check log at {self.output_dir}/master.log")
            
        return output_pdf

@click.command()
@click.option('--name', default=None, help='Name for the final PDF document')
def main(name):
    merger = DocumentMerger()
    print("🔍 Scanning documents...")
    merger.scan_documents()
    
    if not merger.chapters:
        print("❌ No documents in input/ folder!")
        print("📂 Put your PDF/DOCX/TEX files in input/ first")
    else:
        # Determine output filename
        custom_name = name
        if not custom_name:
            # Interactive fallback if no argument provided
            print("-" * 50)
            custom_name = input("Enter name for final PDF (default: FINAL_MERGED_DOCUMENT): ").strip()
            if not custom_name:
                custom_name = "FINAL_MERGED_DOCUMENT"
            print("-" * 50)
        else:
             print(f"📄 Output name provided: {custom_name}")

        print("\n🔄 Converting...")
        for chapter in merger.chapters:
            merger.convert_to_pdf(chapter)
        
        # Keyword Indexing Step
        print("\n📇 Indexing...")
        keyword_map = merger.extract_keywords()
        index_pdf = merger.generate_index_pdf(keyword_map)
        
        print("\n📚 Creating master document...")
        # Use custom_name for title, replacing underscores with spaces
        doc_title = custom_name.replace("_", " ")
        master_tex = merger.create_master_latex(index_pdf, doc_title)
        
        print("\n⚙️  Compiling...")
        final_pdf = merger.compile_final_pdf(master_tex, custom_name)
        print(f"\n✅ SUCCESS! {final_pdf}")

if __name__ == '__main__':
    main()
