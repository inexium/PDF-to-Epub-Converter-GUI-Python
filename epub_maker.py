import customtkinter as ctk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import os
import io
from PIL import Image
import tempfile
import shutil
from ebooklib import epub
from datetime import datetime


class PDFtoEPUBConverter:
    def __init__(self):
        # Set the appearance mode and color theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        self.window = ctk.CTk()
        self.window.title("PDF to EPUB Converter")
        self.window.geometry("500x300")

        # Create and configure GUI elements
        self.setup_gui()

    def setup_gui(self):
        # Main frame
        main_frame = ctk.CTkFrame(self.window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Title Label
        title_label = ctk.CTkLabel(
            main_frame,
            text="PDF to EPUB Converter",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        title_label.pack(pady=20)

        # Description Label
        desc_label = ctk.CTkLabel(
            main_frame,
            text="Convert your PDF files to EPUB format",
            font=ctk.CTkFont(size=14)
        )
        desc_label.pack(pady=10)

        # Convert Button
        convert_button = ctk.CTkButton(
            main_frame,
            text="Select PDF and Convert",
            command=self.convert_pdf_to_epub,
            font=ctk.CTkFont(size=16),
            width=200,
            height=40
        )
        convert_button.pack(pady=20)

    def extract_images_from_page(self, page):
        """Extract images from a page using multiple methods."""
        try:
            # Method 1: Try to get embedded images first
            image_list = page.get_images()
            if image_list:
                # Get the first image (assuming one image per page)
                img_index = 0
                xref = image_list[img_index][0]
                base_image = self.pdf_document.extract_image(xref)

                if base_image:
                    # Convert the raw image data to PIL Image
                    image_bytes = base_image["image"]
                    img = Image.open(io.BytesIO(image_bytes))

                    # Convert to RGB if necessary
                    if img.mode != 'RGB':
                        img = img.convert('RGB')

                    # Save to bytes
                    img_bytes = io.BytesIO()
                    img.save(img_bytes, format='JPEG', quality=95)
                    return img_bytes.getvalue()

            # Method 2: If no embedded images, render page as image
            zoom = 2.0  # Increase for higher quality
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat, alpha=False)

            # Convert pixmap to PIL Image
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

            # Save to bytes
            img_bytes = io.BytesIO()
            img.save(img_bytes, format='JPEG', quality=95)
            return img_bytes.getvalue()

        except Exception as e:
            print(f"Error in extract_images_from_page: {str(e)}")
            return None

    def create_epub(self, title, images, output_file):
        """Create EPUB file from images."""
        # Create temporary directory for EPUB contents
        with tempfile.TemporaryDirectory() as temp_dir:
            # Create EPUB structure
            epub_dir = os.path.join(temp_dir, 'epub')
            meta_inf_dir = os.path.join(epub_dir, 'META-INF')
            os.makedirs(meta_inf_dir, exist_ok=True)

            # Create mimetype file
            with open(os.path.join(epub_dir, 'mimetype'), 'w', encoding='utf-8') as f:
                f.write('application/epub+zip')

            # Create container.xml
            container_xml = '''<?xml version="1.0"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
   <rootfiles>
      <rootfile full-path="content.opf" media-type="application/oebps-package+xml"/>
   </rootfiles>
</container>'''
            with open(os.path.join(meta_inf_dir, 'container.xml'), 'w', encoding='utf-8') as f:
                f.write(container_xml)

            # Create CSS files
            page_styles_css = '''@page {
  margin-bottom: 5pt;
  margin-top: 5pt;
}'''
            with open(os.path.join(epub_dir, 'page_styles.css'), 'w', encoding='utf-8') as f:
                f.write(page_styles_css)

            stylesheet_css = '''.calibre {
  display: block;
  font-size: 1em;
  padding-left: 0;
  padding-right: 0;
  margin: 0 5pt;
}
.calibre1 {
  display: block;
  margin: 0 0;
}
.calibre2 {
  height: auto;
  width: auto;
}'''
            with open(os.path.join(epub_dir, 'stylesheet.css'), 'w', encoding='utf-8') as f:
                f.write(stylesheet_css)

            # Save images with proper naming
            html_content = ['<?xml version=\'1.0\' encoding=\'utf-8\'?>',
                          '<html xmlns="http://www.w3.org/1999/xhtml" lang="en" xml:lang="en">',
                          '  <head>',
                          '    <title>index</title>',
                          '    <meta name="generator" content="pdftohtml 0.36"/>',
                          '    <meta name="author" content="python-docx"/>',
                          f'    <meta name="date" content="{datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")}"/>',
                          '    <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>',
                          '  <link rel="stylesheet" type="text/css" href="stylesheet.css"/>',
                          '<link rel="stylesheet" type="text/css" href="page_styles.css"/>',
                          '</head>',
                          '  <body class="calibre">']

            for i, img_data in enumerate(images, 1):
                # Save image
                img_filename = f'index-{i}_1.jpg'
                with open(os.path.join(epub_dir, img_filename), 'wb') as f:
                    f.write(img_data)
                
                # Add image reference to HTML
                html_content.extend([
                    f'<p class="calibre1"><a id="p{i}"></a><img src="{img_filename}" '
                    f'alt="Image {i}" class="calibre2"/></p>',
                    '<p class="calibre1"></p>'
                ])

            html_content.extend(['  </body>', '</html>'])

            # Write index.html
            with open(os.path.join(epub_dir, 'index.html'), 'w', encoding='utf-8') as f:
                f.write('\n'.join(html_content))

            # Create toc.ncx
            toc_ncx = f'''<?xml version='1.0' encoding='utf-8'?>
<ncx xmlns="http://www.daisy.org/z3986/2005/ncx/" version="2005-1" xml:lang="eng">
  <head>
    <meta name="dtb:uid" content="{datetime.now().strftime('%Y%m%d%H%M%S')}"/>
    <meta name="dtb:depth" content="2"/>
    <meta name="dtb:generator" content="calibre (7.1.0)"/>
    <meta name="dtb:totalPageCount" content="0"/>
    <meta name="dtb:maxPageNumber" content="0"/>
  </head>
  <docTitle>
    <text>{title}</text>
  </docTitle>
  <navMap>
    <navPoint id="start" playOrder="1">
      <navLabel>
        <text>Start</text>
      </navLabel>
      <content src="index.html"/>
    </navPoint>
  </navMap>
</ncx>'''
            with open(os.path.join(epub_dir, 'toc.ncx'), 'w', encoding='utf-8') as f:
                f.write(toc_ncx)

            # Create content.opf
            content_opf = f'''<?xml version='1.0' encoding='utf-8'?>
<package xmlns="http://www.idpf.org/2007/opf" version="2.0" unique-identifier="uuid_id">
  <metadata xmlns:opf="http://www.idpf.org/2007/opf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:calibre="http://calibre.kovidgoyal.net/2009/metadata">
    <dc:language>en</dc:language>
    <dc:title>{title}</dc:title>
    <dc:creator opf:file-as="python-docx" opf:role="aut">python-docx</dc:creator>
    <dc:contributor opf:role="bkp">calibre (7.1.0) [https://calibre-ebook.com]</dc:contributor>
    <dc:date>{datetime.now().strftime("%Y-%m-%dT%H:%M:%S+00:00")}</dc:date>
    <meta name="calibre:timestamp" content="{datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f+00:00")}"/>
    <dc:identifier id="uuid_id" opf:scheme="uuid">{datetime.now().strftime('%Y%m%d%H%M%S')}</dc:identifier>
    <dc:identifier opf:scheme="calibre">{datetime.now().strftime('%Y%m%d%H%M%S')}</dc:identifier>
    <meta name="calibre:title_sort" content="{title}"/>
  </metadata>
  <manifest>
    <item id="id1" href="index.html" media-type="application/xhtml+xml" properties="nav"/>'''

            # Add image items to manifest
            for i in range(1, len(images) + 1):
                content_opf += f'\n    <item id="id{i+1}" href="index-{i}_1.jpg" media-type="image/jpeg"/>'

            content_opf += '''
    <item id="page_css" href="page_styles.css" media-type="text/css"/>
    <item id="css" href="stylesheet.css" media-type="text/css"/>
    <item id="ncx" href="toc.ncx" media-type="application/x-dtbncx+xml"/>
  </manifest>
  <spine toc="ncx">
    <itemref idref="id1"/>
  </spine>
</package>'''

            with open(os.path.join(epub_dir, 'content.opf'), 'w', encoding='utf-8') as f:
                f.write(content_opf)

            # Create EPUB file
            epub_name = os.path.basename(output_file)
            epub_dir_name = os.path.dirname(output_file)
            
            if os.path.exists(output_file):
                os.remove(output_file)

            # Create EPUB zip with mimetype first (uncompressed)
            import zipfile
            with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zf:
                # Add mimetype first (uncompressed)
                zf.write(os.path.join(epub_dir, 'mimetype'), 'mimetype', compress_type=zipfile.ZIP_STORED)
                
                # Add all other files
                for root, dirs, files in os.walk(epub_dir):
                    for file in files:
                        if file != 'mimetype':  # Skip mimetype as it's already added
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, epub_dir)
                            zf.write(file_path, arcname)

    def convert_pdf_to_epub(self):
        # Select input PDF file
        pdf_file = filedialog.askopenfilename(
            title="Select PDF file",
            filetypes=[("PDF files", "*.pdf")]
        )

        if not pdf_file:
            return

        # Generate output filename in the same directory as the PDF
        pdf_dir = os.path.dirname(pdf_file)
        pdf_name = os.path.splitext(os.path.basename(pdf_file))[0]
        output_file = os.path.join(pdf_dir, f"{pdf_name}.epub")


        try:
            # Create progress window
            progress_window = ctk.CTkToplevel(self.window)
            progress_window.title("Converting...")
            progress_window.geometry("400x200")
            progress_window.transient(self.window)
            progress_window.grab_set()

            # Center the progress window
            progress_window.update_idletasks()
            x = self.window.winfo_x() + (self.window.winfo_width() - progress_window.winfo_width()) // 2
            y = self.window.winfo_y() + (self.window.winfo_height() - progress_window.winfo_height()) // 2
            progress_window.geometry(f"+{x}+{y}")

            # Progress frame
            progress_frame = ctk.CTkFrame(progress_window)
            progress_frame.pack(fill="both", expand=True, padx=20, pady=20)

            progress_label = ctk.CTkLabel(
                progress_frame,
                text="Processing page: 0/0",
                font=ctk.CTkFont(size=14)
            )
            progress_label.pack(pady=20)

            # Progress bar
            progress_bar = ctk.CTkProgressBar(progress_frame)
            progress_bar.pack(pady=10, padx=20, fill="x")
            progress_bar.set(0)

            # Open PDF document
            self.pdf_document = fitz.open(pdf_file)
            total_pages = self.pdf_document.page_count

            if total_pages == 0:
                raise ValueError("PDF document appears to be empty")

            progress_label.configure(text=f"Processing page: 0/{total_pages}")
            progress_window.update()

            # Store images
            images = []

            # Process each page
            for page_num in range(total_pages):
                # Update progress
                progress = (page_num + 1) / total_pages
                progress_bar.set(progress)
                progress_label.configure(text=f"Processing page: {page_num + 1}/{total_pages}")
                progress_window.update()

                # Get the page
                page = self.pdf_document[page_num]

                # Extract image data
                img_data = self.extract_images_from_page(page)

                # Clear page from memory
                page = None

                if img_data is None:
                    print(f"Failed to extract image from page {page_num + 1}")
                    continue

                images.append(img_data)

            # Create EPUB
            title = os.path.splitext(os.path.basename(pdf_file))[0]
            self.create_epub(title, images, output_file)

            # Close progress window
            progress_window.destroy()

            # Show success message
            messagebox.showinfo("Success", "PDF has been successfully converted to EPUB!")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            if 'progress_window' in locals():
                progress_window.destroy()

        finally:
            if hasattr(self, 'pdf_document'):
                self.pdf_document.close()

    def run(self):
        self.window.mainloop()


if __name__ == "__main__":
    app = PDFtoEPUBConverter()
    app.run()