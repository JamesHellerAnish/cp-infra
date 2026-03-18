import fitz
import os

pdf_path = r'c:\Users\naben\OneDrive\Desktop\CP sytems\C P SYSTEMS PVT LTD.pdf'
out_dir = r'c:\Users\naben\OneDrive\Desktop\CP sytems\website\public\images'

doc = fitz.open(pdf_path)
count = 0

for page_index in range(len(doc)):
    page = doc[page_index]
    image_list = page.get_images(full=True)
    
    for image_index, img in enumerate(image_list, start=1):
        xref = img[0]
        base_image = doc.extract_image(xref)
        image_bytes = base_image["image"]
        image_ext = base_image["ext"]
        
        if base_image["width"] < 100 or base_image["height"] < 100:
            continue
        
        # We will save it with its native extension. If it's a PNG, it will be a valid PNG.
        filename = f"pdf_page_{page_index+1}_img_{count}.{image_ext}"
        filepath = os.path.join(out_dir, filename)
        
        with open(filepath, "wb") as wfd:
            wfd.write(image_bytes)
        print(f"Saved {filename}")
        count += 1
