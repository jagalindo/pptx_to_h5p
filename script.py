from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
import os
import json

def emu_to_px(emu):
    """Convert EMU (English Metric Unit) to pixels (assuming 96 DPI)."""
    return int(emu / 9525)  # 1 pixel = 9525 EMUs at 96 DPI

def convert_pptx_to_h5p(input_pptx, output_dir='h5p_content'):
    """
    Converts a PPTX file into an H5P Course Presentation package structure.
    
    Parameters:
    - input_pptx: Path to the .pptx file.
    - output_dir: Directory where the H5P content folder will be created.
    
    After running, you'll have:
    - output_dir/
      - h5p.json
      - content/
        - content.json
        - images/
          - slideX_imgY.ext
    Use 'h5p-cli pack <output_dir>' to create the final .h5p file.
    """
    prs = Presentation(input_pptx)
    content_dir = os.path.join(output_dir, 'content')
    images_dir = os.path.join(content_dir, 'images')
    os.makedirs(images_dir, exist_ok=True)

    slides = []
    for idx, slide in enumerate(prs.slides, start=1):
        slide_dict = {'elements': []}
        for shape_idx, shape in enumerate(slide.shapes, start=1):
            left  = emu_to_px(shape.left)
            top   = emu_to_px(shape.top)
            width = emu_to_px(shape.width)
            height= emu_to_px(shape.height)

            # Handle images
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = shape.image
                ext   = image.ext
                blob  = image.blob
                img_filename = f'images/slide{idx}_img{shape_idx}.{ext}'
                img_path     = os.path.join(content_dir, img_filename)
                with open(img_path, 'wb') as img_file:
                    img_file.write(blob)
                slide_dict['elements'].append({
                    'type':   'image',
                    'path':   img_filename,
                    'x':      left,
                    'y':      top,
                    'width':  width,
                    'height': height
                })

            # Handle text
            elif shape.has_text_frame:
                text = "\n".join(
                    "".join(run.text for run in para.runs)
                    for para in shape.text_frame.paragraphs
                ).strip()
                if text:
                    slide_dict['elements'].append({
                        'type':   'text',
                        'text':   text,
                        'x':      left,
                        'y':      top,
                        'width':  width,
                        'height': height
                    })

        slides.append(slide_dict)

    # Write content.json
    content = {'slides': slides}
    with open(os.path.join(content_dir, 'content.json'), 'w', encoding='utf-8') as f:
        json.dump(content, f, ensure_ascii=False, indent=2)

    # Write h5p.json (package definition)
    h5p_json = {
        "title": os.path.splitext(os.path.basename(input_pptx))[0],
        "mainLibrary": "H5P.CoursePresentation",
        "language": "en",
        "preloadedDependencies": [
            {"machineName": "H5P.CoursePresentation", "majorVersion": 1, "minorVersion": 23},
            {"machineName": "H5P.Text",               "majorVersion": 1, "minorVersion": 5},
            {"machineName": "H5P.Image",              "majorVersion": 1, "minorVersion": 3}
        ],
        "embedTypes": ["div"]
    }
    with open(os.path.join(output_dir, 'h5p.json'), 'w', encoding='utf-8') as f:
        json.dump(h5p_json, f, ensure_ascii=False, indent=2)

    print(f"H5P package structure generated in '{output_dir}'.")
    print("Run 'h5p-cli pack', for example:")
    print(f"    h5p-cli pack {output_dir}")

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Convert PPTX to H5P Course Presentation")
    parser.add_argument("pptx_file", help="Path to the .pptx file to convert")
    parser.add_argument("-o", "--output", default="h5p_content",
                        help="Output directory for the H5P package structure")
    args = parser.parse_args()

    convert_pptx_to_h5p(args.pptx_file, args.output)

