# PPTX to H5P Converter

This tool converts a PowerPoint presentation into an H5P Course Presentation package. It extracts images, text, simple shapes and media files, generating a directory ready for `h5p-cli pack`.

## Requirements
- Python 3.8+
- `python-pptx`
- `h5p-cli` (optional, for packaging)

Install dependencies with:
```bash
pip install -r requirements.txt
```

## Usage
```bash
python script.py myslides.pptx -o output_dir --pack
```
The `--pack` flag calls `h5p-cli` to produce a `.h5p` archive automatically. Without it, the output directory can be packed later using `h5p-cli pack output_dir`.
