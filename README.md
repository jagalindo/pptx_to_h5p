# PPTX to H5P Converter

This tool converts a PowerPoint presentation into an H5P Course Presentation package. It extracts images, text, simple shapes and media files, generating a directory ready for packaging with the `h5p` command (from the `@lumieducation/h5p-cli` package) or the `jagalindo/h5p-cli` Docker image.

## Requirements
- Python 3.8+
- `python-pptx`
- Docker (for packaging with the `jagalindo/h5p-cli` image)

It is recommended to use a virtual environment. Create one with:
```bash
./setup_env.sh
```
This script creates a `.venv` folder and installs the requirements.

If you prefer to do it manually, run:
```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage
```bash
python script.py myslides.pptx -o output_dir --pack
```
The `--pack` flag uses the Docker image `jagalindo/h5p-cli` to produce a `.h5p` archive automatically. It first copies the default extensions from the image and then packs the directory. Without the flag, you can do the same manually with:
```bash
docker run --rm -v /path/to/output_dir:/data jagalindo/h5p-cli \
  sh -c 'mkdir -p /data/.h5p && cp -r /usr/local/lib/h5p/* /data/.h5p/ && h5p pack /data || h5p-cli pack /data'
```
