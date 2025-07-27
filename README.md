# PPTX to H5P Converter

This tool converts a PowerPoint presentation into an H5P Course Presentation package. It extracts images, text, simple shapes and media files, generating a directory ready to be zipped into a `.h5p` archive. Pass `--pack` to create the `.h5p` file automatically. Use `--include-libraries` to copy the libraries referenced in the generated `h5p.json` from the `jagalindo/h5p-cli` Docker image before packaging. Add `-r` to also resolve dependencies recursively.

## Requirements
- Python 3.8+
- `python-pptx`
- Docker (needed when using `--include-libraries`)

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

## Getting the Docker Image
Pull the latest `jagalindo/h5p-cli` image with:
```bash
docker pull jagalindo/h5p-cli:latest
```
You can also build it from source if desired:
```bash
git clone https://github.com/jagalindo/h5p-cli.git
cd h5p-cli
docker build -t jagalindo/h5p-cli .
```
Updating the image ensures the bundled H5P libraries are up to date.

If you encounter messages such as `cp: cannot stat '/usr/local/lib/h5p/...'`,
the Docker image may be missing those libraries. Pull the latest image again:

```bash
docker pull jagalindo/h5p-cli:latest
```

## Usage
```bash
python script.py myslides.pptx -o output_dir --pack
```
This packs the generated directory into `output_dir.h5p` without bundling the
H5P libraries. To also include the libraries listed in `h5p.json` run:
```bash
python script.py myslides.pptx -o output_dir --pack --include-libraries
```
Add `-r` to also copy any dependencies of those libraries recursively.
Libraries are stored under `.h5p/libraries` inside the output directory but the
final archive places them at the package root just like `h5p-cli pack` does.
When packing, common development artifacts such as `.git` folders or `tests`
directories are automatically skipped. Without the flag you can copy the
libraries and zip the directory manually:
```bash
docker run --rm -v /path/to/output_dir:/data jagalindo/h5p-cli \
  sh -c 'mkdir -p /data/.h5p && cp -r /usr/local/lib/h5p/<Lib> /data/.h5p/'
cd output_dir && zip -r ../output_dir.h5p .
```
