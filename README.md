# pdf2pptx

Simple converter from PDF into PowerPoint PPTX via raster images.

I have created this tool to be able to hand over PPTX versions of my presentations created in LaTeX Beamer, where none of the conversions via Adobe Acrobat nor PowerPoint itself preserve full visual fidelity of the presentation. By converting each PDF page into a raster image, we can ensure near-identical rendering within PPTX.

Written in Python, the tool requires [ImageMagick][imagemagick-link] and [python-pptx][python-pptx-link].

---

## Table of Contents

- [Installation](#installation)
  - [1. Install Python and Pip](#1-install-python-and-pip)
  - [2. Install ImageMagick](#2-install-imagemagick)
  - [3. Install python-pptx](#3-install-python-pptx)
  - [4. Download the `pdf_to_ppt.py` Script](#4-download-the-pdf_to_pptpy-script)
- [Usage](#usage)
  - [Basic Command](#basic-command)
  - [Command-Line Arguments](#command-line-arguments)
- [Examples](#examples)
- [License](#license)

---

## Installation

### 1. Install Python and Pip

1. **Windows**  
   Download from the official [Python.org](https://www.python.org/downloads/) website.  
   During installation, make sure to check **"Add Python to PATH"** so you can run Python from the command line.

2. **Linux** (e.g., Ubuntu/Debian)

    sudo apt-get update  
    sudo apt-get install python3 python3-pip

3. **macOS**  
   Python 3 is often included (depending on your version).  
   Otherwise, install via [Homebrew](https://brew.sh/) with:

    brew install python3

### 2. Install ImageMagick

**ImageMagick** provides the `convert` command, which is used to convert PDFs into raster images (JPEG or PNG).

- **Windows**:  
  Download the installer from the [ImageMagick downloads page](https://imagemagick.org/script/download.php#windows) and follow the on-screen instructions.
- **Linux**:

      # Debian/Ubuntu
      sudo apt-get install imagemagick
      # Fedora/CentOS
      sudo dnf install imagemagick

- **macOS**:

      brew install imagemagick

To verify, run `convert -version` in a terminal or Command Prompt.

### 3. Install python-pptx

Use **pip** to install [python-pptx][python-pptx-link]:

    pip install python-pptx

### 4. Download the `pdf_to_ppt.py` Script

1. Save the `pdf_to_ppt.py` script (from this repository) into a local directory.  
2. Make it executable (on Linux/macOS) with:

       chmod +x pdf_to_ppt.py

3. Optionally, place it in a directory on your `PATH` so it can be called from anywhere.

---

## Usage

### Basic Command

    python pdf_to_ppt.py [OPTIONS] <input.pdf> <output.pptx>

Where `input.pdf` is the source PDF and `output.pptx` is the resulting PowerPoint file.

### Command-Line Arguments

| Argument / Option             | Description                                                                                                                                                                      | Default / Notes                         |
|-------------------------------|----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|------------------------------------------|
| `-v, --verbose`               | Show detailed progress information (INFO-level logs). By default, only warnings and errors are shown.                                                                            | Off by default                          |
| `--retain-image-files`        | Do not delete the temporary image files/directory after creating the PPTX. By default, the script removes the temporary images once done.                                        | Off by default                          |
| `-d, --dpi <DPI>`            | Set the DPI (dots per inch) used by **ImageMagick** `convert`. Higher DPI yields higher-resolution images but larger file sizes.                                                  | `1200`                                  |
| `-q, --quality <QUALITY>`     | Set the image compression quality (1-100) for **ImageMagick**. For JPEG, higher means better quality (and bigger files). For PNG, it's mostly useless.                | `95`                                    |
| `-f, --image-format <FORMAT>` | Choose the intermediate image format (`JPEG` or `PNG`).                                                                                                                           | `JPEG`                                  |
| `-t, --temp-dir <DIR>`        | Directory to store the temporary image files (must not already exist). If unspecified, the script creates `<PDF_basename>_temp`.                                                 | `<PDF_basename>_temp`                  |
| `-s, --slide-size <SIZE>`     | Slide size: `'16:9'`, `'4:3'`, or `'WxH'` in **inches**. For example, `16:9` = 13.3333×7.5 inches, `4:3` = 10×7.5 inches, `11.0x8.5` = 11.0 inches by 8.5 inches, etc.            | `16:9` (13.3333×7.5 inches)            |
| `<input.pdf>`                 | The source PDF file to convert.                                                                                                                                                  | Required                                |
| `<output.pptx>`               | The destination PowerPoint file.                                                                                                                                                 | Required                                |

---

## Examples

1. **Default usage**  
   Converts a PDF at 1200 DPI to high-quality JPEG images, 16:9 slides, removing images afterward:

       python pdf_to_ppt.py presentation.pdf presentation.pptx
   
2. **Verbose logs** and **retain all images**:

       python pdf_to_ppt.py -v --retain-image-files presentation.pdf presentation.pptx

   This will show progress for each PDF page and keep the `presentation_temp` folder, so you can inspect or reuse the images.

3. **Use PNG images**, **300 DPI**, **80% quality**, and specify a custom temporary directory:

       python pdf_to_ppt.py -v -f PNG -d 300 -q 80 -t /tmp/my-temp-dir presentation.pdf presentation.pptx

4. **Set the slide size** to 4:3 (10" × 7.5"):

       python pdf_to_ppt.py -s 4:3 presentation.pdf presentation.pptx

5. **Set a custom size** of **11"x8.5"** (for a near Letter-sized slide):

       python pdf_to_ppt.py -s 11.0x8.5 presentation.pdf presentation.pptx

---

## License

This project is distributed under the [GNU General Public License version 3][license-link] (GPLv3).

For more details, see [LICENSE](./LICENSE).

---

Happy converting! If you have any issues or want to contribute, please open an issue or pull request.

[imagemagick-link]: https://imagemagick.org/
[python-pptx-link]: https://pypi.org/project/python-pptx/
[license-link]: https://www.gnu.org/licenses/gpl-3.0.html
