# manim-pptx
An addon for [3b1b/manim](https://github.com/3b1b/manim) that generates a PowerPoint presentation from a scene

## Installation
Since the addon API is still in development, make sure to use a version of `manim` supporting the [proposed API](https://github.com/3b1b/manim/pull/609), e.g. clone [this fork](https://github.com/yoshiask/manim) and run `pip3 install .` there.

To install the addon, copy the `pptx_exporter` folder into `[path to manim repo]/addons/pptx_addon`.

## Usage
Create a PowerPoint presentation from a scene by invoking manim with the `--save_to_pptx` argument, e.g:

`python3 -m manim your_script.py YourScene --save_to_pptx`
