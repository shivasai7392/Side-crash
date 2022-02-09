# Side crash reporting automation

Automating the PPTX output by feeding 2D & 3D data into some functions which spits out a report.

The following image shows a user defined GUI to work with the Side Crash report automation.

<img src=".res/image.png" width="400" >

> **NOTE:** .


## Dev. Setup

Before getting started, one may want to know how relative imports work in Python. In this project we will be using relative imports inside the `src` package. It's important that we use relative imports to package the checks using *BETA Packager Installer*.

1. The script `debug.py` that resides outside the `side_crash_src` folder is the one that pulls all the strings while testing.

2. The task of `side_crash_src/side_crash_gui.py` script is to write the user friendly GUI for input and output operations.

```python