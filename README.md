# Powerpoint_from_pictures
To create a powerpoint from folders with pictures inside. You can use this script, you just need to replace the path in "chemins" by your folder 's path
It will make a powerpoint in your current folder with one picture well-sized per slide
You can also choose the pptx's name by changing the output_filename in main.

You have to install `python-pptx` and `Pillow`, you can use

```bash
pip install python-pptx Pillow 
```
or
```bash
pip install -r requirements.txt
```

And then, after making sure your pictures are in the folder `pictures`, you can run 

```bash
python3 main.py
```
And you will find the file `presentation.pptx` that contains all your pictures
