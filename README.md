## To build environment:
```
cd src
python3 -m venv venv
source venv/bin/activate
pip3 install -r requirements.txt
```
## To run:
```
python3 tobinator.py
```
## To make python script executable:
Make "open.command" double clickable:
- Right click on folder "tobinator" and select "New Terminal at Folder".
- Copy and paste: `chmod u+x open.command` and hit enter.
- The file should now be double clickable and you can close the terminal.

## To build app:
```
pip3 install pyinstaller
pyinstaller tobinator.py --onefile --add-data 'SampleCover.jpg:.'
mkdir ../tobinator.app
mv dist/tobinator ../tobinator.app/tobinator
cp SampleCover.jpg ../tobinator.app/SampleCover.jpg
```
It took me 35 seconds to app the app for the first time and 15 seconds every time after that.
## Use:
- Open the app. If permision is needed to open it, go to System Preferences > Security and Privacy > General > Open app
- Click 'Open files' and select all the 4 data files (can use `shift-click` to select multiple)
- Select the output folder
- Get current conversion rates. For some reason it doesn't update the labels. It did in development mode.
- Run. A dialog should pop up when finished.

Opening it a second time should be faster btw.

## To do:
- Fix blur
- Fix currency rate labels so they update
- Make app smaller