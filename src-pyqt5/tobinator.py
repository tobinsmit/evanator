print("Starting tobinator")

# from PyQt5 import QtWidgets as qw
print("Importing PyQt5")
from PyQt5.QtWidgets import QApplication, QWidget, QLabel, QGridLayout, QPushButton, QFileDialog, QMessageBox

print("Importing sys, re, traceback")
import sys, re, traceback

print("Importing forex_python")
from forex_python.converter import CurrencyRates

print("Importing evanator")
import evanator
print('Done import evanator')

QMESSAGEBOX_NO_ICON = 0
QMESSAGEBOX_INFORMATION = 1
QMESSAGEBOX_WARNING = 2
QMESSAGEBOX_CRITICAL = 3
QMESSAGEBOX_QUESTION = 4

print('Defining window')
class Window(QWidget):
    def __init__(self, parent=None):
        super(Window, self).__init__(parent)
        layout = QGridLayout()
        # layout = QVBoxLayout()
        currentRow = 0

        self.autoDetectButt = QPushButton("Auto Detect files")
        self.autoDetectButt.clicked.connect(self.autoDetectFiles)
        layout.addWidget(self.autoDetectButt, currentRow, 0)
        currentRow += 1

        # Amazon
        self.amazonPath = None
        layout.addWidget(QLabel("Amazon:"), currentRow, 0)
        self.chooseAmazonButt = QPushButton("Choose File")
        self.chooseAmazonButt.clicked.connect(self.chooseAmazonFile)
        layout.addWidget(self.chooseAmazonButt, currentRow, 1)
        self.amazonPathLabel = QLabel("None")
        layout.addWidget(self.amazonPathLabel, currentRow, 2)
        currentRow += 1

        # Ingram AU
        self.ingramAuPath = None
        layout.addWidget(QLabel("Ingram AU:"), currentRow, 0)
        self.chooseIngramAuButt = QPushButton("Choose File")
        self.chooseIngramAuButt.clicked.connect(self.chooseIngramAuFile)
        layout.addWidget(self.chooseIngramAuButt, currentRow, 1)
        self.ingramAuPathLabel = QLabel("None")
        layout.addWidget(self.ingramAuPathLabel, currentRow, 2)
        currentRow += 1

        # Ingram US
        self.ingramUsPath = None
        layout.addWidget(QLabel("Ingram US:"), currentRow, 0)
        self.chooseIngramUsButt = QPushButton("Choose File")
        self.chooseIngramUsButt.clicked.connect(self.chooseIngramUsFile)
        layout.addWidget(self.chooseIngramUsButt, currentRow, 1)
        self.ingramUsPathLabel = QLabel("None")
        layout.addWidget(self.ingramUsPathLabel, currentRow, 2)
        currentRow += 1

        # Ingram UK
        self.ingramUkPath = None
        layout.addWidget(QLabel("Ingram UK:"), currentRow, 0)
        self.chooseIngramUkButt = QPushButton("Choose File")
        self.chooseIngramUkButt.clicked.connect(self.chooseIngramUkFile)
        layout.addWidget(self.chooseIngramUkButt, currentRow, 1)
        self.ingramUkPathLabel = QLabel("None")
        layout.addWidget(self.ingramUkPathLabel, currentRow, 2)
        currentRow += 1

        self.outputButt = QPushButton("Set output folder")
        self.outputButt.clicked.connect(self.getOutputPath)
        layout.addWidget(self.outputButt, currentRow, 0)
        currentRow += 1

        self.outputPath = None
        layout.addWidget(QLabel("Output path:"), currentRow, 0)
        self.outputLabel = QLabel("None")
        layout.addWidget(self.outputLabel, currentRow, 1)
        currentRow += 1

        self.ratesButt = QPushButton("Get rates")
        self.ratesButt.clicked.connect(self.getRates)
        layout.addWidget(self.ratesButt, currentRow, 0)

        self.ratesViewButt = QPushButton("View all rates")
        self.ratesViewButt.clicked.connect(self.viewRates)
        layout.addWidget(self.ratesViewButt, currentRow, 1)
        currentRow += 1

        self.rates = None

        layout.addWidget(QLabel("1 USD ="), currentRow, 0)
        self.usdLabel = QLabel("None")
        layout.addWidget(self.usdLabel, currentRow, 1)
        currentRow += 1

        layout.addWidget(QLabel("1 GBP ="), currentRow, 0)
        self.gbpLabel = QLabel("None")
        layout.addWidget(self.gbpLabel, currentRow, 1)
        currentRow += 1

        self.runButton = QPushButton('Run')
        self.runButton.clicked.connect(self.run)
        layout.addWidget(self.runButton, currentRow, 0)
        currentRow += 1

        self.setLayout(layout)
        # self.setMinimumSize(410, 400)
        self.show()

    def autoDetectFiles(self):
        paths = QFileDialog.getOpenFileNames(self, 'Open file', '', 'Excel (*.xls *.xlsx)')
        print(paths)
        for path in paths[0]: 
            if bool(re.match(r".*KDP-Sales-Dashboard([\w-]+)\.xlsx", path)):
                self.amazonPath = path
                self.amazonPathLabel.setText(path)
            elif bool(re.match(r".*sales_compAU\.xls", path)):
                self.ingramAuPath = path
                self.ingramAuPathLabel.setText(path)
            elif bool(re.match(r".*sales_compUS\.xls", path)):
                self.ingramUsPath = path
                self.ingramUsPathLabel.setText(path)
            elif bool(re.match(r".*sales_compUK\.xls", path)):
                self.ingramUkPath = path
                self.ingramUkPathLabel.setText(path)

    def chooseAmazonFile(self):
        paths = QFileDialog.getOpenFileName(self, 'Open Amazon file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.amazonPath = path
        self.amazonPathLabel.setText(path)
    
    def chooseIngramAuFile(self):
        paths = QFileDialog.getOpenFileName(self, 'Open Ingram AU file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.ingramAuPath = path
        self.ingramAuPathLabel.setText(path)

    def chooseIngramUsFile(self):
        paths = QFileDialog.getOpenFileName(self, 'Open Ingram US file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.ingramUsPath = path
        self.ingramUsPathLabel.setText(path)

    def chooseIngramUkFile(self):
        paths = QFileDialog.getOpenFileName(self, 'Open Ingram UK file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.ingramUkPath = path
        self.ingramUkPathLabel.setText(path)

    def getOutputPath(self):
        path = QFileDialog.getExistingDirectory(self, 'Select output folder')
        print(path)
        self.outputPath = path
        self.outputLabel.setText(path)

    def getRates(self):
        self.rates = CurrencyRates().get_rates('AUD')
        self.rates['AUD'] = 1
        self.usdLabel.setText(f"{1/self.rates['USD']:.5f} AUD")
        self.gbpLabel.setText(f"{1/self.rates['GBP']:.5f} AUD")
             
    def viewRates(self):
        if self.rates == None:
            QMessageBox.about(self,'Rates', 'Rates have not been set yet')
        else:
            ratesArray = [f'1 {key} = {1/value:.5f} AUD' for key, value in self.rates.items()]
            ratesStr = '\n'.join(ratesArray)
            QMessageBox.about(self,'Rates', ratesStr)
             
    def run(self):
        print("Running")
        
        if self.outputPath == None:
            print("Please set outputPath")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("Error")
            msg.setInformativeText("Please set outputPath")
            msg.exec()
            return

        if self.rates == None:
            print("Please set rates")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("Error")
            msg.setInformativeText("Please set rates")
            msg.exec_()
            return

        if self.amazonPath == None:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for amazonPath. Generating output anyway.")
            msg.exec_()
        if self.ingramAuPath == None:
            print("Please set ingramAuPath")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for ingramAuPath. Generating output anyway.")
            msg.exec_()
        if self.ingramUkPath == None:
            print("Please set ingramUkPath")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for ingramUkPath. Generating output anyway.")
            msg.exec_()
        if self.ingramUsPath == None:
            print("Please set ingramUsPath")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for ingramUsPath. Generating output anyway.")
            msg.exec_()

        ingramPaths = [self.ingramAuPath, self.ingramUkPath, self.ingramUsPath]

        try:
            errArr = evanator.main(ingramPaths=ingramPaths, amazonPath=self.amazonPath, outputPath=self.outputPath, rates=self.rates)
        except Exception as err:
            msg = QMessageBox.critical()
            # msg.setIcon(QMessageBox.Icon.Critical)
            msg.setText("Error\t\t\t\t")
            msg.setInformativeText((
                "Unknown error occured. Output was not generated."+
                "\n\n"+
                traceback.format_exc()
            ))
            print(traceback.format_exc())
            msg.exec_()
            return

        if len(errArr) == 0:
            print("Done run")
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.NoIcon)
            msg.setText("Done!")
            msg.setInformativeText("Outputted files")
            msg.exec_()
        else:
            for err in errArr:
                msg = QMessageBox()
                msg.setIcon(QMessageBox.Icon.Critical)
                msg.setText("Error")
                msg.setInformativeText(err)
                msg.exec_()


# Create the Qt Application
print('Making app')
app = QApplication([])
# Create and show the app
print('Making window')
windows = Window()
print('Showing window')
windows.show()
# Run the main Qt loop
sys.exit(app.exec())
