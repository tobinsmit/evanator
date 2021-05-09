print("Starting tobinator")
from PySide2 import QtWidgets as qw
print("Done import PySide2")
import sys, re, traceback
from forex_python.converter import CurrencyRates
print("Done import sys, re, forex_python")
import evanator
print("Done import evanator")

class Window(qw.QWidget):
    def __init__(self, parent=None):
        super(Window, self).__init__(parent)
        layout = qw.QGridLayout()
        # layout = QVBoxLayout()
        currentRow = 0

        self.autoDetectButt = qw.QPushButton("Auto Detect files")
        self.autoDetectButt.clicked.connect(self.autoDetectFiles)
        layout.addWidget(self.autoDetectButt, currentRow, 0)
        currentRow += 1

        # Amazon
        self.amazonPath = None
        layout.addWidget(qw.QLabel("Amazon:"), currentRow, 0)
        self.chooseAmazonButt = qw.QPushButton("Choose File")
        self.chooseAmazonButt.clicked.connect(self.chooseAmazonFile)
        layout.addWidget(self.chooseAmazonButt, currentRow, 1)
        self.amazonPathLabel = qw.QLabel("None")
        layout.addWidget(self.amazonPathLabel, currentRow, 2)
        currentRow += 1

        # Ingram AU
        self.ingramAuPath = None
        layout.addWidget(qw.QLabel("Ingram AU:"), currentRow, 0)
        self.chooseIngramAuButt = qw.QPushButton("Choose File")
        self.chooseIngramAuButt.clicked.connect(self.chooseIngramAuFile)
        layout.addWidget(self.chooseIngramAuButt, currentRow, 1)
        self.ingramAuPathLabel = qw.QLabel("None")
        layout.addWidget(self.ingramAuPathLabel, currentRow, 2)
        currentRow += 1

        # Ingram US
        self.ingramUsPath = None
        layout.addWidget(qw.QLabel("Ingram US:"), currentRow, 0)
        self.chooseIngramUsButt = qw.QPushButton("Choose File")
        self.chooseIngramUsButt.clicked.connect(self.chooseIngramUsFile)
        layout.addWidget(self.chooseIngramUsButt, currentRow, 1)
        self.ingramUsPathLabel = qw.QLabel("None")
        layout.addWidget(self.ingramUsPathLabel, currentRow, 2)
        currentRow += 1

        # Ingram UK
        self.ingramUkPath = None
        layout.addWidget(qw.QLabel("Ingram UK:"), currentRow, 0)
        self.chooseIngramUkButt = qw.QPushButton("Choose File")
        self.chooseIngramUkButt.clicked.connect(self.chooseIngramUkFile)
        layout.addWidget(self.chooseIngramUkButt, currentRow, 1)
        self.ingramUkPathLabel = qw.QLabel("None")
        layout.addWidget(self.ingramUkPathLabel, currentRow, 2)
        currentRow += 1

        self.outputButt = qw.QPushButton("Set output folder")
        self.outputButt.clicked.connect(self.getOutputPath)
        layout.addWidget(self.outputButt, currentRow, 0)
        currentRow += 1

        self.outputPath = None
        layout.addWidget(qw.QLabel("Output path:"), currentRow, 0)
        self.outputLabel = qw.QLabel("None")
        layout.addWidget(self.outputLabel, currentRow, 1)
        currentRow += 1

        self.ratesButt = qw.QPushButton("Get rates")
        self.ratesButt.clicked.connect(self.getRates)
        layout.addWidget(self.ratesButt, currentRow, 0)
        currentRow += 1

        self.rates = None

        layout.addWidget(qw.QLabel("1 USD ="), currentRow, 0)
        self.usdLabel = qw.QLabel("None")
        layout.addWidget(self.usdLabel, currentRow, 1)
        currentRow += 1

        layout.addWidget(qw.QLabel("1 GBP ="), currentRow, 0)
        self.gbpLabel = qw.QLabel("None")
        layout.addWidget(self.gbpLabel, currentRow, 1)
        currentRow += 1

        self.runButton = qw.QPushButton('Run')
        self.runButton.clicked.connect(self.run)
        layout.addWidget(self.runButton, currentRow, 0)
        currentRow += 1

        self.setLayout(layout)
        # self.setMinimumSize(410, 400)
        self.show()

    def autoDetectFiles(self):
        paths = qw.QFileDialog.getOpenFileNames(self, 'Open file', '', 'Excel (*.xls *.xlsx)')
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
        paths = qw.QFileDialog.getOpenFileName(self, 'Open Amazon file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.amazonPath = path
        self.amazonPathLabel.setText(path)
    
    def chooseIngramAuFile(self):
        paths = qw.QFileDialog.getOpenFileName(self, 'Open Ingram AU file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.ingramAuPath = path
        self.ingramAuPathLabel.setText(path)

    def chooseIngramUsFile(self):
        paths = qw.QFileDialog.getOpenFileName(self, 'Open Ingram US file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.ingramUsPath = path
        self.ingramUsPathLabel.setText(path)

    def chooseIngramUkFile(self):
        paths = qw.QFileDialog.getOpenFileName(self, 'Open Ingram UK file', '', 'Excel (*.xls *.xlsx)')
        path = paths[0]
        self.ingramUkPath = path
        self.ingramUkPathLabel.setText(path)

    def getOutputPath(self):
        path = qw.QFileDialog.getExistingDirectory(self, 'Select output folder')
        print(path)
        self.outputPath = path
        self.outputLabel.setText(path)

    def getRates(self):
        self.rates = CurrencyRates().get_rates('AUD')
        self.rates['USD'] = 1/self.rates['USD']
        self.rates['GBP'] = 1/self.rates['GBP']
        self.usdLabel.setText(f"{self.rates['USD']} AUD")
        self.gbpLabel.setText(f"{self.rates['GBP']} AUD")
             
    def run(self):
        print("Running")
        
        if self.outputPath == None:
            print("Please set outputPath")
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText("Please set outputPath")
            msg.exec_()
            return

        if self.rates == None:
            print("Please set rates")
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Critical)
            msg.setText("Error")
            msg.setInformativeText("Please set rates")
            msg.exec_()
            return

        if self.amazonPath == None:
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for amazonPath. Generating output anyway.")
            msg.exec_()
        if self.ingramAuPath == None:
            print("Please set ingramAuPath")
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for ingramAuPath. Generating output anyway.")
            msg.exec_()
        if self.ingramUkPath == None:
            print("Please set ingramUkPath")
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for ingramUkPath. Generating output anyway.")
            msg.exec_()
        if self.ingramUsPath == None:
            print("Please set ingramUsPath")
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Warning)
            msg.setText("Warning")
            msg.setInformativeText("Nothing set for ingramUsPath. Generating output anyway.")
            msg.exec_()

        ingramPaths = [self.ingramAuPath, self.ingramUkPath, self.ingramUsPath]

        try:
            errArr = evanator.main(ingramPaths=ingramPaths, amazonPath=self.amazonPath, outputPath=self.outputPath, rates=self.rates)
        except Exception as err:
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.Critical)
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
            msg = qw.QMessageBox()
            msg.setIcon(qw.QMessageBox.NoIcon)
            msg.setText("Done!")
            msg.setInformativeText("Outputted files")
            msg.exec_()
        else:
            for err in errArr:
                msg = qw.QMessageBox()
                msg.setIcon(qw.QMessageBox.Critical)
                msg.setText("Error")
                msg.setInformativeText(err)
                msg.exec_()


# Create the Qt Application
app = qw.QApplication([])
# Create and show the app
windows = Window()
windows.show()
# Run the main Qt loop
sys.exit(app.exec_())
