import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import *
import docx2txt


class Notepy (QMainWindow):
    def __init__(self):
        super().__init__()
        self.title ="Notepy"
        self.setWindowTitle(self.title)

        self.editor =QTextEdit()
        self.setCentralWidget(self.editor)

        self.showMaximized()
        self.editor.setFontPointSize(20)

        self.font_size_box = QSpinBox()

        self.create_tool_bar()
        self.create_menu_bar()

    def create_menu_bar (self):
        menu_bar = QMenuBar()
        self.setMenuBar(menu_bar)

        #======================file_menu===============

        file_menu =QMenu('File',self)
        menu_bar.addMenu(file_menu)

        open_action = QAction('Open', self)
        open_action.triggered.connect(self.file_open)
        file_menu.addAction(open_action)

        save_action = QAction('Save', self)
        save_action.triggered.connect(self.file_save)
        file_menu.addAction(save_action)

        save_as_pdf_action = QAction('Save As PDF', self)
        save_as_pdf_action.triggered.connect(self.save_as_pdf)
        file_menu.addAction(save_as_pdf_action)

        rename_action = QAction('Rename', self)
        rename_action.triggered.connect(self.file_saveas)
        file_menu.addAction(rename_action)

        edit_menu = QMenu('Edit', self)
        menu_bar.addMenu(edit_menu)

        view_menu = QMenu('View', self)
        menu_bar.addMenu(view_menu)

        insert_menu = QMenu('Insert', self)
        menu_bar.addMenu(insert_menu)

        format_menu = QMenu('Format', self)
        menu_bar.addMenu(format_menu)

        tools_menu = QMenu('Tools', self)
        menu_bar.addMenu(tools_menu)




    def create_tool_bar(self):

        tool_bar = QToolBar()

        undo_action =QAction(QIcon('undo.png'),'Undo', self)
        undo_action.triggered.connect(self.editor.undo)
        tool_bar.addAction(undo_action)

        redo_action =QAction(QIcon('redo.png'),'redo', self)
        redo_action.triggered.connect(self.editor.redo)
        tool_bar.addAction(redo_action)

        cut_action = QAction(QIcon('cut.png'), 'cut', self)
        cut_action.triggered.connect(self.editor.cut)
        tool_bar.addAction(cut_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        copy_action = QAction(QIcon('copy.png'), 'copy', self)
        copy_action.triggered.connect(self.editor.copy)
        tool_bar.addAction(copy_action)

        paste_action = QAction(QIcon('paste.png'), 'paste', self)
        paste_action.triggered.connect(self.editor.paste)
        tool_bar.addAction(paste_action)


        tool_bar.addSeparator()
        tool_bar.addSeparator()

        self.font_size_box.setValue(14)
        self.font_size_box.valueChanged.connect(self.set_font_size)
        tool_bar.addWidget(self.font_size_box)

        self.addToolBar(tool_bar)


         # =================  funcitons  =============


    def set_font_size(self):

        value = self.font_size_box.value()
        self.editor.setFontPointSize(value)

    def file_open(self):
        self.path, _ = QFileDialog.getOpenFileNames(self,"Open File", "","Text documents (*.text); Text documents (*.txt);All file(*.*)")
        try:
            text = docx2txt.process(self.path)
        except Exception as e:
            print(e)
        else:
            self.editor.setText(text)
            self.update_title()

    def file_save(self):
        print(self.path)
        if self.path == '':

            self.file_saveas()
        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "",
                                                   "text documents (*.text);Text documents (*.txt);All files (*.*)")

        if self.path == ' ':
            return
        text = self.editor.toPlainText()

        try:
            with open(path,'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
                print(e)

    def update_title(self):
        self.setWindowTitle(self.title + ' ' + self.path)

    def save_as_pdf(self):
        file_path, _ = QFileDialog.getSaveFileName(self, 'Save PDF', 'Notepy', 'PDF File (*.pdf)')
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        self.editor.document().print(printer)

app = QApplication(sys.argv)
window = Notepy()
window.show()
sys.exit(app.exec_())