#coding = utf-8
import os
import sys

import reportlab
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont as pdfFont
from reportlab.pdfbase import pdfmetrics
from datetime import datetime

from pdfrw import PdfReader, PdfWriter, PageMerge
from cn2an import transform
from fontTools.ttLib import TTFont

from PyQt5 import QtGui, QtWidgets, QtCore
from PyQt5.QtGui import *
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *

class Ui_widget(object):
    def setupUi(self, widget):
        widget.setObjectName("widget")
        widget.resize(446, 326)
        self.gridLayout = QtWidgets.QGridLayout(widget)
        self.gridLayout.setContentsMargins(20, 20, 20, 20)
        self.gridLayout.setObjectName("gridLayout")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_3 = QtWidgets.QLabel(widget)
        self.label_3.setObjectName("label_3")
        self.horizontalLayout_2.addWidget(self.label_3)
        self.file_path_line = QtWidgets.QLineEdit(widget)
        self.file_path_line.setText("")
        self.file_path_line.setReadOnly(True)
        self.file_path_line.setObjectName("file_path_line")
        self.horizontalLayout_2.addWidget(self.file_path_line)
        self.file_path_btn = QtWidgets.QPushButton(widget)
        self.file_path_btn.setFocusPolicy(QtCore.Qt.NoFocus)
        self.file_path_btn.setObjectName("file_path_btn")
        self.horizontalLayout_2.addWidget(self.file_path_btn)
        self.gridLayout.addLayout(self.horizontalLayout_2, 0, 0, 1, 1)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.font_box = QtWidgets.QComboBox(widget)
        self.font_box.setEditable(True)
        self.font_box.setObjectName("font_box")
        self.horizontalLayout.addWidget(self.font_box)
        self.font_size_box = QtWidgets.QComboBox(widget)
        self.font_size_box.setMaximumSize(QtCore.QSize(70, 16777215))
        self.font_size_box.setEditable(True)
        self.font_size_box.setObjectName("font_size_box")
        self.horizontalLayout.addWidget(self.font_size_box)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout.addItem(spacerItem)
        self.label_1 = QtWidgets.QLabel(widget)
        self.label_1.setMaximumSize(QtCore.QSize(10, 16777215))
        self.label_1.setObjectName("label_1")
        self.horizontalLayout.addWidget(self.label_1)
        self.x_box = QtWidgets.QComboBox(widget)
        self.x_box.setMaximumSize(QtCore.QSize(70, 16777215))
        self.x_box.setEditable(True)
        self.x_box.setObjectName("x_box")
        self.horizontalLayout.addWidget(self.x_box)
        self.label_2 = QtWidgets.QLabel(widget)
        self.label_2.setMaximumSize(QtCore.QSize(10, 16777215))
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.y_box = QtWidgets.QComboBox(widget)
        self.y_box.setMaximumSize(QtCore.QSize(70, 16777215))
        self.y_box.setEditable(True)
        self.y_box.setObjectName("y_box")
        self.horizontalLayout.addWidget(self.y_box)
        self.gridLayout.addLayout(self.horizontalLayout, 1, 0, 1, 1)
        self.input_string_view = QtWidgets.QTextEdit(widget)
        self.input_string_view.setObjectName("input_string_view")
        self.gridLayout.addWidget(self.input_string_view, 2, 0, 1, 1)
        self.horizontalLayout_3 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_3.setObjectName("horizontalLayout_3")
        spacerItem1 = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_3.addItem(spacerItem1)
        self.run_btn = QtWidgets.QPushButton(widget)
        self.run_btn.setObjectName("run_btn")
        self.horizontalLayout_3.addWidget(self.run_btn)
        self.gridLayout.addLayout(self.horizontalLayout_3, 3, 0, 1, 1)

        self.retranslateUi(widget)
        QtCore.QMetaObject.connectSlotsByName(widget)

    def retranslateUi(self, widget):
        _translate = QtCore.QCoreApplication.translate
        widget.setWindowTitle(_translate("widget", "pdf文字添加"))
        self.label_3.setText(_translate("widget", "源pdf："))
        self.file_path_btn.setText(_translate("widget", "..."))
        self.label_1.setText(_translate("widget", "x:"))
        self.label_2.setText(_translate("widget", "y:"))
        self.run_btn.setText(_translate("widget", "输出.."))


class ui_main(QDialog, Ui_widget):
	def __init__(self, parent = None):
		super(ui_main, self).__init__(parent)
		self.setupUi(self)

		self.file_path_btn.clicked.connect(self.get_file_path)
		self.run_btn.clicked.connect(self.run)

		self.font_book = []
		self.font_book_path = {}
		self.create_font_book(f"{os.environ['HOME']}/Library/Fonts")
		self.font_box.addItems(self.font_book)
		
		self.font_size_box.addItems([str((x + 1) * 4) for x in range(18)])
		self.x_box.addItems([str(x * 5) for x in range(21)])
		self.y_box.addItems([str(x * 5) for x in range(21)])
		
		self.font_size_box.setCurrentIndex(3)
		self.x_box.setCurrentIndex(10)
		self.y_box.setCurrentIndex(2)

		self.font_size_box.currentTextChanged.connect(self.font_size_change)
		self.x_box.currentTextChanged.connect(self.x_change)
		self.y_box.currentTextChanged.connect(self.y_change)

		s = '在此输入需要添加的文字，将在每一页的(x, y)位置上添加，(0, 0)为左下角，(50, 50)为中心；\n\n'\
		'使用/A来替代逐页递增的自然数，'\
		'在/A后紧跟着的<>中使用/B、/C和数字分别表示按书签递增、中文数字和起始数字；\n\n'\
		'示例：“第/A</C10>页”，将在首页添加“第十页”，次页添加“第十一页”，以此类推。'
		self.input_string_view.setPlaceholderText(s)

	def x_change(self):
		t = self.x_box.currentText()
		if not t.isdigit():
			self.x_box.setCurrentIndex(0)
		elif int(t) not in range(101):
			self.x_box.setCurrentIndex(20)

	def y_change(self):
		t = self.y_box.currentText()
		if not t.isdigit():
			self.y_box.setCurrentIndex(0)
		elif int(t) not in range(101):
			self.y_box.setCurrentIndex(20)

	def font_size_change(self):
		t = self.font_size_box.currentText()
		if not t.isdigit():
			self.font_size_box.setCurrentIndex(3)

	def get_file_path(self):
		file_path = QFileDialog.getOpenFileName(self, 'open', '.', '*.pdf')[0]
		self.file_path_line.setText(file_path)


	def create_font_book(self, font_path):
		files = os.listdir(font_path)
		for file in files:
			if file[-4:].lower() == '.ttf':
				path = os.path.join(font_path,file)
				ttfont = TTFont(path)
				font = ttfont['name'].names[4].string
				try:
					font = font.decode('utf-8')
				except:
					continue
				else:
					if isinstance(font, str) and font not in self.font_book:
						self.font_book.append(font)
						self.font_book_path.update({font: path})

	def run(self):
		self.input_path = self.file_path_line.text()
		if not os.path.exists(self.input_path):
			return

		self.output_path = QFileDialog.getSaveFileName(self, 'save file', 
			f'{self.input_path[:len(self.input_path) - 4]}_new.pdf','*.pdf')[0]
		if not self.output_path:
			return

		self.tmp_path =  f'{self.input_path[:len(self.input_path) - 4]}_'\
		f'{datetime.timestamp(datetime.now())}.pdf'

		self.x = int(self.x_box.currentText())
		self.y = int(self.y_box.currentText())
		self.font = self.font_box.currentText()
		self.fontsize = int(self.font_size_box.currentText())
		self.input_string = self.input_string_view.toPlainText()

		self.get_outline()
		self.create_tmp()
		self.merge_pdf()
		os.remove(self.tmp_path)

	def get_outline(self):
		self.outline = []
		i_pdf = PdfReader(self.input_path)
		if i_pdf.Root.Outlines:
			page = i_pdf.Root.Outlines.First
			for i, p in enumerate(i_pdf.pages):
				if id(page.Dest[0]) == id(p):
					self.outline.append(i)
					page = page.Next
					if not page:
						break

	def generator(self, input_string):
		def strip_it(s, lb):
			i = 0
			while i + 1 < len(s):
				if s[i] + s[i + 1] == lb:
					s = s[:i] + s[i + 2:]
				i += 1
			return s

		s = input_string
		l = []
		i = 0
		while i < len(s):
			if i + 1 < len(s) and s[i] + s[i + 1] == '/A':
				l.append([0, s[:i]])
				s = s[i + 2:]
				t = [1, 1, 0, 0, 0]
				if len(s) > 2 and s[0] == '<':
					j = 0
					while j < len(s) and s[j] != '>':
						j += 1
					if j < len(s) and s[j] == '>':
						k = s[1: j]
						t[2] = 'B' if '/B' in k else 0
						t[3] = 'C' if '/C' in k else 0
						k = strip_it(k, '/B')
						k = strip_it(k, '/C')
						t[1] = int(k) if k.isdigit() else 1
						s = s[j + 1:]
				l.append(t)
				i = 0
			else:
				i += 1
		if s:
			l.append([0, s])

		while True:
			g = ''
			for i, t in enumerate(l):
				if t[0]:
					x = transform(str(t[1]), 'an2cn') if t[3] == 'C' else str(t[1])
					g += x
					if t[2] == 'B':
						l[i][4] += 1
						if l[i][4] in self.outline:
							l[i][1] += 1
					else:
						l[i][1] += 1
				else:
					g += t[1]
			yield str(g)

	def create_tmp(self):
		t = canvas.Canvas(self.tmp_path)
		g = self.generator(self.input_string)
		i_pdf = PdfReader(self.input_path)
		try:
			pdfmetrics.registerFont(pdfFont(self.font, self.font_book_path[self.font]))
		except:
			self.font = 'Times-Roman'
		for page in i_pdf.pages:
			w, h = page.MediaBox[2], page.MediaBox[3]
			t.setPageSize((w, h))
			t.setFont(self.font, self.fontsize)
			t.drawCentredString(float(w) * self.x / 100, float(h) * self.y / 100, next(g))
			t.showPage()
		t.save()

	def merge_pdf(self):
		i_pdf = PdfReader(self.input_path)
		t_pdf = PdfReader(self.tmp_path)
		o_pdf = PdfWriter()
		for i in range(len(i_pdf.pages)):
			page = PageMerge(i_pdf.pages[i])
			page.add(t_pdf.pages[i]).render()
		o_pdf.write(self.output_path, i_pdf)

if __name__ == '__main__':
	app = QApplication(sys.argv)
	window_main = ui_main()
	window_main.show()
	sys.exit(app.exec_())
	
