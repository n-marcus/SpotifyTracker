# -*- coding: utf-8 -*-
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QApplication, QPushButton
from PyQt5.QtCore import pyqtSlot
import backend
import time
import threading
import os
import re
import subprocess
import SpotipyManager
import Excel



indexread = False
index = 0
startTime = 0
endTime = 0
song = ""
artist = ""

if os.name == "nt":
    import ctypes
    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("spotifylyrics.version1")

class Communicate(QtCore.QObject):
    signal = QtCore.pyqtSignal(str, str)

class Ui_Form(object):
    sync = False
    ontop = False
    open_spotify = False

    if os.name == "nt":
        settingsdir = os.getenv("APPDATA") + "\\SpotifyLyrics\\"
    else:
        settingsdir = os.path.expanduser("~") + "/.SpotifyLyrics/"
    def __init__(self):
        super().__init__()

        self.comm = Communicate()
        self.comm.signal.connect(self.change_lyrics)
        self.setupUi(Form)
        self.set_style()
        Excel.setupExcelFormatting()
        self.load_save_settings()
        if self.open_spotify:
            self.spotify()
        self.start_thread()


    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.resize(550, 610)
        Form.setMinimumSize(QtCore.QSize(350, 310))
        self.gridLayout_2 = QtWidgets.QGridLayout(Form)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.label_songname = QtWidgets.QLabel(Form)
        self.label_songname.setObjectName("label_songname")
        self.label_songname.setOpenExternalLinks(True)
        self.horizontalLayout_2.addWidget(self.label_songname, 0, QtCore.Qt.AlignLeft|QtCore.Qt.AlignVCenter)
        spacerItem = QtWidgets.QSpacerItem(40, 20, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Minimum)
        self.horizontalLayout_2.addItem(spacerItem)
        self.comboBox = QtWidgets.QComboBox(Form)
        self.comboBox.setGeometry(QtCore.QRect(160, 120, 69, 22))
        self.comboBox.setObjectName("comboBox")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.horizontalLayout_2.addWidget(self.comboBox, 0, QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.fontBox = QtWidgets.QSpinBox(Form)
        self.fontBox.setMinimum(1)
        self.fontBox.setProperty("value", 10)
        self.fontBox.setObjectName("fontBox")
        self.horizontalLayout_2.addWidget(self.fontBox, 0, QtCore.Qt.AlignRight|QtCore.Qt.AlignVCenter)
        self.verticalLayout_2.addLayout(self.horizontalLayout_2)
        self.textBrowser = QtWidgets.QTextBrowser(Form)
        self.textBrowser.setObjectName("textBrowser")
        self.textBrowser.setAcceptRichText(True)
        self.textBrowser.setStyleSheet("font-size: %spt;" % self.fontBox.value() * 2)
        self.textBrowser.setFontPointSize(self.fontBox.value())
        self.verticalLayout_2.addWidget(self.textBrowser)
        self.gridLayout_2.addLayout(self.verticalLayout_2, 2, 0, 1, 1)
        #self.button = QtGui.QPushButton('Test', self)
        #self.button.clicked.connect(self.handleButton)
        self.retranslateUi(Form)
        self.fontBox.valueChanged.connect(self.update_fontsize)
        self.comboBox.currentIndexChanged.connect(self.optionschanged)

        self.button = QPushButton("&Erase all",Form)
        self.button.setToolTip('This is an example button')
        self.button.move(50,550)
        self.button.clicked.connect(Excel.eraseExcelData)

        QtCore.QMetaObject.connectSlotsByName(Form)
        Form.setTabOrder(self.textBrowser, self.comboBox)
        Form.setTabOrder(self.comboBox, self.fontBox)

    def load_save_settings(self, save=False):
        settingsfile = self.settingsdir + "settings.ini"
        if save is False:
            if os.path.exists(settingsfile):
                with open(settingsfile, 'r') as settings:
                    for line in settings.readlines():
                        lcline = line.lower()
                        if "syncedlyrics" in lcline:
                            if "true" in lcline:
                                self.sync = True
                            else:
                                self.sync = False
                        if "alwaysontop" in lcline:
                            if "true" in lcline:
                                self.ontop = True
                            else:
                                self.ontop = False
                        if "fontsize" in lcline:
                            set = line.split("=",1)[1].strip()
                            try:
                                self.fontBox.setValue(int(set))
                            except ValueError:
                                pass
                        if "openspotify" in lcline:
                            if "true" in lcline:
                                self.open_spotify = True
                            else:
                                self.open_spotify = False

            else:
                directory = os.path.dirname(settingsfile)
                if not os.path.exists(directory):
                    os.makedirs(directory)
                with open(settingsfile, 'w+') as settings:
                    settings.write("[settings]\nAlwaysOnTop=False\nFontSize=10\nOpenSpotify=False")
            if self.ontop is True:
                Form.setWindowFlags(Form.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
                self.comboBox.setItemText(2, ("Always on Top (on)"))
                Form.show()
            if self.open_spotify is True:
                self.comboBox.setItemText(4, ("Start spotify (on)"))
        else:
            with open(settingsfile, 'w+') as settings:
                settings.write("[settings]\n")
                if self.ontop is True:
                    settings.write("AlwaysOnTop=True\n")
                else:
                    settings.write("AlwaysOnTop=False\n")
                if self.open_spotify is True:
                    settings.write("OpenSpotify=True\n")
                else:
                    settings.write("AlwaysOnTop=False\n")
                settings.write("FontSize=%s" % str(self.fontBox.value()))

    def optionschanged(self):
        current_index = self.comboBox.currentIndex()
        if current_index == 1:
            if self.ontop is False:
                self.ontop = True
                Form.setWindowFlags(Form.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
                self.comboBox.setItemText(1, ("Always on Top (on)"))
                Form.show()
            else:
                self.ontop = False
                Form.setWindowFlags(Form.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
                self.comboBox.setItemText(1, ("Always on Top"))
                Form.show()
        elif current_index == 2:
            self.load_save_settings(save=True)
        elif current_index == 3:
            if self.open_spotify is True:
                self.open_spotify = False
                self.comboBox.setItemText(3, ("Open Spotify"))
            else:
                self.open_spotify = True
                self.comboBox.setItemText(3, ("Open Spotify (on)"))
        else:
            pass
        self.comboBox.setCurrentIndex(0)


    def set_style(self):
        if os.path.exists(self.settingsdir + "theme.ini"):
            themefile = self.settingsdir + "theme.ini"
        else:
            themefile = "theme.ini"
        if os.path.exists(themefile):
            with open(themefile, 'r') as theme:
                try:
                    for setting in theme.readlines():
                        lcsetting = setting.lower()
                        try:
                            set = setting.split("=",1)[1].strip()
                        except IndexError:
                            set = ""
                        if "windowopacity" in lcsetting:
                            Form.setWindowOpacity(float(set))
                        if "backgroundcolor" in lcsetting:
                            Form.setStyleSheet("background-color: %s" % set)
                        if "lyricsbackgroundcolor" in lcsetting:
                            style = self.textBrowser.styleSheet()
                            style = style + "background-color: %s;" % set
                            self.textBrowser.setStyleSheet(style)
                        if "lyricstextcolor" in lcsetting:
                            style = self.textBrowser.styleSheet()
                            style = style + "color: %s;" % set
                            self.textBrowser.setStyleSheet(style)
                        if "songnamecolor" in lcsetting:
                            style = self.label_songname.styleSheet()
                            style = style + "color: %s;" % set
                            self.label_songname.setStyleSheet(style)
                        if "fontboxbackgroundcolor" in lcsetting:
                            style = self.fontBox.styleSheet()
                            style = style + "background-color: %s;" % set
                            self.comboBox.setStyleSheet(style)
                            self.fontBox.setStyleSheet(style)
                        if "fontboxtextcolor" in lcsetting:
                            style = self.fontBox.styleSheet()
                            style = style + "color: %s;" % set
                            self.comboBox.setStyleSheet(style)
                            self.fontBox.setStyleSheet(style)
                        if "songnameunderline" in lcsetting:
                            if "true" in set.lower():
                                style = self.label_songname.styleSheet()
                                style = style + "text-decoration: underline;"
                                self.label_songname.setStyleSheet(style)
                except Exception:
                    pass
        else:
            self.label_songname.setStyleSheet("color: black; text-decoration: underline;")
            pass

    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)

    def update_fontsize(self):
        self.textBrowser.setFontPointSize(self.fontBox.value())
        style = self.textBrowser.styleSheet()
        style = style.replace('%s' % style[style.find("font"):style.find("pt;") + 3], '')
        style = style.replace('p ', '')
        self.textBrowser.setStyleSheet(style + "p font-size: %spt;" % self.fontBox.value() * 2)
        lyrics = self.textBrowser.toPlainText()
        self.textBrowser.setText(lyrics)

    def retranslateUi(self, Form):
        _translate = QtCore.QCoreApplication.translate
        Form.setWindowTitle(_translate("Form", "Spotify Lyrics - {}".format(backend.version())))
        Form.setWindowIcon(QtGui.QIcon(self.resource_path('icon.png')))
        self.textBrowser.setText(_translate("Form", "I don't think Spotify is open")) ##THIS Writes stuff
        self.fontBox.setToolTip(_translate("Form", "Font Size"))
        self.comboBox.setItemText(0, _translate("Form", "Options"))
        self.comboBox.setItemText(1, _translate("Form", "Always on Top"))
        self.comboBox.setItemText(2, _translate("Form", "Save Settings"))
        self.comboBox.setItemText(3, _translate("Form", "Open Spotify"))

    def newSong(self, songname):
        Excel.manageIndex(songname)

        Excel.writeNewSongToFile(songname)




    def lyrics_thread(self, comm):
        global song
        global artist
        oldsongname = ""
        style = self.label_songname.styleSheet()
        if style == "":
            color = "color: black"
        else:
            color = style
        while True:
            songname = backend.getwindowtitle()
            if oldsongname != songname and songname != "":
                print("Changed!")
                self.newSong(songname)
            oldsongname = songname

            print("Playing: " + song)

            song, artist = backend.getSongData(songname)

            comm.signal.emit(songname, "Playing song " + song + " by " + artist)
            if songname != "Spotify" and songname != "": #when you switch songs


                lyrics, url, timed = backend.getlyrics(songname)
                if url == "":
                    header = songname
                else:
                    header = '''<style type="text/css">a {text-decoration: none; %s}</style><a href="%s">%s</a>''' % (color, url, songname)
            time.sleep(1)



    def start_thread(self):
        lyricsthread = threading.Thread(target=self.lyrics_thread, args=(self.comm,))
        lyricsthread.daemon = True
        lyricsthread.start()

    def change_lyrics(self, songname, lyrics):
        _translate = QtCore.QCoreApplication.translate
        self.label_songname.setText(_translate("Form", songname))
        self.textBrowser.setText(_translate("Form", lyrics))
        self.textBrowser.scrollToAnchor("#scrollHere")

    def spotify(self):
        if os.name == "nt":
            path = os.getenv("APPDATA") + '\Spotify\Spotify.exe'
            subprocess.Popen(path)


if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    Form = QtWidgets.QWidget()
    ui = Ui_Form()
    Form.show()
    sys.exit(app.exec_())
