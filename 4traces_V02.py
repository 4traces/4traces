import sys
import win32com.shell.shell as shell
import win32gui
import os
import re
import time
import codecs
import pandas as pd
import json
import numpy as np
from neo4j import GraphDatabase
from collections import deque
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QFont, QPen
from PyQt5.QtGui import QColor, QBrush, QLinearGradient, QGradient, QPainter
from PyQt5.QtCore import QPoint
from PyQt5.QtWidgets import QMessageBox, QInputDialog, QDialog
from PyQt5.Qt import QStandardItemModel, QStandardItem, QModelIndex, QIcon, QSortFilterProxyModel, QComboBox
from PyQt5.Qt import QFileSystemModel, QAbstractItemModel, QRegExp, Qt
from PyQt5.QtChart import QPieSeries, QPieSlice, QBarSeries, QBarSet
from PyQt5.QtWidgets import QApplication, QMainWindow
from PyQt5.QtChart import QChart, QChartView, QLineSeries, QCategoryAxis, QBarCategoryAxis
import textwrap
import xmltodict
import winreg as _winreg
import datetime
import matplotlib.pyplot as plt



new_Row = []
new_Row.append(QStandardItem("Test 1"))
new_Row.append(QStandardItem("Test 2"))
root.appendrow(new_Row)

# creating checkable combo box class
class CheckableComboBox(QComboBox):
    def __init__(self):
        super(CheckableComboBox, self).__init__()
        self.view().pressed.connect(self.handle_item_pressed)
        self.setModel(QStandardItemModel(self))

        # when any item get pressed

    def handle_item_pressed(self, index):

        # getting which item is pressed
        item = self.model().itemFromIndex(index)

        # make it check if unchecked and vice-versa
        if item.checkState() == QtCore.Qt.Checked:
            item.setCheckState(QtCore.Qt.Unchecked)
        else:
            item.setCheckState(QtCore.Qt.Checked)

            # calling method
        self.check_items()

        # method called by check_items

    def item_checked(self, index):

        # getting item at index
        item = self.model().item(index, 0)

        # return true if checked else false
        return item.checkState() == QtCore.Qt.Checked

        # calling method

    def checked_item_text(self, index):

        # getting item at index
        item = self.model().item(index, 0)
        if item.checkState() == QtCore.Qt.Checked:
            return item.text()

        # calling method

    def check_items(self):
        # blank list
        checkedScopes = []
        # traversing the items
        for i in range(self.count()):

            # if item is checked add it to the list
            if self.item_checked(i):
                checkedScopes.append(self.checked_item_text(i))

                # call this method
        return checkedScopes


class StandardItem(QStandardItem):
    def __init__(self, txt='', item_type='', edit_item=False, print_item=False, font_size=12, set_bold=False,
                 color=QColor(0, 0, 0)):
        super().__init__()
        fnt = QFont('Open Sans', font_size)
        fnt.setBold(set_bold)
        self.setEditable(edit_item)
        self.setDragEnabled(False)
        self.setDropEnabled(False)

        if print_item:
            self.setCheckable(True)

        if item_type == 'Folder':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-ordner_db.png'))
            self.setForeground(QColor(155, 0, 0))
        elif item_type == 'OS_Folder':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-mappe-48.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Requirement':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-requirement-48.png'))
            self.setForeground(QColor(0, 155, 0))
        elif item_type == 'TestCase':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-scorecard-40.png'))
            self.setForeground(QColor(0, 0, 155))
        elif item_type == 'TestStep':
            if print_item:
                self.setIcon(QIcon('./Icons/test_step.png'))
            self.setForeground(QColor(0, 0, 155))
        elif item_type == 'TestExecution':
            if print_item:
                self.setIcon(QIcon('./Icons/test_execution.png'))
            self.setForeground(QColor(0, 0, 155))
        elif item_type == 'TestResult':
            if print_item:
                self.setIcon(QIcon('./Icons/test_result.png'))
            self.setForeground(QColor(0, 0, 155))
        elif item_type == 'Chapter':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-page-48.png'))
            self.setForeground(QColor(0, 155, 155))
        elif item_type == 'Req_Spec':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-book-stack-48.png'))
            self.setForeground(QColor(155, 155, 155))
        elif item_type == 'Part':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-zahnrad-100.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'CPEM_Project':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-laufband-48.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Function':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-formel-fx-60.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'HookNode':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-crane-hook-64.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Company':
            if print_item:
                self.setIcon(QIcon('./Icons/company.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Role':
            if print_item:
                self.setIcon(QIcon('./Icons/role.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Department':
            if print_item:
                self.setIcon(QIcon('./Icons/gruppe.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Person':
            if print_item:
                self.setIcon(QIcon('./Icons/Person.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'File':
            if print_item:
                self.setIcon(QIcon('./Icons/imageres.dll/imageres_1305.ico'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Team':
            if print_item:
                self.setIcon(QIcon('./Icons/imageres.dll/imageres_1010.ico'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Epic':
            if print_item:
                self.setIcon(QIcon('./Icons/Epic.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Feature':
            if print_item:
                self.setIcon(QIcon('./Icons/Feature.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'UserStory':
            if print_item:
                self.setIcon(QIcon('./Icons/User_Story.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'User_Story':
            if print_item:
                self.setIcon(QIcon('./Icons/User_Story.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Project':
            if print_item:
                self.setIcon(QIcon('./Icons/3dx_project_icon.PNG'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Phase':
            if print_item:
                self.setIcon(QIcon('./Icons/3dx_phase_icon.PNG'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Task':
            if print_item:
                self.setIcon(QIcon('./Icons/Task.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Bug':
            if print_item:
                self.setIcon(QIcon('./Icons/Bug.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Comment':
            if print_item:
                self.setIcon(QIcon('./Icons/shell32/shell32_1001.ico'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Scope':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-angry-eye-64.png'))
            self.setForeground(QColor(24, 55, 55))
            self.setCheckable(False)
            self.setEditable(False)
        elif item_type == 'NewContent':
            if print_item:
                self.setIcon(QIcon('./Icons/shell32/shell32_16752.ico'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Attribute':
            if print_item:
                self.setIcon(QIcon('./Icons/Attribute.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Vendor':
            if print_item:
                self.setIcon(QIcon('./Icons/vendor.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Software':
            if print_item:
                self.setIcon(QIcon('./Icons/Software.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Physic':
            if print_item:
                self.setIcon(QIcon('./Icons/Physic.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Portfolio':
            if print_item:
                self.setIcon(QIcon('./Icons/Portfolio.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Skill':
            if print_item:
                self.setIcon(QIcon('./Icons/Skill.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Method':
            if print_item:
                self.setIcon(QIcon('./Icons/Method.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Product':
            if print_item:
                self.setIcon(QIcon('./Icons/product.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Artifact':
            if print_item:
                self.setIcon(QIcon('./Icons/Artifact.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Information':
            if print_item:
                self.setIcon(QIcon('./Icons/mmcndmgr.dll/mmcndmgr_30560.ico'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Organisation':
            if print_item:
                self.setIcon(QIcon('./Icons/icons8-flussdiagramm-100.png'))
            self.setForeground(QColor(24, 55, 55))
        elif item_type == 'Skript':
            if print_item:
                self.setIcon(QIcon('./Icons/Skript.png'))
            self.setForeground(QColor(24, 55, 55))
        else:
            self.setForeground(QColor(24, 55, 55))

        if not color == QColor(0, 0, 0):
            self.setForeground(QColor(255, 0, 0))
            fnt.setBold(True)

        self.setFont(fnt)
        self.setText(txt)


class StandardTableItem(QStandardItem):
    def __init__(self, txt='', edit_item=False, font_size=12, set_bold=False, color=QColor(0, 0, 0)):
        super().__init__()
        fnt = QFont('Open Sans', font_size)
        fnt.setBold(set_bold)
        self.setEditable(edit_item)
        self.setDragEnabled(False)
        self.setDropEnabled(False)

        self.setCheckable(False)
        self.setForeground(color)
        self.setFont(fnt)
        self.setText(txt)


def extract_table_old(obj):
    """Rekrusion zum abflachen der strukturierten JSON"""
    arr = []
    current_Parent = ''

    def extract_dict(obj, arr, current_Parent):
        """Recursively search for values of key in JSON tree."""
        if isinstance(obj, dict):
            temp_dict = {}
            for k, v in obj.items():
                if isinstance(v, (dict, list)):
                    new_parent = obj.get('_id', '')
                    extract_dict(v, arr, new_parent)
                else:
                    temp_dict.update(({'child_of': current_Parent}))
                    temp_dict.update({k: v})
            arr.append(temp_dict)
        elif isinstance(obj, list):
            for item in obj:
                extract_dict(item, arr, current_Parent)
        return arr

    list_temp = extract_dict(obj, arr, current_Parent)
    results = []
    for item in list_temp:
        if item not in results:
            results.append(item)
    return results


def extract_table(obj):
    print("extract start")
    """Rekrusion zum abflachen der strukturierten JSON"""
    arr = []
    current_Parent = ''
    link_type = ''

    # print(obj)

    def extract_dict(obj, arr, current_Parent, link_type):
        """Recursively search for values of key in JSON tree."""
        if isinstance(obj, dict):
            temp_dict = {}
            new_child_loop = False
            child_lst = []
            new_parent = ''
            parent_link = ''
            for k, v in obj.items():
                if isinstance(v, (dict, list)):
                    child_lst.extend(v)
                    new_parent = obj.get('_id', '')
                    new_child_loop = True
                    parent_link = k
                    # print_all_nodes(v, arr, new_parent, k)
                else:
                    temp_dict.update(({'child_of': current_Parent}))
                    temp_dict.update(({'parent_link': link_type}))
                    if "HookNode" in k:
                        temp_dict.update(({'hooknode': v}))
                    temp_dict.update({k: v})
            arr.append(temp_dict)
            if new_child_loop:
                extract_dict(child_lst, arr, new_parent, parent_link)
        elif isinstance(obj, list):
            for item in obj:
                extract_dict(item, arr, current_Parent, link_type)
        return arr

    list_temp = extract_dict(obj, arr, current_Parent, link_type)

    weitermachen = True
    if len(list_temp) > 4999:
        print(len(list_temp))
        weitermachen = query_info_dialog()
    if weitermachen:
        print("extract done")
        results = []
        for item in list_temp:
            if item not in results:
                results.append(item)
        # print(results)
        return results
    else:
        print("Die Querie liefert zu viel Icons")
        return []


def data_by_id(data_by_id_obejct: list, dict_out: dict):
    for dict_temp in data_by_id_obejct:
        if dict_temp:
            dict_out.update({dict_temp['_id']: dict_temp})

    return dict_out


def modify_windows_libary(libary_name: str, folder_dict: dict):
    if libary_name == "Aktuelle_Bearbeitung":
        libary_file = "C:/Users/i004625/AppData/Roaming/Microsoft/Windows/Libraries/Aktuelle_Bearbeitung.library-ms"
        sym_link_folder = "C:\\Arbeitsbereich\\01_Aktuelle_Bearbeitung\\"
    elif libary_name == "My Workspace":
        libary_file = "C:/Users/i004625/AppData/Roaming/Microsoft/Windows/Libraries/My_Workspace.library-ms"
        sym_link_folder = "C:\\Arbeitsbereich\\02_Workplace\\"


    if False:
        default_folder = {"isDefaultSaveLocation": "true",
                          "isDefaultNonOwnerSaveLocation": "true",
                          "isSupported": "false",
                          "simpleLocation": {"url": "C:\\Arbeitsbereich"}}

        libary_file_tmp = "./Input/windows_libary_template.xml"
        f = codecs.open(libary_file_tmp, "r", "utf-8")
        xml_data = xmltodict.parse(f.read())

        new_list = []
        for k, fol in folder_dict.items():
            if not "http" in fol:
                fol = fol.replace("/", "\\")
            new_dct = dict(default_folder)
            new_dct.update({'simpleLocation': {'url': fol}})
            new_list.append(new_dct.copy())

        xml_data["libraryDescription"]["searchConnectorDescriptionList"]["searchConnectorDescription"] = new_list

        with codecs.open(libary_file, 'w', encoding='utf8') as f:
            f.write(xmltodict.unparse(xml_data))

    if True:
        d = sym_link_folder
        commands = []
        for o in os.listdir(d):
            if os.path.isdir(os.path.join(d, o)):
                print(o)
                new_folder = True
                if o not in folder_dict.keys():
                    commands.append('Rmdir /s/q "' + os.path.join(d, o) + '"')
                    # os.remove(os.path.join(d, o))
            elif os.path.isfile(os.path.join(d, o)):
                print(o)
                new_folder = True
                if o not in folder_dict.keys():
                    commands.append('del /s/q "' + os.path.join(d, o) + '"')

        for k, v in folder_dict.items():
            if not os.path.exists(d + k):
                if "http" in v:
                    v = v.replace("https:", "")
                    v = v.replace("http:", "")
                    v = v.replace("\\", "/")
                    v = v.replace("%20", " ")

                    tst = v.split("/")[-1]
                    if "." in tst:
                        commands.append('mklink "' + d + k + '" "' + v + '"')
                    else:
                        commands.append('mklink /d "' + d + k + '" "' + v + '"')

                elif os.path.isdir(v):
                    commands.append('mklink /d "' + d + k + '" "' + v + '"')
                elif os.path.isfile(v):
                    commands.append('mklink "' + d + k + '" "' + v + '"')


        if len(commands) > 0:
            command = ' & '.join(commands)
            print(command)
            shell.ShellExecuteEx(lpVerb='runas', lpFile='cmd.exe', lpParameters='/c ' + command)


def getDefaultIcon(filename):
    '''Retrieve the default icon of a filename'''
    (root, extension) = os.path.splitext(filename)
    if extension:
        try:
            value_name = _winreg.QueryValue(_winreg.HKEY_CLASSES_ROOT, extension)
        except _winreg.error:
            value_name = None
    else:
        value_name = None
    if value_name:
        try:
            icon = _winreg.QueryValue(_winreg.HKEY_CLASSES_ROOT,
                                      value_name + "\\DefaultIcon")
        except _winreg.error:
            icon = None
    else:
        icon = None
    return icon


def delete_node_dialog():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Question)

    msg.setText("Detach node from structure")
    msg.setInformativeText("This node has no other parent. Do you realy want to delete it from this structure? ")
    msg.setWindowTitle("Datach node")
    msg.setDetailedText("This node has no other parent. If you delete it from your structure it will remain in the "
                        "database but without a connection to a hook node")
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    retval = msg.exec_()
    if retval == QMessageBox.Yes:
        return True
    else:
        return False


def query_info_dialog():
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Question)

    msg.setText("Query to long")
    msg.setInformativeText("This query is to long. Do you want to continue with a reduced amount of info")
    msg.setWindowTitle("Query to long")
    msg.setDetailedText("This Query is to long. To displayed amount of notes has been reduced.")
    msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    retval = msg.exec_()
    if retval == QMessageBox.Yes:
        return True
    else:
        return False


def start_neo4j_api():
    with open('./Input/config.json') as json_file:
        config = json.load(json_file)
    driver = GraphDatabase.driver(config["Server"], auth=(config["Server_Login"], config["Server_PW"]))
    return driver


def cleanup_database(driver):
    # ################################################################################################################
    # ###############    Add common attributes     ###################################################################
    # ################################################################################################################

    common_attributes = []
    common_attributes.append('Details')
    common_attributes.append('Description')
    common_attributes.append('DeepLink')
    common_attributes.append('Key')
    common_attributes.append('Owner')
    common_attributes.append('ModifyDate')

    with driver.session() as session:
        for head in common_attributes:
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (n) WHERE NOT EXISTS(n.' + head + ')')
            qry_cmd_Lst.append('SET n.' + head + ' = \"\"')
            qry = ' '.join(qry_cmd_Lst)
            print(qry)
            records = session.run(qry)
        session.close()

    # ################################################################################################################
    # ###############    Set Key where not set     ###################################################################
    # ################################################################################################################
    lst = []
    with driver.session() as session:
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x) WHERE x.Key = ""')
        qry_cmd_Lst.append('SET x.Key = apoc.create.uuid()')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)
        records = session.run(qry)

    # ################################################################################################################
    # ###############    Set Key where not set     ###################################################################
    # ################################################################################################################
    lst = []
    with driver.session() as session:
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x) WHERE x.Owner = ""')
        qry_cmd_Lst.append('SET x.Owner = \"System\"')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)
        records = session.run(qry)
    session.close()

    with driver.session() as session:
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (n:Folder) WHERE NOT EXISTS(n.Sync)')
        qry_cmd_Lst.append('SET n.Sync = \"False\"')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)
        records = session.run(qry)
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (m:File) WHERE NOT EXISTS(m.Sync)')
        qry_cmd_Lst.append('SET m.Sync = \"False\"')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)
        records = session.run(qry)
    session.close()


class Query_Dialog(QDialog):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Dialog')
        self.setObjectName("Dialog")
        self.resize(774, 361)
        self.buttonBox = QtWidgets.QDialogButtonBox(self)
        self.buttonBox.setGeometry(QtCore.QRect(530, 300, 160, 25))
        self.buttonBox.setOrientation(QtCore.Qt.Horizontal)
        self.buttonBox.setStandardButtons(QtWidgets.QDialogButtonBox.Cancel | QtWidgets.QDialogButtonBox.Ok)
        self.buttonBox.setObjectName("buttonBox")
        self.widget = QtWidgets.QWidget(self)
        self.widget.setGeometry(QtCore.QRect(0, 8, 631, 161))
        self.widget.setObjectName("widget")
        self.gridLayout = QtWidgets.QGridLayout(self.widget)
        self.gridLayout.setContentsMargins(0, 0, 0, 0)
        self.gridLayout.setObjectName("gridLayout")
        self.comboBox_type_1 = QtWidgets.QComboBox(self.widget)
        self.comboBox_type_1.setObjectName("comboBox_type_1")
        self.gridLayout.addWidget(self.comboBox_type_1, 0, 0, 1, 1)
        self.comboBox_type_2 = QtWidgets.QComboBox(self.widget)
        self.comboBox_type_2.setObjectName("comboBox_type_2")
        self.gridLayout.addWidget(self.comboBox_type_2, 0, 1, 1, 1)
        self.comboBox_type_3 = QtWidgets.QComboBox(self.widget)
        self.comboBox_type_3.setObjectName("comboBox_type_3")
        self.gridLayout.addWidget(self.comboBox_type_3, 0, 2, 1, 1)
        self.comboBox_attribute_1 = QtWidgets.QComboBox(self.widget)
        self.comboBox_attribute_1.setObjectName("comboBox_attribute_1")
        self.gridLayout.addWidget(self.comboBox_attribute_1, 1, 0, 1, 1)
        self.comboBox_attribute_2 = QtWidgets.QComboBox(self.widget)
        self.comboBox_attribute_2.setObjectName("comboBox_attribute_2")
        self.gridLayout.addWidget(self.comboBox_attribute_2, 1, 1, 1, 1)
        self.comboBox_attribute_3 = QtWidgets.QComboBox(self.widget)
        self.comboBox_attribute_3.setObjectName("comboBox_attribute_3")
        self.gridLayout.addWidget(self.comboBox_attribute_3, 1, 2, 1, 1)
        self.comboBox_depth_2 = QtWidgets.QComboBox(self.widget)
        self.comboBox_depth_2.setObjectName("comboBox_depth_2")
        self.gridLayout.addWidget(self.comboBox_depth_2, 2, 1, 1, 1)
        self.comboBox_depth_3 = QtWidgets.QComboBox(self.widget)
        self.comboBox_depth_3.setObjectName("comboBox_depth_3")
        self.gridLayout.addWidget(self.comboBox_depth_3, 2, 2, 1, 1)
        self.lineEdit_value_1 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_value_1.setObjectName("lineEdit_value_1")
        self.gridLayout.addWidget(self.lineEdit_value_1, 3, 0, 1, 1)
        self.lineEdit_value_2 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_value_2.setObjectName("lineEdit_value_2")
        self.gridLayout.addWidget(self.lineEdit_value_2, 3, 1, 1, 1)
        self.lineEdit_value_3 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_value_3.setObjectName("lineEdit_value_3")
        self.gridLayout.addWidget(self.lineEdit_value_3, 3, 2, 1, 1)

        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)
        QtCore.QMetaObject.connectSlotsByName(self)

        self.driver = start_neo4j_api()
        self.comboBox_type_1.currentTextChanged.connect(lambda: self.update_attributes(1))
        self.comboBox_type_2.currentTextChanged.connect(lambda: self.update_attributes(2))
        self.comboBox_type_3.currentTextChanged.connect(lambda: self.update_attributes(3))

        self.comboBox_depth_2.addItems(['1', '2', '3', '4', '5', '6', '7', '8'])
        self.comboBox_depth_3.addItems(['shortest', '1', '2', '3', '4', '5', '6', '7', '8'])

    def update_attributes(self, combo_nr: int):
        if combo_nr == 1:
            item_label = self.comboBox_type_1.currentText()
            self.comboBox_attribute_1.clear()
        elif combo_nr == 2:
            item_label = self.comboBox_type_2.currentText()
            self.comboBox_attribute_2.clear()
        elif combo_nr == 3:
            item_label = self.comboBox_type_3.currentText()
            self.comboBox_attribute_3.clear()

        if not item_label == '*':
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (p:' + item_label + ') WITH DISTINCT keys(p) AS keys')
            qry_cmd_Lst.append('UNWIND keys AS keyslisting WITH DISTINCT keyslisting AS allfields')
            qry_cmd_Lst.append('RETURN allfields;')
            qry = ' '.join(qry_cmd_Lst)
            print(qry)

            ret_val = []
            ret_val.append("*")
            with self.driver.session() as session:
                records = session.run(qry)
                print("neo_hook_scope_path_query QUERY DONE")

                for node in records:
                    print(node[0])
                    ret_val.append(node[0])
            session.close()
            ret_val = sorted(ret_val, key=str.casefold)

            if combo_nr == 1:
                self.comboBox_attribute_1.clear()
                self.comboBox_attribute_1.addItems(ret_val)
            elif combo_nr == 2:
                self.comboBox_attribute_2.clear()
                self.comboBox_attribute_2.addItems(ret_val)
            elif combo_nr == 3:
                self.comboBox_attribute_3.clear()
                self.comboBox_attribute_3.addItems(ret_val)
        else:
            ret_val = ["*"]
            if combo_nr == 1:
                self.comboBox_attribute_1.clear()
                self.comboBox_attribute_1.addItems(ret_val)
            elif combo_nr == 2:
                self.comboBox_attribute_2.clear()
                self.comboBox_attribute_2.addItems(ret_val)
            elif combo_nr == 3:
                self.comboBox_attribute_3.clear()
                self.comboBox_attribute_3.addItems(ret_val)

    def get_path_query_string(self):
        start_label = ':' + self.comboBox_type_1.currentText()
        start_label = start_label.replace(":*", "")
        start_attribute = self.comboBox_attribute_1.currentText()
        start_attribute_value = self.lineEdit_value_1.text()

        mid_label = self.comboBox_type_2.currentText()
        mid_attribute = self.comboBox_attribute_2.currentText()
        mid_attribute_value = self.lineEdit_value_2.text()

        end_label = ':' + self.comboBox_type_3.currentText()
        end_label = end_label.replace(":*", "")
        end_attribute = self.comboBox_attribute_3.currentText()
        end_attribute_value = self.lineEdit_value_3.text()

        level_depth_mid = self.comboBox_depth_2.currentText()
        level_depth_end = self.comboBox_depth_3.currentText()

        qry_cmd_Lst = []

        if start_attribute == "*":
            qry_cmd_Lst.append('MATCH (x' + start_label + ')')
            qry_cmd_Lst.append('WHERE (any(prop in keys(x) where x[prop] =~ \"(?i).*' + start_attribute_value + '.*\"))')
        else:
            qry_cmd_Lst.append('MATCH (x' + start_label + ') WHERE x.' + start_attribute +
                               ' =~ \"(?i).*' + start_attribute_value + '.*\"')
        # if not end_attribute_value:
        #     qry_cmd_Lst.append('MATCH (y:' + end_label + ')')
        # else:

        if end_attribute == "*":
            qry_cmd_Lst.append('MATCH (y' + end_label + ')')
            qry_cmd_Lst.append('WHERE (any(prop in keys(y) where y[prop] =~ \"(?i).*' + end_attribute_value + '.*\"))')
        else:
            qry_cmd_Lst.append('MATCH (y' + end_label + ') WHERE y.' + end_attribute +
                               ' =~ \"(?i).*' + end_attribute_value + '.*\"')

        print("mid_label", mid_label)
        if mid_label == "*":
            if level_depth_end == 'shortest':
                qry_cmd_Lst.append('MATCH path = shortestPath((x)-[*0..' + level_depth_mid + ']-(y))')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[*0..' + level_depth_end + ']-(y)')
        else:
            if mid_attribute == "*":
                qry_cmd_Lst.append('MATCH (z' + end_label + ')')
                qry_cmd_Lst.append(
                    'WHERE (any(prop in keys(z) where z[prop] =~ \"(?i).*' + end_attribute_value + '.*\"))')
            else:
                qry_cmd_Lst.append('MATCH (z' + mid_label + ') WHERE z.' + mid_attribute +
                                   ' =~ \"(?i).*' + mid_attribute_value + '.*\"')

            qry_cmd_Lst.append('MATCH path = (x)-[*0..' + level_depth_mid + ']-(z)-[*0..' + level_depth_end + ']-(y)')
        # qry_cmd_Lst.append('WHERE apoc.coll.duplicates(NODES(path)) = []')
        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)

        return qry


class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        self.driver = start_neo4j_api()
        cleanup_database(self.driver)
        self.databyid = {}
        self.treeView_1_selected = []
        self.treeView_2_selected = []
        self.treeView_3_selected = []
        self.treeView_4_selected = []

        self.currentitemindex = QModelIndex

        self.icon_dict = {}

        self.tree_font_size = 10

        df = pd.read_excel('./Input/Data_Model_Hirachie.xlsx', sheet_name='Tabelle1')
        df = df.fillna(0)
        df = df.set_index('Oben_hat_links')
        self.link_data_model = df.astype(int)
        del df

        self.model_tree_1 = QtGui.QStandardItemModel()
        self.model_tree_2 = QtGui.QStandardItemModel()
        self.model_tree_3 = QtGui.QStandardItemModel()
        self.model_tree_4 = QtGui.QStandardItemModel()

        self.proxyModel_tree_1 = QSortFilterProxyModel()
        self.proxyModel_tree_1_1 = QSortFilterProxyModel()
        self.proxyModel_tree_1_2 = QSortFilterProxyModel()

        self.proxyModel_tree_2 = QSortFilterProxyModel()
        self.proxyModel_tree_2_1 = QSortFilterProxyModel()
        self.proxyModel_tree_2_2 = QSortFilterProxyModel()

        self.proxyModel_tree_3 = QSortFilterProxyModel()
        self.proxyModel_tree_4 = QSortFilterProxyModel()

        self.fileSystemModel = QFileSystemModel()
        self.model_tree_2_ready = True

        self.headers = []
        self.headers.append(['Title', 'Type', 'NeoID', 'Parent Link', 'Details', 'Key', 'ModifyDate', 'DeepLink',
                             'Comment', 'Component', 'Description', 'Origin', 'Priority', 'Responsible'])
        self.headers.append(['Title', 'Type', 'NeoID'])
        self.headers.append(['Title', 'Type', 'NeoID', 'Parent Link'])

        self.table_model = QStandardItemModel()
        self.table_model.setHorizontalHeaderLabels(['Attribute', 'Value'])

        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1139, 880)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_5 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName("tabWidget")
        self.tab_1 = QtWidgets.QWidget()
        self.tab_1.setObjectName("tab_1")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.tab_1)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.splitter_3 = QtWidgets.QSplitter(self.tab_1)
        self.splitter_3.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_3.setObjectName("splitter_3")
        self.widget = QtWidgets.QWidget(self.splitter_3)
        self.widget.setObjectName("widget")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.widget)
        self.gridLayout_3.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.treeView_1 = QtWidgets.QTreeView(self.widget)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.MinimumExpanding, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.treeView_1.sizePolicy().hasHeightForWidth())
        self.treeView_1.setSizePolicy(sizePolicy)
        self.treeView_1.setAcceptDrops(True)
        self.treeView_1.setSortingEnabled(True)
        self.treeView_1.setExpandsOnDoubleClick(False)
        self.treeView_1.setObjectName("treeView_1")
        self.treeView_1.header().setDefaultSectionSize(200)
        self.gridLayout_3.addWidget(self.treeView_1, 4, 0, 1, 3)
        self.comboBox_levels_1 = QtWidgets.QComboBox(self.widget)
        self.comboBox_levels_1.setObjectName("comboBox_levels_1")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.comboBox_levels_1.addItem("")
        self.gridLayout_3.addWidget(self.comboBox_levels_1, 1, 1, 1, 1)
        self.lineEdit_1 = QtWidgets.QLineEdit(self.widget)
        self.lineEdit_1.setObjectName("lineEdit_1")
        self.gridLayout_3.addWidget(self.lineEdit_1, 3, 0, 1, 3)
        self.comboBox_multi_1 = CheckableComboBox()
        self.comboBox_multi_1.setObjectName("comboBox_multi_1")
        self.gridLayout_3.addWidget(self.comboBox_multi_1, 0, 1, 1, 2)
        self.pushButton_1_trace = QtWidgets.QPushButton(self.widget)
        self.pushButton_1_trace.setObjectName("pushButton_1_trace")
        self.gridLayout_3.addWidget(self.pushButton_1_trace, 2, 2, 1, 1)
        self.checkBox_respect_related_1 = QtWidgets.QCheckBox(self.widget)
        self.checkBox_respect_related_1.setObjectName("checkBox_respect_related_1")
        self.gridLayout_3.addWidget(self.checkBox_respect_related_1, 2, 0, 1, 1)
        self.comboBox_tree_1_Endnode = QtWidgets.QComboBox(self.widget)
        self.comboBox_tree_1_Endnode.setObjectName("comboBox_tree_1_Endnode")
        self.gridLayout_3.addWidget(self.comboBox_tree_1_Endnode, 1, 2, 1, 1)
        self.pushButton_tree_1_reload = QtWidgets.QPushButton(self.widget)
        self.pushButton_tree_1_reload.setObjectName("pushButton_tree_1_reload")
        self.gridLayout_3.addWidget(self.pushButton_tree_1_reload, 1, 0, 1, 1)
        self.comboBox_Hook_Node_1 = QtWidgets.QComboBox(self.widget)
        self.comboBox_Hook_Node_1.setObjectName("comboBox_Hook_Node_1")
        self.gridLayout_3.addWidget(self.comboBox_Hook_Node_1, 0, 0, 1, 1)
        self.comboBox_expand_1 = QtWidgets.QComboBox(self.widget)
        self.comboBox_expand_1.setObjectName("comboBox_expand_1")
        self.gridLayout_3.addWidget(self.comboBox_expand_1, 2, 1, 1, 1)
        self.splitter = QtWidgets.QSplitter(self.splitter_3)
        self.splitter.setOrientation(QtCore.Qt.Vertical)
        self.splitter.setObjectName("splitter")
        self.tableView = QtWidgets.QTableView(self.splitter)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.tableView.sizePolicy().hasHeightForWidth())
        self.tableView.setSizePolicy(sizePolicy)
        self.tableView.setAcceptDrops(False)
        self.tableView.setSortingEnabled(True)
        self.tableView.setObjectName("tableView")
        self.groupBox = QtWidgets.QGroupBox(self.splitter)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Preferred, QtWidgets.QSizePolicy.Minimum)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.groupBox.sizePolicy().hasHeightForWidth())
        self.groupBox.setSizePolicy(sizePolicy)
        self.groupBox.setObjectName("groupBox")
        self.gridLayout = QtWidgets.QGridLayout(self.groupBox)
        self.gridLayout.setObjectName("gridLayout")
        self.pushButton_add_to_Scope = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_add_to_Scope.setObjectName("pushButton_add_to_Scope")
        self.gridLayout.addWidget(self.pushButton_add_to_Scope, 1, 2, 1, 1)
        self.comboBox_applay_Scope_1 = QtWidgets.QComboBox(self.groupBox)
        self.comboBox_applay_Scope_1.setObjectName("comboBox_applay_Scope_1")
        self.comboBox_applay_Scope_1.addItem("")
        self.comboBox_applay_Scope_1.setItemText(0, "")
        self.gridLayout.addWidget(self.comboBox_applay_Scope_1, 1, 1, 1, 1)
        self.pushButton_save_change = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_save_change.setObjectName("pushButton_save_change")
        self.gridLayout.addWidget(self.pushButton_save_change, 0, 1, 1, 1)
        self.checkBox_use_hooknode = QtWidgets.QCheckBox(self.groupBox)
        self.checkBox_use_hooknode.setObjectName("checkBox_use_hooknode")
        self.gridLayout.addWidget(self.checkBox_use_hooknode, 2, 0, 1, 1)
        self.pushButton_2 = QtWidgets.QPushButton(self.groupBox)
        self.pushButton_2.setObjectName("pushButton_2")
        self.gridLayout.addWidget(self.pushButton_2, 1, 0, 1, 1)
        self.lineEdit_2 = QtWidgets.QLineEdit(self.groupBox)
        self.lineEdit_2.setObjectName("lineEdit_2")
        self.gridLayout.addWidget(self.lineEdit_2, 3, 0, 1, 3)
        self.treeView_2 = QtWidgets.QTreeView(self.splitter)
        sizePolicy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        sizePolicy.setHorizontalStretch(0)
        sizePolicy.setVerticalStretch(0)
        sizePolicy.setHeightForWidth(self.treeView_2.sizePolicy().hasHeightForWidth())
        self.treeView_2.setSizePolicy(sizePolicy)
        self.treeView_2.setSortingEnabled(True)
        self.treeView_2.setExpandsOnDoubleClick(False)
        self.treeView_2.setObjectName("treeView_2")
        self.treeView_4 = QtWidgets.QTreeView(self.splitter)
        self.treeView_4.setSortingEnabled(True)
        self.treeView_4.setExpandsOnDoubleClick(False)
        self.treeView_4.setObjectName("treeView_4")
        self.splitter_2 = QtWidgets.QSplitter(self.splitter_3)
        self.splitter_2.setOrientation(QtCore.Qt.Vertical)
        self.splitter_2.setObjectName("splitter_2")
        self.layoutWidget = QtWidgets.QWidget(self.splitter_2)
        self.layoutWidget.setObjectName("layoutWidget")
        self.gridLayout_2 = QtWidgets.QGridLayout(self.layoutWidget)
        self.gridLayout_2.setContentsMargins(0, 0, 0, 0)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.comboBox_4search_label_1 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_label_1.setObjectName("comboBox_4search_label_1")
        self.comboBox_4search_label_1.addItem("")
        self.gridLayout_2.addWidget(self.comboBox_4search_label_1, 0, 0, 1, 1)
        self.comboBox_4search_label_2 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_label_2.setObjectName("comboBox_4search_label_2")
        self.comboBox_4search_label_2.addItem("")
        self.gridLayout_2.addWidget(self.comboBox_4search_label_2, 0, 1, 1, 1)
        self.comboBox_4search_label_3 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_label_3.setObjectName("comboBox_4search_label_3")
        self.comboBox_4search_label_3.addItem("")
        self.gridLayout_2.addWidget(self.comboBox_4search_label_3, 0, 2, 1, 1)
        self.comboBox_4search_property_1 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_property_1.setObjectName("comboBox_4search_property_1")
        self.gridLayout_2.addWidget(self.comboBox_4search_property_1, 1, 0, 1, 1)
        self.comboBox_4search_property_2 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_property_2.setObjectName("comboBox_4search_property_2")
        self.gridLayout_2.addWidget(self.comboBox_4search_property_2, 1, 1, 1, 1)
        self.comboBox_4search_property_3 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_property_3.setObjectName("comboBox_4search_property_3")
        self.gridLayout_2.addWidget(self.comboBox_4search_property_3, 1, 2, 1, 1)
        self.comboBox_4search_depth_2 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_depth_2.setObjectName("comboBox_4search_depth_2")
        self.gridLayout_2.addWidget(self.comboBox_4search_depth_2, 2, 1, 1, 1)
        self.comboBox_4search_depth_3 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_depth_3.setObjectName("comboBox_4search_depth_3")
        self.gridLayout_2.addWidget(self.comboBox_4search_depth_3, 2, 2, 1, 1)
        self.comboBox_4search_direction_2 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_direction_2.setObjectName("comboBox_4search_direction_2")
        self.comboBox_4search_direction_2.addItem("")
        self.comboBox_4search_direction_2.addItem("")
        self.comboBox_4search_direction_2.addItem("")
        self.gridLayout_2.addWidget(self.comboBox_4search_direction_2, 3, 1, 1, 1)
        self.comboBox_4search_direction_3 = QtWidgets.QComboBox(self.layoutWidget)
        self.comboBox_4search_direction_3.setObjectName("comboBox_4search_direction_3")
        self.comboBox_4search_direction_3.addItem("")
        self.comboBox_4search_direction_3.addItem("")
        self.comboBox_4search_direction_3.addItem("")
        self.gridLayout_2.addWidget(self.comboBox_4search_direction_3, 3, 2, 1, 1)
        self.lineEdit_4search_value_1 = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_4search_value_1.setObjectName("lineEdit_4search_value_1")
        self.gridLayout_2.addWidget(self.lineEdit_4search_value_1, 4, 0, 1, 1)
        self.lineEdit_4search_value_2 = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_4search_value_2.setObjectName("lineEdit_4search_value_2")
        self.gridLayout_2.addWidget(self.lineEdit_4search_value_2, 4, 1, 1, 1)
        self.lineEdit_4search_value_3 = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_4search_value_3.setObjectName("lineEdit_4search_value_3")
        self.gridLayout_2.addWidget(self.lineEdit_4search_value_3, 4, 2, 1, 1)
        self.pushButton_tree_3_reload = QtWidgets.QPushButton(self.layoutWidget)
        self.pushButton_tree_3_reload.setObjectName("pushButton_tree_3_reload")
        self.gridLayout_2.addWidget(self.pushButton_tree_3_reload, 5, 0, 1, 1)
        self.checkBox_respect_HookNode = QtWidgets.QCheckBox(self.layoutWidget)
        self.checkBox_respect_HookNode.setObjectName("checkBox_respect_HookNode")
        self.gridLayout_2.addWidget(self.checkBox_respect_HookNode, 5, 1, 1, 2)
        self.lineEdit_3 = QtWidgets.QLineEdit(self.layoutWidget)
        self.lineEdit_3.setObjectName("lineEdit_3")
        self.gridLayout_2.addWidget(self.lineEdit_3, 6, 0, 1, 3)
        self.treeView_3 = QtWidgets.QTreeView(self.splitter_2)
        self.treeView_3.setSortingEnabled(True)
        self.treeView_3.setExpandsOnDoubleClick(False)
        self.treeView_3.setObjectName("treeView_3")
        self.gridLayout_4.addWidget(self.splitter_3, 0, 0, 1, 1)
        self.tabWidget.addTab(self.tab_1, "")
        self.tab_2 = QtWidgets.QWidget()
        self.tab_2.setObjectName("tab_2")
        self.gridLayout_10 = QtWidgets.QGridLayout(self.tab_2)
        self.gridLayout_10.setObjectName("gridLayout_10")
        self.gridLayout_6 = QtWidgets.QGridLayout()
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.graphicsView = QChartView(self.tab_2)
        self.graphicsView.setObjectName("graphicsView")
        self.gridLayout_6.addWidget(self.graphicsView, 0, 0, 1, 1)
        self.graphicsView_2 = QChartView(self.tab_2)
        self.graphicsView_2.setObjectName("graphicsView_2")
        self.gridLayout_6.addWidget(self.graphicsView_2, 0, 1, 1, 1)
        self.graphicsView_3 = QChartView(self.tab_2)
        self.graphicsView_3.setObjectName("graphicsView_3")
        self.gridLayout_6.addWidget(self.graphicsView_3, 1, 0, 1, 1)
        self.graphicsView_4 = QChartView(self.tab_2)
        self.graphicsView_4.setObjectName("graphicsView_4")
        self.gridLayout_6.addWidget(self.graphicsView_4, 1, 1, 1, 1)
        self.gridLayout_10.addLayout(self.gridLayout_6, 0, 0, 1, 1)
        self.tabWidget.addTab(self.tab_2, "")
        self.gridLayout_5.addWidget(self.tabWidget, 0, 1, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1139, 20))
        self.menubar.setObjectName("menubar")
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuCreate = QtWidgets.QMenu(self.menubar)
        self.menuCreate.setObjectName("menuCreate")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)
        self.actionSave_DB = QtWidgets.QAction(MainWindow)
        self.actionSave_DB.setObjectName("actionSave_DB")
        self.actionSave_to_Excel = QtWidgets.QAction(MainWindow)
        self.actionSave_to_Excel.setObjectName("actionSave_to_Excel")
        self.menuFile.addAction(self.actionSave_DB)
        self.menuFile.addAction(self.actionSave_to_Excel)
        self.menubar.addAction(self.menuFile.menuAction())
        self.menubar.addAction(self.menuCreate.menuAction())

        self.load_deeplink = QtWidgets.QAction()
        self.load_deeplink.setText("Load Deep Link")

        self.searchfornextartefact = QtWidgets.QAction()
        self.searchfornextartefact.setText("Search for next artefact")

        self.test_querie_action = QtWidgets.QAction()
        self.test_querie_action.setText("Test Querie")


        self.comboBox_Hook_Node_1.addItems(["Product Structure", "Fuctional Structure"])

        self.tableView.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)


        self.retranslateUi(MainWindow)
        self.treeView_1.clicked['QModelIndex'].connect(self.show_details_of_node)
        self.treeView_2.clicked['QModelIndex'].connect(self.show_details_of_node)
        self.treeView_3.clicked['QModelIndex'].connect(self.show_details_of_node)
        self.treeView_4.clicked['QModelIndex'].connect(self.show_details_of_node)
        self.treeView_1.doubleClicked['QModelIndex'].connect(self.show_neighbor)
        self.treeView_2.doubleClicked['QModelIndex'].connect(self.show_neighbor)
        self.treeView_3.doubleClicked['QModelIndex'].connect(self.show_neighbor)
        self.treeView_4.doubleClicked['QModelIndex'].connect(self.show_neighbor)
        self.pushButton_tree_1_reload.clicked.connect(self.tree_reload_clicked)
        # self.pushButton_tree_3_reload.clicked.connect(lambda: self.tree_reload_clicked("tree_3"))
        self.pushButton_tree_3_reload.clicked.connect(self.path_query)
        # self.model_tree_1.itemChanged.connect(self.edit_node)
        self.pushButton_add_to_Scope.clicked.connect(self.add_node_to_scope)
        self.pushButton_2.clicked.connect(self.delete_node_from_scope)
        #self.pushButton_add_tree.clicked.connect(self.add_tree)
        self.lineEdit_1.returnPressed.connect(lambda: self.filter_tree_view(1))
        self.lineEdit_3.returnPressed.connect(lambda: self.filter_tree_view(3))
        self.lineEdit_2.returnPressed.connect(lambda: self.filter_tree_view(2))
        self.pushButton_save_change.clicked.connect(self.edit_node)
        self.pushButton_1_trace.clicked.connect(lambda: self.search_for_next_artefact(self.model_tree_1.invisibleRootItem()))

        self.comboBox_4search_label_1.currentTextChanged.connect(lambda: self.update_attributes(1))
        self.comboBox_4search_label_2.currentTextChanged.connect(lambda: self.update_attributes(2))
        self.comboBox_4search_label_3.currentTextChanged.connect(lambda: self.update_attributes(3))

        self.comboBox_4search_depth_3.currentTextChanged.connect(self.ruled_based_design_changes)

        self.actionSave_DB.triggered.connect(self.neo_safe_database)
        self.actionSave_to_Excel.triggered.connect(self.save_model_to_excel)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        # Rechts Klick Menüs
        self.treeView_1.customContextMenuRequested.connect(self.menuContextTree_1)
        self.treeView_2.customContextMenuRequested.connect(self.menuContextTree_2)
        self.treeView_3.customContextMenuRequested.connect(self.menuContextTree_3)
        self.treeView_4.customContextMenuRequested.connect(self.menuContextTree_4)
        self.treeView_1.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.treeView_2.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.treeView_3.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.treeView_4.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)


        self.setup_clear_db_template()

        self.refresh_menus()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "MainWindow"))
        self.comboBox_levels_1.setItemText(0, _translate("MainWindow", "1"))
        self.comboBox_levels_1.setItemText(1, _translate("MainWindow", "2"))
        self.comboBox_levels_1.setItemText(2, _translate("MainWindow", "3"))
        self.comboBox_levels_1.setItemText(3, _translate("MainWindow", "4"))
        self.comboBox_levels_1.setItemText(4, _translate("MainWindow", "5"))
        self.comboBox_levels_1.setItemText(5, _translate("MainWindow", "6"))
        self.comboBox_levels_1.setItemText(6, _translate("MainWindow", "7"))
        self.comboBox_levels_1.setItemText(7, _translate("MainWindow", "8"))
        self.comboBox_levels_1.setItemText(8, _translate("MainWindow", "9"))
        self.comboBox_levels_1.setItemText(9, _translate("MainWindow", "10"))
        self.pushButton_1_trace.setText(_translate("MainWindow", "Trace"))
        self.checkBox_respect_related_1.setText(_translate("MainWindow", "Respect related objects"))
        self.pushButton_tree_1_reload.setText(_translate("MainWindow", "Reload"))
        self.groupBox.setTitle(_translate("MainWindow", "Linking Tool"))
        self.pushButton_add_to_Scope.setText(_translate("MainWindow", "Add to Scope"))
        self.pushButton_save_change.setText(_translate("MainWindow", "Save"))
        self.checkBox_use_hooknode.setText(_translate("MainWindow", "Use HookNode"))
        self.pushButton_2.setText(_translate("MainWindow", "Remove from scope"))
        self.comboBox_4search_label_1.setItemText(0, _translate("MainWindow", "*"))
        self.comboBox_4search_label_2.setItemText(0, _translate("MainWindow", "*"))
        self.comboBox_4search_label_3.setItemText(0, _translate("MainWindow", "*"))
        self.comboBox_4search_direction_2.setItemText(0, _translate("MainWindow", "-->"))
        self.comboBox_4search_direction_2.setItemText(1, _translate("MainWindow", "<--"))
        self.comboBox_4search_direction_2.setItemText(2, _translate("MainWindow", "---"))
        self.comboBox_4search_direction_3.setItemText(0, _translate("MainWindow", "-->"))
        self.comboBox_4search_direction_3.setItemText(1, _translate("MainWindow", "<--"))
        self.comboBox_4search_direction_3.setItemText(2, _translate("MainWindow", "---"))
        self.pushButton_tree_3_reload.setText(_translate("MainWindow", "PushButton"))
        self.checkBox_respect_HookNode.setText(_translate("MainWindow", "Respect HookNode"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_1), _translate("MainWindow", "Explorer"))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), _translate("MainWindow", "Querry"))
        self.menuFile.setTitle(_translate("MainWindow", "File"))
        self.menuCreate.setTitle(_translate("MainWindow", "Create"))
        self.actionSave_DB.setText(_translate("MainWindow", "Save DB"))
        self.actionSave_to_Excel.setText(_translate("MainWindow", "Save to Excel"))

    def refresh_menus(self):
        self.comboBox_expand_1.clear()
        self.comboBox_4search_depth_2.clear()
        self.comboBox_4search_depth_3.clear()

        self.comboBox_expand_1.addItems(['1', '2', '3', '4', '5', '6', '7', '8'])
        self.comboBox_4search_depth_2.addItems(['1', '2', '3', '4', '5', '6', '7', '8'])
        self.comboBox_4search_depth_3.addItems(['0', 'shortest', '1', '2', '3', '4', '5', '6', '7', '8'])

        self.neo_set_all_query_end_nodes()
        self.neo_set_all_hook_nodes()
        self.neo_set_all_scope_nodes()

        self.comboBox_levels_1.setCurrentText("1")
        self.comboBox_expand_1.setCurrentText("3")
        self.checkBox_respect_HookNode.setCheckState(QtCore.Qt.Checked)
        print("Alle Labels gesetzt")
        return True

    def menuContextTree_1(self, point):
        # Infos about the node selected.
        index = self.treeView_1.indexAt(point)
        if not index.isValid():
            return
        qitem_index = index
        qmodel = qitem_index.model()
        while type(qmodel) == QSortFilterProxyModel:
            print("qmodel == QSortFilterProxyModel")
            qitem_index = qmodel.mapToSource(qitem_index)
            qmodel = qitem_index.model()

        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
        item_type = qmodel.itemFromIndex(qitem_index.siblingAtColumn(1)).text()
        qitem = qmodel.itemFromIndex(qitem_index)

        self.show_details_of_node(qitem_index)

        self.load_deeplink.triggered.connect(lambda: self.load_deep_link(index))
        self.searchfornextartefact.triggered.connect(lambda: self.search_for_next_artefact(index))
        self.test_querie_action.triggered.connect(self.querie_test)

        # We build the menu.
        menu = QtWidgets.QMenu()
        menu.addAction(self.load_deeplink)
        menu.addSeparator()
        menu.addAction(self.searchfornextartefact)
        menu.addAction(self.test_querie_action)
        menu.addAction("This Level Status Querry")
        if qitem.checkState() == QtCore.Qt.Checked:
            add_item_menu = menu.addMenu("Add node as child")
            for x in self.link_data_model.index:
                if self.link_data_model.at[x, item_type]:
                    add_item_menu.addAction(x)
            menu.addAction("Remove link to parent")
            menu.addAction("Set as query start node")
            if item_type == "Folder":
                menu.addAction("Sync Folder")
            menu.addSeparator()
            if (item_type == "Epic") or (item_type == "Feature") or (item_type == "User_Story"):
                menu.addAction("Calculate bottom up dates")
                menu.addAction("Plot Gantt Chart")
                menu.addSeparator()
        menu.addAction("Set all file Icons")
        menu.addAction("Switch tree side")
        menu.addSeparator()
        menu.addAction("Andreas Rieping")

        ret = menu.exec_(self.treeView_1.mapToGlobal(point))
        if ret:
            if ret.text() in self.link_data_model.index:
                print("Create new Item")
                self.create_new_node(dbid, qitem, ret.text())
                #qitem.setCheckState(QtCore.Qt.Unchecked)
            elif ret.text() == "Remove link to parent":
                self.delete_link_to_parent(qitem, dbid)
            elif ret.text() == "Sync Folder":
                self.sync_folder_with_os(dbid)
                self.show_neighbor(qitem_index)
            elif ret.text() == "Set all file Icons":
                self.set_all_icons(qitem)
            elif ret.text() == "This Level Status Querry":
                self.first_level_querie(qitem)
            elif ret.text() == "Switch tree side":
                self.switch_tree_side()
            elif ret.text() == "Set as query start node":
                self.set_first_path_query_node(qitem)
            elif ret.text() == "Calculate bottom up dates":
                self.calculate_bottom_up_dates(qitem)
            elif ret.text() == "Plot Gantt Chart":
                self.plot_gantt_chart(qitem)

        self.load_deeplink.disconnect()
        self.searchfornextartefact.disconnect()
        self.test_querie_action.disconnect()

    def menuContextTree_3(self, point):
        # Infos about the node selected.
        index = self.treeView_3.indexAt(point)
        if not index.isValid():
            return
        qitem_index = index
        qmodel = qitem_index.model()
        while type(qmodel) == QSortFilterProxyModel:
            print("qmodel == QSortFilterProxyModel")
            qitem_index = qmodel.mapToSource(qitem_index)
            qmodel = qitem_index.model()

        qmodel = qitem_index.model()
        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
        item_type = qmodel.itemFromIndex(qitem_index.siblingAtColumn(1)).text()
        qitem = qmodel.itemFromIndex(qitem_index)

        self.show_details_of_node(qitem_index)

        self.load_deeplink.triggered.connect(lambda: self.load_deep_link(index))
        self.searchfornextartefact.triggered.connect(lambda: self.search_for_next_artefact(index))
        self.test_querie_action.triggered.connect(self.querie_test)

        # We build the menu.
        menu = QtWidgets.QMenu()
        menu.addAction(self.load_deeplink)
        menu.addSeparator()
        menu.addAction(self.searchfornextartefact)
        menu.addAction(self.test_querie_action)
        if qitem.checkState() == QtCore.Qt.Checked:
            # add_item_menu = menu.addMenu("Add node as child")
            # for x in self.link_data_model.index:
            #     if self.link_data_model.at[x, item_type]:
            #         add_item_menu.addAction(x)
            menu.addSeparator()
            # menu.addAction("Add Tree and dublicate nodes")
            menu.addAction("Add Tree and reference nodes")
            menu.addAction("Add Node as child")
            menu.addSeparator()
            menu.addAction("Set as query start node")
            if item_type == "HookNode":
                menu.addAction("Select in work4traces")
            # menu.addAction("Add Node as reference")
        menu.addSeparator()
        menu.addAction("Switch tree side")
        menu.addAction("Andreas Rieping")

        ret = menu.exec_(self.treeView_3.mapToGlobal(point))
        if ret:
            if ret.text() in self.link_data_model.index:
                print("Create new Item")
                self.create_new_node(dbid, qitem, ret.text())
                # qitem.setCheckState(QtCore.Qt.Unchecked)
            elif ret.text() == "Remove link to parent":
                self.delete_link_to_parent(qitem, dbid)
            elif ret.text() == "Add Tree and dublicate nodes":
                self.dublicate_tree(qitem)
            elif ret.text() == "Add Tree and reference nodes":
                self.add_tree(qitem)
            elif ret.text() == "Add Node as child":
                self.link_checked_items(dbid, self.model_tree_1)
            elif ret.text() == "Add Node as reference":
                self.relate_checked_items(dbid, self.model_tree_1)
            elif ret.text() == "Switch tree side":
                self.switch_tree_side()
            elif ret.text() == "Set as query start node":
                self.set_first_path_query_node(qitem)
            elif ret.text() == "Select in work4traces":
                self.select_in_work4traces(dbid)
        self.load_deeplink.disconnect()
        self.searchfornextartefact.disconnect()
        self.test_querie_action.disconnect()

    def menuContextTree_2(self, point):
        # Infos about the node selected.
        index = self.treeView_2.indexAt(point)
        if not index.isValid():
            return
        model = index.model()
        if type(model) == QSortFilterProxyModel:
            qitem_index = model.mapToSource(index)
        else:
            qitem_index = index

        qmodel = qitem_index.model()
        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
        item_type = qmodel.itemFromIndex(qitem_index.siblingAtColumn(1)).text()
        qitem = qmodel.itemFromIndex(qitem_index)

        self.load_deeplink.triggered.connect(lambda: self.load_deep_link(index))

        # We build the menu.
        menu = QtWidgets.QMenu()
        menu.addAction(self.load_deeplink)
        menu.addSeparator()
        if qitem.checkState() == QtCore.Qt.Checked:
            menu.addAction("Add Node as child")
            menu.addAction("Add Node as reference")
        menu.addAction("Andreas Rieping")

        ret = menu.exec_(self.treeView_2.mapToGlobal(point))
        if ret:
            if ret.text() == "Add Node as child":
                self.link_checked_items(dbid, self.model_tree_1)
            elif ret.text() == "Add Node as reference":
                self.relate_checked_items(dbid, self.model_tree_1)
        self.load_deeplink.disconnect()

    def menuContextTree_4(self, point):
        # Infos about the node selected.
        index = self.treeView_4.indexAt(point)
        if not index.isValid():
            return
        model = index.model()
        if type(model) == QSortFilterProxyModel:
            qitem_index = model.mapToSource(index)
        else:
            qitem_index = index

        qmodel = qitem_index.model()
        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
        item_type = qmodel.itemFromIndex(qitem_index.siblingAtColumn(1)).text()
        qitem = qmodel.itemFromIndex(qitem_index)

        self.load_deeplink.triggered.connect(lambda: self.load_deep_link(index))

        # We build the menu.
        menu = QtWidgets.QMenu()
        menu.addAction(self.load_deeplink)
        menu.addSeparator()
        if qitem.checkState() == QtCore.Qt.Checked:
            menu.addAction("Add Node as child")
            menu.addAction("Add Node as reference")
        menu.addAction("Andreas Rieping")

        ret = menu.exec_(self.treeView_4.mapToGlobal(point))
        if ret:
            if ret.text() == "Add Node as child":
                self.link_checked_items(dbid, self.model_tree_1)
            elif ret.text() == "Add Node as reference":
                self.relate_checked_items(dbid, self.model_tree_1)
        self.load_deeplink.disconnect()

    def ruled_based_design_changes(self):
        search_depth = self.comboBox_4search_depth_3.currentText()
        if search_depth == "shortest":
            self.comboBox_4search_depth_2.setCurrentText("6")

    def set_all_icons(self, qitem):
        def loop_over_tree(qitem):
            qmodel = qitem.model()
            if qitem.hasChildren():
                for ii in range(0, qitem.rowCount()):
                    child = qitem.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    node_type = qmodel.itemFromIndex(child.index().siblingAtColumn(1)).text()
                    if node_type == "File":
                        change_icon = False
                        if cid in self.databyid:
                            dicti = self.databyid[cid]
                        else:
                            dicti = self.get_node_details_by_id(cid)
                            self.databyid.update({cid: dicti})

                        if "system_icon" in dicti.keys():
                            icon = dicti["system_icon"]
                            if icon:
                                change_icon = True
                        if ("DeepLink" in dicti.keys()) and not change_icon:
                            path = dicti["DeepLink"]
                            if "http" in path:
                                path = path.split("?")[0]
                            (root, extension) = os.path.splitext(path)
                            if extension in self.icon_dict.keys():
                                icon = self.icon_dict[extension]
                                change_icon = True
                            else:
                                icon_info = getDefaultIcon(path)

                                if icon_info:
                                    path = icon_info.split(",")[0]
                                    fileInfo = QtCore.QFileInfo(path)
                                    iconProvider = QtWidgets.QFileIconProvider()
                                    icon = iconProvider.icon(fileInfo)
                                    dicti.update({"system_icon": icon})
                                    self.databyid.update({cid: dicti})
                                    self.icon_dict.update({extension: icon})
                                    change_icon = True
                        if change_icon:
                            try:
                                child.setIcon(icon)
                            except:
                                print("Fehler bei Icon zuweisung")
                    loop_over_tree(child)
        loop_over_tree(qitem)

    def switch_tree_side(self):
        tmp_model_1 = self.model_tree_1
        tmp_model_3 = self.model_tree_3

        current_index_1 = self.treeView_1.currentIndex()
        current_index_3 = self.treeView_3.currentIndex()
        current_item_1 = self.model_tree_1.itemFromIndex(current_index_1)
        current_item_3 = self.model_tree_3.itemFromIndex(current_index_3)

        self.treeView_1_selected.clear()
        self.treeView_3_selected.clear()

        self.model_tree_3 = tmp_model_1
        self.model_tree_1 = tmp_model_3

        self.treeView_1.setModel(self.model_tree_1)
        self.treeView_3.setModel(self.model_tree_3)

        current_hook_1 = self.comboBox_Hook_Node_1.currentIndex()
        # current_hook_3 = self.comboBox_Hook_Node_3.currentIndex()
        # self.comboBox_Hook_Node_1.setCurrentIndex(current_hook_3)
        # self.comboBox_Hook_Node_3.setCurrentIndex(current_hook_1)

        if current_item_1:
            current_item_1.setCheckState(QtCore.Qt.Unchecked)
            self.expant_until_root_node(self.treeView_3, current_item_1)
            self.treeView_3.setCurrentIndex(current_index_1)
        if current_item_3:
            current_item_3.setCheckState(QtCore.Qt.Unchecked)
            self.expant_until_root_node(self.treeView_1, current_item_3)
            self.treeView_1.setCurrentIndex(current_index_3)

    def importData(self, data, in_QStandardModel, header):
        print("Start Import Data")
        def create_items_dict(in_list: list):
            output = {}
            for item in in_list:
                item_id = item['_id']
                if not item['child_of']:
                    item['child_of'] = 0
                if item_id not in output:
                    output.update({item_id: [item]})
                else:
                    results = output[item_id]
                    if item not in results:
                        results.append(item)
                    output.update({item_id: results})
            return output

        def create_parent_dict(in_list: list):
            output = {}
            for item in in_list:
                child_of = item['child_of']
                if not child_of:
                    child_of = 0
                if child_of not in output:
                    output.update({child_of: [item['_id']]})
                else:
                    results = output[child_of]
                    if item['_id'] not in results:
                        results.append(item['_id'])
                    output.update({child_of: results})
            return output

        def print_children(id: int, parents: dict, items: dict, seen: dict, parents_chain: list, header: list):
            all_parents = []
            all_parents.extend(parents_chain)
            for children_id in parents[id]:
                for value in items[children_id]:
                    parent = seen[id]
                    dbid = value['_id']
                    pid = value['child_of']
                    if pid == id:
                        parent.appendRow(self.fill_value_table(value, header, editable_bool))
                        seen[dbid] = parent.child(parent.rowCount() - 1)
                        # print(value)
                        all_parents.append(pid)
                if children_id in parents:
                    if children_id not in all_parents:
                        print_children(children_id, parents, items, seen, all_parents, header)

        if not header:
            header.extend(self.headers[1])

        seen = {}
        seen.update({0: in_QStandardModel.invisibleRootItem()})

        editable_bool = False
        parents_in = create_parent_dict(data)
        print("parents_in createn")
        items_in = create_items_dict(data)
        print("items_in createn")
        print_children(0, parents_in, items_in, seen, [0], header)
        self.set_all_icons(in_QStandardModel.invisibleRootItem())
        self.model_tree_3.setHorizontalHeaderLabels(header)

        return True

    def load_deep_link(self, index):
        # Suchfilter berücksichtigen
        office_docs = [".xlsx", ".xltx", ".pptx", ".potx", ".ppsx", ".docx", ".dotx", ]
        qitem_index = index
        qmodel = qitem_index.model()
        while type(qmodel) == QSortFilterProxyModel:
            # print("qmodel == QSortFilterProxyModel")
            qitem_index = qmodel.mapToSource(qitem_index)
            qmodel = qitem_index.model()
        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())

        if dbid in self.databyid:
            dicti = self.databyid[dbid]
        else:
            dicti = self.get_node_details_by_id(dbid)
            self.databyid.update({dbid: dicti})
        if "DeepLink" in dicti.keys():
            file_path = dicti["DeepLink"]
            if len(file_path) > 0:

                if ("//collaboration.claas.com" in file_path) or ("//sharepoint.claas.com" in file_path):
                    try:
                        for doc_type in office_docs:
                            office_type = False
                            if doc_type in file_path:
                                print("1 try to open: ", "https:" + file_path.replace(" ", "%20") + "?web=1")
                                os.startfile("https:" + file_path.replace(" ", "%20") + "?web=1")
                                office_type = True
                                break

                        if not office_type:
                            print("11 try to open: ", "https:" + file_path.replace(" ", "%20"))
                            os.startfile("https:" + file_path.replace(" ", "%20"))
                    except:
                        try:
                            print("3 try to open: ", file_path)
                            os.startfile(file_path)
                        except:
                            print("4 try to open: ", file_path, " ging nicht")
                else:
                    try:
                        print("6 try to open: ", file_path)
                        os.startfile(file_path.replace("/", "\\"))
                    except:
                        try:
                            print("7 try to open: ", file_path)
                            os.startfile(file_path)
                        except:
                            print("öffnen nicht möglich")

    def expand_tree_node(self, data, qitem_index, header):
        print("Start expand_tree_node")
        if not header:
            header = self.headers[1]

        def create_items_dict(in_list: list, id: int):
            output = {}
            for item in in_list:
                item_id = item['_id']
                if not item['child_of']:
                    item['child_of'] = id
                if item_id not in output:
                    output.update({item_id: [item]})
                else:
                    results = output[item_id]
                    if item not in results:
                        results.append(item)
                    output.update({item_id: results})
            return output

        def create_parent_dict(in_list: list):
            output = {}
            for item in in_list:
                child_of = item['child_of']
                if not child_of:
                    continue
                if child_of not in output:
                    output.update({child_of: [item['_id']]})
                else:
                    results = output[child_of]
                    if item['_id'] not in results:
                        results.append(item['_id'])
                    output.update({child_of: results})
            return output

        def print_children(id: int, parents: dict, items: dict, seen: dict, parents_chain: list, header: list):
            all_parents = []
            all_parents.extend(parents_chain)
            if id in parents.keys():
                for children_id in parents[id]:
                    for value in items[children_id]:
                        parent = seen[id]
                        dbid = value['_id']
                        pid = value['child_of']
                        if pid == id:
                            parent.appendRow(self.fill_value_table(value, header, editable_bool))
                            seen[dbid] = parent.child(parent.rowCount() - 1)
                            # print(value)
                            all_parents.append(pid)
                    if children_id in parents:
                        if children_id not in all_parents:
                            print_children(children_id, parents, items, seen, all_parents, header)

        seen = {}
        qmodel = qitem_index.model()
        while type(qmodel) == QSortFilterProxyModel:
            print("qmodel == QSortFilterProxyModel")
            qitem_index = qmodel.mapToSource(qitem_index)
            qmodel = qitem_index.model()
        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
        active_item = qmodel.itemFromIndex(qitem_index.siblingAtColumn(0))
        if active_item.hasChildren():
            for ii in range(0, active_item.rowCount()).__reversed__():
                active_item.removeRow(ii)

        if not active_item.hasChildren():
            seen.update({dbid: active_item})
            editable_bool = False

            parents_in = create_parent_dict(data)
            print("parents_in createn")
            items_in = create_items_dict(data, dbid)
            print("items_in createn")
            print_children(dbid, parents_in, items_in, seen, [dbid], header)

            self.set_all_icons(active_item)

        if qmodel == self.model_tree_1:
            self.treeView_1.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_1.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_1.expand(qitem_index)
        if qmodel == self.model_tree_3:
            self.treeView_3.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_3.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_3.expand(qitem_index)

        return True

    def expand_tree_node_printer(self, data, qitem_index):
        seen = {}
        qmodel = qitem_index.model()
        while type(qmodel) == QSortFilterProxyModel:
            print("qmodel == QSortFilterProxyModel")
            qitem_index = qmodel.mapToSource(qitem_index)
            qmodel = qitem_index.model()
        dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
        active_item = qmodel.itemFromIndex(qitem_index.siblingAtColumn(0))
        if active_item.hasChildren():
            for ii in range(0, active_item.rowCount()).__reversed__():
                active_item.removeRow(ii)

        if not active_item.hasChildren():
            seen.update({dbid: active_item})
            editable_bool = False

            parents_in = create_parent_dict(data)
            print("parents_in createn")
            items_in = create_items_dict(data, dbid)
            print("items_in createn")
            print_children(dbid, parents_in, items_in, seen, [dbid], header)

            self.set_all_icons(active_item)

        if qmodel == self.model_tree_1:
            self.treeView_1.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_1.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_1.expand(qitem_index)
        if qmodel == self.model_tree_3:
            self.treeView_3.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_3.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_3.expand(qitem_index)

        return True

    def fill_value_table(self, value: dict, header: list, editable_bool: bool):
        new_Row = []
        color = QColor(0, 0, 0)
        if value:
            if value["_type"] == "Folder" or value["_type"] == "File":
                if value["Sync"] == "ERROR":
                    color = QColor(255, 0, 0)
            for ii, head in enumerate(header):
                if head == 'NeoID':
                    head = '_id'
                elif head == 'Type':
                    head = '_type'
                elif head == 'Parent Link':
                    head = 'parent_link'

                if head in value:
                    tmp = str(value[head])
                    txt = tmp
                    # if ii <= 5:
                    #     txt = textwrap.fill(txt, 50)
                    if head == 'Title':
                        new_Row.append(StandardItem(txt, value['_type'], False, True, self.tree_font_size, True, color))
                    else:
                        new_Row.append(StandardItem(txt, value['_type'], False, False, self.tree_font_size, False, color))
                else:
                    new_Row.append(StandardItem('', value['_type'], False, False, self.tree_font_size, False, color))
            return new_Row

    def setup_clear_db_template(self):
        arti_lst = ["Scope", "HookNode", "Folder", "OS_Folder", "Req_Spec", "Chapter", "Requirement", "CPEM_Project",
                    "Part", "File", "Team", "Person", "UserStory", "Information", "Department", "Role", "Company",
                    "Project", "Phase", "Epic", "Feature", "User_Story", "Task", "Bug", "Comment", "TestCase",
                    "TestStep", "TestExecution", "TestResult", "Domain_Group", "Domain", "Product", "Artifact",
                    "Attribute", "Physic", "Vendor", "Portfolio", "Software"]
        qry_cmd_Lst = []
        inti = 0
        for arti in arti_lst:
            qry_cmd_Lst.append('MERGE (a' + str(inti) + ' :' + arti + ' {Title: \"new ' + arti + '\", Details: \"\"})')
            inti += 1

        qry = ' '.join(qry_cmd_Lst)
        with self.driver.session() as session:
            records = session.run(qry)
        session.close()

    def neo_safe_database(self):
        current_time = str(time.time()).replace('.', '_')
        with self.driver.session() as session:
            session.run('CALL apoc.export.cypher.all("all-plain' + current_time + '.cypher", {format: "plain"})')
            session.run('CALL apoc.export.cypher.all("C:/Arbeitsbereich/' + current_time + '.cypher", {format: "plain"})')
        session.close()

    def update_attributes(self, combo_nr: int):
        if combo_nr == 1:
            item_label = self.comboBox_4search_label_1.currentText()
        elif combo_nr == 2:
            item_label = self.comboBox_4search_label_2.currentText()
        elif combo_nr == 3:
            item_label = self.comboBox_4search_label_3.currentText()

        if not item_label:
            return False

        if not item_label == '*':
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (p:' + item_label + ') WITH DISTINCT keys(p) AS keys')
            qry_cmd_Lst.append('UNWIND keys AS keyslisting WITH DISTINCT keyslisting AS allfields')
            qry_cmd_Lst.append('RETURN allfields;')
            qry = ' '.join(qry_cmd_Lst)
            print(qry)

            ret_val = []
            ret_val.append("*")
            with self.driver.session() as session:
                records = session.run(qry)
                print("neo_hook_scope_path_query QUERY DONE")

                for node in records:
                    print(node[0])
                    ret_val.append(node[0])
            session.close()
            ret_val = sorted(ret_val, key=str.casefold)

            if combo_nr == 1:
                self.comboBox_4search_property_1.clear()
                self.comboBox_4search_property_1.addItems(ret_val)
                self.comboBox_4search_property_1.setCurrentText("Title")
            elif combo_nr == 2:
                self.comboBox_4search_property_2.clear()
                self.comboBox_4search_property_2.addItems(ret_val)
                self.comboBox_4search_property_2.setCurrentText("Title")
            elif combo_nr == 3:
                self.comboBox_4search_property_3.clear()
                self.comboBox_4search_property_3.addItems(ret_val)
                self.comboBox_4search_property_3.setCurrentText("Title")
        else:
            ret_val = ["*"]
            if combo_nr == 1:
                self.comboBox_4search_property_1.clear()
                self.comboBox_4search_property_1.addItems(ret_val)
            elif combo_nr == 2:
                self.comboBox_4search_property_2.clear()
                self.comboBox_4search_property_2.addItems(ret_val)
            elif combo_nr == 3:
                self.comboBox_4search_property_3.clear()
                self.comboBox_4search_property_3.addItems(ret_val)
        return True

    def neo_set_all_query_end_nodes(self):
        with self.driver.session() as session:
            Label_lst = []
            Label_lst.append("*")
            qry1 = ('MATCH (n) RETURN distinct labels(n) ')

            records = session.run(qry1)

            for record in records:
                print(record['labels(n)'])
                for tmp in record['labels(n)']:
                    Label_lst.append(tmp)
            for Lable_tmp in Label_lst:
                print(Lable_tmp)
            Label_lst = sorted(Label_lst, key=str.casefold)

            self.comboBox_tree_1_Endnode.clear()
            self.comboBox_tree_1_Endnode.addItems(Label_lst)

            self.comboBox_4search_label_1.clear()
            self.comboBox_4search_label_2.clear()
            self.comboBox_4search_label_3.clear()

            self.comboBox_4search_label_2.addItems(Label_lst)
            self.comboBox_4search_label_3.addItems(Label_lst)
            Label_lst.insert(0, "HookNode")
            self.comboBox_4search_label_1.addItems(Label_lst)

        session.close()

    def neo_set_all_hook_nodes(self):
        with self.driver.session() as session:
            hook_lst = []
            qry1 = ('MATCH (x:HookNode) '
                    'return x ')

            records = session.run(qry1)

            for record in records:
                print(record[0]['Title'])
                hook_lst.append(record[0]['Title'])
        session.close()
        hook_lst = sorted(hook_lst, key=str.casefold)
        self.comboBox_Hook_Node_1.clear()
        self.comboBox_Hook_Node_1.addItems(hook_lst)

    def neo_set_all_scope_nodes(self):
        hook_lst = []
        hook_lst2 = []
        hook_lst.append("System")
        qry1 = ('MATCH (x:Scope) '
                'return x ')
        with self.driver.session() as session:
            records = session.run(qry1)

            for record in records:
                hook_lst.append(record[0]['Title'])
                hook_lst2.append(record[0]['Title'])
        session.close()

        self.comboBox_applay_Scope_1.clear()
        self.comboBox_multi_1.clear()
        self.comboBox_applay_Scope_1.addItems(hook_lst)

        for i, txt in enumerate(hook_lst):
            # adding item
            self.comboBox_multi_1.addItem(txt)
            item1 = self.comboBox_multi_1.model().item(i, 0)
            item1.setCheckState(QtCore.Qt.Unchecked)

    def neo_hook_scope_path_query(self, HookNode_in: str, Scope_in: list, End_Node: str, level_depth: str, resp_rel: bool):
        qry_cmd_Lst = []

        qry_cmd_Lst.append('MATCH (x:HookNode) WHERE x.Title = \"' + HookNode_in + '\"')
        if Scope_in:
            qry_cmd_Lst.append('MATCH (b:Scope) WHERE')
            for scope in Scope_in:
                qry_cmd_Lst.append('b.Title= \"' + scope + '\"')
                qry_cmd_Lst.append('OR')
            del qry_cmd_Lst[-1]
            qry_cmd_Lst.append('MATCH tmp = (y)-[]-(b)')
        if resp_rel:
            if End_Node == "*":
                qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + level_depth + ']->(y)-[ee]-(z)')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + level_depth + ']->(y:' + End_Node + ')-[ee]-(z)')
        else:
            if End_Node == "*":
                qry_cmd_Lst.append('MATCH path = (x)-[*0..' + level_depth + ']->(y)')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[*0..' + level_depth + ']->(y:' + End_Node + ')')
        qry_cmd_Lst.append('WITH path, relationships(path) AS rr,x ,y')
        qry_cmd_Lst.append('WHERE ALL(r in rr')
        qry_cmd_Lst.append('WHERE r.HookNode = \"' + HookNode_in + '\"')
        # qry_cmd_Lst.append('AND type(r) = "has"')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        if resp_rel:
            qry_cmd_Lst.append('AND type(ee) <> "has"')
        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)

        ret_val = []
        with self.driver.session() as session:
            records = session.run(qry)
            print("neo_hook_scope_path_query QUERY DONE")

            for node in records:
                ret_val = node[0]
                break
            print(ret_val)
            session.close()
        return ret_val

    def printer(self, json_str, in_qitem, header):
        def print_all_nodes(obj, arr, current_Parent, link_type, parent_qstandardmodelitem):
            """Recursively search for values of key in JSON tree."""
            if isinstance(obj, dict):
                temp_dict = {}
                all_childs = []
                new_parent = obj.get('_id', '')
                for k, v in obj.items():
                    if isinstance(v, (dict, list)):
                        parent_link = k

                        tmp_child = {}
                        tmp_child.update({"child_lst": v.copy()})
                        tmp_child.update({"new_parent": new_parent})
                        tmp_child.update({"parent_link": parent_link})
                        all_childs.append(tmp_child)
                    else:
                        temp_dict.update(({'child_of': current_Parent}))
                        temp_dict.update(({'parent_link': link_type}))
                        if "HookNode" in k:
                            temp_dict.update(({'hooknode': v}))
                        temp_dict.update({k: v})
                arr.append(temp_dict)

                try:
                    qmodel = parent_qstandardmodelitem.model()
                    nid = qmodel.itemFromIndex(parent_qstandardmodelitem.index().siblingAtColumn(2)).text()
                    # print("try:", temp_dict)
                    # print("nid: ", nid)
                except:
                    nid = ""
                    # print("except: ", temp_dict)
                # Die Querry startet bei einem Knoten, der nicht geprintet werden soll. Daher muss der Start Knoten vom Print ausgeschlossen werden
                if not str(nid) == str(temp_dict["_id"]):
                    # print("Ausnahme")
                    parent_qstandardmodelitem.appendRow(self.fill_value_table(temp_dict, header, False))
                    parent_qstandardmodelitem = parent_qstandardmodelitem.child(parent_qstandardmodelitem.rowCount() - 1)


                for child in all_childs:
                    print_all_nodes(child["child_lst"], arr, child["new_parent"], child["parent_link"], parent_qstandardmodelitem)
            elif isinstance(obj, list):
                for item in obj:
                    print_all_nodes(item, arr, current_Parent, link_type, parent_qstandardmodelitem)
            return arr

        qmodel = in_qitem.model()
        qitem_index = in_qitem.index()
        if not in_qitem == qmodel.invisibleRootItem():
            while type(qmodel) == QSortFilterProxyModel:
                # print("qmodel == QSortFilterProxyModel")
                qitem_index = qmodel.mapToSource(qitem_index)
                qmodel = qitem_index.model()
            dbid = int(qmodel.itemFromIndex(qitem_index.siblingAtColumn(2)).text())
            active_item = qmodel.itemFromIndex(qitem_index.siblingAtColumn(0))
            in_qitem = active_item
            print("Start acitve Item:", dbid)
            if active_item.hasChildren():
                for ii in range(0, active_item.rowCount()).__reversed__():
                    active_item.removeRow(ii)
            # active_item = active_item.parent()
            # in_qitem = active_item
            # if active_item.hasChildren():
            #     for ii in range(0, active_item.rowCount()).__reversed__():
            #         print("Hallo", active_item.child(ii, 2).text())
            #         if active_item.child(ii, 2).text() == str(dbid):
            #             active_item.removeRow(ii)

        print("Tree Printer starts")
        arr = []
        current_Parent = ''
        link_type = ''
        print_all_nodes(json_str, arr, current_Parent, link_type, in_qitem)
        print("Tree Printer ends")
        self.set_all_icons(in_qitem)
        print("Set Icons")

    def get_node_details_by_id(self, node_id_in: int):
        def clear_dict(input_dict: dict):
            for k, v in input_dict.items():
                if isinstance(v, str):
                    input_dict.update({k: re.sub(r'[\000-\010]|[\013-\014]|[\016-\037]', '', v)})
            return input_dict
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(node_id_in))
        qry_cmd_Lst.append('WITH {_label:labels(a),_id:id(a),properties: properties(a)} as x ')
        qry_cmd_Lst.append('RETURN x')
        qry = ' '.join(qry_cmd_Lst)

        dicti = {}
        with self.driver.session() as session:
            records = session.run(qry)

            for record in records:
                dicti.update({'_id': int(record["x"]['_id'])})
                dicti.update({'_type': record["x"]['_label'][0]})
                # x.decode('utf-8', 'ignore').encode('utf-8')

                prop_dict = {}
                prop_dict.update(record["x"]['properties'])
                prop_dict = clear_dict(prop_dict)
                dicti.update(prop_dict)

                # Alternative für die Zukunft
                # dicti.update({'_id': int(record["x"].id)})
                # dicti.update({'_type': next(iter(record["x"].labels))})
                # dicti.update(dict(record["x"].items()))
        session.close()

        data = [dicti]
        data_by_id(data, self.databyid)
        return dicti

    def neo_create_link(self, From_Detail_in, To_Detail_in, current_hooknode):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(From_Detail_in))        # Da ein Int übergeben wird, hier anders
        qry_cmd_Lst.append('MATCH (b) WHERE ID(b) = ' + str(To_Detail_in))
        # qry_cmd_Lst.append('MATCH (c:Scope) WHERE c.Title = \"' + current_scope + '\"')
        qry_cmd_Lst.append('MERGE (a)-[l:has {HookNode: \"' + current_hooknode + '\"}]->(b)')
        # qry_cmd_Lst.append('MERGE (b)-[:has {Title: \"' + current_scope + '\"}]->(c)')
        qry_cmd_Lst.append('RETURN ID(l)')
        qry = ' '.join(qry_cmd_Lst)

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(From_Detail_in))
        qry_cmd_Lst.append('MATCH (b) WHERE ID(b) = ' + str(To_Detail_in))
        qry_cmd_Lst.append('MATCH path = (b)-[*0..10]->(a)')
        qry_cmd_Lst.append('WHERE ALL(r in relationships(path)')
        qry_cmd_Lst.append('WHERE r.HookNode = \"' + current_hooknode + '\"')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('RETURN count(path) as pathCount')
        conter_qry = ' '.join(qry_cmd_Lst)

        print(qry)
        with self.driver.session() as session:
            records = session.run(conter_qry)
            for record in records:
                num_paths = record['pathCount']

            if num_paths > 0:
                print('Loop erkannt. Link kann nicht erstellt werden')
                session.close()
                return False
            else:
                records = session.run(qry)
                print("create_link NEO4J CMD DONE", records)
                session.close()
                return True

    def neo_dublicate_node(self, From_Detail_in):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (f) WHERE ID(f) = ' + str(From_Detail_in))
        qry_cmd_Lst.append('CALL apoc.refactor.cloneNodes([f])')
        qry_cmd_Lst.append('YIELD input, output')
        qry_cmd_Lst.append('WITH output as n')
        qry_cmd_Lst.append('SET n.Key = ""')
        qry_cmd_Lst.append('RETURN ID(n) as new_id')
        qry = ' '.join(qry_cmd_Lst)

        print(qry)
        new_id = 0
        with self.driver.session() as session:
            records = session.run(qry)
            for record in records:
                return record["new_id"]
            session.close()
        return False

    def neo_create_related(self, From_Detail_in, To_Detail_in, current_hooknode):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(From_Detail_in))        # Da ein Int übergeben wird, hier anders
        qry_cmd_Lst.append('MATCH (b) WHERE ID(b) = ' + str(To_Detail_in))
        # qry_cmd_Lst.append('MATCH (c:Scope) WHERE c.Title = \"' + current_scope + '\"')
        qry_cmd_Lst.append('MERGE (a)-[l:related {HookNode: \"' + current_hooknode + '\"}]->(b)')
        # qry_cmd_Lst.append('MERGE (b)-[:has {Title: \"' + current_scope + '\"}]->(c)')
        qry_cmd_Lst.append('RETURN ID(l)')
        qry = ' '.join(qry_cmd_Lst)


        print(qry)
        with self.driver.session() as session:
            records = session.run(qry)
            print("neo_create_related NEO4J CMD DONE", records)
            session.close()
            return True

    def neo_create_node(self, to_node_id: int, new_node_label: str, current_hooknode:str, current_scope: str):
        qry_cmd_Lst = []

        common_attributes = ', ModifyDate: \"\", Details: \"\", Description: \"\", DeepLink: \"\", Key: apoc.create.uuid()'
        if not current_scope == 'System':
            qry_cmd_Lst.append('MATCH (c:Scope) WHERE c.Title = \"' + current_scope + '\"')
        qry_cmd_Lst.append('MATCH (b) WHERE ID(b) = ' + str(to_node_id))
        if new_node_label == "Folder":
            qry_cmd_Lst.append('CREATE (a:Folder {Title: \"New Folder\", Sync: \"False\"'
                               + common_attributes + '})')
        elif new_node_label == "File":
            qry_cmd_Lst.append('CREATE (a:File {Title: \"New File\", Sync: \"False\"'
                               + common_attributes + '})')
        elif new_node_label == "Team":
            qry_cmd_Lst.append('CREATE (a:Team {Title: \"New Team\"' + common_attributes + '})')
        elif new_node_label == "Chapter":
            qry_cmd_Lst.append('CREATE (a:Chapter {Title: \"New Chapter\"' + common_attributes + '})')
        elif new_node_label == "Part":
            qry_cmd_Lst.append('CREATE (a:Part {Title: \"New Part\"' + common_attributes + '})')
        elif new_node_label == "TestCase":
            qry_cmd_Lst.append('CREATE (a:TestCase {Title: \"New TestCase\"' + common_attributes + '})')
        elif new_node_label == "TestStep":
            qry_cmd_Lst.append('CREATE (a:TestStep {Title: \"New TestStep\"' + common_attributes + '})')
        elif new_node_label == "TestExecution":
            qry_cmd_Lst.append('CREATE (a:TestExecution {Title: \"New TestExecution\", Responsible: \"\"'
                               + common_attributes + '})')
        elif new_node_label == "TestResult":
            qry_cmd_Lst.append('CREATE (a:TestResult {Title: \"New TestExecution\", Status: \"\",'
                               ' Result: \"\"' + common_attributes + '})')
        elif new_node_label == "Req_Spec":
            qry_cmd_Lst.append('CREATE (a:Req_Spec {Title: \"New Req_Spec\"' + common_attributes + '})')
        elif new_node_label == "Requirement":
            qry_cmd_Lst.append('CREATE (a:Requirement {Title: \"New Requirement\"' + common_attributes + '})')
        elif new_node_label == "CPEM_Project":
            qry_cmd_Lst.append('CREATE (a:CPEM_Project {Title: \"New CPEM_Project\"' + common_attributes + '})')
        elif new_node_label == "Function":
            qry_cmd_Lst.append('CREATE (a:Function {Title: \"New Function\"' + common_attributes + '})')
        elif new_node_label == "OS_Folder":
            qry_cmd_Lst.append('CREATE (a:OS_Folder {Title: \"New Function\"' + common_attributes + '})')
        elif new_node_label == "Comment":
            qry_cmd_Lst.append('CREATE (a:Comment {Title: \"New Comment\"' + common_attributes + '})')
        elif new_node_label == "NewContent":
            qry_cmd_Lst.append('CREATE (a:NewContent {Title: \"NewContent\", NCRI: \"\", Change: \"\",'
                               ' Weighting_Factor: \"\"' + common_attributes + '})')
        elif new_node_label == "Task":
            qry_cmd_Lst.append('CREATE (a:Task {Title: \"New Task\", Effort: \"\", Responsible: \"\"'
                               + common_attributes + '})')
        elif new_node_label == "Project":
            qry_cmd_Lst.append('CREATE (a:Project {Title: \"New Project\"' + common_attributes + '})')
        elif new_node_label == "Phase":
            qry_cmd_Lst.append('CREATE (a:Phase {Title: \"New Phase\"' + common_attributes + '})')
        elif new_node_label == "Artifact":
            qry_cmd_Lst.append('CREATE (a:Artifact {Title: \"New Artifact\"' + common_attributes + '})')
        elif new_node_label == "Attribute":
            qry_cmd_Lst.append('CREATE (a:Attribute {Title: \"New Attribute\" , Mandatory: \"False\"'
                               + common_attributes + '})')
        else:
            qry_cmd_Lst.append('CREATE (a:' + new_node_label + ' {Title: \"New ' + new_node_label + '\"' + common_attributes + '})')

        qry_cmd_Lst.append('SET a.CreatedIn = \"' + current_hooknode + '\"')

        if current_scope == 'System':
            qry_cmd_Lst.append('MERGE (b)-[:has {HookNode: \"' + current_hooknode + '\"}]->(a)')
        else:
            qry_cmd_Lst.append('MERGE (a)-[:has {Scope: \"' + current_scope + '\"}]->(c)')
            qry_cmd_Lst.append('MERGE (b)-[:has {HookNode: \"' + current_hooknode + '\"}]->(a)')
        qry_cmd_Lst.append('RETURN a')
        qry = ' '.join(qry_cmd_Lst)

        print(qry)

        new_node_id = -1
        with self.driver.session() as session:
            records = session.run(qry)
            for record in records:
                print(record[0].id)
                new_node_id = record[0].id
            print("create_node NEO4J CMD DONE", records)
        session.close()
        return int(new_node_id)

    def neo_down_query_has(self, node_id: int, level_depth: int, use_hooknode: bool, HookNode_in: str, json_yn: bool):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + str(node_id))
        qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + str(level_depth) + ']->(y)')
        qry_cmd_Lst.append('WHERE ALL(r in rr')
        qry_cmd_Lst.append('WHERE type(r) <> ""')
        if use_hooknode:
            qry_cmd_Lst.append('AND r.HookNode = \"' + HookNode_in + '\"')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)

        if json_yn:
            ret_val = []
            with self.driver.session() as session:
                records = session.run(qry)
                print("neo_hook_scope_path_query QUERY DONE")
                for node in records:
                    ret_val = node[0]
                    break
                print(ret_val)
                session.close()
            return ret_val
        else:
            with self.driver.session() as session:
                query_rst = session.run(qry)
                records = []
                for iterator in query_rst:
                    records.append(iterator[0])
                for record in records:
                    data = extract_table(record)
                session.close()
            print("neo_down_query_has done")
            return data

    def neo_down_query_related(self, node_id: int):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + str(node_id))
        qry_cmd_Lst.append('MATCH path = (x)-[rr*0..1]->(y)')
        qry_cmd_Lst.append('WHERE ALL(r in rr')
        qry_cmd_Lst.append('WHERE type(r) <> ""')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)

        with self.driver.session() as session:
            query_rst = session.run(qry)
            records = []
            for iterator in query_rst:
                records.append(iterator[0])
            for record in records:
                data = extract_table(record)
            session.close()
        print("neo_down_query_related done")
        return data

    def neo_up_query_has(self, node_id: int, level_depth: int):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + str(node_id))
        qry_cmd_Lst.append('MATCH path = (x)<-[rr*0..' + str(level_depth) + ']-(y)')
        qry_cmd_Lst.append('WHERE ALL(r in rr')
        qry_cmd_Lst.append('WHERE type(r) <> ""')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)

        with self.driver.session() as session:
            query_rst = session.run(qry)
            records = []
            print(type(query_rst))
            for iterator in query_rst:
                records.append(iterator[0])
            for record in records:
                data = extract_table(record)
            session.close()
        return data

    def neo_add_to_scope(self, node_ids: list):
        current_scope = self.comboBox_applay_Scope_1.currentText()
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (b:Scope) WHERE b.Title = \"' + current_scope + '\"')
        for ii, node_id in enumerate(node_ids):
            qry_cmd_Lst.append('MATCH (a' + str(ii) + ') WHERE ID(a' + str(ii) + ') = ' + str(node_id))
        for ii, node_id in enumerate(node_ids):
            qry_cmd_Lst.append('MERGE (a' + str(ii) + ')-[:has {Title: \"' + current_scope + '\"}]->(b)')
        qry_cmd_Lst.append('RETURN b')
        qry = ' '.join(qry_cmd_Lst)

        print(qry)
        with self.driver.session() as session:
            records = session.run(qry)
            print("neo_add_to_scope NEO4J CMD DONE", records)
        session.close()
        return True

    def neo_delete_from_scope(self, node_ids: list):
        current_scope = self.comboBox_applay_Scope_1.currentText()
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (b:Scope) WHERE b.Title = \"' + current_scope + '\"')
        for ii, node_id in enumerate(node_ids):
            qry_cmd_Lst.append('MATCH (a' + str(ii) + ') WHERE ID(a' + str(ii) + ') = ' + str(node_id))
            qry_cmd_Lst.append('MATCH (a' + str(ii) + ')-[r' + str(ii) + ']->(b)')
        for ii, node_id in enumerate(node_ids):
            qry_cmd_Lst.append('DELETE r' + str(ii))
        qry = ' '.join(qry_cmd_Lst)

        print(qry)
        with self.driver.session() as session:
            records = session.run(qry)
            print("neo_delete_from_scope NEO4J CMD DONE", records)
        session.close()
        return True

    def neo_delete_link_by_node_ids(self, child_id: int, parent_id: int, link_type: str):
        current_hooknode = self.comboBox_Hook_Node_1.currentText()

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(child_id))
        qry_cmd_Lst.append('MATCH (b) WHERE ID(b) = ' + str(parent_id))
        qry_cmd_Lst.append('MATCH path = (b)-[r:' + link_type + ']->(a)')
        qry_cmd_Lst.append('WHERE r.HookNode = \"' + current_hooknode + '\"')
        qry_cmd_Lst.append('DELETE r')
        qry = ' '.join(qry_cmd_Lst)

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(child_id))
        qry_cmd_Lst.append('MATCH path = (a)-[]->()')
        qry_cmd_Lst.append('WHERE ALL(r in relationships(path)')
        qry_cmd_Lst.append('WHERE r.HookNode = \"' + current_hooknode + '\"')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('RETURN count(path) as pathCount')
        conter_qry_down = ' '.join(qry_cmd_Lst)

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(child_id))
        qry_cmd_Lst.append('MATCH path = (a)<-[]-()')
        qry_cmd_Lst.append('WHERE ALL(r in relationships(path)')
        qry_cmd_Lst.append('WHERE r.HookNode = \"' + current_hooknode + '\"')
        # qry_cmd_Lst.append('AND type(r) = "has"')
        qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('RETURN count(path) as pathCount')
        conter_qry_up_path = ' '.join(qry_cmd_Lst)

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(child_id))
        qry_cmd_Lst.append('MATCH path = (a)<-[]-()')
        # qry_cmd_Lst.append('WHERE ALL(r in relationships(path)')
        # qry_cmd_Lst.append('WHERE type(r) = "has"')
        # qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
        qry_cmd_Lst.append('RETURN count(path) as pathCount')
        conter_qry_up = ' '.join(qry_cmd_Lst)

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(child_id))
        qry_cmd_Lst.append('RETURN a.Sync as rst')
        sync_qry = ' '.join(qry_cmd_Lst)

        with self.driver.session() as session:
            records = session.run(sync_qry)
            for record in records:
                sync_state = record['rst']
            records = session.run(conter_qry_down)
            for record in records:
                num_paths_down = record['pathCount']
            records = session.run(conter_qry_up)
            for record in records:
                num_paths_up = record['pathCount']
            records = session.run(conter_qry_up_path)
            for record in records:
                num_paths_up_path = record['pathCount']

            print("num_paths_up: ", num_paths_up, "  num_paths_down: ", num_paths_down, "sync_state", sync_state)

            if sync_state == "ERROR":
                records = session.run(qry)
                session.close()
                return True
            if (num_paths_up == 1) and (num_paths_down == 0):
                delete_bool = delete_node_dialog()
                if delete_bool:
                    records = session.run(qry)
                    session.close()
                    return True
            elif (num_paths_up == 1) and (num_paths_down > 0):
                print('ERROR Object has children in same HookPath but not another parent')
                session.close()
                return False
            elif (num_paths_up > 1) and (num_paths_down == 0):
                records = session.run(qry)
                session.close()
                return True
                print("neo_delete_link_by_node_ids NEO4J CMD DONE", records)
            elif (num_paths_up_path > 1):
                records = session.run(qry)
                session.close()
                return True
                print("neo_delete_link_by_node_ids NEO4J CMD DONE", records)
            else:
                print("neo_delete_link_by_node_ids Nicht definierter zustand")
                return False

    def neo_edit_node(self, Value):
        HookNode = self.comboBox_Hook_Node_1.currentText()
        forbidden = ['_id', '_type', 'parent_link', 'child_of', 'system_icon', 'Key', 'current_Parent', 'hooknode']
        Value_keys = []
        Value_keys.extend(Value_keys)
        for k in Value_keys:
            if 'HookNode' in k:
                Value.pop(k)

        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(Value['_id']))  # Da ein Int übergeben wird, hier anders
        for k, v in Value.items():
            if not k in forbidden:
                qry_cmd_Lst.append('SET a.' + k + ' = \"' + v + '\"')
        qry_cmd_Lst.append('RETURN ID(a)')
        qry = ' '.join(qry_cmd_Lst)
        print("neo_edit_node: ", qry)

        qry_cmd_Lst = []
        if "child_of" in Value.keys():
            qry_cmd_Lst.append('MATCH (a) WHERE ID(a) = ' + str(Value['_id']))  # Da ein Int übergeben wird, hier anders
            qry_cmd_Lst.append('MATCH (b) WHERE ID(b) = ' + str(Value['child_of']))  # Da ein Int übergeben wird, hier anders
            qry_cmd_Lst.append('MATCH path = (b)-[r1]->(a)')
            qry_cmd_Lst.append('WHERE r1.HookNode = \"' + HookNode + '\"')
            qry_cmd_Lst.append('CALL apoc.refactor.setType(r1, \"' + Value['parent_link'] + '\")')
            qry_cmd_Lst.append('YIELD input, output')
            qry_cmd_Lst.append('RETURN input, output')
        else:
            qry_cmd_Lst.append('RETURN "hallo"')
        qry2 = ' '.join(qry_cmd_Lst)
        print("neo_edit_node_relation: ", qry2)

        with self.driver.session() as session:
            if "child_of" in Value.keys():
                records2 = session.run(qry2)
            records = session.run(qry)
            print("neo_edit_node NEO4J CMD DONE", records)

        return True

    def get_header_of_querry(self, data, header):
        for item in data:
            tmp_keys = item.keys()
            for tmp_key in tmp_keys:
                if tmp_key not in header:
                    header.append(tmp_key)
        return header

    def tree_reload_clicked(self):
        try:
            self.clear_node_selection()
        except:
            print("clear_node_selection GEHT NICHT")
        self.treeView_1_selected.clear()

        selected_Hook_Node = self.comboBox_Hook_Node_1.currentText()
        selected_Scope = self.comboBox_multi_1.check_items()
        selected_Query_End_Node = self.comboBox_tree_1_Endnode.currentText()
        level_depth = self.comboBox_levels_1.currentText()
        # level_depth = "1"
        resp_rel = self.checkBox_respect_related_1.checkState()

        data = self.neo_hook_scope_path_query(selected_Hook_Node, selected_Scope, selected_Query_End_Node,
                                              level_depth, False)
        if resp_rel:
            data2 = self.neo_hook_scope_path_query(selected_Hook_Node, selected_Scope, selected_Query_End_Node,
                                              level_depth, True)
            data2.extend(data)
            sortedlist = []
            for item in data2:
                if item not in sortedlist:
                    sortedlist.append(item)
            data = sortedlist

        if data:
            extracted_data = extract_table(data)
            data_by_id(extracted_data, self.databyid)
            header = []
            header.extend(self.get_header_of_querry(extracted_data, self.headers[0]))

            self.model_tree_1.clear()
            self.model_tree_1.setHorizontalHeaderLabels(header)

            self.treeView_1.header().setDefaultSectionSize(400)
            print("Start import Data for Tree 1")
            # self.importData(data, self.model_tree_1, header)
            self.printer(data, self.model_tree_1.invisibleRootItem(), header)
            self.treeView_1.setModel(self.model_tree_1)
            self.treeView_1.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_1.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_1.expandToDepth(0)
            print("End import Data for Tree 1")
            q_header = self.treeView_1.header()
            for ii in range(1, len(self.headers[0])):
                q_header.setSectionResizeMode(ii, QtWidgets.QHeaderView.ResizeToContents)
        else:
            self.model_tree_1.clear()
            print('Leeres Query Ergebnis')

    def check_selected_item(self, in_qmodel_intex):
        def check_all_childs(qmodel_intex, CheckState):
            qmodel = qmodel_intex.model()
            qitem_active = qmodel.itemFromIndex(qmodel_intex)
            nid = int(qmodel.itemFromIndex(qmodel_intex.siblingAtColumn(2)).text())
            if qitem_active.hasChildren():
                for ii in range(0, qitem_active.rowCount()):
                    child = qitem_active.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    if CheckState:
                        child.setCheckState(QtCore.Qt.Checked)
                    else:
                        child.setCheckState(QtCore.Qt.Unchecked)

                    check_all_childs(child.index(), CheckState)

        qmodel = in_qmodel_intex.model()
        try:
            qitem_active = qmodel.itemFromIndex(in_qmodel_intex)
            if qmodel == self.model_tree_1:
                for item_temp in self.treeView_1_selected:
                    item_temp.setCheckState(0)
                self.treeView_1_selected.clear()
                qitem_active.setCheckState(QtCore.Qt.Checked)
                self.treeView_1_selected.append(qitem_active)
            elif qmodel == self.model_tree_3:
                for item_temp in self.treeView_3_selected:
                    item_temp.setCheckState(0)
                    check_all_childs(item_temp.index(), False)
                self.treeView_3_selected.clear()
                qitem_active.setCheckState(QtCore.Qt.Checked)
                self.treeView_3_selected.append(qitem_active)
            elif qmodel == self.model_tree_2:
                for item_temp in self.treeView_2_selected:
                    item_temp.setCheckState(0)
                    check_all_childs(item_temp.index(), False)
                self.treeView_2_selected.clear()
                qitem_active.setCheckState(QtCore.Qt.Checked)
                self.treeView_2_selected.append(qitem_active)
            elif qmodel == self.model_tree_4:
                for item_temp in self.treeView_4_selected:
                    item_temp.setCheckState(0)
                    check_all_childs(item_temp.index(), False)
                self.treeView_4_selected.clear()
                qitem_active.setCheckState(QtCore.Qt.Checked)
                self.treeView_4_selected.append(qitem_active)
        except:
            print("ERROR to Check Elements")

    def show_details_of_node(self, qmodel_intex):
        def check_all_childs(qmodel_intex, CheckState):
            qmodel = qmodel_intex.model()
            qitem_active = qmodel.itemFromIndex(qmodel_intex)
            nid = int(qmodel.itemFromIndex(qmodel_intex.siblingAtColumn(2)).text())
            if qitem_active.hasChildren():
                for ii in range(0, qitem_active.rowCount()):
                    child = qitem_active.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    if CheckState:
                        child.setCheckState(QtCore.Qt.Checked)
                    else:
                        child.setCheckState(QtCore.Qt.Unchecked)

                    check_all_childs(child.index(), CheckState)

        def check_selected_item(in_qmodel_intex):
            try:
                in_qmodel_intex = in_qmodel_intex.siblingAtColumn(0)
                qitem_active = qmodel.itemFromIndex(in_qmodel_intex)
                if qmodel == self.model_tree_1:
                    for item_temp in self.treeView_1_selected:
                        item_temp.setCheckState(0)
                    self.treeView_1_selected.clear()
                    qitem_active.setCheckState(QtCore.Qt.Checked)
                    self.treeView_1_selected.append(qitem_active)
                elif qmodel == self.model_tree_3:
                    for item_temp in self.treeView_3_selected:
                        item_temp.setCheckState(0)
                        check_all_childs(item_temp.index(), False)
                    self.treeView_3_selected.clear()
                    qitem_active.setCheckState(QtCore.Qt.Checked)
                    self.treeView_3_selected.append(qitem_active)
                elif qmodel == self.model_tree_2:
                    for item_temp in self.treeView_2_selected:
                        item_temp.setCheckState(0)
                        check_all_childs(item_temp.index(), False)
                    self.treeView_2_selected.clear()
                    qitem_active.setCheckState(QtCore.Qt.Checked)
                    self.treeView_2_selected.append(qitem_active)
                elif qmodel == self.model_tree_4:
                    for item_temp in self.treeView_4_selected:
                        item_temp.setCheckState(0)
                        check_all_childs(item_temp.index(), False)
                    self.treeView_4_selected.clear()
                    qitem_active.setCheckState(QtCore.Qt.Checked)
                    self.treeView_4_selected.append(qitem_active)
            except:
                print("ERROR to Check Elements")

        self.currentitemindex = qmodel_intex
        self.table_model.clear()
        qmodel = qmodel_intex.model()
        if type(qmodel) == QSortFilterProxyModel:
            while type(qmodel) == QSortFilterProxyModel:
                qmodel_intex = qmodel.mapToSource(qmodel_intex)
                qmodel = qmodel_intex.model()
        else:
            qmodel_intex = qmodel_intex
            qmodel = qmodel_intex.model()
        qitem = qmodel.itemFromIndex(qmodel_intex)
        in_qmodel_intex = qmodel_intex

        dicti = {}
        if qmodel == self.fileSystemModel:
            print("Hier bin ich")
            file_name = self.fileSystemModel.fileName(in_qmodel_intex)
            file_path = self.fileSystemModel.filePath(in_qmodel_intex)

            dicti = {'Title': file_name, 'Details': file_path}
        else:
            """ Könnte als  Funktion get node deteils ausgebunden werden"""
            dbid = int(qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(2)).text())
            if dbid in self.databyid:
                # dicti = self.databyid[dbid]
                dicti = self.get_node_details_by_id(dbid)
                self.databyid.update({dbid: dicti})
            else:
                dicti = self.get_node_details_by_id(dbid)
                self.databyid.update({dbid: dicti})
            """" Stellt sicher, das nur 1 Item selected ist"""
            HookNode = ""
            check_selected_item(in_qmodel_intex)
        ######################################### #####################################################################
        ###################    Sonderfunktion zum Update des Windows Explorer Libary XML     ##########################
        ######################################### #####################################################################
        if dicti['Title'] == "Aktuelle_Bearbeitung":
            update_libary = True
        elif dicti['Title'] == "My Workspace":
            update_libary = True
        else:
            update_libary = False

        if update_libary:
            folder_dict = {}
            HookNode_in = self.comboBox_Hook_Node_1.currentText()
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + str(dicti['_id']))
            qry_cmd_Lst.append('MATCH path = (x)-[rr*0..1]->(b)')
            qry_cmd_Lst.append('WHERE (b:File OR b:OS_Folder) AND')
            qry_cmd_Lst.append('ALL(r in rr')
            qry_cmd_Lst.append('WHERE type(r) = "has"')
            qry_cmd_Lst.append('AND r.HookNode = \"' + HookNode_in + '\"')
            qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
            qry_cmd_Lst.append('return b')
            qry = ' '.join(qry_cmd_Lst)
            print(qry)
            with self.driver.session() as session:
                records = session.run(qry)

                for record in records:
                    folder_dict.update({record['b'].get("Title"): record['b'].get("DeepLink")})
                session.close()

            print("folder_dict", folder_dict)
            modify_windows_libary(dicti['Title'], folder_dict)
        ######################################### #####################################################################
        ###################    Sonderfunktion zum Update des Windows Explorer Libary XML     ##########################
        ######################################### #####################################################################

        try:
            pid = int(qmodel.itemFromIndex(qitem.parent().index().siblingAtColumn(2)).text())
            parent_link = qmodel.itemFromIndex(qitem.index().siblingAtColumn(3)).text()
            dicti.update({"child_of": pid})
            dicti.update({"parent_link": parent_link})
        except:
            print("Selected Node == Root Node")

        # print('dicti of details', dicti)
        forbidden = ['_id', '_type', 'child_of', 'Key', 'hooknode']
        for k, v in dicti.items():
            if not v:
                v = ""
            if k in forbidden:
                edit_able = False
            else:
                edit_able = True

            self.table_model.invisibleRootItem().appendRow([StandardTableItem(k, False, self.tree_font_size, True, color=QColor(0, 0, 0)),
                                                            StandardTableItem(str(v), edit_able, self.tree_font_size, False, color=QColor(0, 0, 0))])

        self.tableView.setModel(self.table_model)
        self.tableView.resizeColumnToContents(1)
        self.tableView.resizeRowsToContents()
        self.tableView.sortByColumn(0, QtCore.Qt.AscendingOrder)

        header = self.tableView.horizontalHeader()
        header.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeToContents)

        return True

    def show_neighbor(self, qmodel_intex):
        use_hooknode = self.checkBox_use_hooknode.checkState()

        if self.model_tree_2_ready:
            self.model_tree_2_ready = False

            qmodel = qmodel_intex.model()
            if type(qmodel) == QSortFilterProxyModel:
                while type(qmodel) == QSortFilterProxyModel:
                    qmodel_intex = qmodel.mapToSource(qmodel_intex)
                    qmodel = qmodel_intex.model()
            else:
                qmodel_intex = qmodel_intex
                qmodel = qmodel_intex.model()
            qitem = qmodel.itemFromIndex(qmodel_intex.siblingAtColumn(0))
            in_qmodel_intex = qmodel_intex
            qmodel = in_qmodel_intex.model()

            dbid = int(qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(2)).text())
            node_type = qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(1)).text()
            node_title = qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(0)).text()

            if qmodel == self.model_tree_1:
                HookNode_in = self.comboBox_Hook_Node_1.currentText()

            if dbid in self.databyid:
                dicti = self.databyid[dbid]
            else:
                dicti = self.get_node_details_by_id(dbid)
                self.databyid.update({dbid: dicti})

            if node_type == "OS_Folder":
                data = self.neo_up_query_has(dbid, 6)
                header = []
                header.extend(self.headers[0])
                header = self.get_header_of_querry(data, header)

                self.model_tree_2.clear()
                self.treeView_2.setModel(self.model_tree_2)
                self.model_tree_2.setHorizontalHeaderLabels(header)
                self.treeView_2.header().setDefaultSectionSize(300)
                self.importData(data, self.model_tree_2, header)
                self.treeView_2.expandToDepth(0)

                q_header = self.treeView_2.header()
                q_header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
                q_header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)


                node_details = qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(4)).text()
                self.fileSystemModel.setReadOnly(True)
                root = self.fileSystemModel.setRootPath(node_details)
                self.model_tree_4.clear()
                self.treeView_4.setModel(self.fileSystemModel)
                self.treeView_4.setRootIndex(root)
                #self.treeView_4.setRootIndex(self.treeView_4.model().index(QtCore.QDir.currentPath()))

            else:
                # Upwards / Parent Query
                print("# Upwards / Parent Query")
                data = self.neo_up_query_has(dbid, 2)
                header = self.headers[2]
                self.model_tree_2.clear()
                self.treeView_2.setModel(self.model_tree_2)
                self.model_tree_2.setHorizontalHeaderLabels(header)
                self.treeView_2.header().setDefaultSectionSize(300)
                self.importData(data, self.model_tree_2, header)
                self.treeView_2.expandToDepth(0)

                q_header = self.treeView_2.header()
                q_header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
                q_header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

                # Downwards / Parent Query
                print("# Downwards / Child Query")
                HookNode_in = ""
                use_hooknode = False
                if qmodel == self.model_tree_1:
                    HookNode_in = self.comboBox_Hook_Node_1.currentText()
                    use_hooknode = True
                elif qmodel == self.model_tree_3:
                    if node_type == "HookNode":
                        HookNode_in = node_title
                        use_hooknode = self.checkBox_respect_HookNode.checkState()


                data_for_expand = self.neo_down_query_has(dbid, int(self.comboBox_expand_1.currentText()),
                                                          use_hooknode, HookNode_in, True)
                if (qmodel == self.model_tree_1) or (qmodel == self.model_tree_3):
                        # self.expand_tree_node(data_for_expand, qmodel_intex, self.headers[0])
                        extracted_data = extract_table(data_for_expand)
                        data_by_id(extracted_data, self.databyid)
                        header = []
                        header.extend(self.get_header_of_querry(extracted_data, self.headers[0]))
                        print("Start Printer")
                        self.printer(data_for_expand, qitem, self.headers[2])

                        if (qmodel == self.model_tree_1):
                            self.treeView_1.expand(qmodel_intex)
                        elif (qmodel == self.model_tree_3):
                            self.treeView_3.expand(qmodel_intex)

                data2 = self.neo_down_query_has(dbid, 1, False, "", True)
                # # data2 = self.neo_down_query_related(dbid)
                # # data2.extend(data)
                # sortedlist = []
                # for item in data2:
                #     if item:
                #         if item not in sortedlist:
                #            sortedlist.append(item)
                # data = sortedlist
                # data_by_id(data, self.databyid)
                header = self.headers[2]
                self.model_tree_4.clear()
                self.treeView_4.setModel(self.model_tree_4)
                self.model_tree_4.setHorizontalHeaderLabels(header)
                self.treeView_4.header().setDefaultSectionSize(300)
                # self.importData(data, self.model_tree_4, header)
                self.printer(data2, self.model_tree_4.invisibleRootItem(), header)
                self.treeView_4.expandToDepth(0)

                q_header = self.treeView_4.header()
                q_header.setSectionResizeMode(1, QtWidgets.QHeaderView.ResizeToContents)
                q_header.setSectionResizeMode(2, QtWidgets.QHeaderView.ResizeToContents)

            self.model_tree_2_ready = True

    def link_checked_items(self, from_node_id, target_model):
        to_nodes = []
        try:
            if target_model == self.model_tree_1:
                to_item = self.treeView_1_selected[0]  # Als Liste implementiert für spätere Erweiterungen
                # to_item.setCheckState(0)
                current_hooknode = self.comboBox_Hook_Node_1.currentText()
                dbid = int(self.model_tree_1.itemFromIndex(to_item.index().siblingAtColumn(2)).text())
                to_nodes.append(dbid)
                # tmp_item.setCheckState(0)


            from_type = self.databyid[from_node_id]['_type']
            to_type = self.databyid[to_nodes[0]]['_type']
            link_ok = self.link_data_model[to_type][from_type]
            print(link_ok)
        except:
            link_ok = 0

        if not link_ok > 0:
            print("Da ginn was nicht?")
        else:
            print("Alles Cool")
            link_done = self.neo_create_link(to_nodes[0], from_node_id, current_hooknode)

            if link_done:
                value = self.databyid[from_node_id]
                to_item.appendRow([
                    StandardItem(value['Title'], value['_type'], False, True, self.tree_font_size,
                                 color=QColor(0, 0, 0)),
                    StandardItem(value['_type'], value['_type'], False, False, self.tree_font_size,
                                 color=QColor(0, 0, 0)),
                    StandardItem(str(value['_id']), value['_type'], False, False, self.tree_font_size,
                                 color=QColor(0, 0, 0))
                ])
        # self.clear_node_selection()

    def relate_checked_items(self, from_node_id, target_model):
        to_nodes = []
        try:
            if target_model == self.model_tree_1:
                to_items = self.treeView_1_selected  # Als Liste implementiert für spätere Erweiterungen
                to_items[0].setCheckState(0)
                current_hooknode = self.comboBox_Hook_Node_1.currentText()
                for tmp_item in self.treeView_1_selected:
                    dbid = int(self.model_tree_1.itemFromIndex(tmp_item.index().siblingAtColumn(2)).text())
                    to_nodes.append(dbid)
                    tmp_item.setCheckState(0)

            link_ok = 1
        except:
            link_ok = 0

        if not link_ok > 0:
            print("Bist du dumm?")
        else:
            print("Alles Cool")
            link_done = self.neo_create_related(to_nodes[0], from_node_id, current_hooknode)

            if link_done:
                value = self.databyid[from_node_id]
                to_items[0].appendRow([
                    StandardItem(value['Title'], value['_type'], False, True, self.tree_font_size,
                                 color=QColor(0, 0, 0)),
                    StandardItem(value['_type'], value['_type'], False, False, self.tree_font_size,
                                 color=QColor(0, 0, 0)),
                    StandardItem(str(value['_id']), value['_type'], False, False, self.tree_font_size,
                                 color=QColor(0, 0, 0))
                ])
        self.clear_node_selection()

    def delete_link_to_parent(self, qitem, dbid):
        child_row = qitem.row()
        parent_item = qitem.parent()
        qmodel = qitem.index().model()
        if qmodel == self.model_tree_1:
            pid = int(self.model_tree_1.itemFromIndex(parent_item.index().siblingAtColumn(2)).text())
            link_type = self.model_tree_1.itemFromIndex(qitem.index().siblingAtColumn(3)).text()
        elif qmodel == self.model_tree_3:
            pid = int(self.model_tree_3.itemFromIndex(parent_item.index().siblingAtColumn(2)).text())
            link_type = self.model_tree_3.itemFromIndex(qitem.index().siblingAtColumn(3)).text()

        done = self.neo_delete_link_by_node_ids(dbid, pid, link_type)
        self.clear_node_selection()
        if done:
            try:
                parent_item.removeRow(child_row)
            except:
                print("ERROR parent_item.removeRow(child_row)")
            return
        else:
            print("error in delete link")

    def create_new_node(self, pid, item, new_item_type):
        try:
            qmodel = item.index().model()
            if qmodel == self.model_tree_1:
                current_scope = self.comboBox_applay_Scope_1.currentText()
                current_hooknode = self.comboBox_Hook_Node_1.currentText()
            elif qmodel == self.model_tree_3:
                current_scope = self.comboBox_applay_Scope_1.currentText()
                current_hooknode = self.comboBox_Hook_Node_3.currentText()
            new_node_id = self.neo_create_node(pid, new_item_type, current_hooknode, current_scope)
            print("new _node_id", new_node_id)
            value = self.get_node_details_by_id(new_node_id)
            print("value", value)

            if True:
                item.appendRow([
                        StandardItem(value['Title'], value['_type'], False, True, self.tree_font_size,
                                     color=QColor(0, 0, 0)),
                        StandardItem(value['_type'], value['_type'], False, False, self.tree_font_size,
                                     color=QColor(0, 0, 0)),
                        StandardItem(str(value['_id']), value['_type'], False, False, self.tree_font_size,
                                     color=QColor(0, 0, 0)),
                        StandardItem("has", value['_type'], False, False, self.tree_font_size,
                                     color=QColor(0, 0, 0)),
                        StandardItem(value['Details'], value['_type'], False, False, self.tree_font_size,
                                     color=QColor(0, 0, 0))
                    ])
            return True
        except:
            print("Funktion create_new_node ging schief")
            return False

    def edit_node(self):
        Value= {}

        qmodel = self.table_model
        qitem_active = qmodel.invisibleRootItem()
        if qitem_active.hasChildren():
            for ii in range(0, qitem_active.rowCount()):
                key_name = qitem_active.child(ii, 0).text()
                value_name = qitem_active.child(ii, 1).text()
                value_name = value_name.replace("\\", "/")
                value_name = value_name.replace("\"", "")

                # if "//collaboration.claas.com" in value_name:
                #     value_name = value_name.replace("https:", "")
                #     value_name = value_name.replace("%20", " ")

                Value.update({key_name: value_name})

        print(Value)
        self.neo_edit_node(Value)

        qmodel_intex = self.currentitemindex
        qmodel = qmodel_intex.model()
        if type(qmodel) == QSortFilterProxyModel:
            while type(qmodel) == QSortFilterProxyModel:
                qmodel_intex = qmodel.mapToSource(qmodel_intex)
                qmodel = qmodel_intex.model()
        else:
            qmodel_intex = qmodel_intex
            qmodel = qmodel_intex.model()
        in_qmodel_intex = qmodel_intex

        qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(0)).setText(Value['Title'])
        if 'Details' in Value.keys():
            qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(4)).setText(Value['Details'])
        if 'parent_link' in Value.keys():
            qmodel.itemFromIndex(in_qmodel_intex.siblingAtColumn(3)).setText(Value['parent_link'])

    def add_node_to_scope(self):
        all_nodes = []
        for tmp_item in self.treeView_4_selected:
            dbid = int(self.model_tree_4.itemFromIndex(tmp_item.index().siblingAtColumn(2)).text())
            all_nodes.append(dbid)
            tmp_item.setCheckState(0)
        self.treeView_4_selected.clear()
        print(all_nodes)

        self.clear_node_selection()

        if len(all_nodes) > 0:
            self.neo_add_to_scope(all_nodes)

    def delete_node_from_scope(self):
        all_nodes = []
        for tmp_item in self.treeView_1_selected:
            dbid = int(self.model_tree_1.itemFromIndex(tmp_item.index().siblingAtColumn(2)).text())
            all_nodes.append(dbid)
            tmp_item.setCheckState(0)
        self.treeView_1_selected.clear()
        print(all_nodes)

        self.clear_node_selection()

        if len(all_nodes) > 0:
            self.neo_delete_from_scope(all_nodes)

    def clear_node_selection(self):
        try:
            for tmp_item in self.treeView_1_selected:
                tmp_item.setCheckState(0)
            self.treeView_1_selected.clear()
            for tmp_item in self.treeView_3_selected:
                tmp_item.setCheckState(0)
            self.treeView_3_selected.clear()
            for tmp_item in self.treeView_4_selected:
                tmp_item.setCheckState(0)
            self.treeView_4_selected.clear()
        except:
            print("ERROR in clear_node_selection")

    def expant_until_root_node(self, tree_view, qitem: QStandardItem):
        qmodel = qitem.model()
        if not qitem == qmodel.invisibleRootItem():
            tree_view.expand(qitem.index())
            qitem = qitem.parent()
            if qitem:
                self.expant_until_root_node(tree_view, qitem)

    def expand_tree_to_item(self, tree_nr: int):
        if tree_nr == 1:
            if self.treeView_1_selected:
                qitem = self.treeView_1_selected[0]
                qmodel_index = qitem.index()
                qmodel = qmodel_index.model()
                if type(qmodel) == QSortFilterProxyModel:
                    print("Ist SortFilterProxy")
                    qmodel_index = qmodel.mapToSource(qmodel_index)
                qitem = qmodel.itemFromIndex(qmodel_index)
                self.expant_until_root_node(self.treeView_1, qitem)
        elif tree_nr == 3:
            if self.treeView3_selected:
                qitem = self.treeView_3_selected[0]
                qmodel_index = qitem.index()
                qmodel = qmodel_index.model()
                if type(qmodel) == QSortFilterProxyModel:
                    print("Ist SortFilterProxy")
                    qmodel_index = qmodel.mapToSource(qmodel_index)
                qitem = qmodel.itemFromIndex(qmodel_index)
                self.expant_until_root_node(self.treeView_3, qitem)

    def filter_tree_view(self, tree_nr: int):
        if tree_nr == 1:
            search_string = self.lineEdit_1.text()
            if search_string:
                tmp_lst = search_string.split(" ")
                if len(tmp_lst) == 1:
                    self.proxyModel_tree_1.setSourceModel(self.model_tree_1)
                    self.proxyModel_tree_1.setFilterKeyColumn(-1)
                    self.proxyModel_tree_1.setRecursiveFilteringEnabled(True)
                    self.proxyModel_tree_1.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                    self.proxyModel_tree_1.setFilterFixedString(tmp_lst[0])
                elif len(tmp_lst) == 2:
                    self.proxyModel_tree_1_1.setSourceModel(self.model_tree_1)
                    self.proxyModel_tree_1_1.setFilterKeyColumn(-1)
                    self.proxyModel_tree_1_1.setRecursiveFilteringEnabled(True)
                    self.proxyModel_tree_1_1.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                    self.proxyModel_tree_1_1.setFilterFixedString(tmp_lst[0])

                    self.proxyModel_tree_1.setSourceModel(self.proxyModel_tree_1_1)
                    self.proxyModel_tree_1.setFilterKeyColumn(-1)
                    self.proxyModel_tree_1.setRecursiveFilteringEnabled(True)
                    self.proxyModel_tree_1.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                    self.proxyModel_tree_1.setFilterFixedString(tmp_lst[1])

                elif len(tmp_lst) == 3:
                    self.proxyModel_tree_1_1.setSourceModel(self.model_tree_1)
                    self.proxyModel_tree_1_1.setFilterKeyColumn(-1)
                    self.proxyModel_tree_1_1.setRecursiveFilteringEnabled(True)
                    self.proxyModel_tree_1_1.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                    self.proxyModel_tree_1_1.setFilterFixedString(tmp_lst[0])

                    self.proxyModel_tree_1_2.setSourceModel(self.proxyModel_tree_1_1)
                    self.proxyModel_tree_1_2.setFilterKeyColumn(-1)
                    self.proxyModel_tree_1_2.setRecursiveFilteringEnabled(True)
                    self.proxyModel_tree_1_2.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                    self.proxyModel_tree_1_2.setFilterFixedString(tmp_lst[0])

                    self.proxyModel_tree_1.setSourceModel(self.proxyModel_tree_1_2)
                    self.proxyModel_tree_1.setFilterKeyColumn(-1)
                    self.proxyModel_tree_1.setRecursiveFilteringEnabled(True)
                    self.proxyModel_tree_1.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                    self.proxyModel_tree_1.setFilterFixedString(tmp_lst[1])

                # search_reg_ex = "^"
                # for tmp_str in tmp_lst:
                #     search_reg_ex += '(?=.*' + tmp_str + '.*)'
                # search_reg_ex += '.*$'
                #
                # print(search_reg_ex)
                # self.proxyModel_tree_1.setSourceModel(self.model_tree_1)
                # self.proxyModel_tree_1.setFilterKeyColumn(-1)
                # self.proxyModel_tree_1.setRecursiveFilteringEnabled(True)
                # self.proxyModel_tree_1.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                # self.proxyModel_tree_1.setFilterRegularExpression(search_reg_ex)

                self.treeView_1.setModel(self.proxyModel_tree_1)
                self.treeView_1.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_1.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_1.expandAll()
            else:
                self.treeView_1.setModel(self.model_tree_1)
                self.treeView_1.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_1.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_1.expandToDepth(0)

                self.expand_tree_to_item(tree_nr)


        elif tree_nr == 2:
            search_string = self.lineEdit_2.text()
            if search_string:
                tmp_lst = search_string.split(" ")
                search_reg_ex = "^"
                for tmp_str in tmp_lst:
                    search_reg_ex += '(?=.*' + tmp_str + '.*)'
                search_reg_ex += '.*$'

                self.proxyModel_tree_2.setSourceModel(self.model_tree_2)
                self.proxyModel_tree_2.setFilterKeyColumn(-1)
                self.proxyModel_tree_2.setRecursiveFilteringEnabled(True)
                self.proxyModel_tree_2.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                self.proxyModel_tree_2.setFilterRegularExpression(search_reg_ex)

                self.proxyModel_tree_4.setSourceModel(self.model_tree_4)
                self.proxyModel_tree_4.setFilterKeyColumn(-1)
                self.proxyModel_tree_4.setRecursiveFilteringEnabled(True)
                self.proxyModel_tree_4.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                self.proxyModel_tree_4.setFilterRegularExpression(search_reg_ex)

                self.treeView_2.setModel(self.proxyModel_tree_2)
                self.treeView_2.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_2.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_2.expandAll()

                self.treeView_4.setModel(self.proxyModel_tree_4)
                self.treeView_4.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_4.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_4.expandAll()
            else:
                self.treeView_2.setModel(self.model_tree_2)
                self.treeView_2.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_2.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_2.expandToDepth(0)

                self.treeView_4.setModel(self.model_tree_4)
                self.treeView_4.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_4.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_4.expandToDepth(0)

        elif tree_nr == 3:
            search_string = self.lineEdit_3.text()
            if search_string:
                tmp_lst = search_string.split(" ")
                search_reg_ex = "^"
                for tmp_str in tmp_lst:
                    search_reg_ex += '(?=.*' + tmp_str + '.*)'
                search_reg_ex += '.*$'

                self.proxyModel_tree_3.setSourceModel(self.model_tree_3)
                self.proxyModel_tree_3.setFilterKeyColumn(-1)
                self.proxyModel_tree_3.setRecursiveFilteringEnabled(True)
                self.proxyModel_tree_3.setFilterCaseSensitivity(QtCore.Qt.CaseInsensitive)
                self.proxyModel_tree_3.setFilterRegularExpression(search_reg_ex)

                self.treeView_3.setModel(self.proxyModel_tree_3)
                self.treeView_3.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_3.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_3.expandAll()
            else:
                self.treeView_3.setModel(self.model_tree_3)
                self.treeView_3.sortByColumn(0, QtCore.Qt.AscendingOrder)
                self.treeView_3.sortByColumn(1, QtCore.Qt.AscendingOrder)
                self.treeView_3.expandToDepth(0)

        return True

    def save_model_to_excel(self):
        def get_item_dicti(item: QStandardItem, Dict_List: list):
            changed_item_index = item.index()
            changed_item_model = item.model()
            dbid = int(changed_item_model.itemFromIndex(changed_item_index.siblingAtColumn(2)).text())
            parent_link = changed_item_model.itemFromIndex(changed_item_index.siblingAtColumn(3)).text()
            dicti = {}
            if dbid in self.databyid:
                dicti.update(self.databyid[dbid])
            else:
                dicti.update(self.get_node_details_by_id(dbid))

            # print(dicti)
            level = 1
            temp_item = item
            while temp_item.parent():
                temp_item = temp_item.parent()
                level += 1
            dicti.update({'Level': level})
            dicti.update({'parent_link': parent_link})
            Dict_List.append(dicti)
            rows = item.rowCount()
            for row in range(0, rows):
                child = item.child(row, 0)
                get_item_dicti(child, Dict_List)

        with self.driver.session() as session:
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (x) WHERE x.Key = ""')
            qry_cmd_Lst.append('SET x.Key = apoc.create.uuid()')
            qry = ' '.join(qry_cmd_Lst)
            print(qry)
            records = session.run(qry)
        session.close()

        list_out = []
        get_item_dicti(self.model_tree_1.invisibleRootItem().child(0, 0), list_out)
        selected_Hook_Node = self.comboBox_Hook_Node_1.currentText()

        df = pd.DataFrame.from_records(list_out)
        df = df.fillna('')

        headers = df.columns
        print(headers)

        for head in headers:
            if '.HookNode' in head:
                df = df.drop(head, axis=1)



        # df = df.applymap(lambda x: x.encode('unicode_escape').decode('ansi') if isinstance(x, str) else x)
        #df = df.applymap(lambda x: x.decode('utf-8', 'ignore').encode('utf-8') if isinstance(x, str) else x)
        print(df)

        ordner_pfad = './Output/'
        filename = selected_Hook_Node
        filename = filename.replace('(', '')
        filename = filename.replace(')', '')
        filename = filename.replace(':', '_')
        filename = filename.replace(' ', '_')
        filename = filename + '.xlsx'
        df.to_excel(ordner_pfad + 'out_' + filename, sheet_name='Tabelle1', index=False)


        model = self.model_tree_1
        root = model.invisibleRootItem()
        root.columnCount()

    def search_for_next_artefact(self, qmodel_intex):
        def neo_count_level_to_artefact(Neo_ID: str):
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + Neo_ID)
            if Scope_in:
                qry_cmd_Lst.append('MATCH (b:Scope) WHERE')
                for scope in Scope_in:
                    qry_cmd_Lst.append('b.Title= \"' + scope + '\"')
                    qry_cmd_Lst.append('OR')
                del qry_cmd_Lst[-1]
                qry_cmd_Lst.append('MATCH tmp = (y)-[]-(b)')
            if respect_related:
                qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + level_depth + ']->()-[]-(y:' + End_Node + ')')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + level_depth + ']->(y:' + End_Node + ')')
            qry_cmd_Lst.append('WHERE ALL(r in rr')
            qry_cmd_Lst.append('WHERE r.HookNode = \"' + HookNode_in + '\"')
            qry_cmd_Lst.append('AND type(r) = "has"')
            qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
            qry_cmd_Lst.append('RETURN count(path) as pathCount')
            qry = ' '.join(qry_cmd_Lst)

            num_paths = -1
            with self.driver.session() as session:
                records = session.run(qry)
                for record in records:
                    num_paths = record['pathCount']
            session.close()
            return num_paths

        def search_for_next_artefact_recrusion(qitem_active: QStandardItem):
            qmodel = qitem_active.model()
            if not qitem_active == qmodel.invisibleRootItem():
                cid = str(qmodel.itemFromIndex(qitem_active.index().siblingAtColumn(2)).text())
                if End_Node == "*":
                    qitem_active.setBackground(QColor(255, 255, 255))
                else:
                    levels = neo_count_level_to_artefact(cid)
                    if levels == 0:
                        qitem_active.setBackground(QColor(255, 173, 153))
                    else:
                        qitem_active.setBackground(QColor(179, 230, 179))

            if qitem_active.hasChildren():
                for ii in range(0, qitem_active.rowCount()):
                    child = qitem_active.child(ii)
                    search_for_next_artefact_recrusion(child)

        qmodel = qmodel_intex.model()
        qitem = qmodel.itemFromIndex(qmodel_intex)

        if qmodel == self.model_tree_1:
            HookNode_in = self.comboBox_Hook_Node_1.currentText()
            Scope_in = self.comboBox_multi_1.check_items()
            End_Node = self.comboBox_tree_1_Endnode.currentText()
            level_depth = self.comboBox_levels_1.currentText()
            respect_related = self.checkBox_respect_related_1.checkState()
        elif qmodel == self.model_tree_3:
            HookNode_in = self.comboBox_Hook_Node_3.currentText()
            Scope_in = self.comboBox_multi_3.check_items()
            End_Node = self.comboBox_tree_3_Endnode.currentText()
            level_depth = self.comboBox_levels_3.currentText()
            respect_related = self.checkBox_respect_related_2.checkState()


        search_for_next_artefact_recrusion(qitem)

    def add_tree(self, qitem):
        def add_all_childs(qmodel_intex, CheckState, current_hooknode):
            qmodel = qmodel_intex.model()
            qitem_active = qmodel.itemFromIndex(qmodel_intex)
            pid = int(qmodel.itemFromIndex(qmodel_intex.siblingAtColumn(2)).text())
            if qitem_active.hasChildren():
                for ii in range(0, qitem_active.rowCount()):
                    child = qitem_active.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    if CheckState:
                        child.setCheckState(QtCore.Qt.Checked)
                        print("Parent: ", pid, ' Child: ', cid)
                        self.neo_create_link(pid, cid, current_hooknode)
                    else:
                        child.setCheckState(QtCore.Qt.Unchecked)

                    add_all_childs(child.index(), CheckState, current_hooknode)

        if self.treeView_3_selected:
            current_hooknode = self.comboBox_Hook_Node_1.currentText()
            tree_1_node = self.treeView_1_selected[0]
            pid = int(self.model_tree_1.itemFromIndex(tree_1_node.index().siblingAtColumn(2)).text())
            in_qmodel_intex = qitem.index()
            qmodel = in_qmodel_intex.model()
            cid = int(qmodel.itemFromIndex(qitem.index().siblingAtColumn(2)).text())
            self.neo_create_link(pid, cid, current_hooknode)
            add_all_childs(in_qmodel_intex, qitem.checkState(), current_hooknode)

            self.tree_reload_clicked()

    def dublicate_tree(self, qitem):
        def add_all_childs(qmodel_intex, CheckState, parent_dicti, current_hooknode):
            qmodel = qmodel_intex.model()
            qitem_active = qmodel.itemFromIndex(qmodel_intex)
            pid = int(qmodel.itemFromIndex(qmodel_intex.siblingAtColumn(2)).text())
            if qitem_active.hasChildren():
                for ii in range(0, qitem_active.rowCount()):
                    child = qitem_active.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    if CheckState:
                        child.setCheckState(QtCore.Qt.Checked)
                        print("Parent: ", pid, ' Child: ', cid)
                        new_cid = self.neo_dublicate_node(cid)
                        parent_map.update({cid: new_cid})
                        print(new_cid)
                        self.neo_create_link(parent_map[pid], new_cid, current_hooknode)
                    else:
                        child.setCheckState(QtCore.Qt.Unchecked)

                    add_all_childs(child.index(), CheckState, parent_dicti, current_hooknode)

        parent_map = {}
        try:
            current_hooknode = self.comboBox_Hook_Node_1.currentText()
            tree_1_node = self.treeView_1_selected[0]
            pid = int(self.model_tree_1.itemFromIndex(tree_1_node.index().siblingAtColumn(2)).text())
            in_qmodel_intex = qitem.index()
            qmodel = in_qmodel_intex.model()
            cid = int(qmodel.itemFromIndex(qitem.index().siblingAtColumn(2)).text())
            new_cid = self.neo_dublicate_node(cid)
            parent_map.update({cid: new_cid})
            self.neo_create_link(pid, new_cid, current_hooknode)
            add_all_childs(in_qmodel_intex, qitem.checkState(), parent_map, current_hooknode)

            self.tree_reload_clicked('tree_1')
        except:
            print("Dublicate Tree geht nicht")

    def get_path_query_string(self):
        start_label = ':' + self.comboBox_4search_label_1.currentText()
        start_label = start_label.replace(":*", "")
        start_attribute = self.comboBox_4search_property_1.currentText()
        start_attribute_value = self.lineEdit_4search_value_1.text()

        mid_label = ':' + self.comboBox_4search_label_2.currentText()
        mid_label = mid_label.replace(":*", "")
        mid_attribute = self.comboBox_4search_property_2.currentText()
        mid_attribute_value = self.lineEdit_4search_value_2.text()

        end_label = ':' + self.comboBox_4search_label_3.currentText()
        end_label = end_label.replace(":*", "")
        end_attribute = self.comboBox_4search_property_3.currentText()
        end_attribute_value = self.lineEdit_4search_value_3.text()

        level_depth_mid = self.comboBox_4search_depth_2.currentText()
        level_depth_end = self.comboBox_4search_depth_3.currentText()

        qry_cmd_Lst = []
        # Start Node
        if start_attribute == "*":
            qry_cmd_Lst.append('MATCH (x' + start_label + ')')
            tmp_lst = []
            for start_str_part in start_attribute_value.split(" "):
                tmp_lst.append('x[prop] =~ \"(?i).*' + start_str_part + '.*\"')

            qry_cmd_Lst.append('WHERE (any(prop in keys(x) where ' + " AND ".join(tmp_lst) + '))')
        else:
            tmp_lst = []
            for start_str_part in start_attribute_value.split(" "):
                tmp_lst.append('x.' + start_attribute + ' =~ \"(?i).*' + start_str_part + '.*\"')
            qry_cmd_Lst.append('MATCH (x' + start_label + ') WHERE ' + " AND ".join(tmp_lst))

        # End Node
        if end_attribute == "*":
            qry_cmd_Lst.append('MATCH (y' + end_label + ')')
            tmp_lst = []
            for str_part in end_attribute_value.split(" "):
                tmp_lst.append('y[prop] =~ \"(?i).*' + str_part + '.*\"')

            qry_cmd_Lst.append('WHERE (any(prop in keys(y) where ' + " AND ".join(tmp_lst) + '))')
        else:
            tmp_lst = []
            for str_part in end_attribute_value.split(" "):
                tmp_lst.append('y.' + end_attribute + ' =~ \"(?i).*' + str_part + '.*\"')
            qry_cmd_Lst.append('MATCH (y' + end_label + ') WHERE ' + " AND ".join(tmp_lst))

        # if end_attribute == "*":
        #     qry_cmd_Lst.append('MATCH (y' + end_label + ')')
        #     qry_cmd_Lst.append('WHERE (any(prop in keys(y) where y[prop] =~ \"(?i).*' + end_attribute_value + '.*\"))')
        # else:
        #     qry_cmd_Lst.append('MATCH (y' + end_label + ') WHERE y.' + end_attribute +
        #                        ' =~ \"(?i).*' + end_attribute_value + '.*\"')

        print("mid_label", mid_label)
        if mid_label == "":
            if level_depth_end == 'shortest':
                qry_cmd_Lst.append('MATCH path = shortestPath((x)-[*0..' + level_depth_mid + ']-(y))')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[*0..' + level_depth_end + ']-(y)')
        else:
            # End Node
            if mid_attribute == "*":
                qry_cmd_Lst.append('MATCH (z' + mid_label + ')')
                tmp_lst = []
                for str_part in mid_attribute_value.split(" "):
                    tmp_lst.append('z[prop] =~ \"(?i).*' + str_part + '.*\"')

                qry_cmd_Lst.append('WHERE (any(prop in keys(z) where ' + " AND ".join(tmp_lst) + '))')
            else:
                tmp_lst = []
                for str_part in mid_attribute_value.split(" "):
                    tmp_lst.append('z.' + mid_attribute + ' =~ \"(?i).*' + str_part + '.*\"')
                qry_cmd_Lst.append('MATCH (z' + end_label + ') WHERE ' + " AND ".join(tmp_lst))

            qry_cmd_Lst.append('MATCH path = (x)-[*0..' + level_depth_mid + ']-(z)-[*0..' + level_depth_end + ']-(y)')


        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)

        return qry

    def set_first_path_query_node(self, qitem):
        print("Bla")
        qmodel = qitem.model()
        start_item_key = qmodel.itemFromIndex(qitem.index().siblingAtColumn(5)).text()
        start_item_type = qmodel.itemFromIndex(qitem.index().siblingAtColumn(1)).text()

        self.comboBox_4search_label_1.setCurrentText(start_item_type)
        self.comboBox_4search_property_1.setCurrentText("Key")
        self.lineEdit_4search_value_1.setText(start_item_key)

    def path_query(self):
        qry = self.get_path_query_string()
        print(qry)
        data = []
        with self.driver.session() as session:
            records = session.run(qry)
            print("neo_hook_scope_path_query QUERY DONE")

            for node in records:
                # self.printer(node[0])
                # ret_val = extract_table(node[0])
                data.append(node[0])
                # print(data)
                # break
            # print(len(ret_val))
            session.close()

        if data[0]:
            # self.clear_node_selection()
            self.treeView_3_selected.clear()
            extracted_data = extract_table(data)
            data_by_id(extracted_data, self.databyid)
            header = []
            header.extend(self.get_header_of_querry(extracted_data, self.headers[0]))

            self.treeView_3_selected.clear()
            self.model_tree_3.clear()
            self.model_tree_3.setHorizontalHeaderLabels(header)

            self.treeView_3.header().setDefaultSectionSize(400)
            print("Start import Data for Tree 1")
            # self.importData(data, self.model_tree_1, header)
            self.printer(data, self.model_tree_3.invisibleRootItem(), header)
            self.treeView_3.setModel(self.model_tree_3)
            self.treeView_3.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_3.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_3.collapseAll()
            print("End import Data for Tree 1")
            q_header = self.treeView_3.header()
            for ii in range(1, len(self.headers[0])):
                q_header.setSectionResizeMode(ii, QtWidgets.QHeaderView.ResizeToContents)
        else:
            self.model_tree_3.clear()
            print('Leeres Query Ergebnis')

    def select_in_work4traces(self, dbid):
        dicti = self.databyid[dbid]
        title = dicti["Title"]
        self.comboBox_Hook_Node_1.setCurrentText(title)
        self.tree_reload_clicked()

    def querie_test(self):
        qry_cmd_Lst = []
        qry_cmd_Lst.append('MATCH (x:CPEM_Project)')
        qry_cmd_Lst.append('MATCH path_tmp = (x)-[]-(y:TestExecution)<-[*0..15]-(z:HookNode)')
        qry_cmd_Lst.append('WHERE z.Title = "Trion Test Plan"')
        qry_cmd_Lst.append('WITH x, y')
        qry_cmd_Lst.append('MATCH (dep:Department)')
        qry_cmd_Lst.append('MATCH path = (dep)-[:has]->(p:Person)-[:execute_validation]-(x)-[]-(y)')
        qry_cmd_Lst.append('WITH collect(path) as paths ')
        qry_cmd_Lst.append('CALL apoc.convert.toTree(paths) yield value')
        qry_cmd_Lst.append('return value')
        qry = ' '.join(qry_cmd_Lst)
        print(qry)

        ret_val = []
        with self.driver.session() as session:
            records = session.run(qry)
            print("neo_hook_scope_path_query QUERY DONE")

            for node in records:
                ret_val.extend(extract_table(node[0]))
            print(len(ret_val))
            session.close()
        if not ret_val[0]:
            ret_val.clear()
        data = ret_val

        sortedlist = []
        for item in data:
            if item not in sortedlist:
                sortedlist.append(item)
        data = sortedlist

        for dasda in data:
            print(dasda)

        header = self.headers[0]
        if data:
            data_by_id(data, self.databyid)
            self.model_tree_1.clear()
            self.model_tree_1.setHorizontalHeaderLabels(header)
            self.treeView_1.header().setDefaultSectionSize(400)
            print("Start import Data for Tree 1")
            self.importData(data, self.model_tree_1, header)
            self.treeView_1.setModel(self.model_tree_1)
            self.treeView_1.sortByColumn(0, QtCore.Qt.AscendingOrder)
            self.treeView_1.sortByColumn(1, QtCore.Qt.AscendingOrder)
            self.treeView_1.collapseAll()
            print("End import Data for Tree 1")
            q_header = self.treeView_1.header()
            for ii in range(1, len(header)):
                q_header.setSectionResizeMode(ii, QtWidgets.QHeaderView.ResizeToContents)
        else:
            self.model_tree_1.clear()
            print('Leeres Query Ergebnis')

    #######################################
    # Grundlegende Axen Gestaltung
    #######################################
    def set_chart_default_axis(self, axis: QCategoryAxis):
        labelFont = QFont("Bahnschrift Light Condensed")
        labelFont.setPixelSize(20)
        axis.setLabelsFont(labelFont)

        axisPen = QPen(Qt.white)
        axisPen.setWidth(2)
        axis.setLinePen(axisPen)

        axisBrush = QBrush(Qt.white)
        axis.setLabelsBrush(axisBrush)

    #######################################
    # Grundlegende Layout Gestaltung
    #######################################
    def set_chart_default_design(self, chart: QChart):
        chart.setTheme(QChart.ChartThemeBlueCerulean)

        # Hier kann man das Thema einstellen,
        # einfach hinter Theme anfangen zu schreiben,
        # die vorschläge kommen dann
        # im Anschluss habe ich noch Schriftarten/Farben/Größe usw. eingestellt

        font = QFont("Bahnschrift Light Condensed")
        font.setPixelSize(15)
        font.setBold(True)
        chart.setTitleFont(font)
        chart.setTitleBrush(QColor(Qt.white))

        legendFont = QFont("Bahnschrift Light Condensed")
        legendFont.setPixelSize(12)

        chart.legend().setFont(legendFont)
        chart.legend().setVisible(True)
        chart.legend().setAlignment(Qt.AlignTop)

    def get_node_dickt_of_item(self, qitem):
        qmodel = qitem.model()
        dbid = int(qmodel.itemFromIndex(qitem.index().siblingAtColumn(2)).text())
        if dbid in self.databyid:
            dicti = self.databyid[dbid]
        else:
            dicti = self.get_node_details_by_id(dbid)
            # print("get_node_details_by_id")
            self.databyid.update({dbid: dicti})
        return dicti

    def first_level_querie(self, qitem: QStandardItem):
        def add_all_childs(qitem):
            result_dicit = {}
            qmodel = qitem.model()
            if qitem.hasChildren():
                for ii in range(0, qitem.rowCount()):
                    child = qitem.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    title = str(qmodel.itemFromIndex(child.index().siblingAtColumn(0)).text())
                    node_type = str(qmodel.itemFromIndex(child.index().siblingAtColumn(1)).text())
                    temp = {}
                    temp.update({"Title": title})
                    temp.update({"Type": node_type})
                    temp.update({"QItem": child})
                    result_dicit.update({cid: temp})
            return result_dicit

        def neo_count_node_by_attribute_value(node_id, level_depth, HookNode_in, use_hooknode, node_dict: dict,
                                              respect_related: bool):
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + str(node_id))
            if respect_related:
                qry_cmd_Lst.append(
                    'MATCH path = (x)-[rr*0..' + str(level_depth) + ']->()-[]-(y:' + node_dict["_type"] + ')')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + str(level_depth) + ']->(y:' + node_dict["_type"] + ')')
            qry_cmd_Lst.append('WHERE ALL(r in rr')
            qry_cmd_Lst.append('WHERE type(r) = "has"')
            if use_hooknode:
                qry_cmd_Lst.append('AND r.HookNode = \"' + HookNode_in + '\"')
            qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
            if node_dict["attribute"]:
                qry_cmd_Lst.append('AND y.' + node_dict["attribute"] + ' =~ \'.*' + node_dict["value"] + '.*\'')
            qry_cmd_Lst.append('RETURN count(DISTINCT y) as count')
            # qry_cmd_Lst.append('RETURN count(y) as count')
            qry = ' '.join(qry_cmd_Lst)
            # print(qry)

            output = 0
            with self.driver.session() as session:
                records = session.run(qry)
                for record in records:
                    output = record["count"]
                session.close()
            return output

        def neo_sum_node_by_attribute_value(node_id, level_depth, HookNode_in, use_hooknode, node_dict: dict,
                                            respect_related: bool):
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (x) WHERE ID(x) = ' + str(node_id))
            if respect_related:
                qry_cmd_Lst.append(
                    'MATCH path = (x)-[rr*0..' + str(level_depth) + ']->()-[]-(y:' + node_dict["_type"] + ')')
            else:
                qry_cmd_Lst.append('MATCH path = (x)-[rr*0..' + str(level_depth) + ']->(y:' + node_dict["_type"] + ')')
            qry_cmd_Lst.append('WITH *, relationships(path) AS rr')
            qry_cmd_Lst.append('WHERE ALL(r in rr')
            qry_cmd_Lst.append('WHERE type(r) = "has"')
            if use_hooknode:
                qry_cmd_Lst.append('AND r.HookNode = \"' + HookNode_in + '\"')
            qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
            qry_cmd_Lst.append('WITH  DISTINCT y as z')  # Gibt einmalige Knoten zurück
            qry_cmd_Lst.append('RETURN  z.' + node_dict["attribute"] + ' as rst')
            # qry_cmd_Lst.append('RETURN count(y) as count')
            qry = ' '.join(qry_cmd_Lst)
            print(qry)

            output = 0.0
            with self.driver.session() as session:
                records = session.run(qry)
                for record in records:
                    try:
                        output += float(record["rst"])
                    except:
                        output += 0
                        print("ERROR to flaot", record["rst"])
                    print("Test Sum", output)
                session.close()
            return output

        def qitemmodel_count_end_node(qitem, node_dict):
            def count_node_types(qitem, node_dict: dict, match_list: list):
                if qitem.hasChildren():
                    for ii in range(0, qitem.rowCount()):
                        child = qitem.child(ii)
                        tmp_node_dict = self.get_node_dickt_of_item(child)
                        if tmp_node_dict["_type"] == node_dict["_type"]:
                            if node_dict["attribute"]:
                                if node_dict["value"] in tmp_node_dict[node_dict["attribute"]]:
                                    if child not in match_list:
                                        match_list.append(child)
                            else:
                                if child not in match_list:
                                    match_list.append(child)
                        count_node_types(child, node_dict, match_list)
                return match_list

            match_list = count_node_types(qitem, node_dict, [])
            return len(match_list)

        def pie_count_type(result_dicit, node_type, view_nr):
            series = QPieSeries()
            for k, v in result_dicit.items():
                if v[node_type] > 0:
                    print(v)
                    series.append(v["Title"], v[node_type])

            for slice in series.slices():
                slice.setLabel(str(float(slice.value()))[:5] + ": " + slice.label())
                slice.setExploded(False)
                slice.setLabelVisible(True)

            chart = QChart()
            # chart.legend().hide()
            chart.addSeries(series)
            chart.createDefaultAxes()
            chart.setAnimationOptions(QChart.SeriesAnimations)
            chart.setTitle("Anzahl " + node_type)

            chart.legend().setVisible(True)
            chart.legend().setAlignment(Qt.AlignBottom)

            # chartview = QChartView(chart)
            # chartview.setRenderHint(QPainter.Antialiasing)

            self.set_chart_default_design(chart)
            chart.setPlotAreaBackgroundVisible(True)
            if view_nr == 0:
                self.graphicsView.setChart(chart)
            elif view_nr == 1:
                self.graphicsView_2.setChart(chart)
            elif view_nr == 2:
                self.graphicsView_3.setChart(chart)
            else:
                self.graphicsView_4.setChart(chart)
            print(result_dicit)

        def bar_sum_type(result_dicit, node_type, view_nr):
            bar_set = QBarSet("New Content")
            axisX = QBarCategoryAxis()
            for k, v in result_dicit.items():
                if v[node_type] > 0:
                    print(v)
                    bar_set.append(v[node_type])
                    axisX.append(v["Title"])

            #
            # for slice in series.slices():
            #     slice.setLabel(str(float(slice.value()))[:5] + ": " + slice.label())
            #     slice.setExploded(False)
            #     slice.setLabelVisible(True)
            series = QBarSeries()
            series.append(bar_set)
            series.attachAxis(axisX)
            chart = QChart()
            # chart.legend().hide()
            chart.addSeries(series)
            chart.addAxis(axisX, Qt.AlignBottom)
            chart.createDefaultAxes()
            chart.setAnimationOptions(QChart.SeriesAnimations)
            chart.setTitle("Anzahl " + node_type)

            chart.legend().setVisible(True)
            chart.legend().setAlignment(Qt.AlignBottom)

            # chartview = QChartView(chart)
            # chartview.setRenderHint(QPainter.Antialiasing)

            self.set_chart_default_design(chart)
            chart.setPlotAreaBackgroundVisible(True)
            if view_nr == 0:
                self.graphicsView.setChart(chart)
            elif view_nr == 1:
                self.graphicsView_2.setChart(chart)
            elif view_nr == 2:
                self.graphicsView_3.setChart(chart)
            else:
                self.graphicsView_4.setChart(chart)
            print(result_dicit)

        qmodel = qitem.model()

        if qmodel == self.model_tree_1:
            HookNode = self.comboBox_Hook_Node_1.currentText()
            Scope_in = self.comboBox_multi_1.check_items()
            End_Node = self.comboBox_tree_1_Endnode.currentText()
            level_depth = self.comboBox_levels_1.currentText()
            respect_related = self.checkBox_respect_related_1.checkState()
        elif qmodel == self.model_tree_3:
            HookNode = self.comboBox_Hook_Node_3.currentText()
            Scope_in = self.comboBox_multi_3.check_items()
            End_Node = self.comboBox_tree_3_Endnode.currentText()
            level_depth = self.comboBox_levels_3.currentText()
            respect_related = self.checkBox_respect_related_2.checkState()

        empty_chart = QChart()
        self.graphicsView.setChart(empty_chart)
        self.graphicsView_2.setChart(empty_chart)
        self.graphicsView_3.setChart(empty_chart)
        self.graphicsView_4.setChart(empty_chart)

        is_querry_result = False
        if not qitem.parent():
            qitem = qitem.model().invisibleRootItem()
            is_querry_result = True
        else:
            qitem = qitem.parent()

        if End_Node == "NewContent":
            end_types = []
            end_types.append("NCRI")
            end_types.append("Change")

            result_dicit = add_all_childs(qitem)
            for end_type in end_types:
                tmp_dict = {}
                tmp_dict.update({"_type": "NewContent"})
                tmp_dict.update({"attribute": end_type})
                tmp_dict.update({"value": ""})
                for k, v in result_dicit.items():
                    count_res = neo_sum_node_by_attribute_value(k, 10, HookNode, True, tmp_dict, respect_related)
                    v.update({end_type: count_res})

            print("result_dicit", result_dicit)
            pie_count_type(result_dicit, "NCRI", 0)
            bar_sum_type(result_dicit, "NCRI", 1)
            pie_count_type(result_dicit, "Change", 2)
            bar_sum_type(result_dicit, "Change", 3)
            self.tabWidget.setCurrentIndex(1)

        elif End_Node == "Task":
            end_types = []
            end_types.append("Task")

            result_dicit = add_all_childs(qitem)
            for end_type in end_types:
                tmp_dict = {}
                tmp_dict.update({"_type": end_type})

                tmp_dict.update({"value": ""})
                for k, v in result_dicit.items():
                    tmp_dict.update({"attribute": "Effort"})
                    count_res = neo_sum_node_by_attribute_value(k, 10, HookNode, True, tmp_dict, respect_related)
                    tmp_dict.update({"attribute": "Original_Estimate"})
                    count_res += neo_sum_node_by_attribute_value(k, 10, HookNode, True, tmp_dict, respect_related)
                    v.update({end_type: count_res})

            print("result_dicit", result_dicit)
            pie_count_type(result_dicit, "Task", 0)
            bar_sum_type(result_dicit, "Task", 1)
            self.tabWidget.setCurrentIndex(1)

        elif End_Node == "*":
            end_types = []
            end_types.append("Requirement")
            end_types.append("TestCase")
            end_types.append("TestExecution")
            end_types.append("TestResult")

            result_dicit = add_all_childs(qitem)
            for end_type in end_types:
                tmp_dict = {}
                tmp_dict.update({"_type": end_type})
                tmp_dict.update({"attribute": ""})
                tmp_dict.update({"value": ""})
                for k, v in result_dicit.items():
                    if is_querry_result:
                        count_res = qitemmodel_count_end_node(v["QItem"], tmp_dict)
                    else:
                        count_res = neo_count_node_by_attribute_value(k, 10, HookNode, True, tmp_dict, respect_related)
                    v.update({end_type: count_res})

            print(result_dicit)

            pie_count_type(result_dicit, "Requirement", 0)
            pie_count_type(result_dicit, "TestCase", 0)
            pie_count_type(result_dicit, "TestExecution", 1)

            statussen = ["abgeschlossen", "gestartet", "offen", "Berichterstellung", "ausstehend",
                         "Merkmal wird nicht geprüft",
                         "Testobjekt nicht verfügbar", "keine Relevanz", "ungültig,abgeschlossen > Ergebnisdiskussion"]

            result_dicit_status = {}
            for tmp_status in statussen:
                tmp_dict = {}
                tmp_dict.update({"_type": "TestResult"})
                tmp_dict.update({"attribute": "Status"})
                tmp_dict.update({"value": tmp_status})

                tmp_dict_2 = {}
                tmp_dict_2.update({"Title": tmp_status})
                try:
                    dbid = int(qitem.model().itemFromIndex(qitem.index().siblingAtColumn(2)).text())
                    summe = neo_count_node_by_attribute_value(dbid, 10, HookNode, True, tmp_dict, respect_related)
                except:
                    summe = qitemmodel_count_end_node(qitem, tmp_dict)
                tmp_dict_2.update({"Count": summe})
                result_dicit_status.update({tmp_status: tmp_dict_2})

            print(result_dicit_status)
            pie_count_type(result_dicit_status, "Count", 2)

            resultussen = ['n/a', 'bedingt i.O. / limited good', 'nicht bestanden / failed', 'bestanden / passed']
            # resultussen = ['bedingt', 'nicht bestanden', 'bestanden']
            result_dicit_results = {}
            for tmp_status in resultussen:
                tmp_dict = {}
                tmp_dict.update({"_type": "TestResult"})
                tmp_dict.update({"attribute": "Result"})
                tmp_dict.update({"value": tmp_status})

                tmp_dict_2 = {}
                tmp_dict_2.update({"Title": tmp_status})
                try:
                    dbid = int(qitem.model().itemFromIndex(qitem.index().siblingAtColumn(2)).text())
                    summe = neo_count_node_by_attribute_value(dbid, 10, HookNode, True, tmp_dict, respect_related)
                    print("hier")
                except:
                    summe = 0
                    for k, v in result_dicit.items():
                        summe += neo_count_node_by_attribute_value(k, 10, HookNode, True, tmp_dict, respect_related)
                tmp_dict_2.update({"Count": summe})
                result_dicit_results.update({tmp_status: tmp_dict_2})
            print(result_dicit_results)
            pie_count_type(result_dicit_results, "Count", 3)

            self.tabWidget.setCurrentIndex(1)

        else:
            end_types = []
            end_types.append(End_Node)

            result_dicit = add_all_childs(qitem)
            for end_type in end_types:
                tmp_dict = {}
                tmp_dict.update({"_type": end_type})
                tmp_dict.update({"attribute": ""})
                tmp_dict.update({"value": ""})
                for k, v in result_dicit.items():
                    if is_querry_result:
                        count_res = qitemmodel_count_end_node(v["QItem"], tmp_dict)
                    else:
                        count_res = neo_count_node_by_attribute_value(k, 10, HookNode, True, tmp_dict, respect_related)
                    v.update({end_type: count_res})

            print(result_dicit)

            pie_count_type(result_dicit, End_Node, 0)
            bar_sum_type(result_dicit, End_Node, 1)
            self.tabWidget.setCurrentIndex(1)

    def sync_folder_with_os(self, node_id_in):
        def get_neo_paths_list(node_dicti, HookNode, max_depth):
            qry_cmd_Lst = []
            qry_cmd_Lst.append('MATCH (start_n) WHERE start_n.Key = \"' + node_dicti["Key"] + '\"')
            qry_cmd_Lst.append(
                'MATCH (parent_n) WHERE parent_n.DeepLink STARTS WITH \"' + node_dicti["DeepLink"] + '\"')
            qry_cmd_Lst.append('MATCH path = (start_n)-[*0..' + str(max_depth + 1) + ']->(parent_n)')
            qry_cmd_Lst.append('WITH parent_n, relationships(path) AS rr')
            qry_cmd_Lst.append('WHERE ALL(r in rr')
            qry_cmd_Lst.append('WHERE r.HookNode = \"' + HookNode + '\"')
            # qry_cmd_Lst.append('AND type(r) = "has"')
            qry_cmd_Lst.append(')')  # Klammer wird weiter oben geöffnet
            qry_cmd_Lst.append('RETURN parent_n.DeepLink as DeepLink, parent_n.Key as Key, parent_n.ModifyDate as mDate')
            qry = ' '.join(qry_cmd_Lst)
            # print(qry)

            sync_paths_dict = {}
            with self.driver.session() as session:
                records = session.run(qry)

                for record in records:
                    sync_paths_dict.update({record["DeepLink"]: record["Key"]})
            session.close()
            return sync_paths_dict

        def get_os_paths_list_old(node_dicti):
            paths = []
            if node_dicti['DeepLink']:
                start_folder = node_dicti['DeepLink']
                for dirName, subdirs, fileList in os.walk(start_folder):
                    if not "$RECYCLE.BIN" in dirName:
                        dirName = dirName.replace("\\", "/")
                        if dirName not in paths:
                            paths.append(dirName)
                        for filename in fileList:
                            filepath = os.path.join(dirName, filename)
                            filepath = os.path.normpath(filepath).replace("\\", "/")
                            if filepath not in paths:
                                paths.append(filepath)
            return paths

        def get_os_paths_list(arr, path_in, depth, max_depth):
            try:
                tmp_paths_lst = os.listdir(path_in)
                for path in tmp_paths_lst:
                    filepath = os.path.join(path_in, path)
                    filepath = os.path.normpath(filepath).replace("\\", "/")
                    if "$RECYCLE.BIN" not in filepath:
                        if filepath not in arr:
                            arr.append(filepath)

                            if os.path.isdir(filepath):
                                if depth < max_depth:
                                    new_depth = depth + 1
                                    get_os_paths_list(arr, filepath, new_depth, max_depth)
                return arr
            except:
                return arr

        def create_sync_neo_item(path_in):
            if os.path.isdir(path_in):
                is_dir = True
            else:
                is_dir = False

            qry_cmd_Lst = []
            if is_dir:
                qry_cmd_Lst.append('MERGE (child_n:Folder {')

            else:
                qry_cmd_Lst.append('MERGE (child_n:File {')
            qry_cmd_Lst.append('Sync: \"' + "False" + '\", ')
            qry_cmd_Lst.append('Title: \"' + os.path.basename(path_in) + '\", ')
            qry_cmd_Lst.append('DeepLink: \"' + path_in + '\", ')
            qry_cmd_Lst.append('Details: \"' + "" + '\", ')
            qry_cmd_Lst.append('Description: \"' + "" + '\", ')
            qry_cmd_Lst.append('Key: apoc.create.uuid() })')
            qry_cmd_Lst.append('return child_n.Key as Key')
            qry = ' '.join(qry_cmd_Lst)
            # print(qry)

            with self.driver.session() as session:
                records = session.run(qry)
                for record in records:
                    # print(records)
                    new_neo_key = record["Key"]
            session.close()
            return new_neo_key

        def create_sync_neo_link(new_neo_nodes_lst, HookNode):
            def get_key_of_path(path_in, new_neo_nodes_lst):
                for node in new_neo_nodes_lst:
                    if node["DeepLink"] == path_in:
                        return node["Key"]

            for node in new_neo_nodes_lst:
                parent_path = os.path.dirname(node["DeepLink"])
                parent_key = ""
                parent_key = get_key_of_path(parent_path, new_neo_nodes_lst)
                child_key = node["Key"]

                if not parent_key:
                    qry_cmd_Lst = []
                    qry_cmd_Lst.append('MATCH (hocknode:HookNode) WHERE hocknode.Title = \"' + HookNode + '\"')
                    qry_cmd_Lst.append('MATCH (parent) WHERE parent.DeepLink = \"' + parent_path + '\"')
                    qry_cmd_Lst.append('MATCH path = (hocknode)-[*0..25]->(parent)')
                    qry_cmd_Lst.append('return parent.Key as rst')
                    qry = ' '.join(qry_cmd_Lst)
                    # print(qry)

                    sync_paths_list = []
                    with self.driver.session() as session:
                        records = session.run(qry)
                        for record in records:
                            parent_key = record["rst"]
                    session.close()

                if parent_key:
                    qry_cmd_Lst = []
                    qry_cmd_Lst.append('MATCH (child) WHERE child.Key = \"' + child_key + '\"')
                    qry_cmd_Lst.append('MATCH (parent) WHERE parent.Key = \"' + parent_key + '\"')
                    qry_cmd_Lst.append('MERGE (parent)-[:has {HookNode: \"' + HookNode + '\"}]->(child)')
                    qry = ' '.join(qry_cmd_Lst)
                    # print(qry)

                    with self.driver.session() as session:
                        records = session.run(qry)
                    session.close()
            return True

        HookNode = self.comboBox_Hook_Node_1.currentText()
        node_dicti = self.get_node_details_by_id(node_id_in)
        if "Sync" in node_dicti.keys():
            sync_cmd = node_dicti["Sync"]
            if not sync_cmd == "False":
                try:
                    max_depth = int(sync_cmd)
                    exe_sync = True
                except:
                    exe_sync = False
            else:
                exe_sync = False
        else:
            exe_sync = False

        if exe_sync:
            neo_paths_list = get_neo_paths_list(node_dicti, HookNode, max_depth)
            os_paths_list = []
            get_os_paths_list(os_paths_list, node_dicti["DeepLink"], 0, max_depth)

            print(node_dicti)
            print("neo_paths_list with length", len(neo_paths_list))# , neo_paths_list)
            print("os_paths_list with length", len(os_paths_list))# , os_paths_list)

            new_neo_nodes_lst = []
            new_neo_nodes_lst.append(node_dicti)
            print("Start creating new nodes")
            for os_path in os_paths_list:
                if os_path not in neo_paths_list.keys():
                    tmp = {}
                    new_neo_node_keys = create_sync_neo_item(os_path)
                    tmp.update({"Key": new_neo_node_keys})
                    tmp.update({"DeepLink": os_path})
                    new_neo_nodes_lst.append(tmp)

            print("Start to link new notes")
            create_sync_neo_link(new_neo_nodes_lst, HookNode)

            print("Start to compare folders and file info")
            for neo_path, neo_key in neo_paths_list.items():
                if neo_path == node_dicti["DeepLink"]:
                    continue
                if neo_path not in os_paths_list:
                    qry_cmd_Lst = []
                    qry_cmd_Lst.append('MATCH (child) WHERE child.Key = \"' + neo_key + '\"')
                    qry_cmd_Lst.append('SET child.Sync = "ERROR"')
                    qry = ' '.join(qry_cmd_Lst)
                    # print(qry)

                    with self.driver.session() as session:
                        records = session.run(qry)
                    session.close()
                else:
                    try:
                        mdate = str(datetime.datetime.fromtimestamp(os.path.getmtime(neo_path)))
                    except:
                        mdate = ""
                    qry_cmd_Lst = []
                    qry_cmd_Lst.append('MATCH (child) WHERE child.Key = \"' + neo_key + '\"')
                    qry_cmd_Lst.append('SET child.ModifyDate = \"' + mdate + '\"')
                    if os.path.isdir(neo_path):
                        qry_cmd_Lst.append('SET child.Sync = "1"')
                    else:
                        qry_cmd_Lst.append('SET child.Sync = "False"')
                    qry = ' '.join(qry_cmd_Lst)
                    # print(qry)

                    with self.driver.session() as session:
                        records = session.run(qry)
                    session.close()

        else:
            print("Befehl nicht für diesen Knoten ausführbar. Vielleicht setzte Sync auf True")

    def calculate_bottom_up_dates(self, qitem):
        def add_all_childs(result_dicit, qitem):
            qmodel = qitem.model()
            if qitem.hasChildren():
                for ii in range(0, qitem.rowCount()):
                    child = qitem.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    title = str(qmodel.itemFromIndex(child.index().siblingAtColumn(0)).text())
                    node_type = str(qmodel.itemFromIndex(child.index().siblingAtColumn(1)).text())
                    temp = {}
                    temp.update({"Title": title})
                    temp.update({"Type": node_type})
                    temp.update({"QItem": child})
                    result_dicit.update({cid: temp})
                    if child.hasChildren():
                        add_all_childs(result_dicit, child)
            return result_dicit

        def neo_get_start_due_date(nid, HookNode: str):
            qry = 'MATCH (x) WHERE ID(x) = ' + str(nid) + ' '\
                  'MATCH path = (x)-[*0..3]->(y:Task) ' \
                  'WITH path, relationships(path) AS rr,x ,y WHERE ALL(r in rr WHERE r.HookNode = \"' + HookNode + '\" ) ' \
                  'return y.Due_Date as rst'

            all_dates = []
            with self.driver.session() as session:
                records = session.run(qry)

                for record in records:
                    date_time_str = record["rst"]
                    # print(date_time_str)
                    if date_time_str:
                        date_time_obj = datetime.datetime.strptime(date_time_str, '%d.%m.%Y %H:%M')
                        all_dates.append(date_time_obj)
            session.close()
            if all_dates:
                rst = {}
                rst.update({"Start_Date": min(all_dates)})
                rst.update({"Due_Date": max(all_dates)})
                return rst
            else:
                return False

        HookNode = self.comboBox_Hook_Node_1.currentText()

        qitem = qitem.parent()
        result_dicit = {}
        result_dicit = add_all_childs(result_dicit, qitem)

        for k, v in result_dicit.items():
            rst = neo_get_start_due_date(k, HookNode)
            # print(v["Title"], rst)
            #
            # print(v)
            if rst:
                Value = {}
                Value.update({"_id": k})
                Value.update({"hooknode": HookNode})
                Value.update({"Start_Date": rst["Start_Date"].strftime("%d.%m.%Y %H:%M")})
                Value.update({"Due_Date": rst["Due_Date"].strftime("%d.%m.%Y %H:%M")})
                # print(Value)

                self.neo_edit_node(Value)

    def plot_gantt_chart(self, qitem):
        def add_all_childs(result_dicit, qitem, n_type: str, recursion: bool):
            qmodel = qitem.model()
            if qitem.hasChildren():
                for ii in range(0, qitem.rowCount()):
                    child = qitem.child(ii)
                    cid = int(qmodel.itemFromIndex(child.index().siblingAtColumn(2)).text())
                    title = str(qmodel.itemFromIndex(child.index().siblingAtColumn(0)).text())
                    node_type = str(qmodel.itemFromIndex(child.index().siblingAtColumn(1)).text())

                    if (n_type == node_type) or (n_type == "*"):
                        dicti = self.get_node_dickt_of_item(child)
                        dicti.update({"QItem": child})
                        result_dicit.update({cid: dicti})
                    if child.hasChildren() and recursion:
                        add_all_childs(result_dicit, child, n_type, recursion)
            return result_dicit


        qitem = qitem.parent()
        first_level_dicit = {}
        first_level_dicit = add_all_childs(first_level_dicit, qitem, "*", False)

        print(first_level_dicit)
        list_to_print = []
        for k, v in first_level_dicit.items():
            print(v["Title"])
            tmp_dicti = {}
            tmp_dicti =  add_all_childs(tmp_dicti, v["QItem"], "Task", True)
            # print(tmp_dicti)

            first_run = True
            for key, node in tmp_dicti.items():
                # print(node)
                if "Due_Date" in node.keys():
                    due_date_timestamp = datetime.datetime.strptime(node["Due_Date"], "%d.%m.%Y %H:%M").timestamp()

                    if first_run:
                        total_min = due_date_timestamp
                        total_max = due_date_timestamp
                        first_run = False
                    # print("huhu")
                    if due_date_timestamp < total_min:
                        total_min = due_date_timestamp
                    if due_date_timestamp > total_max:
                        total_max = due_date_timestamp
            if not first_run:
                tmp = {}
                tmp.update({"Title": v["Title"]})
                tmp.update({"Start": total_min})
                tmp.update({"End": total_max})
                tmp.update({"Length": total_max - total_min})
                tmp.update({"_type": v["_type"]})
                list_to_print.append(tmp)
                print(tmp)

        list_of_timestamps = []
        for tmp in list_to_print:
            if tmp["Start"] not in list_of_timestamps:
                list_of_timestamps.append(tmp["Start"])
            if tmp["End"] not in list_of_timestamps:
                list_of_timestamps.append(tmp["End"])

        if list_of_timestamps:
            total_max = max(list_of_timestamps)
            total_min = min(list_of_timestamps)


            # Declaring a figure "gnt"
            fig, gnt = plt.subplots()
            figManager = plt.get_current_fig_manager()
            figManager.window.showMaximized()

            # Setting Y-axis limits
            gnt.set_ylim(0, len(list_to_print) * 10)

            # Setting X-axis limits
            gnt.set_xlim(total_min, total_max)

            # Setting labels for x-axis and y-axis
            gnt.set_xlabel('seconds since start')
            gnt.set_ylabel('Processor')

            # Setting ticks on y-axis
            # gnt.set_yticks([15, 25, 35])
            # Labelling tickes of y-axis
            xlabels = []
            xticks = []
            len_ticks = 5
            for ii in range(0, len_ticks + 1):
                tick_stamp = (total_max - total_min) / len_ticks * ii + total_min
                xlabels.append(datetime.datetime.fromtimestamp(tick_stamp).strftime("%d.%m.%Y"))
                xticks.append(tick_stamp)

            ylabels = []
            yticks = []
            ytick = 0
            for tmp_item in list_to_print:
                ylabels.append(tmp_item["Title"])
                yticks.append(ytick)
                ytick += 10

            gnt.set_yticklabels(ylabels)
            gnt.set_yticks(yticks)
            gnt.set_xticklabels(xlabels)
            gnt.set_xticks(xticks)

            # Setting graph attribute
            gnt.grid(True)

            # "https://www.geeksforgeeks.org/python-basic-gantt-chart-using-matplotlib/"
            for ii, tmp_item in enumerate(list_to_print):
                # Declaring a bar in schedule
                if tmp_item["Length"] == 0:
                    tmp_item["Length"] = 10

                if tmp_item["_type"] == "Epic":
                    gnt.broken_barh([(tmp_item["Start"], tmp_item["Length"])], (ii*10-5, 9), facecolors='orange')
                elif tmp_item["_type"] == "Feature":
                    gnt.broken_barh([(tmp_item["Start"], tmp_item["Length"])], (ii*10-5, 9), facecolors='blue')
                elif tmp_item["_type"] == "User_Story":
                    gnt.broken_barh([(tmp_item["Start"], tmp_item["Length"])], (ii*10-5, 9), facecolors='orange')
                else:
                    gnt.broken_barh([(tmp_item["Start"], tmp_item["Length"])], (ii * 10 - 5, 9), facecolors='orange')

            fig.tight_layout()
            figManager = plt.get_current_fig_manager()
            figManager.window.showMaximized()
            plt.show()
            plt.savefig("gantt1.png")




if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
