#!/usr/bin/env python3
# Tool Database Manager for qtvcp
# Copyright (c) 2022  Jim Sloot <persei802@gmail.com>
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.

import sys
import os

from PyQt5 import QtCore, QtWidgets, QtSql, QtGui, QtPrintSupport
from PyQt5.QtWidgets import (QTabWidget, QTableView, QAbstractItemView, QSpinBox, QDoubleSpinBox, QFileDialog,
                             QStyledItemDelegate, QAbstractItemDelegate, QItemEditorFactory, QInputDialog)
from PyQt5.QtCore import Qt, QVariant, QSortFilterProxyModel
from PyQt5.QtGui import (QColor, QFont, QTextDocument, QTextCursor, QTextCharFormat, QTextTableFormat,
                         QTextFrameFormat, QTextFormat)
from PyQt5.QtPrintSupport import QPrinter

from qtvcp import logger
from qtvcp.core import Status, Info, Action, Path, Tool

STATUS = Status()
INFO = Info()
ACTION = Action()
TOOL = Tool()
PATH = Path()
LOG = logger.getLogger(__name__)
LOG.setLevel(logger.DEBUG) # One of DEBUG, INFO, WARNING, ERROR, CRITICAL

VERSION = '1.0.1'

offset_headers = {}
tool_headers = {}


class ItemEditorFactory(QItemEditorFactory):
    def __init__(self):
        super(ItemEditorFactory, self).__init__()

    def createEditor(self, userType, parent):
        if userType == QVariant.Double:
            doubleSpinBox = QDoubleSpinBox(parent)
            doubleSpinBox.setDecimals(4)
            doubleSpinBox.setMaximum(99999)
            doubleSpinBox.setMinimum(-99999)
            return doubleSpinBox
        elif userType == QVariant.Int:
            spinBox = QSpinBox(parent)
            spinBox.setMaximum(20000)
            spinBox.setMinimum(0)
            return spinBox
        else:
            return super(ItemEditorFactory,self).createEditor(userType, parent)


class MyOffsetModel(QtSql.QSqlTableModel):
    def __init__(self, parent=None):
        super(MyOffsetModel, self).__init__(parent)
        self.parent = parent
        self.setObjectName('offset_model')
        self.setTable('offsets')
        self.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.metric_display = True
        self.highlight_color = '#20A0A0'
        self.tool_list = []
        rec = self.record()
        for i in range(rec.count()):
            hdr = rec.fieldName(i)
            self.setHeaderData(i, Qt.Horizontal, hdr)
            if i > 1: offset_headers[hdr] = i
        self.select()
        self.update()

    def data(self, index, role=Qt.DisplayRole):
        value = super(MyOffsetModel, self).data(index, role)
        if index.column() == 1:
            if role == Qt.BackgroundRole:
                return QColor('#809090')
            if role == Qt.CheckStateRole:
                checked = super(MyOffsetModel, self).data(index, Qt.DisplayRole)
                if int(checked) == 1: return Qt.Checked
                else: return Qt.Unchecked
            else: return QVariant()
        if role == Qt.EditRole:
            return value
        if role == Qt.DisplayRole:
            if isinstance(value, float):
                if self.metric_display:
                    return "{:.3f}".format(value)
                else: return "{:.4f}".format(value)
            return value
        if role == Qt.TextAlignmentRole:
            if index.column() == offset_headers['Comment']:
                return (Qt.AlignLeft | Qt.AlignVCenter)
            else:
                return Qt.AlignCenter
        if role == Qt.BackgroundRole:
            if self.parent.current_row is None: return QVariant()
            if index.row() == self.parent.current_row:
                return QColor(self.highlight_color)
        return QVariant()

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid(): return False
        if role == Qt.CheckStateRole and index.column() == 1:
            if value == Qt.Checked:
                self.uncheck_all_tools()
                return self.setData(index, 1)
            elif value == Qt.Unchecked:
                return self.setData(index, 0)
        if role == Qt.EditRole:
            return super(MyOffsetModel, self).setData(index, value, role)
        return super(MyOffsetModel, self).setData(index, value, role)

    def rowCount(self, parent):
        return QtSql.QSqlTableModel.rowCount(self)

    def flags(self, index):
        if not index.isValid():
            return None
        if index.column() == 1:
#            return Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable | Qt.ItemIsUserCheckable
            return Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsUserCheckable
        return Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def sort(self, column, order):
        if column != 1: super(MyOffsetModel, self).sort(column, order)

    def list_checked_tools(self):
        checked_list = []
        for row in range(super().rowCount()):
            idx = self.index(row, 1)
            checked = super().data(idx)
            if checked == 1:
                idx = self.index(row, offset_headers['Tool'])
                tno = super().data(idx)
                checked_list.append(tno)
        return checked_list

    def update(self):
        LOG.debug("Updating offset model")
        self.tool_list = []
        tool_array, wear_model = TOOL.GET_TOOL_MODELS()
        for line in tool_array:
            tno = line[0]
            self.tool_list.append(tno)
            # look for lines to add
            row = self.get_index(tno)
            if row is None:
                self.addrow(line)
            else:
                self.update_row(row, line)
        # look for lines to delete
        delete_list = []
        col = offset_headers['Tool']
        for row in range(self.rowCount(self)):
            idx = self.index(row, col)
            tno = self.data(idx)
            if tno not in self.tool_list:
                delete_list.append(tno)
        LOG.debug(f"Tools to delete {delete_list}")
        if delete_list:
            if len(delete_list) > 1: delete_list.reverse()
            for tno in delete_list:
                self.delrow(tno)
        return True

    def addrow(self, data):
        # add row to offset table
        row = self.rowCount(self)
        rec = self.record()
        for i, key in enumerate(offset_headers):
            rec.setValue(key, data[i])
        rec.setGenerated('idn', False)
        if not self.insertRecord(row, rec): LOG.debug(f"Error: {self.lastError().text()}")
        pkey = self.data(self.index(row, 0))
        tno = data[0]
        # add row to tool table
        self.parent.tool_model.addrow(pkey, tno)

    def update_row(self, row, data):
        for i, key in enumerate(offset_headers):
            idx = self.index(row, offset_headers[key])
            self.setData(idx, data[i])

    def delrow(self, tno):
        row = self.get_index(tno)
        if row is None: return
        LOG.debug(f"Deleting row {row}")
        self.removeRow(row)
        self.select()
        self.parent.tool_model.removeRow(row)
        self.parent.tool_model.select()

    def get_index(self, tno):
        col = offset_headers['Tool']
        count = self.rowCount(self)
        found = None
        for row in range(count):
            idx = self.index(row, col)
            if self.data(idx) == tno:
                found = row
                break
        return found

    def uncheck_all_tools(self):
        rows = super().rowCount()
        for row in range(rows):
            idx = self.index(row, 1)
            if super().data(idx) == 1:
                super().setData(idx, 0)

class MyToolModel(QtSql.QSqlTableModel):
    def __init__(self, parent=None):
        super(MyToolModel, self).__init__(parent)
        self.parent = parent
        self.setObjectName('tool_model')
        self.setTable('tools')
        self.setEditStrategy(QtSql.QSqlTableModel.OnFieldChange)
        self.metric_display = True
        self.highlight_color = '#20A0A0'
        rec = self.record()
        for i in range(rec.count()):
            hdr = rec.fieldName(i)
            self.setHeaderData(i, Qt.Horizontal, hdr)
            tool_headers[hdr] = i
        self.select()

    def data(self, index, role=Qt.DisplayRole):
        value = super(MyToolModel, self).data(index, role)
        if role == Qt.EditRole:
            return value
        if role == Qt.DisplayRole:
            if isinstance(value, float):
                if self.metric_display:
                    return "{:.3f}".format(value)
                else: return "{:.4f}".format(value)
            return value
        if role == Qt.BackgroundRole and index.column() == tool_headers['TOOL']:
            checked = self.parent.get_checked_tools()
            if checked and index.data() == checked[0]:
                return QColor(self.highlight_color)
            else: return QVariant()
        if role == Qt.TextAlignmentRole:
            if index.column() == tool_headers['ICON']: return (Qt.AlignLeft | Qt.AlignVCenter)
            else: return Qt.AlignCenter

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid(): return False
        if role == Qt.EditRole:
            return super(MyToolModel, self).setData(index, value, role)
        return True

    def rowCount(self, parent):
        return QtSql.QSqlTableModel.rowCount(self)

    def flags(self, index):
        if not index.isValid():
            return None
        field = self.record().fieldName(index.column())
        if field in ['TOOL', 'TIME', 'ICON']:
            return Qt.ItemIsEditable | Qt.ItemIsEnabled
        else:
            return Qt.ItemIsEditable | Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def update_tool_no(self, old_tno, new_tno):
        row = self.get_index(old_tno)
        col = tool_headers['TOOL']
        idx = self.index(row, col)
        self.setData(idx, new_tno)

    def get_index(self, tno):
        col = tool_headers['TOOL']
        count = self.rowCount(self)
        found = None
        for row in range(count):
            idx = self.index(row, col)
            if self.data(idx) == tno:
                found = row
                break
        return found

    def addrow(self, pkey, tno):
        row = self.rowCount(self)
        rec = self.record()
        rec.setValue('idn', pkey)
        rec.setValue('TOOL', tno)
        rec.setValue('TIME', 0.0)
        if not self.insertRecord(row, rec): LOG.debug(f"Error: {self.lastError().text()}")
        

class Tool_Database(QTabWidget):
    def __init__(self, parent=None):
        super(Tool_Database, self).__init__(parent)
        self.database = os.path.join(PATH.CONFIGPATH, 'tool_database.db')
        self.tables = []
        self.query = None
        self.enable_edit = False
        self.current_row = 0
        self.timer_dict = {'running': False, 'tool': 0, 'time': 0}
        self.timer_tenths = 0
        self.next_available = 0
        self.axis_list = INFO.AVAILABLE_AXES
        self.hdr_list = []    
        self.styledItemDelegate = QStyledItemDelegate()
        self.styledItemDelegate.setItemEditorFactory(ItemEditorFactory())

        if not self.create_connection():
            return None
        if 'offsets' not in self.tables:
            LOG.debug("Creating offsets table")
            self.create_offset_table()
        if 'tools' not in self.tables:
            LOG.debug("Creating tools table")
            self.create_tool_table()
 
        LOG.debug("Database tables: {}".format(self.tables))
        self.tool_model = MyToolModel(self)
        self.offset_model = MyOffsetModel(self)
        self.offset_view = QTableView()
        self.tool_view = QTableView()
        self.proxyModel = QSortFilterProxyModel()
        self.proxyModel.setSourceModel(self.offset_model)

        self.init_tool_view()
        self.init_offset_view()
        self.addTab(self.offset_view, 'OFFSETS')
        self.addTab(self.tool_view, 'TOOLS')

    def hal_init(self):
        STATUS.connect('all-homed', lambda w: self.setEnabled(True))
        STATUS.connect('not-all-homed', lambda w, axis: self.setEnabled(False))
        STATUS.connect('periodic', lambda w: self.tool_timer())
        STATUS.connect('interp-idle', lambda w: self.interp_state(False))
        STATUS.connect('interp-run', lambda w: self.interp_state(True))
        STATUS.connect('tool-in-spindle-changed', lambda w, tool: self.tool_changed(tool))

    def create_connection(self):
        db = QtSql.QSqlDatabase.addDatabase('QSQLITE')
        db.setDatabaseName(self.database)
        self.query = QtSql.QSqlQuery(db)
        if not db.open():
            LOG.debug("Database Error: {}".format(db.lastError().databaseText()))
            return False
        self.tables = db.tables()
        return True

    def create_offset_table(self):
        if self.query.exec_(
            '''
            CREATE TABLE offsets (
                idn INTEGER PRIMARY KEY AUTOINCREMENT UNIQUE NOT NULL,
                Chk INT DEFAULT 0,
                Tool INT,
                Pocket INT,
                X REAL DEFAULT 0.0,
                Y REAL DEFAULT 0.0,
                Z REAL DEFAULT 0.0,
                A REAL DEFAULT 0.0,
                B REAL DEFAULT 0.0,
                C REAL DEFAULT 0.0,
                U REAL DEFAULT 0.0,
                V REAL DEFAULT 0.0,
                W REAL DEFAULT 0.0,
                Diameter REAL DEFAULT 0.0,
                I REAL DEFAULT 0.0,
                J REAL DEFAULT 0.0,
                Q INT DEFAULT 0,
                Comment TEXT DEFAULT "new tool"
            )
            '''
        ) is True:
            LOG.debug("Create offsets success")
        else:
            LOG.debug("Create offsets error: {}".format(self.query.lastError().text()))

    def create_tool_table(self):
        if self.query.exec_(
            '''
            CREATE TABLE tools (
                idn INTEGER PRIMARY KEY,
                TOOL INTEGER,
                TIME REAL DEFAULT 0.0,
                RPM INTEGER DEFAULT 0,
                CPT REAL DEFAULT 0.0,
                LENGTH REAL DEFAULT 0.0,
                FLUTES INTEGER DEFAULT 0,
                FEED INTEGER DEFAULT 0,
                MFG TEXT,
                ICON TEXT DEFAULT "not_found.png"
            )
            '''
        ) is True:
            LOG.debug("Create tool table success")
        else:
            LOG.debug("Create tool table error: {}".format(self.query.lastError().text()))

    def init_tool_view(self):
        self.tool_view.setModel(self.tool_model)
        self.tool_view.hideColumn(0)
        self.tool_view.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.tool_view.horizontalHeader().setSortIndicator(1,Qt.AscendingOrder)
        self.tool_view.horizontalHeader().setMinimumSectionSize(70)
        self.tool_view.resizeColumnsToContents()
        self.tool_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.tool_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.tool_view.setCornerButtonEnabled(False)
        self.tool_view.horizontalHeader().setStretchLastSection(True)
        self.tool_view.verticalHeader().hide()
        self.tool_view.setSortingEnabled(True)
        self.tool_view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.tool_view.setItemDelegate(self.styledItemDelegate)
        self.tool_view.clicked.connect(self.showToolSelection)

    def init_offset_view(self):
        self.offset_view.setModel(self.offset_model)
        self.offset_view.hideColumn(0)
        self.offset_view.horizontalHeader().setDefaultAlignment(Qt.AlignCenter)
        self.offset_view.horizontalHeader().setSortIndicator(1,Qt.AscendingOrder)
        self.offset_view.horizontalHeader().setMinimumSectionSize(70)
        self.offset_view.resizeColumnsToContents()
        self.offset_view.setSelectionMode(QAbstractItemView.SingleSelection)
        self.offset_view.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.offset_view.setCornerButtonEnabled(False)
        self.offset_view.horizontalHeader().setStretchLastSection(True)
        self.offset_view.verticalHeader().hide()
        self.offset_view.setSortingEnabled(True)
        self.offset_view.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.hdr_list = ['Chk', 'Tool', 'Pocket'] + self.axis_list + ['Diameter', 'Comment']
        for hdr in offset_headers:
            if hdr not in self.hdr_list:
                self.offset_view.hideColumn(offset_headers[hdr])
        self.offset_view.setItemDelegate(self.styledItemDelegate)
        self.offset_view.clicked.connect(self.showOffsetSelection)

    def start_tool_timer(self):
        tno = self.timer_dict['tool']
        if tno == 0: return
        LOG.debug(f"Starting timer for tool {tno}")
        self.timer_dict['running'] = True
        row = self.offset_model.get_index(tno)
        if row is None: return
        idx = self.tool_model.index(row, tool_headers['TIME'])
        pre_time = float(self.tool_model.data(idx))
        self.timer_dict['time'] = int(pre_time * 60)

    def stop_tool_timer(self):
        if not self.timer_dict['running']: return
        tno = self.timer_dict['tool']
        LOG.debug(f"Stopping timer for tool {tno}")
        self.timer_dict['running'] = False
        row = self.offset_model.get_index(tno)
        if row is None: return
        total_time = self.timer_dict['time'] / 60
        total_time = f"{total_time:.3f}"
        idx = self.tool_model.index(row, tool_headers['TIME'])
        self.tool_model.setData(idx, total_time)

## callbacks from STATUS
    def tool_changed(self, tool):
        self.current_row = None if tool == 0 else self.offset_model.get_index(tool)
        if tool > 0:
            self.timer_dict['tool'] = tool
            if STATUS.is_auto_running():
                self.start_tool_timer()

    def interp_state(self, state):
        if not STATUS.is_auto_mode(): return
        if state:
            self.start_tool_timer()
        else:
            self.stop_tool_timer()

    def data_changed(self):
        # update linuxcnc to the changed data
        LOG.debug("Saving tool table to file")
        try:
            if STATUS.is_status_valid():
                array = self.save_table()
                error = TOOL.SAVE_TOOLFILE(array)
                if error:
                    raise
                ACTION.RECORD_CURRENT_MODE()
                ACTION.CALL_MDI('g43')
                ACTION.RESTORE_RECORDED_MODE()
                STATUS.emit('reload-display')
        except Exception as e:
            LOG.exception("offsetpage widget error: MDI call error", exc_info=e)

## callbacks from widgets
    def showOffsetSelection(self, item):
        if self.enable_edit is False: return
        col = item.column()
        field = self.offset_model.record().fieldName(col)
        if field == 'Chk':
            pass
        elif field in offset_headers:
            self.callOffsetDialog(item, field)

    def showToolSelection(self, item):
        if self.enable_edit is False: return
        col = item.column()
        field = self.tool_model.record().fieldName(col)
        if field in ['TOOL', 'TIME', 'ICON']: return
        elif field in tool_headers:
            self.callToolDialog(item, field)

    def callToolDialog(self, item, field):
        idx = self.offset_model.index(item.row(), offset_headers['Tool'])
        tool = self.offset_model.data(idx)
        idx = self.tool_model.index(item.row(), tool_headers[field])
        header = f'Tool {tool} Data'
        preload = self.tool_model.data(idx)
        if field == 'RPM':
            ret_val, ok = QInputDialog.getInt(self, header, field, int(preload), 0, 24000, 100)
            if ok: self.tool_model.setData(idx, ret_val)
        elif field == 'FLUTES':
            ret_val, ok = QInputDialog.getInt(self, header, field, int(preload), 0, 8, 1)
            if ok: self.tool_model.setData(idx, ret_val)
        elif field == 'FEED':
            ret_val, ok = QInputDialog.getInt(self, header, field, int(preload), 0, 6000, 100)
            if ok: self.tool_model.setData(idx, ret_val)
        elif field == 'CPT':
            ret_val, ok = QInputDialog.getDouble(self, header, field, float(preload), decimals=3)
            if ok: self.tool_model.setData(idx, ret_val)
        elif field == 'LENGTH':
            ret_val, ok = QInputDialog.getDouble(self, header, field, float(preload), decimals=3)
            if ok: self.tool_model.setData(idx, ret_val)
        elif field == 'MFG':
            ret_val, ok = QInputDialog.getText(self, header, field, text=preload)
            if ok: self.tool_model.setData(idx, ret_val)

    def callOffsetDialog(self, item, field):
        idx = self.offset_model.index(item.row(), offset_headers['Tool'])
        tool = self.offset_model.data(idx)
        idx = self.offset_model.index(item.row(), offset_headers[field])
        header = f'Tool {tool} Offsets'
        preload = self.offset_model.data(idx)
        changed = True
        if field == 'Tool':
            ret_val, ok = QInputDialog.getInt(self, header, field, int(preload), 0, 100, 1)
            if ok:
                self.offset_model.setData(idx, ret_val)
                self.offset_model.tool_list = [ret_val if item == preload else item for item in self.offset_model.tool_list]
                self.tool_model.update_tool_no(preload, ret_val)
            else: changed = False
        elif field == 'Pocket':
            ret_val, ok = QInputDialog.getInt(self, header, field, int(preload), 0, 100, 1)
            if ok: self.offset_model.setData(idx, ret_val)
            else: changed = False
        elif field == 'Diameter':
            ret_val, ok = QInputDialog.getDouble(self, header, field, float(preload), decimals=3)
            if ok: self.offset_model.setData(idx, ret_val)
            else: changed = False
        elif field == 'Comment':
            ret_val, ok = QInputDialog.getText(self, header, field, text=preload)
            if ok: self.offset_model.setData(idx, ret_val)
            else: changed = False
        else:
            axis = self.offset_model.record().fieldName(item.column())
            title = f"Axis {axis} Offset {item.data()}"
            ret_val, ok = QInputDialog.getDouble(self, header, title, float(preload), decimals=3)
            if ok: self.offset_model.setData(idx, ret_val)
            else: changed = False
        if changed: self.data_changed()

    def tool_timer(self):
        self.timer_tenths += 1
        if self.timer_tenths == 10:
            self.timer_tenths = 0
            if self.timer_dict['running'] is True:
                self.timer_dict['time'] += 1

## calls from host
    def add_tool(self, tno):
        LOG.debug(f"Add tool {tno}")
        array = [tno, tno, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0, 'New tool']
        TOOL.ADD_TOOL(array)
        self.offset_model.update()

    def delete_tool(self, tno):
        LOG.debug(f"Deleting tool {tno}")
        TOOL.DELETE_TOOLS(tno)
        self.offset_model.update()

    def save_table(self):
        array = []
        for row in range(self.offset_model.rowCount(self)):
            line = []
            for hdr in offset_headers:
                idx = self.offset_model.index(row, offset_headers[hdr])
                data = self.offset_model.data(idx)
                line.append(data)
            array.append(line)
        return array

    def set_edit_mode(self, mode):
        self.enable_edit = mode

    def get_checked_tools(self):
        return self.offset_model.list_checked_tools()

    def set_metric_mode(self, state):
        self.offset_model.metric_display = state
        self.tool_model.metric_display = state

    def set_tool_icon(self, icon):
        checked = self.get_checked_tools()
        row = None
        if checked:
            row = self.tool_model.get_index(checked[0])
        if row is None: return
        idx = self.tool_model.index(row, tool_headers['ICON'])
        self.tool_model.setData(idx, icon)

    def get_tool_icon(self, tool):
        row = self.tool_model.get_index(tool)
        if row is None: return None
        idx = self.tool_model.index(row, tool_headers['ICON'])
        icon = self.tool_model.data(idx)
        if icon: return icon
        return None

    def get_next_available(self):
        for tno in range(0, 100):
            if tno not in self.offset_model.tool_list: break
        self.next_available = tno
        return tno

    def get_maxz(self, tool):
        row = self.tool_model.get_index(tool)
        if row is None: return ""
        idx = self.tool_model.index(row, tool_headers['LENGTH'])
        return self.tool_model.data(idx)

    def export_data(self):
        doc = QTextDocument()
        cursor = QTextCursor(doc)
        root = doc.rootFrame()
        cursor.setPosition(root.lastPosition())
        # top level frames
        frame0_format = QTextFrameFormat()
        frame1_format = QTextFrameFormat()
        frame2_format = QTextFrameFormat()
        frame1_format.setPageBreakPolicy(QTextFormat.PageBreak_AlwaysAfter)
        heading = cursor.insertFrame(frame0_format)
        first_page = cursor.insertFrame(frame1_format)
        cursor.setPosition(root.lastPosition())
        second_page = cursor.insertFrame(frame2_format)
        textOption = QtGui.QTextOption()
        textOption.setAlignment(Qt.AlignHCenter)
        doc.setDefaultTextOption(textOption)
        table_format = QTextTableFormat()
        table_format.setHeaderRowCount(1)
        table_format.setAlignment(Qt.AlignHCenter)
        hdr_format = QTextCharFormat()
        hdr_format.setFont(QFont('Lato', 10))
        std_format = QTextCharFormat()
        std_format.setFont(QFont('Lato', 8))
        std_format.setVerticalAlignment(QTextCharFormat.AlignMiddle)
        # heading
        cursor.setPosition(heading.firstPosition())
        cursor.insertText('QtDragon Tool Table', hdr_format)
        # first page
        cursor.setPosition(first_page.firstPosition())
        rows = self.offset_model.rowCount(self) + 1
        cols = len(self.hdr_list) - 1
        table1 = cursor.insertTable(rows, cols, table_format)
        i = 0
        for hdr in self.hdr_list:
            if hdr == 'Chk': continue
            table1.cellAt(0, i).firstCursorPosition().insertText(hdr, hdr_format)
            i += 1
        for row in range(1, rows):
            i = 0
            for hdr in self.hdr_list:
                if hdr == 'Chk': continue
                col = offset_headers[hdr]
                idx = self.offset_model.index(row-1, col)
                table1.cellAt(row, i).firstCursorPosition().insertText(str(self.offset_model.data(idx)), std_format)
                i += 1
        # second page
        cursor.setPosition(second_page.firstPosition())
        rows = self.tool_model.rowCount(self) + 1
        cols = len(tool_headers) - 1
        table2 = cursor.insertTable(rows, cols, table_format)
        i = 0
        for hdr in tool_headers:
            if hdr == 'idn': continue
            table2.cellAt(0, i).firstCursorPosition().insertText(hdr, hdr_format)
            i += 1
        for row in range(1, rows):
            i = 0
            for hdr in tool_headers:
                if hdr == 'idn': continue
                col = tool_headers[hdr]
                idx = self.tool_model.index(row-1, col)
                table2.cellAt(row, i).firstCursorPosition().insertText(str(self.tool_model.data(idx)), std_format)
                i += 1
    
        saveName = self.get_file_save("Select Save Filename")
        if saveName != '':
            printer = QPrinter()
            printer.setPageSize(QPrinter.Letter)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(saveName)
            doc.print_(printer)

    def get_file_save(self, caption):
        dialog = QFileDialog()
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        _filter = "Acrobat Files (*.pdf)"
        _dir = INFO.SUB_PATH
        fname, _ =  dialog.getSaveFileName(None, caption, _dir, _filter, options=options)
        return fname

    def get_tool_list(self):
        return self.offset_model.tool_list

    def get_version(self):
        return VERSION

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    w = Tool_Database()
    w.initialize()
    w.show()
    timer = QtCore.QTimer()
    timer.setInterval(100)
    timer.timeout.connect(w.tool_timer)
    timer.start()
    style = 'tooldb.qss'
    if os.path.isfile(style):
        file = QtCore.QFile(style)
        file.open(QtCore.QFile.ReadOnly)
        styleSheet = QtCore.QTextStream(file)
        w.setStyleSheet("")
        w.setStyleSheet(styleSheet.readAll())
    sys.exit( app.exec_() )
