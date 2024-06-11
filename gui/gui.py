import os
import sys
from collections import defaultdict
from setupUi import Ui_MainWindow
from PyQt5 import QtWidgets, QtGui, QtCore
from concurrent.futures import ThreadPoolExecutor

curr_path = os.path.abspath(__file__)
proj_path = os.sep.join(curr_path.split(os.sep)[:-2])  # 工程目录
sys.path.append(proj_path)

from script.merge_table import MergeData, SaveExcel


class MainWindow(QtWidgets.QMainWindow):

    HANDLE_COL = {
        "销售总额": "C", "退单金额": "G", "退单数": "H", "上新数": "J",
        "访客": "K", "下单人数": "L", "下单总件数": "M"
    }
    # COLUMNS = ["日期", "销售总额", "上新数", "访客", "下单人数", "下单总件数", "退单金额", "退单数"]
    COLUMNS = ["日期", "销售总额", "退单金额", "退单数", "上新数", "访客", "下单人数", "下单总件数"]

    curr_store = ["艾莫克", "乘雀", "卡维妲", "欧密坊", "银雪龙", "屿笙栀"]
    EMOKE = 0
    CHENG = 1
    KA = 2
    OU = 3
    YIN = 4
    YU = 5
    REPORT = 0
    OVERVIEW = 1
    SPU = 2
    DATA = 3

    def __init__(self):
        super(MainWindow, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.currTable = -1
        self.storeTableWidget = None
        self.threadPool = ThreadPoolExecutor(1)
        self.exportTimer = QtCore.QTimer()
        self.importTimer = QtCore.QTimer()
        self.saveTimer = QtCore.QTimer()
        self.tableMap = defaultdict(list)
        self.initStatus()
        self.initSlot()
        self.merge = MergeData()
        self.lastSavePath = "/"

    def initSlot(self):
        self.ui.emokeAction.triggered.connect(self.on_emoke_action)
        self.ui.resetEAction.triggered.connect(lambda : self.on_reset_action(self.EMOKE))

        self.ui.chengAction.triggered.connect(self.on_cheng_action)
        self.ui.resetCAction.triggered.connect(lambda : self.on_reset_action(self.CHENG))

        self.ui.kaAction.triggered.connect(self.on_ka_action)
        self.ui.resetKAction.triggered.connect(lambda : self.on_reset_action(self.KA))

        self.ui.ouAction_4.triggered.connect(self.on_ou_action)
        self.ui.resetOAction.triggered.connect(lambda : self.on_reset_action(self.OU))

        self.ui.yinAction_5.triggered.connect(self.on_yin_action)
        self.ui.resetYinAction.triggered.connect(lambda : self.on_reset_action(self.YIN))

        self.ui.yuAction_6.triggered.connect(self.on_yu_action)
        self.ui.resetYuAction.triggered.connect(lambda : self.on_reset_action(self.YU))

        self.ui.exportAction.triggered.connect(self.on_export_action)
        self.exportTimer.timeout.connect(self.export_timer)
        self.importTimer.timeout.connect(self.isImpSPU)
        self.saveTimer.timeout.connect(self.saveFileTimer)
        self.ui.importSPUAction.triggered.connect(self.importSPU)

    def initStatus(self):
        self.exportTimer.start(100)
        self.importTimer.start(100)
        self.saveTimer.start(5000)
        self.ui.statusbar.showMessage("请选择需要操作的店铺！")
        # 初始化表格数据库
        for store in self.curr_store:
            # [absPath, ]
            # self.tableMap[store] = ["日报表", "交易概览", "SPU", "DATA", "退单"]
            self.tableMap[store] = [None, None, None, None, None]

    def saveFile(self):
        # 原路径保存数据
        for store_name, value in self.tableMap.items():
            if value[3] is not None and self.curr_store.index(store_name) == self.currTable:
                reportPath = value[0]
                overviewPath = value[1]
                file_name = overviewPath.split("/")[-1].split(".")[0]
                prefix, suffix = file_name.split("年")
                sheet_name = ".".join([prefix[-2:], suffix[:3]])
                SaveExcel.tmpSave(reportPath, sheet_name, store_name, value[3])
                self.ui.statusbar.showMessage(f"{store_name} 暂存成功！")

    def saveFileTimer(self):
        if not self.ui.exportAction.isEnabled():
            return
        self.threadPool.submit(self.saveFile)

    def isImpSPU(self):
        if self.currTable == -1:
            return
        store_name = self.curr_store[self.currTable]
        if self.tableMap[store_name][2] is None or self.tableMap[store_name][4] is None:
            self.ui.importSPUAction.setEnabled(True)
        else:
            self.ui.importSPUAction.setDisabled(True)

    def export_timer(self):
        flag = False
        for store_name, value in self.tableMap.items():
            if value[3] is not None:
                flag = True
        if flag:
            self.ui.exportAction.setEnabled(True)
        else:
            self.ui.exportAction.setDisabled(True)

    def importSPU(self):
        self.on_add_file_action()

    def changeCheckable(self):
        if self.currTable == self.EMOKE:
            self.ui.emokeAction.setCheckable(False)
            self.ui.emokeAction.setChecked(False)
        elif self.currTable == self.CHENG:
            self.ui.chengAction.setCheckable(False)
            self.ui.chengAction.setChecked(False)
        elif self.currTable == self.KA:
            self.ui.kaAction.setCheckable(False)
            self.ui.kaAction.setChecked(False)
        elif self.currTable == self.OU:
            self.ui.ouAction_4.setCheckable(False)
            self.ui.ouAction_4.setChecked(False)
        elif self.currTable == self.YIN:
            self.ui.yinAction_5.setCheckable(False)
            self.ui.yinAction_5.setChecked(False)
        elif self.currTable == self.YU:
            self.ui.yuAction_6.setCheckable(False)
            self.ui.yuAction_6.setChecked(False)

    def switchTable(self, store_eum):
        if store_eum != self.currTable:
            if self.storeTableWidget is not None:
                self.ui.verticalLayout.removeWidget(self.storeTableWidget)
                self.storeTableWidget = None
            self.changeCheckable()
        self.currTable = store_eum
        store_name = self.curr_store[store_eum]
        self.ui.statusbar.showMessage(f"{store_name}，已导入文件：空")
        info = self.tableMap[store_name]
        if info[self.OVERVIEW] is None or info[self.DATA] is None:
            return False
        self.imported(store_name)
        return True

    def on_export_action(self):
        try:
            file_dialog = QtWidgets.QFileDialog()
            if absPath := file_dialog.getExistingDirectory(None, "选择保存路径", self.lastSavePath):
                print(absPath)
                self.lastSavePath = absPath
                if sys.platform == "win32":
                    absPath = absPath[1:]
                elif sys.platform == "darwin":
                    absPath = absPath
                for store_name, value in self.tableMap.items():
                    if value[3] is not None:
                        reportPath = value[0]
                        toPath = os.sep.join([absPath, value[0].split(os.sep)[-1]])
                        overviewPath = value[1]
                        file_name = overviewPath.split("/")[-1].split(".")[0]
                        prefix, suffix = file_name.split("年")
                        sheet_name = ".".join([prefix[-2:], suffix[:3]])
                        SaveExcel.tmpSave(reportPath, sheet_name, store_name, value[3], toPath)
                QtWidgets.QMessageBox.about(None, "提示", "保存成功")
        except Exception as e:
            print("一键保存", e)
            QtWidgets.QMessageBox.warning(None, "提示", "保存失败")

    def on_reset_action(self, store_eum):
        try:
            store_name = self.curr_store[store_eum]
            message = QtWidgets.QMessageBox.warning(None, "提示", f"重置{store_name}数据！", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
            if message != 16384:
                # No
                return
            if store_eum == self.currTable and self.storeTableWidget is not None:
                self.storeTableWidget.deleteLater()
                self.storeTableWidget = None
                self.changeCheckable()
                self.ui.label.setText("选择一个店铺")
                self.ui.statusbar.showMessage("请选择需要操作的店铺！")
            # self.tableMap[store_name] = ["日报表", "交易概览", "SPU", "DATA"]
            self.tableMap[store_name] = [None, None, None, None, None]
            QtWidgets.QMessageBox.about(None, "提示", f"已重置{store_name}！")
        except Exception as e:
            print("重置数据: ", e)
            QtWidgets.QMessageBox.warning(None, "提示", f"重置失败！")

    def on_emoke_action(self):
        self.ui.emokeAction.setCheckable(True)
        self.ui.emokeAction.setChecked(True)
        # 检查改店铺是否存在数据（只要未导入交易概况肯定会打开文件）
        if not self.switchTable(self.EMOKE):
            self.on_add_file_action()
            return
        self.showTable()

    def on_cheng_action(self):
        self.ui.chengAction.setCheckable(True)
        self.ui.chengAction.setChecked(True)
        if not self.switchTable(self.CHENG):
            self.on_add_file_action()
            return
        self.showTable()

    def on_ka_action(self):
        self.ui.kaAction.setCheckable(True)
        self.ui.kaAction.setChecked(True)
        if not self.switchTable(self.KA):
            self.on_add_file_action()
            return
        self.showTable()

    def on_ou_action(self):
        self.ui.ouAction_4.setCheckable(True)
        self.ui.ouAction_4.setChecked(True)
        if not self.switchTable(self.OU):
            self.on_add_file_action()
            return
        self.showTable()

    def on_yin_action(self):
        self.ui.yinAction_5.setCheckable(True)
        self.ui.yinAction_5.setChecked(True)
        if not self.switchTable(self.YIN):
            self.on_add_file_action()
            return
        self.showTable()

    def on_yu_action(self):
        self.ui.yuAction_6.setCheckable(True)
        self.ui.yuAction_6.setChecked(True)
        if not self.switchTable(self.YU):
            self.on_add_file_action()
            return
        self.showTable()

    def showTable(self):
        store_eum = self.currTable
        store_name = self.curr_store[store_eum]
        try:
            self.ui.label.setText(store_name)
            # 判断是否更新table
            report_file = self.tableMap[store_name][self.REPORT]
            overview_file = self.tableMap[store_name][self.OVERVIEW]
            SPU_file = self.tableMap[store_name][self.SPU]
            refund_file = self.tableMap[store_name][4]
            merge = MergeData()
            if report_file and overview_file and SPU_file:
                # 整合1
                data = merge.merge3file(overview_file, report_file, SPU_file, store_name, refund_file)
                self.tableMap[store_name][self.DATA] = data[self.COLUMNS]
                pass
            elif report_file and overview_file:
                # 整合2
                data = merge.merge2file(overview_file, report_file, store_name, refund_file)
                self.tableMap[store_name][self.DATA] = data[self.COLUMNS]
            else:
                return
            tableWidget = self.initTableWidget(data)
            if self.storeTableWidget is None:
                self.storeTableWidget = tableWidget
            else:
                self.ui.verticalLayout.removeWidget(self.storeTableWidget)
                self.storeTableWidget = tableWidget
            self.ui.verticalLayout.addWidget(tableWidget)
        except Exception as e:
            self.tableMap[store_name] = [None] * 5
            self.changeCheckable()
            self.imported(store_name)
            self.ui.label.setText("选择一个店铺")
            print("展示表格", e)

    def on_add_file_action(self):
        filename = ""
        try:
            store_eum = self.currTable
            store_name = self.curr_store[store_eum]
            if not self.imported(store_name):  # 检查导入状态
                file_dialog = QtWidgets.QFileDialog(self)
                file_dialog.setFileMode(QtWidgets.QFileDialog.AnyFile)
                if files := file_dialog.getOpenFileNames(None, "支持多选文件", filter="Files(*.xlsx *.xls *.excel)")[0]:
                    for absPath in files:
                        # if sys.platform == "win":
                        #     absPath = absPath[1:]
                        filename = absPath.split(os.sep)[-1]
                        if "日报表" in filename:
                            self.tableMap[store_name][0] = absPath
                        elif "SPU" in filename:
                            self.tableMap[store_name][2] = absPath
                        elif "交易概况" in filename:
                            self.tableMap[store_name][1] = absPath
                        elif "退单" in filename:
                            self.tableMap[store_name][4] = absPath
                        else:
                            QtWidgets.QMessageBox.warning(None, "提示",
                                                          f"{filename}文件名格式有误！\n文件名期望包含(SPU, 交易概况, 存在店铺名+日报表)")
                    self.imported(store_name)  # 更新导入状态
            self.showTable()

        except Exception as e:
            print("文件导入: ", e)
            QtWidgets.QMessageBox.warning(None, "提示", f"{filename}文件名格式有误！\n文件名期望包含(SPU, 交易概况, 存在店铺名+日报表)")

    def imported(self, store_name):
        info = self.tableMap[store_name]
        imported = ""
        file_type = []
        cnt = 0
        if info[self.REPORT] is not None:
            name = info[self.REPORT].split(os.sep)[-1] + ";"
            file_type.append(f"日报表：{name}")
            imported += name
            cnt += 1
        if info[self.OVERVIEW] is not None:
            name = info[self.OVERVIEW].split(os.sep)[-1] + ";"
            file_type.append(f"交易概况：{name}")
            imported += name
            cnt += 1
        if info[self.SPU] is not None:
            name = info[self.SPU].split(os.sep)[-1] + ";"
            file_type.append(f"SPU：{name}")
            imported += name
            cnt += 1
        if info[4] is not None:
            name = info[4].split(os.sep)[-1] + ";"
            file_type.append(f"退单：{name}")
            imported += name
            cnt += 1
        exist = '\n'.join(file_type)
        self.ui.label.setText(f"{store_name}\n\n{exist}")
        self.ui.statusbar.showMessage(f"{store_name}，已导入文件：{imported}")
        if cnt == 4:
            return True
        return False

    def initTableWidget(self, row_data):
        try:
            cols = self.HANDLE_COL
            col_num = len(cols)
            row_num = len(row_data)
            tableWidget = QtWidgets.QTableWidget(self.ui.centralwidget)
            tableWidget.setObjectName("tableWidget")
            tableWidget.setColumnCount(col_num)
            tableWidget.setRowCount(row_num)
            # 头部
            for i, col in enumerate(cols):
                item = QtWidgets.QTableWidgetItem()
                item.setText(col)
                tableWidget.setHorizontalHeaderItem(i, item)

            # 数据本身
            row_data["日期"] = row_data["日期"].dt.date.astype("string")
            for idx, rows in enumerate(row_data[self.COLUMNS].values):
                # 添加row的头部
                item = QtWidgets.QTableWidgetItem()
                font = QtGui.QFont()
                font.setPointSize(12)
                item.setFont(font)
                item.setText(rows[0])  # 日期
                tableWidget.setVerticalHeaderItem(idx, item)

                # 销售总额 # 上新数 # 访客 # 下单人数 # 下单总件数 # 退单金额 # 退单数
                for i, row in enumerate(rows[1:]):
                    item = QtWidgets.QTableWidgetItem()
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    item.setText(str(int(row)))
                    tableWidget.setItem(idx, i, item)
            tableWidget.itemChanged.connect(self.handleEditTable)
            return tableWidget
        except Exception as e:
            QtWidgets.QMessageBox.warning(None, "提示", "导入失败！")
            print("表格制作", e)
            raise e
        # self.verticalLayout.addWidget(self.tableWidget)  # 将组件添加到界面

    def handleEditTable(self, item: QtWidgets.QTableWidgetItem):
        store_name = self.curr_store[self.currTable]
        data = self.tableMap[store_name][3]
        # print(data)
        if item.column() == 0 or item.column() == 1:
            data.iloc[item.row(), item.column() + 1] = round(float(item.text()), 2)
        else:
            data.iloc[item.row(), item.column() + 1] = item.text()
        print(data.iloc[item.row(), item.column()])
        self.tableMap[store_name][3] = data
        self.ui.statusbar.showMessage(f"{store_name} 未保存")


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

