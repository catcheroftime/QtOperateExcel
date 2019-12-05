#include "mainwindow.h"
#include "ui_mainwindow.h"

#include "exportexcel.h"
#include "importexcel.h"

#include <QDebug>
#include <QFileDialog>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    initView();
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::initView()
{
    ui->ptn_selectall->setCheckable(true);
    ui->ptn_selectall->setChecked(false);

    m_headerList = QStringList{ "姓名", "性别" , "年龄"};
    ui->treeWidget_showinfo->setHeaderLabels(m_headerList);
}

void MainWindow::on_ptn_export_clicked()
{
    QString filename = QFileDialog::getSaveFileName(this, tr("导出Excel"), QCoreApplication::applicationDirPath(), tr("excel(*.xlsx)"));
    if (filename.isEmpty())
        return;

    QTreeWidgetItemIterator it(ui->treeWidget_showinfo, QTreeWidgetItemIterator::Checked);
    QList<QStringList> storeinfo;
    while ( *it ) {
        QStringList iteminfo;
        for (int i=0; i < (*it)->columnCount(); i++) {
            iteminfo.append( (*it)->text(i) );
        }
        storeinfo.append(iteminfo);
        ++it;
    }

    ExportExcel excel(storeinfo, m_headerList, filename);
    ExportExcel::ExportError error = excel.exportStatus();
    if (error == ExportExcel::NoError || error == ExportExcel::TableInfoNotMatch)
        qDebug() << "导出成功";
    else if (error ==  ExportExcel::StoreInfoNull)
        qDebug() << "未发现需要导出的条目";
    else if (error == ExportExcel::FileExists)
        qDebug() << "目前不支持在已存在文件上覆盖导出，请在导出Excel弹窗中的文件名上填写不存在的文件名称！";
    else
        qDebug() <<"导出失败";
}

void MainWindow::on_ptn_import_clicked()
{
    QString filename = QFileDialog::getOpenFileName(this, tr("打开文件"), QCoreApplication::applicationDirPath() ,"excel(*.xls *.xlsx)");
    if(filename.isEmpty())
        return;

    ImportExcel excel{filename};

    QList<QStringList> importData = excel.getImportExcelData();

    for (auto itemsinfo: importData) {
        QTreeWidgetItem *item = new QTreeWidgetItem{ itemsinfo };
        item->setCheckState(0, Qt::Unchecked);
        ui->treeWidget_showinfo->addTopLevelItem(item);
    }
}


void MainWindow::on_ptn_selectall_clicked(bool checked)
{
    Qt::CheckState state;
    if (checked) {
        state = Qt::Checked;
        ui->ptn_selectall->setText("取消全选");
    }
    else {
        state = Qt::Unchecked;
        ui->ptn_selectall->setText("全选");
    }

    QTreeWidgetItemIterator it(ui->treeWidget_showinfo);
    while ( *it ) {
        (*it)->setCheckState(0, state);
        ++it;
    }
}
