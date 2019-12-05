#include "exportexcel.h"

#include <QDir>
#include <QFile>
#include <QCoreApplication>

ExportExcel::ExportExcel(const QList<QStringList> &storeinfo, const QStringList &header,const  QString &storagepath, QWidget *parent)
    : QDialog(parent)
    , m_status(NoError)
    , m_pProgressDialog(0)
{
    int count = storeinfo.size();
    if (count != 0) {
        if (storeinfo.at(0).size() != header.size() ) {
            m_status = TableInfoNotMatch;
//            return;
        }
    } else {
        m_status = StoreInfoNull;
        return;
    }

    initProgress(count+1);
    showProgress(0);

    if (newExcel(storagepath)) {
        m_pProgressDialog->setLabelText(tr("导出中..."));
        QCoreApplication::processEvents();

        setCellsInfo(storeinfo, header);
        saveExcel(storagepath);
        delete m_pApp;
    } else {
        m_status = NewFileError;
        return;
    }

    releaseProgress();
}

ExportExcel::~ExportExcel()
{

}

ExportExcel::ExportError ExportExcel::exportStatus()
{
    return m_status;
}

void ExportExcel::initProgress(const int &size)
{
    if (!m_pProgressDialog)
        m_pProgressDialog = new QProgressDialog();//其实这一步就已经开始显示进度条了

    m_pProgressDialog->setAutoClose(false);
    m_pProgressDialog->setWindowFlags(Qt::Tool | Qt::FramelessWindowHint);//去掉标题栏
    m_pProgressDialog->setLabelText(tr("生成文件中..."));
    m_pProgressDialog->setCancelButton(0);
    m_pProgressDialog->setRange(0,size);
    m_pProgressDialog->setModal(true);
    m_pProgressDialog->setWindowModality(Qt::WindowModal);
    m_pProgressDialog->setMinimumDuration(0);
    m_pProgressDialog->show();
    QCoreApplication::processEvents();
}

void ExportExcel::showProgress(const int &index)
{
    int show_index = index;
    if (show_index == m_pProgressDialog->maximum())
        show_index -= 1;
    m_pProgressDialog->setValue(show_index);
    QCoreApplication::processEvents();
}

void ExportExcel::releaseProgress()
{
    m_pProgressDialog->close();
    m_pProgressDialog->deleteLater();
    m_pProgressDialog = 0;
}



bool ExportExcel::newExcel(const QString &storagepath)
{
    m_pApp = new QAxObject();

    m_pApp->setControl("Excel.Application");
    m_pApp->dynamicCall("SetVisible(bool)", false);

    m_pApp->setProperty("DisplayAlerts", false);

    m_pWorkbooks = m_pApp->querySubObject("Workbooks");

    QFile file(storagepath);
    if (file.exists()) {
        m_status = ExportExcel::FileExists;
        return false;
    } else {
        m_pWorkbooks->dynamicCall("Add");
        m_pWorkbook = m_pApp->querySubObject("ActiveWorkBook");
    }

    m_pSheet = m_pWorkbook->querySubObject("Sheets(int)",1);
    return true;
}

void ExportExcel::setCellsInfo(const QList<QStringList> &storeinfo, const QStringList &header)
{
    // create title
    for (int col=2; col< header.size()+2; ++col) {
        setCellValue(col, 1, header.at(col-2));
    }
    showProgress(1);

    // create row info
    for (int row=2; row< storeinfo.size()+2; ++row) {
        QStringList rowinfo = storeinfo.at(row-2);

        setCellValue(1, row, QString::number(row-1));
        for (int col=2; col<rowinfo.size()+2; ++col) {
            QString info = rowinfo.at(col-2);
            setCellValue(col,row,info);
        }
        showProgress(row);
    }

}

void ExportExcel::setCellValue(const int &column, const int &row, const QString &value)
{
    QAxObject *range = m_pSheet->querySubObject("Cells(int,int)", row, column);
    QString savevalue = value;
    if (value.size() >= 15) {
        savevalue.insert(0, '\'');
    }
    range->setProperty("Value", savevalue);
}

void ExportExcel::saveExcel(const QString &filename)
{
    m_pWorkbook->dynamicCall("SaveAs(const QString &)",
        QDir::toNativeSeparators(filename));

    m_pApp->dynamicCall("Quit(void)");
}
