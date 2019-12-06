#include "exportexcel.h"
#include "progressrate.h"


#include <QDir>
#include <QFile>
#include <QCoreApplication>
#include <ActiveQt\QAxWidget>
#include <ActiveQt\QAxObject>


ExportExcel::ExportExcel(const QList<QStringList> &storeinfo, const QStringList &header,const  QString &storagepath, QWidget *parent)
    : QObject(parent)
    , m_status(NoError)
    , m_pProgress(new ProgressRate)
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

    m_pProgress->initProgress(count+1, "生成文件中...");
    m_pProgress->showProgress(0);

    if (newExcel(storagepath)) {
        m_pProgress->updateDescription(tr("导出中..."));


        setCellsInfo(storeinfo, header);
        saveExcel(storagepath);
        delete m_pApp;
    } else {
        m_status = NewFileError;
        return;
    }

    m_pProgress->releaseProgress();
}

ExportExcel::~ExportExcel()
{
    delete m_pProgress;
}

ExportExcel::ExportError ExportExcel::exportStatus()
{
    return m_status;
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
    m_pProgress->showProgress(1);

    // create row info
    for (int row=2; row< storeinfo.size()+2; ++row) {
        QStringList rowinfo = storeinfo.at(row-2);

        setCellValue(1, row, QString::number(row-1));
        for (int col=2; col<rowinfo.size()+2; ++col) {
            QString info = rowinfo.at(col-2);
            setCellValue(col,row,info);
        }
        m_pProgress->showProgress(row);
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
