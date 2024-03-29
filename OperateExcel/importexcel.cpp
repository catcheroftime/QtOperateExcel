#include "importexcel.h"
#include "progressrate.h"

#include <QDebug>
#include <QCoreApplication>
#include <ActiveQt\QAxWidget>
#include <ActiveQt\QAxObject>

#ifndef SAFE_DELETE
#define SAFE_DELETE(p) { if(p){delete(p);  (p)=NULL;} }
#endif


ImportExcel::ImportExcel(const QString &filepath, QWidget *parent)
    : QObject(parent)
    , m_pProgress(new ProgressRate)
{
    readExcel(filepath);
}

ImportExcel::~ImportExcel()
{

}

QList<QStringList> ImportExcel::getImportExcelData()
{
    return m_result;
}

void ImportExcel::readExcel(const QString &filepath)
{
    m_pProgress->initProgress(1000, "分析文件中...");
    m_pProgress->showProgress(0);

    QString xlsFile = filepath;
    xlsFile.replace("/","\\");//获取文件目录并斜杠转成双反斜杠

    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false);
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)",xlsFile);

    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");
    int sheet_count = work_sheets->property("Count").toInt();

    m_pProgress->updateDescription(tr("导入中..."));
    int content_count = getExcelContentCount(work_book,sheet_count);

    int index = 1;
    for(int sheet_i =1 ;sheet_i<= sheet_count; sheet_i++)
    {
        QAxObject * work_sheet = work_book->querySubObject("Sheets(int)", sheet_i);
        QAxObject * used_range = work_sheet->querySubObject("UsedRange");//获取该sheet的使用范围对象

        QVariant var = used_range->dynamicCall("Value");
        QVariantList varRows = var.toList();
        if(varRows.isEmpty())
            continue;

        int row_count = varRows.size();
        QVariantList rowData;
        // 默认从Excel第二行第二列开始读取
        for(int i=1; i< row_count; ++i){
            rowData = varRows[i].toList();
            QStringList info;
            for(int j=1; j<rowData.size(); ++j){
                QString cell_info = rowData.at(j).toString();
                info.append(cell_info);
            }
            m_result.append(info);
            m_pProgress->showProgress(index*1000/content_count);
            index++;
        }
        SAFE_DELETE(used_range);
        SAFE_DELETE(work_sheet);

    }
    work_books->dynamicCall("Close()");
    excel.dynamicCall("Quit()");

    SAFE_DELETE(work_sheets);
    SAFE_DELETE(work_book);
    SAFE_DELETE(work_books);

    m_pProgress->releaseProgress();

}

int ImportExcel::getExcelContentCount(QAxObject *work_book, const int &sheet_count)
{
    int count =0;
    for(int i =1 ;i<= sheet_count; i++)
    {
        QAxObject * work_sheet = work_book->querySubObject("Sheets(int)", i);
        QAxObject * used_range = work_sheet->querySubObject("UsedRange");
        QAxObject * rows = used_range->querySubObject("Rows");
        int intRows = rows->property("Count").toInt()-1;
        count = intRows + count;
    }
    return count;
}
