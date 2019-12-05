#include "importexcel.h"

#include <QDebug>
#include <QCoreApplication>
#include <ActiveQt\QAxWidget>
#include <ActiveQt\QAxObject>
#include <QProgressDialog>

#ifndef SAFE_DELETE
#define SAFE_DELETE(p) { if(p){delete(p);  (p)=NULL;} }
#endif


ImportExcel::ImportExcel(const QString &filepath, QWidget *parent)
    : QDialog(parent)
    , m_pProgressDialog(0)
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

void ImportExcel::initProgress(const int &size)
{
    if (!m_pProgressDialog)
        m_pProgressDialog = new QProgressDialog();//其实这一步就已经开始显示进度条了

    m_pProgressDialog->setAutoClose(false);
    m_pProgressDialog->setWindowFlags(m_pProgressDialog->windowFlags() | Qt::FramelessWindowHint);//去掉标题栏
    m_pProgressDialog->setLabelText(tr("分析文件中..."));
    m_pProgressDialog->setCancelButton(0);
    m_pProgressDialog->setRange(0,size);
    m_pProgressDialog->setModal(true);
    m_pProgressDialog->setWindowModality(Qt::WindowModal);
    m_pProgressDialog->setMinimumDuration(0);
    m_pProgressDialog->show();
    QCoreApplication::processEvents();
}

void ImportExcel::showProgress(const int &index)
{
    int show_index = index;
    if (show_index == m_pProgressDialog->maximum())
        show_index -= 1;

    m_pProgressDialog->setValue(show_index);
    QCoreApplication::processEvents();
}

void ImportExcel::releaseProgress()
{
    m_pProgressDialog->close();
    m_pProgressDialog->deleteLater();
    m_pProgressDialog = 0;
}


void ImportExcel::readExcel(const QString &filepath)
{
    initProgress(1000);
    showProgress(0);


    QString xlsFile = filepath;
    xlsFile.replace("/","\\");//获取文件目录并斜杠转成双反斜杠

    QAxObject excel("Excel.Application");
    excel.setProperty("Visible", false);
    QAxObject *work_books = excel.querySubObject("WorkBooks");
    work_books->dynamicCall("Open (const QString&)",xlsFile);

    QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
    int sheet_count = work_sheets->property("Count").toInt();  //获取工作表数目

    int content_count = getExcelContentCount(work_book,sheet_count);
    qDebug() << content_count ;

    m_pProgressDialog->setLabelText(tr("导入中..."));
    QCoreApplication::processEvents();

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
            this->showProgress(index*1000/content_count);
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

    releaseProgress();

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
