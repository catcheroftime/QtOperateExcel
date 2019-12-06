#ifndef IMPORTEXCEL_H
#define IMPORTEXCEL_H

#include "progressrate.h"

class QAxObject;
class QProgressDialog;

class ImportExcel :public ProgressRate
{
    Q_OBJECT

public:
    ImportExcel(const QString &filepath, QWidget *parent = 0);
    ~ImportExcel();

    QList<QStringList> getImportExcelData();

private:
    void readExcel(const QString &filepath);
    int getExcelContentCount(QAxObject *work_book,const int &sheet_count);

private:
    QList<QStringList> m_result;
};

#endif // IMPORTEXCEL_H
