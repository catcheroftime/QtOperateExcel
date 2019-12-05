#ifndef IMPORTEXCEL_H
#define IMPORTEXCEL_H

#include <QDialog>

class QAxObject;
class QProgressDialog;

class ImportExcel :public QDialog
{
    Q_OBJECT

public:
    ImportExcel(const QString &filepath, QWidget *parent = 0);
    ~ImportExcel();

    QList<QStringList> getImportExcelData();

private:
    void initProgress(const int &size);
    void showProgress(const int &index);
    void releaseProgress();

    void readExcel(const QString &filepath);
    int getExcelContentCount(QAxObject *work_book,const int &sheet_count);

private:
    QList<QStringList> m_result;
    QProgressDialog * m_pProgressDialog;

};

#endif // IMPORTEXCEL_H
