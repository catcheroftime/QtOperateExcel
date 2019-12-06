#ifndef IMPORTEXCEL_H
#define IMPORTEXCEL_H

#include <QObject>

class QAxObject;
class ProgressRate;

class ImportExcel :public QObject
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
    ProgressRate       *m_pProgress;
};

#endif // IMPORTEXCEL_H
