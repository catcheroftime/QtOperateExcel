#ifndef EXPORTEXCEL_H
#define EXPORTEXCEL_H


#include "QObject"

class QAxObject;
class ProgressRate;

class ExportExcel : public QObject
{
    Q_OBJECT

public:
    ExportExcel(const QList<QStringList> &storeinfo, const QStringList &header, const QString &storagepath, QWidget *parent = 0);
    ~ExportExcel();

    enum ExportError
    {
        NoError = 0,
        NewFileError,
        StoreInfoNull,
        TableInfoNotMatch,
        FileExists,
    };

    ExportError exportStatus();

private:
    ExportError   m_status;
    ProgressRate *m_pProgress;
    QAxObject    *m_pApp;
    QAxObject    *m_pWorkbooks;
    QAxObject    *m_pWorkbook;
    QAxObject    *m_pSheet;

private:
    bool newExcel(const QString &storagepath);
    void setCellsInfo(const QList<QStringList> &storeinfo, const QStringList &header);
    void setCellValue(const int &column, const int &row, const QString &value);
    void saveExcel(const QString &filename);
};

#endif // EXPORTEXCEL_H
