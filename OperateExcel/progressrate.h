#ifndef PROGRESSRATE_H
#define PROGRESSRATE_H

#include <QObject>
class QProgressDialog;

class ProgressRate : public QObject
{
    Q_OBJECT
public:
    explicit ProgressRate(QObject *parent = 0);
    ~ProgressRate();

    void initProgress(const int &size, const QString &description);
    void updateDescription(const QString &description);
    void showProgress(const int &index);
    void releaseProgress();

private:
    QProgressDialog * m_pProgressDialog;

};

#endif // PROGRESSRATE_H
