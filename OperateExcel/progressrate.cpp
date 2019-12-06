#include "progressrate.h"

#include <QProgressDialog>
#include <QCoreApplication>

ProgressRate::ProgressRate(QObject *parent) :
    QObject(parent)
  , m_pProgressDialog(0)
{

}

ProgressRate::~ProgressRate()
{
    if (m_pProgressDialog)
        releaseProgress();
}


void ProgressRate::initProgress(const int &size, const QString &description)
{
    if (!m_pProgressDialog)
        m_pProgressDialog = new QProgressDialog();//其实这一步就已经开始显示进度条了

    m_pProgressDialog->setAutoClose(false);
    m_pProgressDialog->setWindowFlags(Qt::Tool | Qt::FramelessWindowHint);//去掉标题栏
    m_pProgressDialog->setLabelText(description);
    m_pProgressDialog->setCancelButton(0);
    m_pProgressDialog->setRange(0,size);
    m_pProgressDialog->setModal(true);
    m_pProgressDialog->setWindowModality(Qt::WindowModal);
    m_pProgressDialog->setMinimumDuration(0);
    m_pProgressDialog->show();
    QCoreApplication::processEvents();
}

void ProgressRate::updateDescription(const QString &description)
{
    if (m_pProgressDialog) {
        m_pProgressDialog->setLabelText(description);
        QCoreApplication::processEvents();
    }
}

void ProgressRate::showProgress(const int &index)
{
    int show_index = index;
    if (show_index == m_pProgressDialog->maximum())
        show_index -= 1;

    m_pProgressDialog->setValue(show_index);
    QCoreApplication::processEvents();
}

void ProgressRate::releaseProgress()
{
    m_pProgressDialog->close();
    m_pProgressDialog->deleteLater();
    m_pProgressDialog = 0;
}
