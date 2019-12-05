#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>

namespace Ui {
class MainWindow;
}

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = 0);
    ~MainWindow();

private:
    void initView();


private slots:
    void on_ptn_export_clicked();
    void on_ptn_import_clicked();
    void on_ptn_selectall_clicked(bool checked);

private:
    Ui::MainWindow *ui;

    QStringList m_headerList;
};

#endif // MAINWINDOW_H
