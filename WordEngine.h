#ifndef WORDENGINE_H
#define WORDENGINE_H
#include <QObject>
#include <ActiveQt/QAxObject>
#include <QtCore>
class WordEngine : public QObject
{
    Q_OBJECT
public:
    WordEngine();
    ~WordEngine();
 
    /// 打开Word文件，如果sFile路径为空或错误，则打开新的Word文档
    bool Open(QString sFile, bool bVisible = true);
 
    void save(QString sSavePath);
    void SaveAsPDF(QString sSavePath);
    void close(bool bSave = true);
 
    bool replaceText(QString sLabel,QString sText);
    bool replacePic(QString sLabel,QString sFile, int nScaleRatio);
    bool AddPicToTable(QAxObject *table, int row, int column, QString sFile);

    //插入一个几行几列表格
    QAxObject *insertTable(QString sLabel,int row,int column);
    //插入一个几行几列表格 并设置表头
    QAxObject *insertTable(QString sLabel,int row,int column,QStringList headList);
    //设置列宽
    void setColumnWidth(QAxObject *table,int column, int width);
    void SetTableCellString(QAxObject *table, int row,int column,QString text);

	//****************************************************************
	//FunctionName : setTableCellFontBold 设置表格单元字体加粗
	//Param : QAxObject *table 表格对象指针
	//Param : int row 行号（从1开始）
	//Param : int column 列号（从1开始）
	//Param : bool isBold  是否加粗（true 加粗，false 不加粗）
	//****************************************************************
	bool setTableCellFontBold(QAxObject *table, int row, int column, bool isBold);

	//****************************************************************
	//FunctionName : setTableCellFontBold 设置表格单元字体大小
	//Param : QAxObject *table 表格对象指针
	//Param : int row 行号（从1开始）
	//Param : int column 列号（从1开始）
	//Param : int size  字体大小
	//****************************************************************
	bool setTableCellFontSize(QAxObject *table, int row, int column, int size);

	//****************************************************************
	//FunctionName : setTableCellFontName 设置表格单元字体
	//Param : QAxObject *table 表格对象指针
	//Param : int row 行号（从1开始）
	//Param : int column 列号（从1开始）
	//Param : QString& fontName 字体名
	//****************************************************************
	bool setTableCellFontName(QAxObject *table, int row, int column, QString& fontName);

	//****************************************************************
	//FunctionName : setTableCellFontName 设置表格列宽
	//Param : QAxObject *table 表格对象指针
	//Param : int column 列号（从1开始）
	//Param : int width 宽度
	//****************************************************************
	bool setTableColumnWidth(QAxObject *table, int column, int width);

	//****************************************************************
	//FunctionName : setTableCellFontName 设置表格列行高
	//Param : QAxObject *table 表格对象指针
	//Param : int row 行号（从1开始）
	//Param : int height 高度
	//****************************************************************
	bool setTableColumnHeight(QAxObject *table, int row, int height);

	//****************************************************************
	//FunctionName : setTableCellFontName 合并表格单元格
	//Param : QAxObject *table 表格对象指针
	//Param : int startRow 起始行号（从1开始）
	//Param : int startColumn 起始列号（从1开始）
	//Param : int endRow 终止行号（从1开始）
	//Param : int endColumn 终止列号（从1开始）
	//****************************************************************
	void mergeTableCells(QAxObject *table, int startRow, int startColumn, int endRow, int endColumn);

private:
 
    QAxObject *m_pWord;      //指向整个Word应用程序
    QAxObject *m_pWorkDocuments;  //指向文档集,Word有很多文档
    QAxObject *m_pWorkDocument;   //指向m_sFile对应的文档，就是要操作的文档
 
    QString m_sFile;
    bool m_bIsOpen;
    bool m_bNewFile;
};
 
#endif // WORDENGINE_H
