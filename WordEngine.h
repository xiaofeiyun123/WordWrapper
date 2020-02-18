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
 
    /// ��Word�ļ������sFile·��Ϊ�ջ��������µ�Word�ĵ�
    bool Open(QString sFile, bool bVisible = true);
 
    void save(QString sSavePath);
    void SaveAsPDF(QString sSavePath);
    void close(bool bSave = true);
 
    bool replaceText(QString sLabel,QString sText);
    bool replacePic(QString sLabel,QString sFile, int nScaleRatio);
    bool AddPicToTable(QAxObject *table, int row, int column, QString sFile);

    //����һ�����м��б��
    QAxObject *insertTable(QString sLabel,int row,int column);
    //����һ�����м��б�� �����ñ�ͷ
    QAxObject *insertTable(QString sLabel,int row,int column,QStringList headList);
    //�����п�
    void setColumnWidth(QAxObject *table,int column, int width);
    void SetTableCellString(QAxObject *table, int row,int column,QString text);

	//****************************************************************
	//FunctionName : setTableCellFontBold ���ñ��Ԫ����Ӵ�
	//Param : QAxObject *table ������ָ��
	//Param : int row �кţ���1��ʼ��
	//Param : int column �кţ���1��ʼ��
	//Param : bool isBold  �Ƿ�Ӵ֣�true �Ӵ֣�false ���Ӵ֣�
	//****************************************************************
	bool setTableCellFontBold(QAxObject *table, int row, int column, bool isBold);

	//****************************************************************
	//FunctionName : setTableCellFontBold ���ñ��Ԫ�����С
	//Param : QAxObject *table ������ָ��
	//Param : int row �кţ���1��ʼ��
	//Param : int column �кţ���1��ʼ��
	//Param : int size  �����С
	//****************************************************************
	bool setTableCellFontSize(QAxObject *table, int row, int column, int size);

	//****************************************************************
	//FunctionName : setTableCellFontName ���ñ��Ԫ����
	//Param : QAxObject *table ������ָ��
	//Param : int row �кţ���1��ʼ��
	//Param : int column �кţ���1��ʼ��
	//Param : QString& fontName ������
	//****************************************************************
	bool setTableCellFontName(QAxObject *table, int row, int column, QString& fontName);

	//****************************************************************
	//FunctionName : setTableCellFontName ���ñ���п�
	//Param : QAxObject *table ������ָ��
	//Param : int column �кţ���1��ʼ��
	//Param : int width ���
	//****************************************************************
	bool setTableColumnWidth(QAxObject *table, int column, int width);

	//****************************************************************
	//FunctionName : setTableCellFontName ���ñ�����и�
	//Param : QAxObject *table ������ָ��
	//Param : int row �кţ���1��ʼ��
	//Param : int height �߶�
	//****************************************************************
	bool setTableColumnHeight(QAxObject *table, int row, int height);

	//****************************************************************
	//FunctionName : setTableCellFontName �ϲ����Ԫ��
	//Param : QAxObject *table ������ָ��
	//Param : int startRow ��ʼ�кţ���1��ʼ��
	//Param : int startColumn ��ʼ�кţ���1��ʼ��
	//Param : int endRow ��ֹ�кţ���1��ʼ��
	//Param : int endColumn ��ֹ�кţ���1��ʼ��
	//****************************************************************
	void mergeTableCells(QAxObject *table, int startRow, int startColumn, int endRow, int endColumn);

private:
 
    QAxObject *m_pWord;      //ָ������WordӦ�ó���
    QAxObject *m_pWorkDocuments;  //ָ���ĵ���,Word�кܶ��ĵ�
    QAxObject *m_pWorkDocument;   //ָ��m_sFile��Ӧ���ĵ�������Ҫ�������ĵ�
 
    QString m_sFile;
    bool m_bIsOpen;
    bool m_bNewFile;
};
 
#endif // WORDENGINE_H
