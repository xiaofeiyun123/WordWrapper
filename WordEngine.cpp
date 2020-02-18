#include "Common/Utility/WordEngine.h"
//#include "qt_windows.h"
WordEngine::WordEngine()
{
    m_pWord = NULL;
    m_pWorkDocuments = NULL;
    m_pWorkDocument = NULL;
 
    m_bIsOpen = false;
    m_bNewFile = false;
 
    HRESULT result = OleInitialize(0);
 
    if (result != S_OK && result != S_FALSE)
    {
        qDebug()<<QString("Could not initialize OLE (error %x)").arg((unsigned int)result);
    }
}
 
WordEngine::~WordEngine()
{
    //if(m_bIsOpen)
    //    close();
 
    OleUninitialize();
}
 
bool WordEngine::Open(QString sFile, bool bVisible)
{
     //新建一个word应用程序
     m_pWord = new QAxObject();
     bool bFlag = m_pWord->setControl( "word.Application" );
     if(!bFlag)
     {
        return false;
     }
     m_pWord->setProperty("Visible", bVisible);
     //获取所有的工作文档
     QAxObject *document = m_pWord->querySubObject("Documents");
     if(!document)
     {
        return false;
     }
     //以文件template.dot为模版新建一个文档
     document->dynamicCall("Add(QString)", sFile);
     //获取当前激活的文档
     m_pWorkDocument = m_pWord->querySubObject("ActiveDocument");
     if(m_pWorkDocument)
         m_bIsOpen = true;
     else
         m_bIsOpen = false;
 
     return m_bIsOpen;
}
 
void WordEngine::save(QString sSavePath)
{
    if(m_bIsOpen && m_pWorkDocument)
    {
        if(m_bNewFile){
            m_pWorkDocument->dynamicCall("Save()");
        }
        else{
            //m_pWorkDocument->dynamicCall("SaveAs (const QString&,int,const QString&,const QString&,bool,bool)",
             //                           m_sFile,56,QString(""),QString(""),false,false);
            m_pWorkDocument->dynamicCall("SaveAs (const QString&)", sSavePath);
        }
    }
    qDebug()<<"save Done.";
}
 
void WordEngine::SaveAsPDF(QString sSavePath)
{
    if(m_bIsOpen && m_pWorkDocument)
    {
            QVariant OutputFileName(sSavePath);
            QVariant ExportFormat(17);      //17是pdf
            QVariant OpenAfterExport(false); //保存后是否自动打开
            //转成pdf
            m_pWorkDocument->querySubObject("ExportAsFixedFormat(const QVariant&,const QVariant&,const QVariant&)",
                OutputFileName,
                ExportFormat,
                OpenAfterExport);

    }
    qDebug()<<"save Done.";
}

void WordEngine::close(bool bSave)
{
    if(bSave){
        //save();
    }
    if(m_pWord){
        m_pWord->setProperty("DisplayAlerts", false);
    }
    if(m_pWorkDocument){
        m_pWorkDocument->dynamicCall("Close(bool)", false);
    }
    if(m_pWord){
        m_pWord->dynamicCall("Quit()");
    }
    if(m_pWorkDocuments)
    {
        delete m_pWorkDocuments;
    }
//     if(m_pWord)
//     {
//         delete m_pWord;
//     }
    m_pWorkDocument = NULL;
    m_pWorkDocuments = NULL;
    m_pWord = NULL;
 
    m_bIsOpen   = false;
    m_bNewFile  = false;
 
}
 
bool WordEngine::replaceText(QString sLabel,QString sText)
{
    if(!m_pWorkDocument){
        return false;
    }
    //获取文档中名字为sLabel的标签
    QAxObject *pBookmark = m_pWorkDocument->querySubObject("Bookmarks(QString)",sLabel);
    if(pBookmark)
    {
        pBookmark->dynamicCall("Select(void)");
        pBookmark->querySubObject("Range")->setProperty("Text",sText);
        delete pBookmark;
    }
    return true;
}
 
bool WordEngine::replacePic( QString sLabel, QString sFile, int nScaleRatio )
{
    if(!m_pWorkDocument)
        return false;
 
    QAxObject *bookmark_pic = m_pWorkDocument->querySubObject("Bookmarks(QString)",sLabel);
    if(bookmark_pic)
    {
        bookmark_pic->dynamicCall("Select(void)");
        QAxObject *Inlineshapes = m_pWorkDocument->querySubObject("InlineShapes");
        Inlineshapes->dynamicCall("AddPicture(const QString&)",sFile);
        QAxObject* shape = m_pWorkDocument->querySubObject("InlineShapes(int)", 2);
        if (NULL == shape)
        {
            return false;
        }
        shape->dynamicCall( "ScaleHeight", nScaleRatio );
        shape->dynamicCall( "ScaleWidth", nScaleRatio );
        //delete Inlineshapes;
    }
    return true;
}

bool WordEngine::AddPicToTable(QAxObject *table, int row, int column, QString sFile)
{
    if(!table)
        return false;
    QAxObject *cell = table->querySubObject("Cell(int,int)",row,column);
    if(!cell)
        return false;
    //cell->dynamicCall("Select(void)");
    QAxObject *range = cell->querySubObject("Range");

    QAxObject *Inlineshapes = range->querySubObject("InlineShapes");
    Inlineshapes->dynamicCall("AddPicture(const QString&)",sFile);
    //delete Inlineshapes;

    return true;
}

QAxObject *WordEngine::insertTable(QString sLabel, int row, int column)
{
     QAxObject *bookmark = m_pWorkDocument->querySubObject("Bookmarks(QVariant)", sLabel);
     if(bookmark)
     {
       bookmark->dynamicCall("Select(void)");
       QAxObject *selection = m_pWord->querySubObject("Selection"); 
       //selection->dynamicCall("MoveDown(int)", 1);
       selection->dynamicCall("TypeParagraph(void)");
       selection->querySubObject("ParagraphFormat")->dynamicCall("Alignment", "wdAlignParagraphCenter");
       //selection->dynamicCall("TypeText(QString&)", "Table Test");//设置标题
 
       QAxObject *range = selection->querySubObject("Range");
       QAxObject *tables = m_pWorkDocument->querySubObject("Tables");
       QAxObject *table = tables->querySubObject("Add(QVariant,int,int)",range->asVariant(),row,column);

       for(int i=1;i<=6;i++)
       {
           QString str = QString("Borders(-%1)").arg(i);
           QAxObject *borders = table->querySubObject(str.toLocal8Bit().constData());
           borders->dynamicCall("SetLineStyle(int)",1);
       }

       return table;
     }

     return nullptr;
}
 
QAxObject *WordEngine::insertTable(QString sLabel, int row, int column, QStringList headList)
{
    QAxObject *bookmark = m_pWorkDocument->querySubObject("Bookmarks(QVariant)", sLabel);
    if(headList.size() != column){
        return NULL;
    }
    if(bookmark)
    {
      bookmark->dynamicCall("Select(void)");
      QAxObject *selection = m_pWord->querySubObject("Selection");
 
      selection->dynamicCall("InsertAfter(QString&)", "\r\n");
      //selection->dynamicCall("MoveLeft(int)", 1);
      selection->querySubObject("ParagraphFormat")->dynamicCall("Alignment", "wdAlignParagraphCenter");
      //设置标题
      //selection->dynamicCall("TypeText(QString&)", "Table Test");
 
      QAxObject *range = selection->querySubObject("Range");
      QAxObject *tables = m_pWorkDocument->querySubObject("Tables");
      QAxObject *table = tables->querySubObject("Add(QVariant,int,int)",range->asVariant(),row,column);
      //表格自动拉伸列 0固定  1根据内容调整  2 根据窗口调整
      table->dynamicCall("AutoFitBehavior(WdAutoFitBehavior)", 2);
 
      //设置表头
      for(int i=0;i<headList.size();i++){
          table->querySubObject("Cell(int,int)",1,i+1)->querySubObject("Range")->dynamicCall("SetText(QString)", headList.at(i));
          //加粗
          table->querySubObject("Cell(int,int)",1,i+1)->querySubObject("Range")->dynamicCall("SetBold(int)", true);
      }
 
      for(int i=1;i<=6;i++)
      {
          QString str = QString("Borders(-%1)").arg(i);
          QAxObject *borders = table->querySubObject(str.toLocal8Bit().constData());
          borders->dynamicCall("SetLineStyle(int)",1);
      }
      return table;
    }

    return nullptr;
}
 
void WordEngine::setColumnWidth(QAxObject *table, int column, int width)
{
    if(!table){
        return;
    }
    table->querySubObject("Columns(int)",column)->setProperty("Width",width);
}

#include <QColor>
void WordEngine::SetTableCellString(QAxObject *table, int row,int column,QString text)
{
    if(!table)
        return;
    QAxObject *cell = table->querySubObject("Cell(int,int)",row,column);
    if(!cell)
        return ;
    cell->dynamicCall("Select(void)");
    cell->querySubObject("Range")->setProperty("Text", text);
    //cell->querySubObject("Range")->setProperty("BgColor", QColor(0, 255, 0));   //设置单元格背景色（绿色）

}

void WordEngine::mergeTableCells(QAxObject *table, int startRow, int startColumn, int endRow, int endColumn)
{
    if (!table)
    {
        return;
    }
    QAxObject* StartCell = table->querySubObject("Cell(int, int)", startRow, startColumn);
    QAxObject* EndCell = table->querySubObject("Cell(int, int)", endRow, endColumn);
    if (!StartCell)
    {
        return;
    }
    if (!EndCell)
    {
        return;
    }
    StartCell->querySubObject("Merge(QAxObject *)", EndCell->asVariant());
}

bool WordEngine::setTableCellFontBold(QAxObject *table, int row, int column, bool isBold)
{
    if (!table)
        return false;
    table->querySubObject("Cell(int,int )", row, column)->querySubObject("Range")->dynamicCall("SetBold(int)", isBold);
    return true;
}
bool WordEngine::setTableCellFontSize(QAxObject *table, int row, int column, int size)
{
    if (!table)
        return false;
    table->querySubObject("Cell(int,int)", row, column)->querySubObject("Range")->querySubObject("Font")->setProperty("Size", size);
    return true;
}
bool WordEngine::setTableCellFontName(QAxObject *table, int row, int column, QString& fontName)
{
    if (!table)
        return false;
    table->querySubObject("Cell(int,int )", row, column)->querySubObject("Range")->querySubObject("Font")->setProperty("Name", fontName);
    return true;
}

bool WordEngine::setTableColumnWidth(QAxObject *table, int column, int width)
{
    if (!table)
    {
        return false;
    }
    table->querySubObject("Columns(int)", column)->setProperty("Width", width);
    return true;
}

bool WordEngine::setTableColumnHeight(QAxObject *table, int row, int height)
{
    if (!table)
    {
        return false;
    }
    table->querySubObject("Rows(int)", row)->setProperty("Height", height);
    return true;
}
