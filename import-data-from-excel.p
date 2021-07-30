DEFINE TEMP-TABLE tt-carrega-excel
   FIELD c-item         AS CHARACTER
   FIELD c-estabel      AS CHARACTER
   FIELD c-deposito     AS CHARACTER
   FIELD c-localiz      AS CHARACTER
   FIELD i-tipo         AS INTEGER
   FIELD d-capacidade   AS DECIMAL
   FIELD i-setor-pgto   AS INTEGER
   FIELD i-tp-localiz   AS INTEGER.

/*defini‡Æo de vari veis locais*/
DEF VAR c-caminho       AS CHARACTER.
DEF VAR i-linha         AS INTEGER INITIAL 1.

/*defini‡Æo de vari veis do Excel*/
DEF VAR chExcelApplication  AS COM-HANDLE.  
DEF VAR chWorkbook          AS COM-HANDLE.  
DEF VAR chWorksheet         AS COM-HANDLE.  
DEF VAR chWorkSheetRange    AS COM-HANDLE.

/*Busca do arquivo excel*/
ASSIGN c-caminho = trim(search('C:\tmp\local-esp.xlsx')).

/*API do Excel*/
CREATE "Excel.Application" chExcelApplication. 
chExcelApplication:Visible = no.

/*definicoes da planilha*/
chWorkbook = chExcelApplication:Workbooks:OPEN(c-caminho). 
chWorksheet = chWorkbook:sheets:item(1). /*Seleciona a primeira Planilha do arquivo*/

REPEAT :
    ASSIGN i-linha = i-linha + 1.

    /*faz o teste para encerrar la‡o de repeti‡Æo*/
    IF chWorksheet:Range('A' + string(i-linha)):text   = ''    or 
       chWorksheet:Range('A' + string(i-linha)):text   = ?     then leave. /*Se a c‚lula A for vazia, sai do la‡o de repeti‡Æo*/
        /*Carrega os valores da planilha*/
    CREATE tt-carrega-excel.
    ASSIGN tt-carrega-excel.c-item         = chWorksheet:Range('A' + STRING(i-linha)):TEXT.      
           tt-carrega-excel.c-estabel      = chWorksheet:Range('B' + STRING(i-linha)):TEXT.  
           tt-carrega-excel.c-deposito     = chWorksheet:Range('C' + STRING(i-linha)):TEXT.  
           tt-carrega-excel.c-localiz      = chWorksheet:Range('D' + STRING(i-linha)):TEXT.  
           tt-carrega-excel.i-tipo         = INT(chWorksheet:Range('E' + STRING(i-linha)):TEXT).
           tt-carrega-excel.d-capacidade   = DEC(chWorksheet:Range('F' + STRING(i-linha)):TEXT).
           tt-carrega-excel.i-setor-pgto   = INT(chWorksheet:Range('G' + STRING(i-linha)):TEXT).
           tt-carrega-excel.i-tp-localiz   = INT(chWorksheet:Range('H' + STRING(i-linha)):TEXT).
END.


/*l¢gica de grava‡Æo dos Dados coletados da planilha*/
FOR EACH tt-carrega-excel:
       

END.

RELEASE OBJECT chExcelApplication     NO-ERROR.
RELEASE OBJECT chWorksheet            NO-ERROR.
