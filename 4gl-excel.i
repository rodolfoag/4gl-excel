/*
 * Include Funcoes/Procedures uteis p/ geracao de relatorios p/ Excel
 *
 * Autor: Rodolfo Goncalves
 * Data.: janeiro/2012
 *
 * Observacao: p/ utilizacao desta include, eh necessario a definicao de uma temp-table
 *             que contera os dados do relatorio, os valores de column-label e format
 *             sao utilizados para formatacao dos campos e a linha de cabecalho.
 *
 * Current Version: 0.1.0
 */


/* Retorna a string de conteudo de um buffer field de acordo com o seu formato e tipo */
function fi-fieldString returns char (p-field as handle):
    def var c-valor as char no-undo.

    if string(p-field:buffer-value) = ? then
        assign c-valor = "".

    else if p-field:data-type = "decimal" then
        assign c-valor = string(p-field:buffer-value, p-field:format).

    else
        assign c-valor = string(p-field:buffer-value).

    return c-valor.
end.


/* Gera um arquivo csv com os dados do relatorio */
procedure pi-cria-arquivo-csv:
    def input  parameter p-buffer  as handle no-undo. /* buffer da temp-table com os dados do relatorio */
    def input  parameter p-arquivo as char   no-undo. /* nome completo do arquivo destino do relatorio */
    def output parameter p-arq-csv as char   no-undo. /* arquivo CSV com os dados do relatorio */

    def var i-cont          as int    no-undo.
    def var h-query         as handle no-undo.

    /* Nome do Arquivo de Destino */
    if p-arquivo = "" then do:
        assign p-arq-csv = entry(num-entries(program-name(1),"\"), program-name(1), "\") + ".csv".
    end.
    else do:
        assign p-arq-csv = replace(p-arquivo, ".xls", ".csv").
    end.


    /* Cria o arquivo csv no SO */
    output to value(p-arq-csv).

    /* Imprime o cabecalho do relatorio */
    do i-cont = 1 to p-buffer:num-fields:
        put unformatted
            p-buffer:buffer-field(i-cont):column-label
            if not i-cont = p-buffer:num-fields then ";" else "".
    end.

    put skip.

    /* Imprime dados da temp-table */
    create query h-query.
    h-query:set-buffers(p-buffer).
    h-query:query-prepare("for each " + p-buffer:name).
    h-query:query-open().

    h-query:get-first().
    do while not h-query:query-off-end:

        do i-cont = 1 to p-buffer:num-fields:

            put unformatted
                fi-fieldString(input p-buffer:buffer-field(i-cont)).

            if i-cont < p-buffer:num-fields then
                put ";".

        end.

        put skip.

        h-query:get-next().
    end.

    h-query:query-close() no-error.

    output close.
end.


procedure pi-cria-arquivo-xls:
    def input  parameter p-buffer  as handle no-undo. /* buffer da temp-table com os dados do relatorio */
    def input  parameter p-arq-csv as char   no-undo. /* Nome do Arquivo csv com os dados do relatorio */
    def output parameter p-arq-xls as char   no-undo. /* Nome do Arquivo xls criado contendo o relatorio formatado */

    def var ch-excel      as com-handle no-undo.
    def var ch-wrk        as com-handle no-undo.
    def var ch-query      as com-handle no-undo.
    def var raw-array     as raw        no-undo.
    def var i-cont        as int        no-undo.
    def var i-num-cols    as int        no-undo.

    /* Abre excel */
    create "Excel.Application" ch-excel no-error.
    ch-excel:visible       = no.
    ch-excel:DisplayAlerts = no.

    ch-wrk = ch-excel:workbooks:add.

    /* Importa arquivo CSV */
    ch-query = ch-wrk:ActiveSheet:QueryTables:add("TEXT;" + p-arq-csv,
                                                  ch-excel:Range("$A$1")).
    assign ch-query:name = "data"
           ch-query:FieldNames = true
           ch-query:RowNumbers = false
           ch-query:FillAdjacentFormulas = false
           ch-query:PreserveFormatting = true
           ch-query:RefreshOnFileOpen = false
           ch-query:RefreshStyle = 1 /* xlInsertDeleteCells */
           ch-query:SavePassword = false
           ch-query:SaveData = true
           ch-query:AdjustColumnWidth = true
           ch-query:RefreshPeriod = 0
           ch-query:TextFilePromptOnRefresh = false
           ch-query:TextFilePlatform = 850
           ch-query:TextFileStartRow = 1
           ch-query:TextFileParseType = 1 /* xlDelimited */
           ch-query:TextFileTextQualifier = -4142 /* xlTextQualifierNone */
           ch-query:TextFileConsecutiveDelimiter = false
           ch-query:TextFileTabDelimiter = false
           ch-query:TextFileSemicolonDelimiter = true
           ch-query:TextFileCommaDelimiter = false
           ch-query:TextFileSpaceDelimiter = false
           ch-query:TextFileTrailingMinusNumbers = true.

    /* Configura o tipo de formatacao das colunas, para isso, utiliza o tipo de variavel raw p/ passar array de tipos p/ o Excel */
    do i-cont = 1 to p-buffer:num-fields:

        if p-buffer:buffer-field(i-cont):column-label = "" then do:
            put-byte(raw-array, i-cont) = 9. /* 9 = xlSkipColumn */
            next.
        end.

        case p-buffer:buffer-field(i-cont):data-type:
            when "character" then put-byte(raw-array, i-cont) = 2. /* 2 = xlTextFormat Excel */
            when "date"      then put-byte(raw-array, i-cont) = 4. /* 4 = xlDMYFormat */
            otherwise  put-byte(raw-array, i-cont) = 1. /* 1 = xlGeneralFormat */
        end case.

        assign i-num-cols = i-num-cols + 1.

    end.

    assign ch-query:TextFileColumnDataTypes = raw-array.

    /* Atualiza os dados */
    ch-query:refresh().


    /* Configura excel visualmente
     * - Negrito p/ primeira linha e Auto Filtro
     * - Insere borda
     */

    /* Negrito */
    ch-excel:range("A1", ch-excel:cells(1, i-num-cols)):select().
    ch-excel:selection:font:bold = true.
    ch-excel:selection:Interior:ColorIndex = 34.
    ch-excel:selection:Interior:Pattern = 1.

    /* Auto Filtro */
    ch-excel:selection:AutoFilter(,,).

    /* Ajusta as colunas de acordo com tamanho do conteudo */
    ch-excel:Cells:select().
    ch-excel:selection:columns:AutoFit().
    ch-excel:Range("A1"):select(). /* tira selecao */

    /* Verifica se pi-customiza-excel esta definido, e caso esteja
       a excuta p/ customizacoes no arquivo criado */
    if lookup("pi-customiza-excel",this-procedure:internal-entries) > 0 then
        run pi-customiza-excel(input ch-excel).

    /* Salva o arquivo XLS */
    assign p-arq-xls = replace(p-arq-csv, ".csv", ".xls").
    ch-wrk:SaveAs(p-arq-xls,-4143,"","",false,false,).

    /* Encerra excel e elimna handles */
    ch-excel:DisplayAlerts = yes.
    ch-excel:quit().
    release object ch-wrk.
    release object ch-query.
    release object ch-excel.
    assign ch-wrk   = ?
           ch-query = ?
           ch-excel = ?.

    /* Elimina arquivo csv gerado */
    os-delete value(p-arq-csv) no-error.
end.
