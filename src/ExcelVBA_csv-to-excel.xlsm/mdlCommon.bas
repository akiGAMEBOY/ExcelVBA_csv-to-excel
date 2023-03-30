Attribute VB_Name = "mdlCommon"
Option Explicit

' �萔
'   ���C��
Public Const MAIN_RANGE As String = "D3,D4,D5,D6,D7,D8,D9,D10"
' �@�G���[���
Public Const ERROR_RANGE As String = "D13,D14"
' �@���͏��
Public Const INPUT_RANGE As String = "D17,D18,D19,D20,D21"
' �@�o�͏��
Public Const OUTPUT_RANGE As String = "D24,D25"
' �@�}�X�^�[
Public Const MASTER_RANGE As String = "D28,D29,D30"
' �@��\������V�[�g���
Public Const HIDDEN_RANGE As String = "D33,D34,D35,D36,D37,D38,D39,D40,D41,D42,D43,D44,D45"
'   �V�[�g��
Public Const SETINFO_SHEETNAME As String = "�ݒ���"
'   �p�����[�^�[
Public Const MAIN_PARA As String = "MAIN"
Public Const ERROR_PARA As String = "ERROR"
Public Const INPUT_PARA As String = "INPUT"
Public Const OUTPUT_PARA As String = "OUTPUT"
Public Const MASTER_PARA As String = "MASTER"
Public Const HIDDEN_PARA As String = "HIDDEN"
'   �t�@�C���^�C�v
Public Const PDF_FILETYPE As String = "PDF�t�@�C��,*.pdf"
Public Const EXCEL_FILETYPE As String = "Excel 97-2003 �u�b�N (*.xls),*.xls,Excel �u�b�N (*.xlsx),*.xlsx"
Public Const CSV_FILETYPE As String = "CSV�t�@�C��,*.csv"
Public Const TEXT_FILETYPE As String = "�e�L�X�g�t�@�C��,*.txt"

' Function�v���V�[�W��
'********************************************************************************
'* �������@�bIsEmptyText
'* �@�\�@�@�b�󕶎��`�F�b�N
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=�l�Ȃ�, False=�l����j
'* �����@�@�bstrValue�F�Ώە�����
'********************************************************************************
Function IsEmptyText(strValue As String) As Boolean
    IsEmptyText = (strValue = "")

End Function

'********************************************************************************
'* �������@�bIsNumericEx
'* �@�\�@�@�b���l�`�F�b�N�i���p���l�̔���j
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=���p���l�AFalse=���p���l�ȊO�j
'* �����@�@�bstrValue�F�Ώە�����
'********************************************************************************
Function IsNumericEx(strValue As String) As Boolean
    Dim objReg As New RegExp

    ' ���������F���p���l�̂�
    objReg.Pattern = "^[+,-]?([1-9]\d*|0)(\.\d+)?$"
    objReg.Global = True

    IsNumericEx = objReg.Test(strValue)

End Function

'********************************************************************************
'* �������@�bIsExistsFile
'* �@�\�@�@�b�t�@�C���̑��݃`�F�b�N
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=���݂���, False=���݂��Ȃ��j
'* �����@�@�bstrPath�F�Ώۃt�@�C���̐�΃p�X
'********************************************************************************
Function IsExistsFile(strPath As String) As Boolean
    IsExistsFile = Dir(strPath) <> ""

End Function

'********************************************************************************
'* �������@�bIsExistsFolder
'* �@�\�@�@�b�t�H���_�̑��݃`�F�b�N
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=���݂���, False=���݂��Ȃ��j
'* �����@�@�bstrPath�F�Ώۃt�H���_�̐�΃p�X
'********************************************************************************
Function IsExistsFolder(strPath As String) As Boolean
    IsExistsFolder = Dir(strPath, vbDirectory) <> ""

End Function

'********************************************************************************
'* �������@�bIsReadonlyFile
'* �@�\�@�@�b�t�@�C���̓ǂݎ���p���`�F�b�N
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=�ǂݎ���p�ł���, False=�ǂݎ���p�ł͂Ȃ��j
'* �����@�@�bstrPath�F�Ώۃt�@�C���̐�΃p�X
'********************************************************************************
Function IsReadonlyFile(strPath As String) As Boolean
    Dim vbResult As VbFileAttribute
    
    ' �t�@�C���̑������擾
    vbResult = GetAttr(strPath)
    
    ' �ǂݎ���p�̔���
    IsReadonlyFile = ((vbResult And vbReadOnly) = vbReadOnly)

End Function

'********************************************************************************
'* �������@�bIsOpenedFile
'* �@�\�@�@�b�t�@�C���̊J����Ԃ��`�F�b�N
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=�Ђ炫���, False=�Ƃ�����ԁj
'* �����@�@�bstrPath�F�Ώۃt�@�C���̐�΃p�X
'********************************************************************************
Function IsOpenedFile(strPath As String) As Boolean
    Dim intFileno As Integer
    
    intFileno = FreeFile
    
    ' �����t�@�C����ǋL���[�h�ŊJ�������̖߂�l�Ŕ���
    On Error Resume Next
    Open strPath For Append As #intFileno
    Close #intFileno
    
    IsOpenedFile = (Err.Number >= 1)

End Function

'********************************************************************************
'* �������@�bIsExistsSheet
'* �@�\�@�@�b�V�[�g�̑��݃`�F�b�N
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=���݂���, False=���݂��Ȃ��j
'* �����@�@�bstrSheetname�F�ΏۃV�[�g��
'********************************************************************************
Function IsExistsSheet(strSheetname As String) As Boolean
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    Set wbTarget = ThisWorkbook
    Set wsTarget = Nothing

    On Error Resume Next
    Set wsTarget = wbTarget.Worksheets(strSheetname)
    On Error GoTo 0

    IsExistsSheet = (Not (wsTarget Is Nothing))
    
    Set wsTarget = Nothing
    Set wbTarget = Nothing

End Function

'********************************************************************************
'* �������@�bIsEmptyTablerequired
'* �@�\�@�@�b�K�{���ڃ`�F�b�N�i�Ώۂ�Excel�̕\�j
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=�����ꂩ�󕶎�����, False=�󕶎��Ȃ� or 0���j
'* �����@�@�bstrSheetname�F�ΏۃV�[�g��, strRange�F�Z���ʒu, aryCheckcol�F�K�{����
'*-------------------------------------------------------------------------------
'* ���ӎ����F1.�ΏۃV�[�g�͑��݂��鎖���O��ƂȂ�B
'* �@�@�@�@�@�@���̃v���V�[�W�����Ăяo���O�ɃV�[�g�̑��݃`�F�b�N�iIsExistsSheet�j���A
'* �@�@�@�@�@�@���s���G���[���䂷�邱�ƁB
'* �@�@�@�@�@2.��̊J�n�ʒu�̍ő�s����ΏۂɃ`�F�b�N�����s�����B
'********************************************************************************
Function IsEmptyTablerequired(strSheetname As String, strRange As String, aryCheckcol() As String) As Boolean
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    
    Dim rngTarget As Range
    Dim lngBeginrow As Long
    Dim lngBegincol As Long
    Dim lngEndrow As Long
    Dim lngEndcol As Long
    Dim strAddress As String
    Dim lngRow As Long
    Dim lngCount As Long
    Dim strValue As String
    
    Set rngTarget = Range(strRange)
    lngBeginrow = rngTarget.Row
    lngBegincol = rngTarget.Column
    
    Set wbTarget = ThisWorkbook
    Set wsTarget = wbTarget.Worksheets(strSheetname)
    
    With wsTarget
        lngEndrow = .Cells(Rows.Count, lngBegincol).End(xlUp).Row
        If lngBeginrow > lngEndrow Then
            ' �f�[�^0���̏ꍇ�͏I��
            IsEmptyTablerequired = False
            Exit Function
        End If
        lngEndcol = .Cells(lngBeginrow, Columns.Count).End(xlToLeft).Column
        
        For lngRow = lngBeginrow To lngEndrow
            For lngCount = 0 To UBound(aryCheckcol)
                strValue = .Cells(lngRow, CLng(aryCheckcol(lngCount)))
                strAddress = .Cells(lngRow, CLng(aryCheckcol(lngCount))).Address(False, False)
                Call SubSelectCell(strSheetname, strAddress)
                If IsEmptyText(strValue) Then
                    IsEmptyTablerequired = True
                    Exit Function
                End If
            Next
        Next
    End With
    
    IsEmptyTablerequired = False
    
    Set wbTarget = Nothing
    Set wsTarget = Nothing

End Function

'********************************************************************************
'* �������@�bFuncReadSetinfo
'* �@�\�@�@�b�ݒ���̓ǂݍ���
'*-------------------------------------------------------------------------------
'* �߂�l�@�bString()�F�w�肵�������敪�̏��
'* �����@�@�bstrClass�F�����敪�̖��O
'********************************************************************************
Function FuncReadSetinfo(strClass As String) As String()
    Dim arySetinfo() As String
    
    Select Case strClass
        Case mdlCommon.MAIN_PARA
            arySetinfo = Split(MAIN_RANGE, ",")
        Case mdlCommon.ERROR_PARA
            arySetinfo = Split(ERROR_RANGE, ",")
        Case mdlCommon.INPUT_PARA
            arySetinfo = Split(INPUT_RANGE, ",")
        Case mdlCommon.OUTPUT_PARA
            arySetinfo = Split(OUTPUT_RANGE, ",")
        Case mdlCommon.MASTER_PARA
            arySetinfo = Split(MASTER_RANGE, ",")
        Case mdlCommon.HIDDEN_PARA
            arySetinfo = Split(HIDDEN_RANGE, ",")
    End Select
    
    FuncReadSetinfo = arySetinfo
    
End Function

'********************************************************************************
'* �������@�bFuncExtractFolderpath
'* �@�\�@�@�b�t�@�C���̐�΃p�X����t�H���_�̃p�X�𒊏o
'*-------------------------------------------------------------------------------
'* �߂�l�@�bString�F�t�H���_�̃p�X
'* �����@�@�b-
'********************************************************************************
Public Function FuncExtractFolderpath(strPath As String) As String
    Dim aryPath() As String

    aryPath = Split(strPath, "\")
    If UBound(aryPath) > 0 Then
        ReDim Preserve aryPath(UBound(aryPath) - 1)
    End If

    FuncExtractFolderpath = Join(aryPath, "\")
    
End Function

'********************************************************************************
'* �������@�bFuncShowBreakmessage
'* �@�\�@�@�b�������f�̃��b�Z�[�W�\��
'*-------------------------------------------------------------------------------
'* �߂�l�@�bBoolean�iTrue=���f����, False=���f���Ȃ��j
'* �����@�@�b-
'********************************************************************************
Public Function FuncShowBreakmessage() As Boolean
    Dim strMessage As String
    
    strMessage = "�����𒆒f���܂����H" & vbCrLf & _
                  vbCrLf & _
                  "�@�u�������v��I�������f���L�����Z�������ꍇ�ł��A" & vbCrLf & _
                  "�@���f�����^�C�~���O�ɂ���ăf�[�^�̕s�������������܂��B" & vbCrLf & _
                  vbCrLf & _
                  "�@�K���ŏ�����ď������Ă��������B"
    FuncShowBreakmessage = (MsgBox(strMessage, vbQuestion + vbYesNo, "�m�F") = vbYes)

End Function

'********************************************************************************
'* �������@�bFuncRetrieveMessage
'* �@�\�@�@�b���b�Z�[�W�̎擾
'*-------------------------------------------------------------------------------
'* �߂�l�@�bString�F���b�Z�[�W���e
'* �����@�@�blngCode�F�Ώۃ��x���R�[�h
'********************************************************************************
Public Function FuncRetrieveMessage(strCode As String) As String
    Dim aryMessages(10, 1) As String
    ' ����I���R�[�h
    aryMessages(0, 0) = "0"
    aryMessages(0, 1) = "����I���B"
    aryMessages(1, 0) = "10"
    aryMessages(1, 1) = "�����������s�B"
    aryMessages(2, 0) = "20"
    aryMessages(2, 1) = "�t�H���_�̍쐬�����������B"
    aryMessages(3, 0) = "999"
    aryMessages(3, 1) = ""
    ' �G���[���b�Z�[�W
    ' �����̓`�F�b�N
    aryMessages(4, 0) = "-111"
    aryMessages(4, 1) = "�K�{���ڂ������́B"
    ' ���݃`�F�b�N
    aryMessages(5, 0) = "-211"
    aryMessages(5, 1) = "�Q�Ƃł��Ȃ��t�@�C��������B"
    aryMessages(6, 0) = "-212"
    aryMessages(6, 1) = "�Q�Ƃł��Ȃ��t�H���_������B"
    ' 0���`�F�b�N
    aryMessages(7, 0) = "-311"
    aryMessages(7, 1) = "��荞�񂾃f�[�^��0���B"
    ' �}�X�^�[�̃G���[
    aryMessages(8, 0) = "-411"
    aryMessages(8, 1) = "�}�X�^�[�f�[�^�̕K�{���ڂ������́B"
    ' ���̑��G���[
    aryMessages(9, 0) = "-901"
    aryMessages(9, 1) = "���s���ɒ��f�B"
    aryMessages(10, 0) = "-999"
    aryMessages(10, 1) = "��O�������B"
    
    Dim lngCount As Long
    lngCount = 0
    For lngCount = LBound(aryMessages, 1) To UBound(aryMessages, 1)
        If aryMessages(lngCount, 0) = strCode Then
            FuncRetrieveMessage = aryMessages(lngCount, 1)
            Exit Function
        End If
    Next

    FuncRetrieveMessage = ""

End Function

' Sub�v���V�[�W��
'********************************************************************************
'* �������@�bSubSelectCell
'* �@�\�@�@�b�w�肵���Z����I��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b�Ȃ�
'* �����@�@�bstrSheetname�F�ΏۃV�[�g���AstrRange�F�ΏۃZ��
'********************************************************************************
Sub SubSelectCell(strSheetname As String, strRange As String)
    With Worksheets(strSheetname)
        .Select
        .Range(strRange).Select
    End With

End Sub

'********************************************************************************
'* �������@�bSubClearSheet
'* �@�\�@�@�b�V�[�g�̃N���A
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bstrSheetname�F�ΏۃV�[�g, strRange�F�J�n�Z���ʒu
'********************************************************************************
Sub SubClearSheet(strSheetname As String, strRange As String)
    Dim wbTarget As Workbook
    Dim wsTarget As Worksheet
    
    Dim rngTarget As Range
    Dim lngBeginrow As Long
    Dim lngBegincol As Long
    Dim lngEndrow As Long
    
    Set wbTarget = ThisWorkbook
    
    ' �ΏۃV�[�g���Ȃ��ꍇ�ł��G���[�𔭐������������𑱍s
    On Error Resume Next
    Set wsTarget = wbTarget.Worksheets(strSheetname)
    On Error GoTo 0
    
    Set rngTarget = Range(strRange)
    lngBeginrow = rngTarget.Row
    lngBegincol = rngTarget.Column
    
    ' �V�[�g������ꍇ
    If (IsExistsSheet(strSheetname)) Then
        With wsTarget
            lngEndrow = .Cells(Rows.Count, lngBegincol).End(xlUp).Row
            ' �\�ɒl���Ȃ��ꍇ�A�ő�s�����J�n�s�ɕύX
            If lngBeginrow > lngEndrow Then
                lngEndrow = lngBeginrow
            End If
            With .Range(.Cells(lngBeginrow, lngBegincol), .Cells(lngEndrow, Columns.Count))
                .ClearContents
            End With
            With .Range(.Cells(lngBeginrow, lngBegincol), .Cells(Rows.Count, Columns.Count))
                .Interior.ColorIndex = 0
                .Borders.LineStyle = False
            End With
        End With
    End If
    
    Set wbTarget = Nothing
    Set wsTarget = Nothing

End Sub

'********************************************************************************
'* �������@�bSubCopySheet
'* �@�\�@�@�bExcel�V�[�g�̃R�s�[
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�b�R�s�[���istrFmName=�V�[�g��, lngFmBeginrow=�J�n�s, lngFmBegincol=�J�n��j
'* �@�@�@�@�b�R�s�[��istrToName=�V�[�g��, lngToBeginrow=�J�n�s, lngToBegincol=�J�n��j
'********************************************************************************
Sub SubCopySheet(strFmName As String, lngFmBeginrow As Long, lngFmBegincol As Long, _
                 strToName As String, lngToBeginrow As Long, lngToBegincol As Long)
    Dim wbTarget As Workbook
    Dim wsFmSheet As Worksheet
    Dim wsToSheet As Worksheet
    
    Dim lngFmEndrow As Long
    Dim lngFmEndcol As Long
    Dim lngToEndrow As Long
    Dim lngToEndcol As Long
    
    Dim lngFmRow As Long
    Dim lngFmCol As Long
    Dim lngToRow As Long
    Dim lngToCol As Long
    
    Dim lngRow As Long
    
    Dim varFmCopydata As Variant
    Dim varToCopydata As Variant
    
    Set wbTarget = ThisWorkbook
    
    On Error Resume Next
    Set wsFmSheet = wbTarget.Worksheets(strFmName)
    Set wsToSheet = wbTarget.Worksheets(strToName)
    On Error GoTo 0
    
    ' �V�[�g������ꍇ
    If Not (wsFmSheet Is Nothing) And _
       Not (wsToSheet Is Nothing) Then
        With wsFmSheet
            lngFmEndrow = .Cells(Rows.Count, lngFmBegincol + 1).End(xlUp).Row
            If lngFmBeginrow > lngFmEndrow Then
                lngFmEndrow = lngFmBeginrow
            End If
            
            lngFmRow = lngFmBeginrow
            lngFmCol = lngFmBegincol
            lngToRow = lngToBeginrow
            lngToCol = lngToBegincol
            
            ReDim varToCopydata(lngFmEndrow - 1, 30)
            
            varFmCopydata = .Range(.Cells(lngFmRow, lngFmCol), .Cells(lngFmEndrow, lngFmCol + 30)).Value
        End With
        
        varToCopydata = varFmCopydata
        
        ' �ҏW�����i����΁j
        '   ���t��YYYY-MM-DD�ŕ\��
        For lngRow = 1 To UBound(varToCopydata)
            varToCopydata(lngRow, 30) = "'" & varToCopydata(lngRow, 30)
        Next
        
        With wsToSheet
            .Range(.Cells(lngToRow, lngToCol), .Cells(lngFmEndrow, lngToCol + 30)).Value = varToCopydata
            ' �r���̐ݒ�
            lngToEndrow = .Cells(Rows.Count, lngToBegincol).End(xlUp).Row
            lngToEndcol = .Cells(1, Columns.Count).End(xlToLeft).Column
            With .Range(.Cells(lngToBeginrow, lngToBegincol), .Cells(lngToEndrow, lngToEndcol))
                .Borders.LineStyle = True
            End With
        End With
    End If
    
    Set wsFmSheet = Nothing
    Set wsToSheet = Nothing
    Set wbTarget = Nothing
    
End Sub

'********************************************************************************
'* �������@�bSubLoadCsv
'* �@�\�@�@�bCSV�t�@�C���̎�荞��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bstrPath�F�Ώۃt�@�C���̐�΃p�X, lngCsvrow�FCSV�J�n�s,
'* �@�@�@�@�bstrSheetname�F�ΏۃV�[�g, strExcelrange�FExcel�J�n�Z���ʒu,
'* �@�@�@�@�b[�C��]strCharcode�F�����R�[�h
'********************************************************************************
Sub SubLoadCsv(strPath As String, lngCsvrow As Long, _
               strSheetname As String, strExcelrange As String, _
               Optional strCharcode As String = "UTF8")
    Dim qtTarget As QueryTable
    Dim rngExcel As Range
    Dim lngExcelbeginrow As Long
    Dim lngExcelbegincol As Long
    Dim lngEndrow As Long
    Set rngExcel = Range(strExcelrange)
    lngExcelbeginrow = rngExcel.Row
    lngExcelbegincol = rngExcel.Column
    
    With Worksheets(strSheetname)
        lngEndrow = .Cells(Rows.Count, lngExcelbegincol).End(xlUp).Row
    End With
    
    If lngEndrow > lngExcelbeginrow Then
        lngExcelbeginrow = lngEndrow + 1
    End If
    
    Set qtTarget = Worksheets(strSheetname).QueryTables.Add(Connection:="TEXT;" & strPath, _
        Destination:=Worksheets(strSheetname).Cells(lngExcelbeginrow, lngExcelbegincol))
    
    With qtTarget
        .TextFileCommaDelimiter = True                          ' �J���}��؂�̎w��
        .TextFileParseType = xlDelimited                        ' ��؂蕶���̌`��
        .TextFileTextQualifier = xlTextQualifierDoubleQuote     ' ���p������_�u���N�H�[�e�[�V�������w��
        If strCharcode = "UTF8" Then
            .TextFilePlatform = 65001                           ' �����R�[�hUTF-8���w��
        Else
            .TextFilePlatform = 932                             ' �����R�[�hShift_JIS���w��
        End If
       .TextFileStartRow = lngCsvrow                            ' �J�n�s�̎w��
       .RefreshStyle = xlOverwriteCells                         ' �Z���͒ǉ������㏑������
       .Refresh                                                 ' QueryTables�I�u�W�F�N�g���X�V���A�V�[�g��ɏo��
       .Delete                                                  ' QueryTables.Add���\�b�h�Ŏ�荞��CSV�Ƃ̐ڑ�������
    End With
    
    Set qtTarget = Nothing

End Sub

'********************************************************************************
'* �������@�bSubSaveCsv
'* �@�\�@�@�bCSV�t�@�C���̕ۑ��i�����R�[�h�FUTF-8�j
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bstrSheetname�F�ΏۃV�[�g, strExcelrange�FExcel�J�n�Z���ʒu,
'* �@�@�@�@�bstrSavepath�F�ۑ���, lngCsvrow�FCSV�J�n�s
'********************************************************************************
Sub SubSaveCsv(strSheetname As String, strExcelrange As String, _
               strSavepath As String, lngCsvbeginrow As Long)
    Dim rngExcel As Range
    Dim lngExcelbeginrow As Long
    Dim lngExcelbegincol As Long
    Dim lngEndrow As Long
    Dim lngEndcol As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strCsvdata As String
    Dim strLine As String
    Dim strValue As String
    Dim strDelimiter As String
    Dim aryItems() As String

    Set rngExcel = Range(strExcelrange)
    lngExcelbeginrow = rngExcel.Row
    lngExcelbegincol = rngExcel.Column

    Dim objUtf8 As Object
    Dim objNonbom As Object
    Set objUtf8 = CreateObject("ADODB.Stream")
    Set objNonbom = CreateObject("ADODB.Stream")

    ' CSV�`���̏���
    With Worksheets(strSheetname)
        lngEndrow = .Cells(Rows.Count, lngExcelbegincol).End(xlUp).Row
        lngEndcol = .Cells(lngExcelbeginrow - 1, Columns.Count).End(xlToLeft).Column
    End With
    
    If lngCsvbeginrow > lngEndrow Then
        lngEndrow = lngCsvbeginrow
    End If

    strDelimiter = ","

    For lngRow = lngExcelbeginrow To lngEndrow
        strLine = ""
        For lngCol = lngExcelbegincol To lngEndcol
            With Worksheets(strSheetname)
                strValue = .Cells(lngRow, lngCol).Value
            End With
            If strLine = "" Then
                strLine = strValue
            Else
                ReDim aryItems(1)
                aryItems(0) = strLine
                aryItems(1) = strValue
                strLine = Join(aryItems, strDelimiter)
            End If
        Next

        If strCsvdata = "" Then
            strCsvdata = strLine
        Else
            ReDim aryItems(1)
            aryItems(0) = strCsvdata
            aryItems(1) = strLine
            strCsvdata = Join(aryItems, vbCrLf)
        End If
    Next
    ' �ŏI�s�ɋ󕶎��s
    ReDim aryItems(1)
    aryItems(0) = strCsvdata
    aryItems(1) = ""
    strCsvdata = Join(aryItems, vbCrLf)

    ' �ۑ�
    With objUtf8
        .Charset = "UTF-8"
        .Open
        .WriteText strCsvdata
        .Position = 0
        .Type = 1                       ' Binary
        .Position = 3
    End With

    With objNonbom
        .Type = 1                       ' Binary
        .Open
        objUtf8.CopyTo objNonbom
        .SaveToFile strSavepath, 2      ' SaveCreateOverWrite
    End With

    objNonbom.Close
    objUtf8.Close

    Set objNonbom = Nothing
    Set objUtf8 = Nothing

End Sub

'********************************************************************************
'* �������@�bSubShowMessagebox
'* �@�\�@�@�b���b�Z�[�W�̕\��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�blngCode�F�Ώۃ��b�Z�[�W�R�[�h, [�C��]strAppend�F�ǉ����b�Z�[�W
'********************************************************************************
Sub SubShowMessagebox(lngCode As Long, Optional strAppend As String = "")
    Dim strMessage As String
    Dim lngLevel As Long
    Dim strLevel As String

    strMessage = FuncRetrieveMessage(CStr(lngCode))
    strMessage = strMessage & vbCrLf & strAppend & vbCrLf
    
    If Left(CStr(lngCode), 1) = "-" Then
        lngLevel = vbCritical
        strLevel = "�x��"
    Else
        lngLevel = vbInformation
        strLevel = "���"
    End If
    
    strMessage = strMessage & vbCrLf & "Message Code�F" & "[" & CStr(lngCode) & "]"
    
    MsgBox strMessage, vbOKOnly + lngLevel, strLevel
    
End Sub

'********************************************************************************
'* �������@�bSubDisplayMessage
'* �@�\�@�@�b�ʒm�p�̃��b�Z�[�W���V�[�g�\��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�blngCode�F�Ώۃ��b�Z�[�W�R�[�h�A[�C��]strAppend�F�ǉ����b�Z�[�W
'********************************************************************************
Public Sub SubDisplayMessage(lngCode As Long, Optional strAppend As String = "")
    Dim arySetinfo() As String
    arySetinfo = FuncReadSetinfo(mdlCommon.MAIN_PARA)
    
    Dim strSheetname As String
    Dim strRange As String
    Dim strMessage As String
    Dim strDatetime As String
    Dim lngCount As Long
    
    strDatetime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    
    With Worksheets(SETINFO_SHEETNAME)
        strSheetname = .Range(arySetinfo(0)).Value
        strRange = .Range(arySetinfo(1)).Value
    End With
    
    strMessage = FuncRetrieveMessage(CStr(lngCode))
    If Not (mdlCommon.IsEmptyText(strMessage)) Then
        strMessage = strDatetime & " " & _
                     strMessage & vbCrLf & _
                     strAppend
    End If
    
    With Worksheets(strSheetname)
        .Range(strRange).Font.ColorIndex = 11
        If Left(CStr(lngCode), 1) = "-" Then
            .Range(strRange).Font.ColorIndex = 3
        End If
        .Range(strRange).Value = strMessage
    End With
    
End Sub

'********************************************************************************
'* �������@�bSubWriteError
'* �@�\�@�@�b�G���[���ւ̏�������
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�blngCode�F�Ώۃ��b�Z�[�W�R�[�h�A[�C��]strAppend�F�ǉ����b�Z�[�W
'********************************************************************************
Public Sub SubWriteError(lngCode As Long, Optional strAppend As String = "")
    Dim strDatetime As String
    Dim strLevel As String
    Dim strMessage As String
    
    strDatetime = Format(Now, "yyyy/mm/dd hh:mm:ss")
    If Left(CStr(lngCode), 1) = "-" Then
        strLevel = "�x��"
    Else
        strLevel = "���"
    End If
    strMessage = FuncRetrieveMessage(CStr(lngCode))

    Dim arySetinfo() As String
    Dim strErrorsheet As String
    Dim strRange As String
    Dim rngTarget As Range
    Dim lngBeginrow As Long
    Dim lngBegincol As Long
    Dim lngEndrow As Long
    arySetinfo = mdlCommon.FuncReadSetinfo(mdlCommon.ERROR_PARA)
    With Worksheets(mdlCommon.SETINFO_SHEETNAME)
        strErrorsheet = .Range(arySetinfo(0)).Value
        strRange = .Range(arySetinfo(1)).Value
    End With

    Set rngTarget = Range(strRange)
    lngBeginrow = rngTarget.Row
    lngBegincol = rngTarget.Column

    With Worksheets(strErrorsheet)
        lngEndrow = .Cells(Rows.Count, lngBegincol).End(xlUp).Row

        .Cells(lngEndrow + 1, lngBegincol).Value = strDatetime
        .Cells(lngEndrow + 1, lngBegincol + 1).Value = strLevel
        .Cells(lngEndrow + 1, lngBegincol + 2).Value = strMessage
        .Cells(lngEndrow + 1, lngBegincol + 3).Value = strAppend

        With .Range(.Cells(lngBeginrow, lngBegincol), .Cells(lngEndrow + 1, lngBegincol + 3))
            .Borders.LineStyle = True
        End With
    End With
End Sub

'********************************************************************************
'* �������@�bSubDisplayStatusbar
'* �@�\�@�@�b�X�e�[�^�X�o�[�̕\��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bblValid�iTrue=�X�e�[�^�X�o�[�L��, False=�����j, [�C��]strAppend�F�ǉ����b�Z�[�W
'********************************************************************************
Sub SubDisplayStatusbar(blValid As Boolean, Optional strAppend As String)
    If blValid Then
        Application.StatusBar = "���s�� ..." & strAppend
    Else
        Application.StatusBar = False
    End If

End Sub

'********************************************************************************
'* �������@�bSubOnSpeedup
'* �@�\�@�@�bVBA�����X�s�[�h�A�b�v�ݒ�
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bblValid�iTrue=�������L��, False=����)
'********************************************************************************
Sub SubOnSpeedup(blValid As Boolean)
    Dim strPath As String
    If blValid Then
        ' ����������
        ' �J�[�\����ҋ@���ɐݒ�
        Application.Cursor = xlWait
    
        ' ��ʕ`�ʂ̒�~��ݒ�
        Application.ScreenUpdating = False
        ' �����v�Z�̒�~��ݒ�
        Application.Calculation = xlCalculationManual
        ' ���[�U����֎~��ݒ�
        Application.Interactive = False
    Else
        ' ��ʕ`�ʂ��ĊJ
        Application.ScreenUpdating = True
        ' �����v�Z���ĊJ
        Application.Calculation = xlCalculationAutomatic
        ' ���[�U������ĊJ
        Application.Interactive = True
        ' �J�����g�f�B���N�g���ړ�
        strPath = ThisWorkbook.Path
        ChDrive strPath
        ChDir strPath
        
        ' �J�[�\�������ɖ߂�
        Application.Cursor = xlDefault
    End If

End Sub

'********************************************************************************
'* �������@�bSubVisibleSheet
'* �@�\�@�@�b�V�[�g�\���^��\���̐ݒ�
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bstrClass�F�ΏۃO���[�v, blVisible�iTrue=�\��, False=��\���j
'********************************************************************************
Sub SubVisibleSheet(strClass As String, blVisible As Boolean)
    Dim arySetinfo() As String
    arySetinfo = FuncReadSetinfo(strClass)
    
    Dim lngCount As Long
    Dim strSheetname As String
    
    For lngCount = 0 To UBound(arySetinfo)
        strSheetname = Worksheets(SETINFO_SHEETNAME).Range(arySetinfo(lngCount)).Value
        Worksheets(strSheetname).Visible = blVisible
    Next

End Sub

'********************************************************************************
'* �������@�bSubOpenFolder
'* �@�\�@�@�b�w��t�H���_���J��
'*-------------------------------------------------------------------------------
'* �߂�l�@�b-
'* �����@�@�bstrPath�F�Ώۃt�H���_�̐�΃p�X
'********************************************************************************
Sub SubOpenFolder(strPath As String)
On Error GoTo CATCH
    Shell "C:\Windows\explorer.exe " & strPath, vbNormalFocus
CATCH:

End Sub
