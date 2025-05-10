Attribute VB_Name = "FormatBab"
Sub FormatBab()
    Dim r As Range
    Dim babStyle As Style
    Dim selectedText As String

    ' Cek apakah ada teks yang dipilih
    If Selection.Range.Text = "" Then
        MsgBox "Pilih teks dulu cuy untuk dibikin heading level 1!"
        Exit Sub
    End If

    ' Ambil teks yang dipilih dan simpan
    selectedText = Trim(Selection.Text)

    ' Buat range dari teks yang dipilih
    Set r = Selection.Range

    ' Cek apakah style Heading 1 sudah ada
    On Error Resume Next
    Set babStyle = ActiveDocument.Styles("Heading 1")
    If babStyle Is Nothing Then
        ' Buat Heading 1 jika belum ada
        Set babStyle = ActiveDocument.Styles.Add(Name:="Heading 1", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0

    ' Atur format Heading 1
    With babStyle.Font
        .Name = "Times New Roman"
        .Size = 14
        .Bold = True
        .Color = wdColorBlack
    End With
    With babStyle.ParagraphFormat
        .Alignment = wdAlignParagraphCenter
        .FirstLineIndent = 0
        .LeftIndent = 0
        .SpaceAfter = 0
    End With

    ' Buat multi-level list dengan format "BAB x"
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "BAB %1 "
        .NumberStyle = wdListNumberStyleArabic
        .TrailingCharacter = wdTrailingNone ' Follow number with: Nothing
        .NumberPosition = 0
        .TextPosition = 0
        .TabPosition = 0
        .ResetOnHigher = 0
        .StartAt = 1
        .LinkedStyle = babStyle.NameLocal
    End With

    ' Hapus teks yang dipilih agar tidak terduplikasi
    r.Text = ""

    ' Sisipkan heading level 1 x di posisi kursor
    r.InsertAfter " "
    r.Collapse Direction:=wdCollapseEnd
    r.InsertAfter selectedText

    ' Terapkan format heading
    r.Style = babStyle
    r.ParagraphFormat.Alignment = wdAlignParagraphCenter
    r.ParagraphFormat.FirstLineIndent = 0
    r.ParagraphFormat.LeftIndent = 0

    ' Terapkan numbering
    r.ListFormat.ApplyListTemplateWithLevel _
        listTemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToSelection, DefaultListBehavior:=wdWord10ListBehavior
End Sub




