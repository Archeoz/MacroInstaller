Attribute VB_Name = "FormatSubBab"
Sub FormatSubBab()
    Dim r As Range
    Dim babStyle As Style
    Dim selectedText As String

    ' Cek apakah ada teks yang dipilih
    If Selection.Range.Text = "" Then
        MsgBox "Pilih teks dulu cuy untuk dibikin heading level 2!"
        Exit Sub
    End If

    ' Ambil teks yang dipilih dan simpan
    selectedText = Trim(Selection.Text)

    ' Buat range dari teks yang dipilih
    Set r = Selection.Range

    ' Cek apakah style Heading 2 sudah ada
    On Error Resume Next
    Set babStyle = ActiveDocument.Styles("Heading 2")
    If babStyle Is Nothing Then
        ' Buat Heading 2 jika belum ada
        Set babStyle = ActiveDocument.Styles.Add(Name:="Heading 2", Type:=wdStyleTypeParagraph)
    End If
    On Error GoTo 0

    ' Atur format Heading 2
    With babStyle.Font
        .Name = "Times New Roman"
        .Size = 14
        .Bold = True
        .Color = wdColorBlack
    End With
    With babStyle.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .FirstLineIndent = 0
        .LeftIndent = 36 ' Menjorok lebih jauh daripada Heading 1
        .SpaceAfter = 0
    End With

    ' Buat multi-level list dengan format "BAB x.x"
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2 "
        .NumberStyle = wdListNumberStyleArabic
        .TrailingCharacter = wdTrailingNone ' Follow number with: Nothing
        .NumberPosition = 0
        .TextPosition = 0
        .TabPosition = 36 ' Menyesuaikan indentasi
        .ResetOnHigher = 1 ' Reset numbering ke 1 ketika Heading Level 1 berubah
        .StartAt = 1
        .LinkedStyle = babStyle.NameLocal
    End With

    ' Hapus teks yang dipilih agar tidak terduplikasi
    r.Text = ""

    ' Sisipkan heading level 2 x.x di posisi kursor
    r.InsertAfter " "
    r.Collapse Direction:=wdCollapseEnd
    r.InsertAfter selectedText

    ' Terapkan format heading
    r.Style = babStyle
    r.ParagraphFormat.Alignment = wdAlignParagraphLeft
    r.ParagraphFormat.FirstLineIndent = 0
    r.ParagraphFormat.LeftIndent = 36

    ' Terapkan numbering
    r.ListFormat.ApplyListTemplateWithLevel _
        listTemplate:=ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=False, ApplyTo:=wdListApplyToSelection, DefaultListBehavior:=wdWord10ListBehavior
End Sub




