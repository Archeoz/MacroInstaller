Sub BuatDaftarIsi()
    Dim tocRange As Range
    Dim toc As TableOfContents
    Dim para As Paragraph
    Dim tocPosition As Range

    ' 1) Hapus semua TOC yang ada (kalau ada)
    On Error Resume Next
    ActiveDocument.TablesOfContents(1).Delete
    On Error GoTo 0

    ' 2) Atur format judul "DAFTAR ISI"
    With Selection.Paragraphs(1).Range
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = wdColorBlack
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
    End With

    ' 3) Simpan posisi kursor di paragraf kosong
    Set tocPosition = Selection.Range.Duplicate
    
    ' 4) Sisipkan tulisan "DAFTAR ISI"
    Selection.TypeText Text:="DAFTAR ISI"
    
    ' 4) Sisipkan paragraf kosong
    Selection.TypeParagraph

    ' 5) Sisipkan tabel daftar isi otomatis (Automatic Table 2)
    Set tocRange = Selection.Range
    Set toc = ActiveDocument.TablesOfContents.Add( _
        Range:=tocRange, _
        UseHeadingStyles:=True, _
        UpperHeadingLevel:=1, _
        LowerHeadingLevel:=3, _
        IncludePageNumbers:=True, _
        RightAlignPageNumbers:=True, _
        UseHyperlinks:=True)

    ' 6) Format semua paragraf dalam daftar isi
    For Each para In toc.Range.Paragraphs
        With para.Range
            .Font.Name = "Times New Roman"
            .Font.Size = 12
            .Font.Bold = False
            .Font.Color = wdColorBlack
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
        End With
    Next para

    ' 7) Update daftar isi
    toc.Update

    ' 8) Kembalikan kursor ke paragraf kosong setelah judul
    tocPosition.Select
End Sub




