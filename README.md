[DataTransfer.txt](https://github.com/user-attachments/files/16578749/DataTransfer.txt)Data Transfer with one button

If you have never used Add-ins, you are working slowly in Excel
Have you ever used add-ins in Excel? It's a useful method, but you need to know how to write code in VBA. When I started in the HR Department, I noticed that there was a lot of manual work. So I focused on learning VBA. VBA has a simple language. Do you want to learn how to write code? If not, you can use my code. But first, you should see how it works.
In this article, I am going to show you how you can use add-ins. Because this is our first step, let's dive in.
How to Add an Add-in in Excel

Step 1: Open Excel
Open Microsoft Excel on your computer.

Step 2: Go to the Add-ins Menu
Click on the File tab in the top-left corner.
Select Options at the bottom of the left-hand menu. This will open the Excel Options window.

Step 3: Access the Add-ins Options
In the Excel Options window, select Add-ins from the list on the left side.
At the bottom of the window, you'll see a drop-down menu labeled Manage. Select Excel Add-ins from this menu and click Go.

Step 4: Add or Browse for Add-ins
A new window called Add-Ins will pop up. Here you can see a list of available add-ins.
To enable an add-in, check the box next to its name and click OK.
If the add-in you need is not listed, click Browse to locate it on your computer. Once found, select the add-in file (usually with an .xlam extension) and click OK.

Step 5: Install the Add-in
Follow any additional prompts or installation instructions that may appear. Some add-ins may require additional steps or configurations.

![image](https://github.com/user-attachments/assets/2fbbd22e-6fd4-4438-be05-4412ae0025df)

Writing a Code in VBA
Now we should write a code in VBA. If you don't want to, you can find the attachments. When we write the code, we should save it in this Excel file here: C:\Users\User\AppData\Roaming\Microsoft\AddIns (For the second user, it should be your computer name). Our Excel file should be .xlsm
Adding an Add-in in Excel
Actually, we are close to the end. We have to follow the steps below.

Step 1: Open Excel
Launch Microsoft Excel on your computer.

Step 2: Access Excel Options
Click on the File tab located at the top-left corner of the screen.
From the File menu, select Options. This will open the Excel Options dialog box.

Step 3: Navigate to Add-Ins
In the Excel Options dialog box, click on Add-Ins from the list on the left side.

Step 4: Manage Add-Ins
At the bottom of the Add-Ins section, you will see a dropdown menu labeled Manage.
Select Excel Add-ins from this dropdown menu and click Go.

Step 5: Enable Add-Ins
In the Add-Ins dialog box that appears, you will see a list of available add-ins.
Check the boxes next to the add-ins you want to enable.
If the add-in you need is not listed, click Browse to locate it on your computer. Select the add-in file and click OK.
Click OK to close the Add-Ins dialog box.

![image](https://github.com/user-attachments/assets/836cdbe5-b931-4b24-ad43-76bb7ed34c9c)


If you have completed all the steps, you can add a button in your Excel. Then you can use this code fast and easily. Let's see how to add a button. If you want to see how it will look, you can find a screenshot below. (If you click the button, it will import the data from a different Excel sheet to the active sheet.)

![image](https://github.com/user-attachments/assets/ba5e14d9-23c2-4874-8332-dc09c4585a11)
![image](https://github.com/user-attachments/assets/ce4be2c0-ab35-45c2-ac37-cdbe447ac75b)

You can find Excel File here;
[UploSub DataTransfer()
    Dim kaynakWb As Workbook
    Dim hedefWb As Workbook
    Dim kaynakWs As Worksheet
    Dim hedefWs As Worksheet
    Dim kaynakLastRow As Long
    Dim hedefLastRow As Long
    Dim i As Long, j As Long
    Dim kaynakFilePath As String
    Dim colMatched As Boolean
    Dim listeCol As Long
    Dim dosyaAdi As String
    Dim sheetName As String
    Dim hedefColumns As Object
    
    ' Aktif olan sayfayı hedef olarak al
    Set hedefWs = ActiveSheet
    Set hedefWb = hedefWs.Parent
    
    ' Kullanıcıdan kaynak dosyayı seçmesini iste
    kaynakFilePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Kaynak dosyayı seçin")
    If kaynakFilePath = "False" Then Exit Sub
    
    Set kaynakWb = Workbooks.Open(kaynakFilePath)
    
    ' Eğer birden fazla sheet varsa, kullanıcıdan sheet adını sor
    If kaynakWb.Sheets.Count > 1 Then
        sheetName = Application.InputBox("Sheet adını girin", "Sheet Seçimi", kaynakWb.Sheets(1).Name, , , , , 2)
        On Error Resume Next
        Set kaynakWs = kaynakWb.Sheets(sheetName)
        On Error GoTo 0
        If kaynakWs Is Nothing Then
            MsgBox "Geçersiz sheet adı", vbCritical
            kaynakWb.Close SaveChanges:=False
            Exit Sub
        End If
    Else
        Set kaynakWs = kaynakWb.Sheets(1)
    End If
    
    dosyaAdi = Mid(kaynakWb.Name, 1, InStrRev(kaynakWb.Name, ".") - 1)
    
    kaynakLastRow = kaynakWs.Cells(kaynakWs.Rows.Count, "A").End(xlUp).Row
    
    ' İlk satırdaki tüm başlıkları ve onların son satırlarını belirle
    Set hedefColumns = CreateObject("Scripting.Dictionary")
    For j = 1 To hedefWs.Cells(1, hedefWs.Columns.Count).End(xlToLeft).Column
        hedefColumns(hedefWs.Cells(1, j).Value) = j
        If hedefWs.Cells(1, j).Value = "Liste" Then
            listeCol = j
        End If
    Next j
    
    ' Eğer "Liste" kolonu yoksa, oluştur
    If listeCol = 0 Then
        listeCol = hedefWs.Cells(1, hedefWs.Columns.Count).End(xlToLeft).Column + 1
        hedefWs.Cells(1, listeCol).Value = "Liste"
    End If
    
    ' Her bir başlık için kaynak verilerini kopyala
    For i = 1 To kaynakWs.Cells(1, kaynakWs.Columns.Count).End(xlToLeft).Column
        colMatched = False
        If Not IsEmpty(kaynakWs.Cells(1, i)) Then
            If hedefColumns.exists(kaynakWs.Cells(1, i).Value) Then
                colMatched = True
                j = hedefColumns(kaynakWs.Cells(1, i).Value)
                hedefLastRow = hedefWs.Cells(hedefWs.Rows.Count, j).End(xlUp).Row + 1
                kaynakWs.Range(kaynakWs.Cells(2, i), kaynakWs.Cells(kaynakLastRow, i)).Copy
                hedefWs.Cells(hedefLastRow, j).PasteSpecial Paste:=xlPasteValues
                
                With hedefWs.Cells(hedefLastRow, listeCol).Resize(kaynakLastRow - 1)
                    .Value = dosyaAdi
                End With
                Application.CutCopyMode = False
            End If
        End If
    Next i
    
    ' Gerekiyorsa sayısal değerleri güncelle
    Dim toplam As Double
    Dim sayiAdedi As Long
    toplam = 0
    sayiAdedi = 0
    
    For i = 2 To hedefWs.Cells(hedefWs.Rows.Count, "A").End(xlUp).Row
        If IsNumeric(hedefWs.Cells(i, "A").Value) Then
            hedefWs.Cells(i, "A").Value = hedefWs.Cells(i, "A").Value * 1
            toplam = toplam + hedefWs.Cells(i, "A").Value
            sayiAdedi = sayiAdedi + 1
        End If
    Next i
    
    kaynakWb.Close SaveChanges:=False

End Sub

ading DataTransfer.txt…]()

