VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search File(s)"
   ClientHeight    =   5700
   ClientLeft      =   2205
   ClientTop       =   1515
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      Caption         =   "Search results"
      Height          =   5655
      Left            =   2520
      TabIndex        =   7
      Top             =   0
      Width           =   5055
      Begin VB.ListBox List1 
         Height          =   5325
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4815
      End
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   3
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Se&arch Now"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   3480
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Choose drive"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Choose a directory to search"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

If Trim(Text2.Text) = "" Then
MsgBox "You must specify a Filename"
Exit Sub
Else
File1.Pattern = Text2.Text
End If

Drive1.Visible = False
Dir1.Visible = False
DoEvents
SearchSubs
Drive1.Visible = True
Dir1.Visible = True
Drive1.SetFocus
MsgBox "There is " & List1.ListCount & " file(s) in your search", , "File(s) found"

End Sub


Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
MsgBox "Don't forget to vote Please!"
End Sub

Private Sub Text2_GotFocus()
List1.Clear 'This clear the list
End Sub

Private Sub SearchSubs()
Dim el_diR(1 To 15) As String
Dim The_File As String
Dim a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p As Integer
On Error Resume Next

'you can search multiple files separating like this:
'"file1.exe;file2.txt;file3.zip"
'the search begins
'in the search there is only one basic estructure

File1.Path = Dir1.Path 'sets the filebox path with the dir's path
el_diR(1) = Dir1.Path 'the first array gots the path to don't lose the parents directory

'this estructure only gets the file(s) if it found it
        For i = 0 To File1.ListCount - 1
            If Len(Dir1.Path) > 3 Then
                The_File = Dir1.Path & "\" & File1.List(i)
                Else
                The_File = Dir1.Path & File1.List(i)
                End If
        Form1.List1.AddItem The_File
        Next i

        
For a = 0 To (Dir1.ListCount - 1)
DoEvents 'gives a refresh to the OS
File1.Path = Dir1.List(a) 'gets the specific subdirectory
Dir1.Path = File1.Path 'sets the dir path same that the file path

                For i = 0 To File1.ListCount - 1
                    If Len(Dir1.Path) > 3 Then
                    The_File = Dir1.Path & "\" & File1.List(i)
                    Else
                    The_File = Dir1.Path & File1.List(i)
                    End If
                Form1.List1.AddItem The_File
                Next i
           
     If Dir1.ListCount >= 1 Then 'if there is some more subdirectories...
     el_diR(2) = Dir1.Path
     
        For b = 0 To Dir1.ListCount - 1
            File1.Path = Dir1.List(b) 'set the specified subdirectory
            Dir1.Path = File1.Path
            
            For i = 0 To File1.ListCount - 1
                    If Len(Dir1.Path) > 3 Then
                    The_File = Dir1.Path & "\" & File1.List(i)
                    Else
                    The_File = Dir1.Path & File1.List(i)
                    End If
            Form1.List1.AddItem The_File
            Next i
                          
                    If Dir1.ListCount >= 1 Then
                    el_diR(3) = Dir1.Path
                    
                        For c = 0 To Dir1.ListCount - 1
                        
                        File1.Path = Dir1.List(c)
                        Dir1.Path = File1.Path
        
                            For i = 0 To File1.ListCount - 1
                                If Len(Dir1.Path) > 3 Then
                                The_File = Dir1.Path & "\" & File1.List(i)
                                Else
                                The_File = Dir1.Path & File1.List(i)
                                End If
                            Form1.List1.AddItem The_File
                            Next i
                            
                                        If Dir1.ListCount >= 1 Then
                                        el_diR(4) = Dir1.Path
                                        
                                            For d = 0 To Dir1.ListCount - 1
                                            
                                            File1.Path = Dir1.List(d)
                                            Dir1.Path = File1.Path
        
                                                    For i = 0 To File1.ListCount - 1
                                                        If Len(Dir1.Path) > 3 Then
                                                        The_File = Dir1.Path & "\" & File1.List(i)
                                                        Else
                                                        The_File = Dir1.Path & File1.List(i)
                                                        End If
                                                    Form1.List1.AddItem The_File
                                                    Next i
                                                                
                                                                If Dir1.ListCount >= 1 Then
                                                                el_diR(5) = Dir1.Path
                    
                                                                        For e = 0 To Dir1.ListCount - 1
                                                                            
                                                                            File1.Path = Dir1.List(e)
                                                                            Dir1.Path = File1.Path
        
                                                                                    For i = 0 To File1.ListCount - 1
                                                                                        If Len(Dir1.Path) > 3 Then
                                                                                        The_File = Dir1.Path & "\" & File1.List(i)
                                                                                        Else
                                                                                        The_File = Dir1.Path & File1.List(i)
                                                                                        End If
                                                                                    Form1.List1.AddItem The_File
                                                                                    Next i
                                                                                    
                                                                                            If Dir1.ListCount >= 1 Then
                                                                                            el_diR(6) = Dir1.Path
                    
                                                                                                    For f = 0 To Dir1.ListCount - 1
                                                                                                       
                                                                                                        File1.Path = Dir1.List(f)
                                                                                                        Dir1.Path = File1.Path
        
                                                                                                                For i = 0 To File1.ListCount - 1
                                                                                                                    If Len(Dir1.Path) > 3 Then
                                                                                                                    The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                    Else
                                                                                                                    The_File = Dir1.Path & File1.List(i)
                                                                                                                    End If
                                                                                                                Form1.List1.AddItem The_File
                                                                                                                Next i
                                                                                                                        
                                                                                                                        If Dir1.ListCount >= 1 Then
                                                                                                                        el_diR(7) = Dir1.Path
                    
                                                                                                                                For g = 0 To Dir1.ListCount - 1
                                                                                                                                   
                                                                                                                                    File1.Path = Dir1.List(g)
                                                                                                                                    Dir1.Path = File1.Path
        
                                                                                                                                        For i = 0 To File1.ListCount - 1
                                                                                                                                            If Len(Dir1.Path) > 3 Then
                                                                                                                                            The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                            Else
                                                                                                                                            The_File = Dir1.Path & File1.List(i)
                                                                                                                                            End If
                                                                                                                                        Form1.List1.AddItem The_File
                                                                                                                                        Next i
                                                                                                                                                If Dir1.ListCount >= 1 Then
                                                                                                                                                el_diR(8) = Dir1.Path
                    
                                                                                                                                                    For h = 0 To Dir1.ListCount - 1
                                                                                                                                                        
                                                                                                                                                        File1.Path = Dir1.List(h)
                                                                                                                                                        Dir1.Path = File1.Path
        
                                                                                                                                                            For i = 0 To File1.ListCount - 1
                                                                                                                                                                If Len(Dir1.Path) > 3 Then
                                                                                                                                                                The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                Else
                                                                                                                                                                The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                End If
                                                                                                                                                            Form1.List1.AddItem The_File
                                                                                                                                                            Next i
                                                                                                                                                                If Dir1.ListCount >= 1 Then
                                                                                                                                                                el_diR(9) = Dir1.Path
                    
                                                                                                                                                                    For j = 0 To Dir1.ListCount - 1
                                                                                                                                                                        
                                                                                                                                                                        File1.Path = Dir1.List(j)
                                                                                                                                                                        Dir1.Path = File1.Path
        
                                                                                                                                                                            For i = 0 To File1.ListCount - 1
                                                                                                                                                                                If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                Else
                                                                                                                                                                                The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                End If
                                                                                                                                                                            Form1.List1.AddItem The_File
                                                                                                                                                                            Next i
                                                                                                                                                                                If Dir1.ListCount >= 1 Then
                                                                                                                                                                                el_diR(10) = Dir1.Path
                    
                                                                                                                                                                                        For k = 0 To Dir1.ListCount - 1
                                                                                                                                                                                            
                                                                                                                                                                                            File1.Path = Dir1.List(k)
                                                                                                                                                                                            Dir1.Path = File1.Path
        
                                                                                                                                                                                            For i = 0 To File1.ListCount - 1
                                                                                                                                                                                                If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                                The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                                Else
                                                                                                                                                                                                The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                                End If
                                                                                                                                                                                            Form1.List1.AddItem The_File
                                                                                                                                                                                            Next i
                                                                                                                                                                                                If Dir1.ListCount >= 1 Then
                                                                                                                                                                                                el_diR(11) = Dir1.Path
                    
                                                                                                                                                                                                    For l = 0 To Dir1.ListCount - 1
                                                                                                                                                                                                        
                                                                                                                                                                                                        File1.Path = Dir1.List(l)
                                                                                                                                                                                                        Dir1.Path = File1.Path
        
                                                                                                                                                                                                        For i = 0 To File1.ListCount - 1
                                                                                                                                                                                                            If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                                            The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                                            Else
                                                                                                                                                                                                            The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                                            End If
                                                                                                                                                                                                        Form1.List1.AddItem The_File
                                                                                                                                                                                                        Next i
                                                                                                                                                                                                            If Dir1.ListCount >= 1 Then
                                                                                                                                                                                                            el_diR(12) = Dir1.Path
                    
                                                                                                                                                                                                                For m = 0 To Dir1.ListCount - 1
                                                                                                                                                                                                                    
                                                                                                                                                                                                                    File1.Path = Dir1.List(m)
                                                                                                                                                                                                                    Dir1.Path = File1.Path
        
                                                                                                                                                                                                                        For i = 0 To File1.ListCount - 1
                                                                                                                                                                                                                            If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                                                            The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                            The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                        Form1.List1.AddItem The_File
                                                                                                                                                                                                                        Next i
                                                                                                                                                                                                                            If Dir1.ListCount >= 1 Then
                                                                                                                                                                                                                            el_diR(13) = Dir1.Path
                    
                                                                                                                                                                                                                                For n = 0 To Dir1.ListCount - 1
                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                    File1.Path = Dir1.List(n)
                                                                                                                                                                                                                                    Dir1.Path = File1.Path
        
                                                                                                                                                                                                                                        For i = 0 To File1.ListCount - 1
                                                                                                                                                                                                                                            If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                                                                            The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                            The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                        Form1.List1.AddItem The_File
                                                                                                                                                                                                                                        Next i
                                                                                                                                                                                                                                            If Dir1.ListCount >= 1 Then
                                                                                                                                                                                                                                            el_diR(14) = Dir1.Path
                    
                                                                                                                                                                                                                                                For o = 0 To Dir1.ListCount - 1
                                                                                                                                                                                                                                                    
                                                                                                                                                                                                                                                    File1.Path = Dir1.List(o)
                                                                                                                                                                                                                                                    Dir1.Path = File1.Path
        
                                                                                                                                                                                                                                                        For i = 0 To File1.ListCount - 1
                                                                                                                                                                                                                                                            If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                                                                                            The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                                            The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                                        Form1.List1.AddItem The_File
                                                                                                                                                                                                                                                        Next i
                                                                                                                                                                                                                                                            If Dir1.ListCount >= 1 Then 'last structure
                                                                                                                                                                                                                                                            el_diR(15) = Dir1.Path
                    
                                                                                                                                                                                                                                                            For p = 0 To Dir1.ListCount - 1
                                                                                                                                                                                                                                                              File1.Path = Dir1.List(p)
                                                                                        
                                                                                                                                                                                                                                                        For i = 0 To File1.ListCount - 1
                                                                                                                                                                                                                                                            If Len(Dir1.Path) > 3 Then
                                                                                                                                                                                                                                                            The_File = Dir1.Path & "\" & File1.List(i)
                                                                                                                                                                                                                                                            Else
                                                                                                                                                                                                                                                            The_File = Dir1.Path & File1.List(i)
                                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                                        Form1.List1.AddItem The_File
                                                                                                                                                                                                                                                        Next i
                                                                                                                                                                                                                                                        
                                                                                                                                                                                                                                                            Dir1.Path = el_diR(15)
                                                                                                                                                                                                                                                            Next p
                                                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                                                            
                                                                                                                                                                                                                                                    Dir1.Path = el_diR(14)
                                                                                                                                                                                                                                                    Next o
                                                                                                                                                                                                                                                End If
                                                                                                                                                                                                                                Dir1.Path = el_diR(13)
                                                                                                                                                                                                                                Next n
                                                                                                                                                                                                                            End If
                                                                                                                                                                                                                Dir1.Path = el_diR(12)
                                                                                                                                                                                                                Next m
                                                                                                                                                                                                            End If
                                                                                                                                                                                                    Dir1.Path = el_diR(11)
                                                                                                                                                                                                    Next l
                                                                                                                                                                                                End If
                                                                                                                                                                                        Dir1.Path = el_diR(10)
                                                                                                                                                                                        Next k
                                                                                                                                                                                End If
                                                                                                                                                                    Dir1.Path = el_diR(9)
                                                                                                                                                                    Next j
                                                                                                                                                                End If
                                                                                                                                                    Dir1.Path = el_diR(8)
                                                                                                                                                    Next h
                                                                                                                                                End If
                                                                                                                                Dir1.Path = el_diR(7)
                                                                                                                                Next g
                                                                                                                            End If
                                                                                                    Dir1.Path = el_diR(6)
                                                                                                    Next f
                                                                                                End If
                                                                                            
                                                                        Dir1.Path = el_diR(5)
                                                                        Next e
                                                                    End If
                                            Dir1.Path = el_diR(4)
                                            Next d
                                        End If
                        Dir1.Path = el_diR(3)
                        Next c
                    End If
             
       Dir1.Path = el_diR(2)     'gives the second old path to the dir
       Next b
     End If
   
                
Dir1.Path = el_diR(1) 'gives the dir path the old path saved in the first variable
Next a

Dir1.Path = el_diR(1) 'this only sets the first old directory path in the finish

End Sub

