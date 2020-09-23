VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      DragMode        =   1  'Automatic
      Height          =   3135
      Left            =   1575
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3075
      ScaleWidth      =   0
      TabIndex        =   2
      Top             =   0
      Width           =   60
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3195
      Left            =   1635
      TabIndex        =   1
      Top             =   0
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   5636
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"Form1.frx":0000
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   3200
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   5636
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'Picture1.BackColor = &H8000000F
    
    'add sample nodes
    TreeView1.Nodes.Add , , "key1", "Treeview1"
    TreeView1.Nodes.Add , tvwChild, , "Sample Node"
End Sub

Private Sub Form_Resize()
    'set properties for top,height of
    'treeview
    'picturebox
    'richtextbox
    Picture1.Height = Me.ScaleHeight
    TreeView1.Height = Me.ScaleHeight
    RichTextBox1.Height = Me.ScaleHeight
    RichTextBox1.Width = Me.ScaleWidth - TreeView1.Width - 30
End Sub

Private Sub RichTextBox1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
If X > 0 Then
    Picture1.Left = Picture1.Left + X
    TreeView1.Width = TreeView1.Width + X
    RichTextBox1.Left = Picture1.Left + Picture1.Width
    RichTextBox1.Width = Abs(RichTextBox1.Left - Me.ScaleWidth)
End If
End Sub

Private Sub TreeView1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Picture1.Left = X
    TreeView1.Width = X
    RichTextBox1.Left = X + Picture1.Width
    RichTextBox1.Width = Abs(RichTextBox1.Left - Me.ScaleWidth)
End Sub
