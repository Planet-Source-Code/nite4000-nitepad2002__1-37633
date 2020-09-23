VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "NitePad2002"
   ClientHeight    =   7080
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9945
   HelpContextID   =   10
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   1440
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComCtl3.CoolBar cbrBar 
      Align           =   1  'Align Top
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   10
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   2646
      BandCount       =   6
      _CBWidth        =   9945
      _CBHeight       =   1500
      _Version        =   "6.7.8862"
      Child1          =   "tbrStandard"
      MinHeight1      =   330
      Width1          =   30
      NewRow1         =   0   'False
      Child2          =   "picFont"
      MinHeight2      =   360
      Width2          =   3300
      NewRow2         =   -1  'True
      Child3          =   "tbrFormat"
      MinHeight3      =   330
      Width3          =   2175
      NewRow3         =   0   'False
      Child4          =   "tbrFile"
      MinHeight4      =   330
      Width4          =   2850
      NewRow4         =   -1  'True
      Visible4        =   0   'False
      Child5          =   "tbrEdit"
      MinHeight5      =   330
      Width5          =   5085
      NewRow5         =   0   'False
      Visible5        =   0   'False
      Child6          =   "tbrWindow"
      MinHeight6      =   330
      Width6          =   2205
      NewRow6         =   -1  'True
      Visible6        =   0   'False
      Begin VB.PictureBox picFont 
         BorderStyle     =   0  'None
         Height          =   360
         HelpContextID   =   600
         Left            =   165
         ScaleHeight     =   360
         ScaleWidth      =   3105
         TabIndex        =   8
         Top             =   390
         WhatsThisHelpID =   10
         Width           =   3105
         Begin VB.ComboBox cboFontSize 
            Height          =   315
            HelpContextID   =   610
            Left            =   2250
            TabIndex        =   10
            Top             =   23
            WhatsThisHelpID =   10
            Width           =   690
         End
         Begin VB.ComboBox cboFontName 
            Height          =   315
            HelpContextID   =   620
            Left            =   75
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   23
            WhatsThisHelpID =   10
            Width           =   2115
         End
      End
      Begin MSComctlLib.Toolbar tbrWindow 
         Height          =   330
         Left            =   165
         TabIndex        =   7
         Top             =   1140
         WhatsThisHelpID =   10
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         HelpContextID   =   630
         Style           =   1
         ImageList       =   "imgWindow"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "THorizontally"
               Object.ToolTipText     =   "Tile windows horizontally"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "TVertically"
               Object.ToolTipText     =   "Tile windows vertically"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cascade"
               Object.ToolTipText     =   "Cascade windows"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Minimize"
               Object.ToolTipText     =   "Minimize windows"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Restore"
               Object.ToolTipText     =   "Restore windows"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Close"
               Object.ToolTipText     =   "Close the active window"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
         Begin VB.OLE OLE1 
            Class           =   "Paint.Picture"
            Height          =   255
            HelpContextID   =   640
            Left            =   2640
            OleObjectBlob   =   "frmMDI.frx":030A
            TabIndex        =   14
            Top             =   0
            WhatsThisHelpID =   10
            Width           =   855
         End
      End
      Begin MSComctlLib.Toolbar tbrFile 
         Height          =   330
         Left            =   165
         TabIndex        =   6
         Top             =   780
         WhatsThisHelpID =   10
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         HelpContextID   =   650
         Style           =   1
         ImageList       =   "imgFile"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Create a new document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open an existing document"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save the active document"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "SaveAll"
               Object.ToolTipText     =   "Save all documents"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Printsetup"
               Object.ToolTipText     =   "printsetup"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print the active document"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Email"
               Description     =   "Email"
               Object.ToolTipText     =   "Email"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrEdit 
         Height          =   330
         Left            =   3045
         TabIndex        =   5
         Top             =   780
         WhatsThisHelpID =   10
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         HelpContextID   =   660
         Style           =   1
         ImageList       =   "imgStandard"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   10
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut the selection to the Clipboard"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy the selection to the Clipboard"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Insert Clipboard contents"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo the last action"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo the previously undone action"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Outdent"
               Object.ToolTipText     =   "Reduce indentation of selected lines"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Indent"
               Object.ToolTipText     =   "Increase indentation of selected lines"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrFormat 
         Height          =   330
         Left            =   3495
         TabIndex        =   4
         Top             =   405
         WhatsThisHelpID =   10
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         HelpContextID   =   670
         Style           =   1
         ImageList       =   "imgFormat"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find the specified text"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paint"
               Description     =   "paint"
               Object.ToolTipText     =   "Paint"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bold"
               Object.ToolTipText     =   "Make selected text bold"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Italic"
               Object.ToolTipText     =   "Make selected text italic"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Underline"
               Object.ToolTipText     =   "Make selected text underline"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Left"
               Object.ToolTipText     =   "Align selected lines to left"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Center"
               Object.ToolTipText     =   "Align selected lines to center"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Right"
               Object.ToolTipText     =   "Align selected lines to right"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bullet"
               Object.ToolTipText     =   "Set bullets"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tbrStandard 
         Height          =   330
         Left            =   165
         TabIndex        =   3
         Top             =   30
         WhatsThisHelpID =   10
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         HelpContextID   =   680
         Style           =   1
         ImageList       =   "imgStandard"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   22
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "Create a new document"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open an existing document"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save the active document"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print the active document"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "FullScreen"
               Object.ToolTipText     =   "View the active document in full screen mode"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "WordWrap"
               Object.ToolTipText     =   "Toggle Word Wrap"
               ImageIndex      =   6
               Style           =   1
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut the selection to the Clipboard"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy the selection to the Clipboard"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Insert Clipboard contents"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo the last action"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo the previously undone action"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Outdent"
               Object.ToolTipText     =   "Reduce indentation of selected lines"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Indent"
               Object.ToolTipText     =   "Increase indentation of selected lines"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Help"
               Object.ToolTipText     =   "Display help contents"
               ImageIndex      =   14
            EndProperty
            BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Tips"
               Description     =   "Tips"
               Object.ToolTipText     =   "Tips"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgFormat 
      Left            =   720
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18322
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1847E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":185DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18736
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18892
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":189B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18B0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18C6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18DC6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":18F22
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":19464
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgWindow 
      Left            =   120
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1A4B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1A612
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1A76E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1A8D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1AA32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1AB8E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFile 
      Left            =   720
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1ACEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1AE46
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1AFA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1B53E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1B69A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1BC36
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1C112
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1C26E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1C5C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1CC12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgStandard 
      Left            =   120
      Top             =   2040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1CF64
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1D0C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1D21C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1D378
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1D4D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1D630
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1D9CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1DB28
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1DC84
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1DDE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1DF3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1E098
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1E274
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1E450
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1E5AC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrRuler 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   1500
      WhatsThisHelpID =   10
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   741
      BandCount       =   1
      _CBWidth        =   9945
      _CBHeight       =   420
      _Version        =   "6.7.8862"
      MinHeight1      =   360
      Width1          =   1440
      NewRow1         =   0   'False
      Begin VB.PictureBox picInsert 
         Height          =   255
         HelpContextID   =   740
         Left            =   1125
         ScaleHeight     =   195
         ScaleWidth      =   390
         TabIndex        =   12
         Top             =   75
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   450
      End
      Begin RichTextLib.RichTextBox rtfTemp 
         CausesValidation=   0   'False
         Height          =   240
         HelpContextID   =   750
         Left            =   825
         TabIndex        =   11
         Top             =   75
         Visible         =   0   'False
         WhatsThisHelpID =   10
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   423
         _Version        =   393217
         TextRTF         =   $"frmMDI.frx":1E8FE
      End
      Begin VB.PictureBox picRuler 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         HelpContextID   =   760
         Left            =   120
         Picture         =   "frmMDI.frx":1E980
         ScaleHeight     =   270
         ScaleWidth      =   11490
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   80
         WhatsThisHelpID =   10
         Width           =   11490
      End
   End
   Begin MSComctlLib.StatusBar SB 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   13
      Top             =   6810
      WhatsThisHelpID =   10
      Width           =   9945
      _ExtentX        =   17542
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "For Help, press F1"
            TextSave        =   "For Help, press F1"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "Line #:"
            TextSave        =   "Line #:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "Total lines:"
            TextSave        =   "Total lines:"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1588
            MinWidth        =   1587
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1766
            TextSave        =   "9:35 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      HelpContextID   =   210
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
         HelpContextID   =   320
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
         Shortcut        =   +^{F12}
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
         HelpContextID   =   330
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "Page Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "&PrintPreview"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
         HelpContextID   =   340
      End
      Begin VB.Menu mnufilesendemail 
         Caption         =   "&Send Email"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      HelpContextID   =   220
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
         HelpContextID   =   350
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
         HelpContextID   =   360
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "&Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
         HelpContextID   =   370
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      HelpContextID   =   230
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSearchBar0 
         Caption         =   "-"
         HelpContextID   =   380
      End
      Begin VB.Menu mnuSearchReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSearchBar1 
         Caption         =   "-"
         HelpContextID   =   390
      End
      Begin VB.Menu mnuSearchGoTo 
         Caption         =   "&Go To..."
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      HelpContextID   =   240
      Begin VB.Menu mnuViewwebbrowser 
         Caption         =   "&Web Broswer"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "&Status Bar"
      End
      Begin VB.Menu mnuViewRuler 
         Caption         =   "&Ruler"
      End
      Begin VB.Menu mnuViewFullScreen 
         Caption         =   "&Full Screen"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
         HelpContextID   =   400
      End
      Begin VB.Menu mnuViewToolbars 
         Caption         =   "T&oolbars"
         Begin VB.Menu mnuViewToolbarStandard 
            Caption         =   "&Standard"
         End
         Begin VB.Menu mnuViewToolbarFile 
            Caption         =   "&File"
         End
         Begin VB.Menu mnuViewToolbarEdit 
            Caption         =   "&Edit"
         End
         Begin VB.Menu mnuViewToolbarFormat 
            Caption         =   "Fo&rmat"
         End
         Begin VB.Menu mnuViewToolbarFont 
            Caption         =   "F&ont"
         End
         Begin VB.Menu mnuViewToolbarWindow 
            Caption         =   "&Window"
         End
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
         HelpContextID   =   410
      End
      Begin VB.Menu mnuViewStayonTop 
         Caption         =   "Stay on &Top"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
         HelpContextID   =   420
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&No Wrap"
         Index           =   0
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "&Word Wrap"
         Index           =   1
      End
      Begin VB.Menu mnuViewMode 
         Caption         =   "WY&SIWYG"
         Index           =   2
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
         HelpContextID   =   430
      End
      Begin VB.Menu mnuViewDocumentProperties 
         Caption         =   "Document P&roperties"
         Shortcut        =   +{F3}
      End
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "&Insert"
      HelpContextID   =   250
      Begin VB.Menu mnuInsertPicture 
         Caption         =   "&Picture..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuInsertBar0 
         Caption         =   "-"
         HelpContextID   =   440
      End
      Begin VB.Menu mnuInsertTimeDate 
         Caption         =   "Time and &Date..."
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuInsertBar1 
         Caption         =   "-"
         HelpContextID   =   450
      End
      Begin VB.Menu mnuInsertTextFile 
         Caption         =   "&Text File..."
      End
      Begin VB.Menu mnuInsertPathandFile 
         Caption         =   "Path and &File"
      End
      Begin VB.Menu mnuInsertBar2 
         Caption         =   "-"
         HelpContextID   =   460
      End
      Begin VB.Menu mnuInsertSymbols 
         Caption         =   "&Symbols..."
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuformat 
      Caption         =   "F&ormat"
      HelpContextID   =   260
      Begin VB.Menu mnuFormatFont 
         Caption         =   "&Font..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuFormatColor 
         Caption         =   "&Color"
      End
      Begin VB.Menu mnuformatbullet 
         Caption         =   "&Bullet"
      End
      Begin VB.Menu filebar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFormatBold 
         Caption         =   "&Bold"
      End
      Begin VB.Menu mnuFormatItalic 
         Caption         =   "&Italic"
      End
      Begin VB.Menu mnuFormatunderline 
         Caption         =   "&underline"
      End
      Begin VB.Menu mnuformatparagraph 
         Caption         =   "&Paragraph"
      End
      Begin VB.Menu mnuFormatBar0 
         Caption         =   "-"
         HelpContextID   =   470
      End
      Begin VB.Menu mnuFormatUpper 
         Caption         =   "To &Upper Case"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuFormatLower 
         Caption         =   "To &Lower Case"
         Shortcut        =   +{F5}
      End
      Begin VB.Menu mnuFormatBar1 
         Caption         =   "-"
         HelpContextID   =   480
      End
      Begin VB.Menu mnuFormatScript 
         Caption         =   "&Script"
         Begin VB.Menu mnuFormatScriptNoScript 
            Caption         =   "&No Scripting"
         End
         Begin VB.Menu mnuFormatScriptSuperScript 
            Caption         =   "&SuperScript"
         End
         Begin VB.Menu mnuFormatScriptSubScript 
            Caption         =   "S&ubScript"
         End
      End
      Begin VB.Menu mnuFormatBar2 
         Caption         =   "-"
         HelpContextID   =   490
      End
      Begin VB.Menu mnuFormatIndent 
         Caption         =   "Increase &Indent"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuFormatOutdent 
         Caption         =   "&Reduce Indent"
         Shortcut        =   +{F1}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      HelpContextID   =   270
      Begin VB.Menu mnutoolscalendar 
         Caption         =   "&Calendar"
      End
      Begin VB.Menu mnutooldcalculator 
         Caption         =   "&Calcualator"
      End
      Begin VB.Menu mnuToolscapturescreen 
         Caption         =   "&captureScreen"
      End
      Begin VB.Menu mnutoolsclock 
         Caption         =   "&Clock"
      End
      Begin VB.Menu mnuToolshtmleditor 
         Caption         =   "&Htmleditor"
      End
      Begin VB.Menu mnutoolspaint 
         Caption         =   "&Paint"
      End
      Begin VB.Menu mnutoolsmp3player 
         Caption         =   "&MP3 Player"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      HelpContextID   =   280
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowBar0 
         Caption         =   "-"
         HelpContextID   =   500
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
         HelpContextID   =   510
      End
      Begin VB.Menu mnuWindowMinimizeAll 
         Caption         =   "&Minimize All"
      End
      Begin VB.Menu mnuWindowRestoreAll 
         Caption         =   "&Restore All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      HelpContextID   =   290
      Begin VB.Menu mnuHelpHelpContents 
         Caption         =   "&HelpContents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
         HelpContextID   =   520
      End
      Begin VB.Menu mnuHelpTips 
         Caption         =   "&Tips"
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
         HelpContextID   =   530
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About NitePad2002..."
      End
   End
   Begin VB.Menu mnuPop 
      Caption         =   "Popup"
      HelpContextID   =   300
      Visible         =   0   'False
      Begin VB.Menu mnuPopUndo 
         Caption         =   "&Undo"
      End
      Begin VB.Menu mnuPopBar0 
         Caption         =   "-"
         HelpContextID   =   540
      End
      Begin VB.Menu mnuPopCut 
         Caption         =   "Cu&t"
      End
      Begin VB.Menu mnuPopCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPopPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuPopBar1 
         Caption         =   "-"
         HelpContextID   =   550
      End
      Begin VB.Menu mnuPopDelete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu mnuPopSelectAll 
         Caption         =   "&Select All"
      End
      Begin VB.Menu mnuPopPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuPopBar2 
         Caption         =   "-"
         HelpContextID   =   560
      End
      Begin VB.Menu mnuPopCase 
         Caption         =   "C&hange Case"
         Begin VB.Menu mnuPopCaseUpper 
            Caption         =   "&Upper Case"
         End
         Begin VB.Menu mnuPopCaseLower 
            Caption         =   "&Lower Case"
         End
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' For the custom MsgBox
Public bYes, bNo, bCancel As Boolean
Private Sub cboFontName_Click()
    ActiveForm.rtfText.SelFontName = cboFontName.Text 'Set selected font name
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Font Name", cboFontName.Text
End Sub
Private Sub cboFontSize_Click()
    ActiveForm.rtfText.SelFontSize = cboFontSize.Text 'Set selected font size
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, SettingsPath, "Font Size", cboFontSize.Text
End Sub
Private Sub MDIForm_Load()

    RGCreateKey HKEY_LOCAL_MACHINE, ViewPath 'Create View key
    RGCreateKey HKEY_LOCAL_MACHINE, SettingsPath 'Create Settings key
    CreateNewDocument
    DisableAll 'Disable All menus and toolbars
   
    lngMenu = GetMenu(frmMDI.hwnd)

'File menu
lngSubMenu = GetSubMenu(lngMenu, 1)
lngMenuItemID = GetMenuItemID(lngSubMenu, 0)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picNew.Picture, frmPics.picNew.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 1)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picopen.Picture, frmPics.picopen.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 5)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picsaveall.Picture, frmPics.picsaveall.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 4)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picsave.Picture, frmPics.picsave.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 9)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picprint.Picture, frmPics.picprint.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 7)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picprintsetup.Picture, frmPics.picprintsetup.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 8)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picpreview.Picture, frmPics.picpreview.Picture)


'edit menu

lngSubMenu = GetSubMenu(lngMenu, 2)
lngMenuItemID = GetMenuItemID(lngSubMenu, 0)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picundo.Picture, frmPics.picundo.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 1)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picredo.Picture, frmPics.picredo.Picture)
    lngMenuItemID = GetMenuItemID(lngSubMenu, 3)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.piccut.Picture, frmPics.piccut.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 4)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.piccopy.Picture, frmPics.piccopy.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 5)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picpaste.Picture, frmPics.picpaste.Picture)
  lngMenuItemID = GetMenuItemID(lngSubMenu, 7)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picdelete.Picture, frmPics.picdelete.Picture)

'Search Menu
lngSubMenu = GetSubMenu(lngMenu, 3)
lngMenuItemID = GetMenuItemID(lngSubMenu, 0)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picfind.Picture, frmPics.picfind.Picture)

'view menu



'insert menu
lngSubMenu = GetSubMenu(lngMenu, 5)
lngMenuItemID = GetMenuItemID(lngSubMenu, 0)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picpic.Picture, frmPics.picpic.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 4)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 3, _
    frmPics.pictxt.Picture, frmPics.pictxt.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 2)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 4, _
    frmPics.pictime.Picture, frmPics.pictime.Picture)



'Format Menu
lngSubMenu = GetSubMenu(lngMenu, 6)

lngMenuItemID = GetMenuItemID(lngSubMenu, 1)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 1, _
    frmPics.picolor.Picture, frmPics.picolor.Picture)



'tools menu

lngSubMenu = GetSubMenu(lngMenu, 7)
lngMenuItemID = GetMenuItemID(lngSubMenu, 4)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.pichtml.Picture, frmPics.pichtml.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 1)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.piccalculator.Picture, frmPics.piccalculator.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 2)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.piccamera.Picture, frmPics.piccamera.Picture)

'windows menu

lngSubMenu = GetSubMenu(lngMenu, 8)
lngMenuItemID = GetMenuItemID(lngSubMenu, 2)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.piccascade.Picture, frmPics.piccascade.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 5)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picarrange.Picture, frmPics.picarrange.Picture)

'help menu

lngSubMenu = GetSubMenu(lngMenu, 9)
lngMenuItemID = GetMenuItemID(lngSubMenu, 0)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.pichelp.Picture, frmPics.pichelp.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 2)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.pictipday.Picture, frmPics.pictipday.Picture)
lngMenuItemID = GetMenuItemID(lngSubMenu, 4)
lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, _
    frmPics.picabout.Picture, frmPics.picabout.Picture)
End Sub
    Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim OLEFilename As String
    Dim fType As String
    Dim i As Integer
    
    For i = 1 To Data.Files.Count
        If Data.GetFormat(vbCFFiles) Then
            OLEFilename = Data.Files(i)
        End If
    
        'Get file extension
        Select Case UCase(Right(OLEFilename, 3))
            Case "RTF"
                fType = rtfRTF
            Case Else
                fType = rtfText
        End Select
        
        On Error GoTo errexit
        CreateNewDocument
        ActiveForm.rtfText.LoadFile OLEFilename, fType 'Load file
        ActiveForm.Caption = OLEFilename 'Set caption
        ActiveForm.bChanged = False 'Set bChanged flag to false
    Next i
errexit:
    Exit Sub
End Sub
Private Sub MDIForm_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    If Not Data.GetFormat(vbCFFiles) Then Effect = vbDropEffectNone
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub
Private Sub mnuFilesendemail_Click()
 ShellExecute frmMDI.hwnd, "open", "mailto:deunice@sc.rr.com", 0&, 0&, vbNormal
 End Sub

Private Sub mnuFormatBold_Click()
 mnuFormatBold.Checked = Not mnuFormatBold.Checked
  frmDocument.rtfText.SelBold = mnuFormatBold.Checked
End Sub

Private Sub mnuformatbullet_Click()
Bullet
End Sub

Private Sub mnuFormatColor_Click()
cmdlg.Flags = cdlCFBoth Or cdlCFEffects
       
    With ActiveForm.rtfText
         cmdlg.Color = .SelColor
         cmdlg.ShowColor
        .SelColor = cmdlg.Color
        
    End With
 
End Sub
Private Sub mnuFormatItalic_Click()
 mnuFormatItalic.Checked = Not mnuFormatItalic.Checked
  frmDocument.rtfText.SelItalic = mnuFormatItalic.Checked
End Sub

Private Sub mnuformatparagraph_Click()
frmParagraph.Show
End Sub

Private Sub mnuFormatunderline_Click()
 mnuFormatunderline.Checked = Not mnuFormatunderline.Checked
  frmDocument.rtfText.SelUnderline = mnuFormatunderline.Checked
End Sub
Private Sub mnuHelpHelpContents_Click()
Help.Show
End Sub
Private Sub mnuHelpTips_Click()
frmTip.Show
End Sub
Private Sub mnuInsertPicture_Click()
 
 
        ' Setting CancelError = True forces the program to errhandler: if the Cancel Button is clicked,
        ' giving the ActiveForm focus once again...
    cmdlg.CancelError = True
 
    cmdlg.DialogTitle = "Select Picture..."
    cmdlg.Filter = "Bitmaps (*.bmp;*.dib)|*.bmp;*.dib|GIF Images (*.gif)|*.gif|JPEG Images (*.jpg)|*.jpg|"
    cmdlg.ShowOpen
    
                        '  Load picture into the Active Form... and NOT into a new Document.
     frmMDI.ActiveForm.Picture = LoadPicture(cmdlg.FileName)
                        ' Copy the picture into the clipboard.
    Clipboard.Clear
    Clipboard.SetData frmMDI.ActiveForm.Picture
    
                         ' Paste the picture into the RichTextBox.
     SendMessage frmMDI.ActiveForm.rtfText.hwnd, WM_PASTE, 0, 0&
     frmMDI.ActiveForm.rtfText.SetFocus
     Exit Sub
    

End Sub
Private Sub mnutooldcalculator_Click()
frmCalculator.Show
End Sub
Private Sub mnuToolsAddressbook_Click()
frmaddress.Show
End Sub


Private Sub mnutoolscalendar_Click()
frmcalendar.Show
End Sub
Private Sub mnuToolscapturescreen_Click()
Capture.Show
End Sub
Private Sub mnutoolsclock_Click()
frmClock.Show
End Sub
Private Sub mnuToolshtmleditor_Click()
HTMLEditor.Show
End Sub
Private Sub mnutoolsmp3player_Click()
mp3.Show
End Sub

Private Sub mnuToolsOptions_Click()

End Sub

Private Sub mnutoolspaint_Click()
paintmain.Show
End Sub
Private Sub mnuViewwebbrowser_Click()
frmBrowser.Show
End Sub

'//////  TOOLBARS //////'
'// STANDARD
Private Sub tbrStandard_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Save"
        mnuFileSave_Click
    Case "Print"
        mnuFilePrint_Click
    Case "FullScreen"
        mnuViewFullScreen_Click
    Case "WordWrap"
        If tbrStandard.Buttons("WordWrap").Value = tbrPressed Then
            SetViewMode 1
            tbrStandard.Buttons("WordWrap").Value = tbrPressed
        Else
            SetViewMode 0
            tbrStandard.Buttons("WordWrap").Value = tbrUnpressed
        End If
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Undo"
        mnuEditUndo_Click
   Case "Outdent"
        mnuFormatOutdent_Click
    Case "Indent"
        mnuFormatIndent_Click
    Case "Help"
        mnuHelpHelpContents_Click
     Case "Tips"
           mnuHelpTips_Click
           End Select
End Sub
'// FORMAT
Private Sub tbrFormat_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
   Case "Find"
        mnuSearchFind_Click
  Case "Paint"
        paintmain.Show
   Case "Bold"
         ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
         Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
   Case "Italic"
         ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
         Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
   Case "Underline"
         ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
         Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
   Case "Left"
        AlignLeft
   Case "Center"
        AlignCenter
   Case "Right"
        AlignRight

    End Select
End Sub
'// FILE
Private Sub tbrFile_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "New"
        mnuFileNew_Click
    Case "Open"
        mnuFileOpen_Click
    Case "Save"
        mnuFileSave_Click
    Case "SaveAll"
        mnuFileSaveAs_Click
    Case "Printsetup"
          mnuFilePageSetup_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Email"
        mnuFilesendemail_Click
       End Select
End Sub
'// EDIT
Private Sub tbrEdit_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Cut"
        mnuEditCut_Click
    Case "Copy"
        mnuEditCopy_Click
    Case "Paste"
        mnuEditPaste_Click
    Case "Undo"
        mnuEditUndo_Click
   Case "Outdent"
        mnuFormatOutdent_Click
    Case "Indent"
        mnuFormatIndent_Click
    End Select
End Sub
'// WINDOW
Private Sub tbrWindow_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "Cascade"
        mnuWindowCascade_Click
    Case "THorizontally"
        mnuWindowTileHorizontal_Click
    Case "TVertically"
        mnuWindowTileVertical_Click
    Case "Minimize"
        mnuWindowMinimizeAll_Click
    Case "Restore"
        mnuWindowRestoreAll_Click
      End Select
End Sub
'// FILE MENU
Private Sub mnuFileNew_Click()
    CreateNewDocument 'Call CreateNewDocument function
End Sub
Private Sub mnuFileOpen_Click()
   Dim sFile As String
  

    With cmdlg
        .DialogTitle = "Open"
        .CancelError = False
        'ToDo: set the flags and attributes of the common dialog control
        .Filter = "Txt Files (*.txt)|*.txt|RTF Files (*.rtf)|*.rtf|All Files (*.*)|*.*"
        .ShowOpen
        If Len(.FileName) = 0 Then
            Exit Sub
        End If
        sFile = .FileName
    End With
    frmDocument.rtfText.LoadFile sFile
    frmDocument.Caption = sFile

End Sub
   
Public Sub mnuFileSave_Click()
     Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With cmdlg
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Text (*.txt)|*.txt|Rtf (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
            myCancel = True
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
   Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub
Private Sub mnuFileSaveAs_Click()
     Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With cmdlg
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Text (*.txt)|*.txt|Rtf (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
            myCancel = True
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
   Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub
Private Sub mnuFileSaveAll_Click()
     Dim sFile As String
    If Left$(ActiveForm.Caption, 8) = "Document" Then
        With cmdlg
            .DialogTitle = "Save"
            .CancelError = False
            'ToDo: set the flags and attributes of the common dialog control
            .Filter = "Text (*.txt)|*.txt|Rtf (*.rtf)|*.rtf|All Files (*.*)|*.*"
            .ShowSave
            If Len(.FileName) = 0 Then
            myCancel = True
                Exit Sub
            End If
            sFile = .FileName
        End With
        ActiveForm.rtfText.SaveFile sFile
   Else
        sFile = ActiveForm.Caption
        ActiveForm.rtfText.SaveFile sFile
    End If

End Sub
Private Sub mnuFilePageSetup_Click()
 cmdlg.Flags = cdlPDPrintSetup
            cmdlg.ShowPrinter
            DoEvents
End Sub
Private Sub mnuFilePrintSetup_Click()
    cmdlg.ShowPrinter 'Show Printer dialog
End Sub
Private Sub mnuFilePrint_Click()
   PrintRTF ActiveForm.rtfText, 720, 720, 720, 720 'Call PrintRTF sub
End Sub
Private Sub mnuFileExit_Click()
    End
End Sub
'// EDIT MENU
Private Sub mnuEditUndo_Click()
   SendMessage ActiveForm.rtfText.hwnd, EM_UNDO, 0, 0&
End Sub
Private Sub mnuEditCut_Click()
    SendMessage ActiveForm.rtfText.hwnd, WM_CUT, 0&, 0& 'Cut
End Sub
Private Sub mnuEditCopy_Click()
   SendMessage ActiveForm.rtfText.hwnd, WM_COPY, 0&, 0& 'Copy
End Sub
Private Sub mnuEditPaste_Click()
   SendMessage ActiveForm.rtfText.hwnd, WM_PASTE, 0&, 0& 'Paste
End Sub
Private Sub mnuEditDelete_Click()
    SendMessage ActiveForm.rtfText.hwnd, WM_CLEAR, 0&, 0& 'Delete
End Sub
Private Sub mnuEditSelectAll_Click()
    ActiveForm.rtfText.SelStart = 0 'Set the start pos of the selection
    ActiveForm.rtfText.SelLength = Len(ActiveForm.rtfText) 'Set length of the selection
End Sub
'// SEARCH MENU
Private Sub mnuSearchFind_Click()
    frmFind.Show , Me
End Sub
Private Sub mnuSearchFindNext_Click()
    On Error GoTo FindNextError
    Dim lngResult As Integer
    Dim lngPos As Integer
    Dim intOptions As Integer
    ' Set search options
    If frmFind.chkNoHighlight.Value = 1 Then intOptions = intOptions + 8
    If frmFind.chkWholeWord.Value = 1 Then intOptions = intOptions + 2
    If frmFind.chkMatchCase.Value = 1 Then intOptions = intOptions + 4

    lngPos = ActiveForm.rtfText.SelStart + ActiveForm.rtfText.SelLength
    ' Get position of the searched word
    lngResult = ActiveForm.rtfText.Find(frmFind.cboFind.Text, lngPos, , intOptions)

    If lngResult = -1 Then 'Text not found
        MsgE "Text not found", "Nitepad2002 - FindNext", 1, True 'Show message
        frmFind.cmdFind.Caption = "&Find" 'Set caption
        frmFind.cmdReplace.Enabled = False 'Disable Replace button
        frmFind.cmdReplaceAll.Enabled = False 'Disable ReplaceAll button
        mnuSearchFindNext.Enabled = False 'Disable Find Next menu
    Else
        ActiveForm.rtfText.SetFocus 'Set focus
    End If
FindNextError:
    ErrorLog "frmMDI\mnuEditFindNext_Click"
End Sub
Private Sub mnuSearchReplace_Click()
   With frmFind
        .cmdReplace.Top = 150 'Set cmdReplace top
        .cmdReplace.Caption = "&Replace" 'Set caption
        .lblReplace.Visible = True 'Show lblReplace
        .cboReplace.Visible = True 'Show cboReplace
        .cmdReplaceAll.Visible = True 'Show cmdReplaceAll
        .Show , Me
    End With
End Sub
Private Sub mnuSearchGoTo_Click()
    frmGoTo.Show , Me
End Sub

'// VIEW MENU
Private Sub mnuViewStatusBar_Click()
    mnuViewStatusBar.Checked = Not mnuViewStatusBar.Checked
    sb.Visible = mnuViewStatusBar.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Status Bar", mnuViewStatusBar.Checked
End Sub
Private Sub mnuViewRuler_Click()
    mnuViewRuler.Checked = Not mnuViewRuler.Checked
    cbrRuler.Visible = mnuViewRuler.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Ruler", mnuViewRuler.Checked
End Sub

Private Sub mnuViewFullScreen_Click()
    On Error GoTo FullScreenError
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "Nitepad2002", 1, True: Exit Sub
    
    ' Select all text
    mnuEditSelectAll_Click
    ' Copy text
    SendMessage ActiveForm.rtfText.hwnd, WM_COPY, 0, 0&
    ' Paste text
    SendMessage frmFScreen.FSRTB.hwnd, WM_PASTE, 0, 0&
    ' Show full screen
    frmFScreen.Show , Me
    frmFScreen.FSRTB.SelStart = 0
FullScreenError:
    ErrorLog "frmMDI\mnuViewFullScreen_Click"
End Sub
Private Sub mnuViewToolbarEdit_Click()
    mnuViewToolbarEdit.Checked = Not mnuViewToolbarEdit.Checked
    cbrBar.Bands(5).Visible = mnuViewToolbarEdit.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Edit Toolbar", mnuViewToolbarEdit.Checked
End Sub
Private Sub mnuViewToolbarFile_Click()
    mnuViewToolbarFile.Checked = Not mnuViewToolbarFile.Checked
    cbrBar.Bands(4).Visible = mnuViewToolbarFile.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "File Toolbar", mnuViewToolbarFile.Checked
End Sub
Private Sub mnuViewToolbarFont_Click()
    mnuViewToolbarFont.Checked = Not mnuViewToolbarFont.Checked
    cbrBar.Bands(2).Visible = mnuViewToolbarFont.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Font Toolbar", mnuViewToolbarFont.Checked
End Sub
Private Sub mnuViewToolbarFormat_Click()
    mnuViewToolbarFormat.Checked = Not mnuViewToolbarFormat.Checked
    cbrBar.Bands(3).Visible = mnuViewToolbarFormat.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Format Toolbar", mnuViewToolbarFormat.Checked
End Sub
Private Sub mnuViewToolbarStandard_Click()
    mnuViewToolbarStandard.Checked = Not mnuViewToolbarStandard.Checked
    cbrBar.Bands(1).Visible = mnuViewToolbarStandard.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Standard Toolbar", mnuViewToolbarStandard.Checked
End Sub
Private Sub mnuViewToolbarWindow_Click()
    mnuViewToolbarWindow.Checked = Not mnuViewToolbarWindow.Checked
    cbrBar.Bands(6).Visible = mnuViewToolbarWindow.Checked
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Window Toolbar", mnuViewToolbarWindow.Checked
End Sub
Private Sub mnuViewStayonTop_Click()
    mnuViewStayonTop.Checked = Not mnuViewStayonTop.Checked
    If mnuViewStayonTop.Checked Then
        OnTop Me 'Put Nitepad2002 on top
        ' Save to registry
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Stay On Top", 1
    Else
        NotOnTop Me 'Remove Nitepad2002 from top
        ' Save to registry
        RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "Stay On Top", 0
    End If
End Sub
Private Sub mnuViewMode_Click(Index As Integer)
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "Nitepad2002", 1, True: Exit Sub
    Dim i As Integer
    SetViewMode Index 'Set selected view mode
    For i = 0 To 2 'Uncheck all items
        mnuViewMode(i).Checked = False
    Next
    mnuViewMode(Index).Checked = True 'Check selected item
    ' Save to registry
    RGSetKeyValue HKEY_LOCAL_MACHINE, ViewPath, "ViewMode", Str(Index)
End Sub
Private Sub mnuViewDocumentProperties_Click()
    frmDocInfo.Show , Me
End Sub
'// INSERT MENU
Private Sub mnuInsertTimeDate_Click()
   frmTimeDate.Show , Me
End Sub
Private Sub mnuInsertPathandFile_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "Nitepad2002", 1, True: Exit Sub
    
    If Left(ActiveForm.Caption, 8) = "Document" Then 'Dont insert if doesn't exists
        Exit Sub
    Else
        ActiveForm.rtfText.SelText = ActiveForm.Caption 'Insert path anf file
    End If
End Sub
Private Sub mnuInsertSymbols_Click()
   frmSymbols.Show , Me
End Sub
'// FORMAT MENU
Private Sub mnuFormatFont_Click()
    ' cdlCFBoth will Display both Screen and Printer Fonts
    ' cdlCFEffects will display all the attributes
    ' including Underline , StrikeThru and Color
    cmdlg.Flags = cdlCFBoth Or cdlCFEffects
    
    
    With ActiveForm.rtfText
        ' Update the Font Dialog Box with
        ' the Current Font Format of the selected text
       
        cmdlg.FontName = .SelFontName
        cmdlg.FontSize = .SelFontSize
        cmdlg.FontBold = .SelBold
        cmdlg.FontItalic = .SelItalic
        cmdlg.FontUnderline = .SelUnderline
        cmdlg.Color = .SelColor
        
       
        cmdlg.ShowFont
        
        ' Apply the settings in the
        ' Font Dialog Box to the selected text
        .SelFontName = cmdlg.FontName
        .SelFontSize = cmdlg.FontSize
        .SelBold = cmdlg.FontBold
        .SelItalic = cmdlg.FontItalic
        .SelStrikeThru = cmdlg.FontStrikethru
        .SelUnderline = cmdlg.FontUnderline
        .SelColor = cmdlg.Color
        
    End With
 
End Sub
Private Sub mnuFormatUpper_Click()
 ActiveForm.rtfText.SelText = UCase(ActiveForm.rtfText.SelText)
End Sub
Private Sub mnuFormatLower_Click()
  ActiveForm.rtfText.SelText = LCase(ActiveForm.rtfText.SelText)
End Sub
Private Sub mnuFormatScriptNoScript_Click()
   ActiveForm.rtfText.SelCharOffset = 0
End Sub
Private Sub mnuFormatScriptSubscript_Click()
   ActiveForm.rtfText.SelCharOffset = -55
End Sub
Private Sub mnuFormatScriptSuperScript_Click()
   ActiveForm.rtfText.SelCharOffset = 55
End Sub
Private Sub mnuFormatIndent_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "Nitepad2002", 1, True: Exit Sub
    'Set the forms scale mode to Millimeters
    ActiveForm.ScaleMode = vbMillimeters
    'Change the indent
    ActiveForm.rtfText.SelIndent = ActiveForm.rtfText.SelIndent + 13
    'Return form scale mode to Twips
    ActiveForm.ScaleMode = vbTwips
End Sub
Private Sub mnuFormatOutdent_Click()
    ' Check to see if there are any open documents, if not show error
    If ActiveForm Is Nothing Then MsgE "No open documents !", "Nitepad2002", 1, True: Exit Sub
    'Set the forms scale mode to Millimeters
    ActiveForm.ScaleMode = vbMillimeters
    'Change the indent
    ActiveForm.rtfText.SelIndent = ActiveForm.rtfText.SelIndent - 13
    'Return form scale mode to Twips
    ActiveForm.ScaleMode = vbTwips
End Sub
'// WINDOW MENU
Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub
Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub
Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub
Private Sub mnuWindowNewWindow_Click()
    CreateNewDocument 'Call CreateNewDocument function
End Sub
Private Sub mnuWindowMinimizeAll_Click()
  Dim i As Integer
    ' Minimize all documents
    For i = 1 To Forms.Count - 1
        Forms(i).WindowState = vbMinimized
    Next i
End Sub
Private Sub mnuWindowRestoreAll_Click()
 Dim i As Integer
    ' Restore all documents
    For i = 1 To Forms.Count - 1
        Forms(i).WindowState = vbNormal
    Next i
End Sub
'// POPUP MENU
Private Sub mnuPopUndo_Click()
    'mnueditundo_click
End Sub
Private Sub mnuPopCut_Click()
    mnuEditCut_Click
End Sub
Private Sub mnuPopCopy_Click()
    mnuEditCopy_Click
End Sub
Private Sub mnuPopPaste_Click()
    mnuEditPaste_Click
End Sub
Private Sub mnuPopDelete_Click()
    mnuEditDelete_Click
End Sub
Private Sub mnuPopSelectAll_Click()
    mnuEditSelectAll_Click
End Sub
Private Sub mnuPopPrint_Click()
    mnuFilePrint_Click
End Sub
Private Sub mnuPopCaseLower_Click()
    mnuFormatLower_Click
End Sub
Private Sub mnuPopCaseUpper_Click()
    mnuFormatUpper_Click
End Sub


Private Sub Timer1_Timer()
frmMDI.sb.Panels(2).Text = Format(Now, "HH:MM:SS")
End Sub

Private Sub Timer2_Timer()
frmMDI.sb.Panels(3).Text = Format(Now, "DD/MM/YY")

End Sub

Private Sub tmrautosave_Timer()
txtrepitition = Val(txtrepitition) + 1
On Error Resume Next
If Val(txtrepitition) = txtuserinput * 60 Then
txtrepitition = "0"
If frmMDI(frm).Text1.Text = "" Then Exit Sub
mnuFileSave_Click
End If
End Sub

Private Sub txtuserinput_Change()
If Not IsNumeric(txtuserinput) Then
MsgBox "Invalid number!"
txtuserinput = "0"
txtrepitition = "0"
tmrautosave.Enabled = False
Else
tmrautosave.Enabled = True
End If

End Sub
