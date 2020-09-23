VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Layout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crimson Cubes"
   ClientHeight    =   8415
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   11910
   Icon            =   "Layout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   Picture         =   "Layout.frx":030A
   ScaleHeight     =   8415
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList img_Dice 
      Left            =   7845
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   63
      ImageHeight     =   61
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":15088C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1536A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1563C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1590E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":15BE00
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":15EB20
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic_GoldBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   10020
      MousePointer    =   1  'Arrow
      ScaleHeight     =   405
      ScaleWidth      =   810
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   810
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9510
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   54
      ImageHeight     =   27
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   200
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":161840
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1629E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":163B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":164D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":165EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":167060
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":168200
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1693A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":16A540
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":16B6E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":16C880
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":16DA20
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":16EBC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":16FD60
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":170F00
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1720A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":173240
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1743E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":175580
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":176720
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1778C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":178A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":179C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":17ADA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":17BF40
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":17D0E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":17E280
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":17F420
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1805C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":181760
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":182900
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":183AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":184C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":185DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":186F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":188120
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1892C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":18A460
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":18B600
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":18C7A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":18D940
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":18EAE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":18FC80
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":190E20
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":191FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":193160
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":194300
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1954A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":196640
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1977E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":198980
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":199B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":19ACC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":19BE60
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":19D000
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":19E1A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":19F340
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A04E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A1680
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A2820
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A39C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A4B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A5D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A6EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A8040
            Key             =   ""
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1A91E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1AA380
            Key             =   ""
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1AB520
            Key             =   ""
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1AC6C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1AD860
            Key             =   ""
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1AEA00
            Key             =   ""
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1AFBA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B0D40
            Key             =   ""
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B1EE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B3080
            Key             =   ""
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B4220
            Key             =   ""
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B53C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B6560
            Key             =   ""
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B7700
            Key             =   ""
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B88A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1B9A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1BABE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1BBD80
            Key             =   ""
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1BCF20
            Key             =   ""
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1BE0C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1BF260
            Key             =   ""
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C0400
            Key             =   ""
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C15A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C2740
            Key             =   ""
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C38E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C4A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C5C20
            Key             =   ""
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C6DC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C7F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1C9100
            Key             =   ""
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1CA2A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1CB440
            Key             =   ""
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1CC5E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1CD780
            Key             =   ""
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1CE920
            Key             =   ""
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1CFAC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D0C60
            Key             =   ""
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D1E00
            Key             =   ""
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D2FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D4140
            Key             =   ""
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D52E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D6480
            Key             =   ""
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D7620
            Key             =   ""
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D87C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1D9960
            Key             =   ""
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1DAB00
            Key             =   ""
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1DBCA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1DCE40
            Key             =   ""
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1DDFE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1DF180
            Key             =   ""
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E0320
            Key             =   ""
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E2660
            Key             =   ""
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E3800
            Key             =   ""
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E49A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E5B40
            Key             =   ""
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E6CE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E7E80
            Key             =   ""
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1E9020
            Key             =   ""
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1EA1C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1EB360
            Key             =   ""
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1EC500
            Key             =   ""
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1ED6A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1EE840
            Key             =   ""
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1EF9E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F0B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F1D20
            Key             =   ""
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F2EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F4060
            Key             =   ""
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F5200
            Key             =   ""
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F63A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F7540
            Key             =   ""
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F86E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1F9880
            Key             =   ""
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1FAA20
            Key             =   ""
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1FBBC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1FCD60
            Key             =   ""
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1FDF00
            Key             =   ""
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":1FF0A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":200240
            Key             =   ""
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2013E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":202580
            Key             =   ""
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":203720
            Key             =   ""
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2048C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":205A60
            Key             =   ""
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":206C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":207DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":208F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":20A0E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":20B280
            Key             =   ""
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":20C420
            Key             =   ""
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":20D5C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":20E760
            Key             =   ""
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":20F900
            Key             =   ""
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":210AA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":211C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":212DE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":213F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":215120
            Key             =   ""
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2162C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":217460
            Key             =   ""
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":218600
            Key             =   ""
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2197A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":21A940
            Key             =   ""
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":21BAE0
            Key             =   ""
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":21CC80
            Key             =   ""
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":21DE20
            Key             =   ""
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":21EFC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":220160
            Key             =   ""
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":221300
            Key             =   ""
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2224A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":223640
            Key             =   ""
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2247E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":225980
            Key             =   ""
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":226B20
            Key             =   ""
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":227CC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":228E60
            Key             =   ""
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":22A000
            Key             =   ""
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":22B1A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":22C340
            Key             =   ""
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":22D4E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":22E680
            Key             =   ""
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":22F820
            Key             =   ""
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2309C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":231B60
            Key             =   ""
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":232D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":233EA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":235040
            Key             =   ""
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2361E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":237380
            Key             =   ""
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":238520
            Key             =   ""
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":2396C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":23A860
            Key             =   ""
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":23BA00
            Key             =   ""
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Layout.frx":23CBA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image img_WhatsThis 
      Height          =   450
      Left            =   10080
      Picture         =   "Layout.frx":23DD40
      ToolTipText     =   "Enable 'What Is This Bet?' mode"
      Top             =   7500
      Width           =   375
   End
   Begin VB.Label lbl_BankRoll 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008000&
      Caption         =   "250"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   525
      Left            =   9480
      TabIndex        =   2
      Top             =   1290
      Width           =   1935
   End
   Begin VB.Image img_Button 
      Height          =   1035
      Left            =   105
      Picture         =   "Layout.frx":23E66A
      Top             =   75
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Image img_D2 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   8220
      Top             =   1035
      Width           =   930
   End
   Begin VB.Image img_D1 
      Appearance      =   0  'Flat
      Height          =   915
      Left            =   7140
      Top             =   1035
      Width           =   930
   End
   Begin VB.Label lbl_Message 
      BackColor       =   &H00008000&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   390
      Left            =   465
      TabIndex        =   0
      Top             =   375
      Width           =   10995
   End
   Begin VB.Image img_Help 
      Height          =   450
      Left            =   10500
      Picture         =   "Layout.frx":241EBC
      ToolTipText     =   "Bring up the main Help "
      Top             =   7500
      Width           =   915
   End
   Begin VB.Image img_CashOut 
      Height          =   450
      Left            =   9660
      Picture         =   "Layout.frx":24348E
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   7
      Left            =   10545
      Picture         =   "Layout.frx":2459D8
      Top             =   3795
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   6
      Left            =   9555
      Picture         =   "Layout.frx":246B66
      Top             =   3795
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   5
      Left            =   10545
      Picture         =   "Layout.frx":247CF4
      Top             =   3255
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   4
      Left            =   9555
      Picture         =   "Layout.frx":248E82
      Top             =   3255
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   3
      Left            =   10545
      Picture         =   "Layout.frx":24A010
      Top             =   2715
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   2
      Left            =   9555
      Picture         =   "Layout.frx":24B19E
      Top             =   2715
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   1
      Left            =   10530
      Picture         =   "Layout.frx":24C32C
      Top             =   2175
      Width           =   810
   End
   Begin VB.Image img_Bet 
      Height          =   405
      Index           =   0
      Left            =   9555
      Picture         =   "Layout.frx":24D4BA
      Top             =   2175
      Width           =   810
   End
   Begin VB.Image img_Roll 
      Appearance      =   0  'Flat
      Height          =   450
      Left            =   7140
      Picture         =   "Layout.frx":24E648
      Top             =   2130
      Width           =   2010
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnu_Recall 
         Caption         =   "Recall the Bet"
      End
      Begin VB.Menu mnu_Cancel 
         Caption         =   "Cancel"
      End
   End
End
Attribute VB_Name = "Layout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private X_OffSet As Integer 'The horizontal offset (mouse in relation to object's left)
Private Y_OffSet As Integer 'The vertical offset (mouse in relation to object's top)
Private RecallBet As Integer 'Holds the bet number on a Right click.  Used when recalling Place bets.
Private MouseDrag As Boolean 'Set to true when dragging a bet

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Test for the question mark or the F1 function key getting pressed
    If KeyCode = 191 And Shift = 1 Then
        'Player has pressed the question mark so enable "What is This" help
        Me.MousePointer = 14
        lbl_Message.Caption = "Click on a bet for a little help...(press Esc to exit help mode)"
    Else
        'See if the F1 key was pressed
        If KeyCode = 112 Then
            Main_Help.Show 1
        Else
            'Any other key resets the mouse and disables help
            Me.MousePointer = 0
            lbl_Message.Caption = vbNullString
        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Determine which button was pressed
    Select Case Button
        Case 1 'Left
            'See if in Help mode or placing a bet
            If Me.MousePointer = 14 Then
                'Player is in Help mode so determine which bet the mouse is currently over and
                'display the info about that bet.
                '(I like to use an explicit "Call" because it makes it obvious.
                ' In this Call there are parameters which  help but let's say there were none and I
                ' left off "Call".  We would just have:
                '   What_Is_This_Bet
                ' It could be a function call or a variable that I forgot to assign a value to - there's
                ' no way to know by just looking at it.  On the other hand:
                '   Call What_Is_This_Bet
                ' is very clear.)
                Call What_Is_This_Bet(Find_The_Mouse(X!, Y!))
            Else
                'See if there is a valid bet amount
                If pic_GoldBar(0).Visible Then
                    'Validate the bet - If Is_legal_Bet returns false then the bet is not placed
                    Call Place_New_Bet(Is_Legal_Bet(X!, Y!))
                End If
            End If
        Case 2 'Right
            'See if this is a working Place bet or the last bet made
            RecallBet% = Find_The_Mouse(X!, Y!)
            If (RecallBet% >= Bet.IsPlace4 And RecallBet% <= Bet.IsPlace10 And gPlacedBet%(RecallBet%)) Or RecallBet% = gLastBet% Then
                'Ok to bring up the menu
                PopupMenu mnuPopUp
            End If
    End Select
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'If the MouseDrag flag has been set then the player is dragging a bet so use the Move
    'method to reposition the bet the the mouse's location (the X and Y parameters)
    If MouseDrag Then
        pic_GoldBar(0).Move X! - X_OffSet%, Y! - Y_OffSet%
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseDrag Then
        'Is_Legal_Bet will determine which bet the player wants to place.
        'If it is a legal bet then the bet number will get passed along to Place_New_Bet which will process it.
        'If the bet is not legal then Is_Legal_Bet will return a zero and Place_New_Bet will reset everything.
        Call Place_New_Bet(Is_Legal_Bet(X!, Y!))
        'Reset the Drag flag in any case
        MouseDrag = False
    End If
End Sub

Private Sub img_Bet_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Press_Button(img_Bet(Index))
End Sub

Private Sub img_Bet_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Release_Button(img_Bet(Index))
    'Generate the graphics for the bet
    Call Create_The_Bet_Amount(Index%)
End Sub

Private Sub img_CashOut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Press_Button(img_CashOut)
End Sub

Private Sub img_CashOut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Release_Button(img_CashOut)
    'The player has cashed out so call the game ending routine
    Call End_The_Game
End Sub

Private Sub img_Help_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Press_Button(img_Help)
End Sub

Private Sub img_Help_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Release_Button(img_Help)
    'Bring up the help form
    Main_Help.Show 1
End Sub

Private Sub img_Roll_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Press_Button(img_Roll)
End Sub

Private Sub img_Roll_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
     Call Release_Button(img_Roll)
    'Roll the dice
    gDiceRoll% = Roll_The_Dice
    'See if there are winners and/or losers for this roll
    Call Test_Win_Or_Lose
    'Place or remove the button as needed
    Call Button_Control
End Sub

Private Sub img_WhatsThis_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Press_Button(img_WhatsThis)
End Sub

Private Sub img_WhatsThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call Release_Button(img_WhatsThis)
    'Pass in a question mark key press to the form to enable the What's This help mode
    Call Form_KeyDown(191, 1)
End Sub

Private Sub mnu_Recall_Click()
    'Place the bet amount back in the bank
    lbl_BankRoll.Caption = Val(lbl_BankRoll.Caption) + gPlacedBet%(RecallBet%)
    'Clear the bet from the layout
    Call Clear_This_Bet(RecallBet%)
End Sub

Private Sub pic_GoldBar_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Record the offsets
    X_OffSet% = X!
    Y_OffSet% = Y!
    'Enable the Drag
    MouseDrag = True
    'Clear the message bar
    lbl_Message.Caption = vbNullString
    lbl_Message.Refresh
    'Make sure this the top most image
    pic_GoldBar(0).ZOrder
    pic_GoldBar(0).DragMode = 1
    pic_GoldBar(0).Enabled = False
End Sub
