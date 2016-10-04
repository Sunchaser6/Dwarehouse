VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{B16553C3-06DB-101B-85B2-0000C009BE81}#1.0#0"; "SPIN32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTSGDataWarehouse 
   BackColor       =   &H80000000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8775
   ClientLeft      =   1485
   ClientTop       =   1425
   ClientWidth     =   11655
   FillColor       =   &H8000000F&
   ForeColor       =   &H80000005&
   Icon            =   "TSG Data Warehouse.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   8775
   ScaleWidth      =   11655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8385
      TabIndex        =   4
      Top             =   540
      Visible         =   0   'False
      Width           =   960
   End
   Begin ComctlLib.StatusBar stb 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   176
      Top             =   8415
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   635
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab tabMain 
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   14817
      _Version        =   393216
      Tabs            =   11
      TabsPerRow      =   11
      TabHeight       =   794
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Data Capture"
      TabPicture(0)   =   "TSG Data Warehouse.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblLocalDriveFranFolder"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblVersion"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblDisplayedFranchise"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblEventLog"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lstDataCaptureFranchiseBusinessName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "tdpEventLogDate"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lvwEventLog"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cdlTSGDataWarehouse"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tmrAutoDataCapture"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtActiveDatabase"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtNewFranchiseBusinessName"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Frame(16)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Frame(17)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Sales Reports"
      TabPicture(1)   =   "TSG Data Warehouse.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl(12)"
      Tab(1).Control(1)=   "lvwProductReport"
      Tab(1).Control(2)=   "Frame(15)"
      Tab(1).Control(3)=   "Frame(4)"
      Tab(1).Control(4)=   "Frame(10)"
      Tab(1).Control(5)=   "Frame(9)"
      Tab(1).Control(6)=   "Frame(3)"
      Tab(1).Control(7)=   "Frame(2)"
      Tab(1).Control(8)=   "Frame(1)"
      Tab(1).Control(9)=   "Frame(0)"
      Tab(1).Control(10)=   "lstProductReportsFranchiseBusinessName"
      Tab(1).Control(11)=   "chkSalesRptTab_IncludeClosedFrans"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Stick Reports"
      TabPicture(2)   =   "TSG Data Warehouse.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl(26)"
      Tab(2).Control(1)=   "lbl(25)"
      Tab(2).Control(2)=   "lvwStickReport"
      Tab(2).Control(3)=   "Frame(11)"
      Tab(2).Control(4)=   "lstStickReportRecipient"
      Tab(2).Control(5)=   "Frame(8)"
      Tab(2).Control(6)=   "lstStickReportsFranchiseBusinessName"
      Tab(2).Control(7)=   "Frame(7)"
      Tab(2).Control(8)=   "Frame(6)"
      Tab(2).Control(9)=   "Frame(5)"
      Tab(2).Control(10)=   "cmdStickReport"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Stock"
      TabPicture(3)   =   "TSG Data Warehouse.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame(12)"
      Tab(3).Control(1)=   "Frame(13)"
      Tab(3).Control(2)=   "chkStockTab_IncludeDeletedStock"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "BATA"
      TabPicture(4)   =   "TSG Data Warehouse.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraBataTab"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "AZTEC"
      TabPicture(5)   =   "TSG Data Warehouse.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "rtxNielsenReportContents"
      Tab(5).Control(1)=   "cmdCreateNielsenReports"
      Tab(5).Control(2)=   "lstNielsenReportDisplayDate"
      Tab(5).Control(3)=   "cmdPurgeNielsenReportList"
      Tab(5).Control(4)=   "Frame(14)"
      Tab(5).Control(5)=   "fraAztec"
      Tab(5).ControlCount=   6
      TabCaption(6)   =   "Versions"
      TabPicture(6)   =   "TSG Data Warehouse.frx":03B2
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "lvwVersions"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Product Report"
      TabPicture(7)   =   "TSG Data Warehouse.frx":03CE
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "lblProducts"
      Tab(7).Control(1)=   "lbl(33)"
      Tab(7).Control(2)=   "lvwPRProductReport"
      Tab(7).Control(3)=   "Frame(19)"
      Tab(7).Control(4)=   "lstPRProductList"
      Tab(7).Control(5)=   "Frame(18)"
      Tab(7).Control(6)=   "Frame(20)"
      Tab(7).Control(7)=   "Frame(21)"
      Tab(7).Control(8)=   "Frame(22)"
      Tab(7).Control(9)=   "Frame(23)"
      Tab(7).Control(10)=   "cmdPRPrint"
      Tab(7).Control(11)=   "lstPRProductReportsFranchiseBusinessName"
      Tab(7).ControlCount=   12
      TabCaption(8)   =   "Settings"
      TabPicture(8)   =   "TSG Data Warehouse.frx":03EA
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "lbl(39)"
      Tab(8).Control(1)=   "txtConfirmPassword"
      Tab(8).Control(2)=   "cmdSettingsPassword(0)"
      Tab(8).Control(3)=   "Frame(24)"
      Tab(8).Control(4)=   "Frame(27)"
      Tab(8).Control(5)=   "Frame(29)"
      Tab(8).ControlCount=   6
      TabCaption(9)   =   "Uploads"
      TabPicture(9)   =   "TSG Data Warehouse.frx":0406
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "lblUploadsPending"
      Tab(9).Control(1)=   "Frame(32)"
      Tab(9).Control(2)=   "Frame(34)"
      Tab(9).Control(3)=   "Frame(30)"
      Tab(9).Control(4)=   "btnPurgeUploadsPending"
      Tab(9).Control(5)=   "Frame(33)"
      Tab(9).Control(6)=   "Frame(35)"
      Tab(9).Control(7)=   "lvwUploadsPending"
      Tab(9).ControlCount=   8
      TabCaption(10)  =   "Promotions"
      TabPicture(10)  =   "TSG Data Warehouse.frx":0422
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "lvwPromo"
      Tab(10).Control(1)=   "cmdPromotion(8)"
      Tab(10).Control(2)=   "cmdPromotion(6)"
      Tab(10).Control(3)=   "cmdPromotion(2)"
      Tab(10).Control(4)=   "cmdPromotionRecall"
      Tab(10).Control(5)=   "fraAddPromotion"
      Tab(10).Control(6)=   "frameNonCompliant"
      Tab(10).ControlCount=   7
      Begin VB.Frame Frame 
         Height          =   2850
         Index           =   17
         Left            =   9690
         TabIndex        =   330
         Top             =   690
         Width           =   1815
         Begin VB.OptionButton optDialupResults 
            Caption         =   "ALL"
            Enabled         =   0   'False
            Height          =   195
            Index           =   0
            Left            =   1065
            TabIndex        =   338
            Top             =   2565
            Width           =   600
         End
         Begin VB.OptionButton optDialupResults 
            Caption         =   "Failures"
            Enabled         =   0   'False
            Height          =   195
            Index           =   1
            Left            =   165
            TabIndex        =   331
            Top             =   2565
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdTfrPosLiveToPreLive 
            Caption         =   "PosLive -> PreLive ..."
            Height          =   330
            Left            =   75
            TabIndex        =   332
            ToolTipText     =   "Transfer PosLiveData To PreLiveData"
            Top             =   1372
            Width           =   1665
         End
         Begin VB.CommandButton cmdCaptureSelected 
            Caption         =   "Capture &Selected"
            Enabled         =   0   'False
            Height          =   330
            Left            =   75
            TabIndex        =   334
            Top             =   525
            Width           =   1665
         End
         Begin VB.CommandButton cmdCaptureData 
            Caption         =   "&Capture All"
            Enabled         =   0   'False
            Height          =   330
            Left            =   75
            TabIndex        =   333
            ToolTipText     =   "Retrieve data from all included franchises"
            Top             =   180
            Width           =   1665
         End
         Begin VB.CommandButton cmdDisplayDialupResults 
            Caption         =   "Show Dialup Results"
            Height          =   330
            Left            =   75
            TabIndex        =   337
            Top             =   2160
            Width           =   1665
         End
         Begin VB.CommandButton cmdPrintRejectedData 
            Caption         =   "Show Rejects"
            Height          =   330
            Index           =   0
            Left            =   75
            TabIndex        =   336
            Top             =   1800
            Width           =   1665
         End
         Begin VB.CommandButton cmdCloseSelectedFranchises 
            Caption         =   "Close Selected"
            Enabled         =   0   'False
            Height          =   330
            Left            =   75
            TabIndex        =   335
            ToolTipText     =   "Close Selected Franchise(s)"
            Top             =   945
            Width           =   1665
         End
      End
      Begin VB.Frame Frame 
         Height          =   2850
         Index           =   16
         Left            =   2340
         TabIndex        =   287
         Top             =   690
         Width           =   7305
         Begin VB.ComboBox cboDCTabRegion 
            Height          =   315
            Left            =   3015
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   306
            Top             =   285
            Width           =   1635
         End
         Begin VB.TextBox txtVpnIpAddress 
            Height          =   285
            Left            =   4710
            Locked          =   -1  'True
            TabIndex        =   305
            Top             =   1380
            Width           =   1365
         End
         Begin VB.ComboBox cboState 
            Height          =   315
            ItemData        =   "TSG Data Warehouse.frx":043E
            Left            =   6330
            List            =   "TSG Data Warehouse.frx":045A
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   304
            Top             =   840
            Width           =   840
         End
         Begin VB.ComboBox cboFranchiseType 
            Height          =   315
            ItemData        =   "TSG Data Warehouse.frx":0483
            Left            =   5700
            List            =   "TSG Data Warehouse.frx":0485
            Locked          =   -1  'True
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   303
            Top             =   285
            Width           =   1470
         End
         Begin VB.CommandButton cmdAddNewFranchise 
            Caption         =   "< &Add New"
            Height          =   285
            Left            =   150
            TabIndex        =   302
            Top             =   285
            Width           =   1005
         End
         Begin VB.CommandButton cmdSaveFranchiseDetails 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   285
            Left            =   1335
            TabIndex        =   301
            Top             =   285
            Width           =   975
         End
         Begin VB.TextBox txtBATAFranchiseID 
            Height          =   285
            Left            =   5325
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   300
            Top             =   1920
            Width           =   615
         End
         Begin VB.TextBox txtContact 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   299
            Top             =   1380
            Width           =   2250
         End
         Begin VB.TextBox txtPhysicalAddress 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            TabIndex        =   298
            Top             =   840
            Width           =   4500
         End
         Begin VB.TextBox txtSuburb 
            Height          =   285
            Left            =   4710
            Locked          =   -1  'True
            TabIndex        =   297
            Top             =   840
            Width           =   1500
         End
         Begin VB.TextBox txtAreaCode 
            Height          =   285
            Left            =   150
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   296
            Top             =   1920
            Width           =   415
         End
         Begin VB.TextBox txtPhone 
            Height          =   285
            Left            =   750
            Locked          =   -1  'True
            TabIndex        =   295
            Top             =   2460
            Width           =   975
         End
         Begin VB.TextBox txtModem 
            Height          =   285
            Left            =   750
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   294
            Top             =   1920
            Width           =   975
         End
         Begin VB.CheckBox chkIncludeInDataCaptureCycle 
            Caption         =   "&Included in capture"
            Height          =   345
            Left            =   2730
            TabIndex        =   293
            Top             =   1350
            Width           =   1665
         End
         Begin VB.TextBox txtNodename 
            Height          =   285
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   292
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtFaxNum 
            Height          =   285
            Left            =   1830
            Locked          =   -1  'True
            TabIndex        =   291
            Top             =   2460
            Width           =   1335
         End
         Begin VB.TextBox txtRASUsername 
            Height          =   285
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   290
            Top             =   1920
            Width           =   1335
         End
         Begin VB.TextBox txtRASPassword 
            Height          =   285
            Left            =   3285
            Locked          =   -1  'True
            TabIndex        =   289
            Top             =   2460
            Width           =   1335
         End
         Begin VB.ComboBox cboDCTabPromoGrade 
            Height          =   315
            Left            =   6075
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   288
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label lblRegion 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            Height          =   195
            Index           =   20
            Left            =   2460
            TabIndex        =   329
            Top             =   330
            Width           =   510
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "BATA-ID"
            Height          =   195
            Index           =   7
            Left            =   5325
            TabIndex        =   328
            Top             =   1725
            Width           =   660
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "VPN IP Address"
            Height          =   195
            Index           =   2
            Left            =   4710
            TabIndex        =   327
            Top             =   1185
            Width           =   1170
         End
         Begin VB.Label lblFranchiseType 
            AutoSize        =   -1  'True
            Caption         =   "Fran' Type"
            Height          =   195
            Left            =   4860
            TabIndex        =   326
            Top             =   360
            Width           =   750
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "PriceModule"
            Height          =   195
            Index           =   29
            Left            =   6075
            TabIndex        =   325
            Top             =   2265
            Width           =   1035
         End
         Begin VB.Label lblPriceModuleVersion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   6075
            TabIndex        =   324
            Top             =   2460
            Width           =   1095
         End
         Begin VB.Label lblRemoteModuleVersion 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4710
            TabIndex        =   323
            Top             =   2460
            Width           =   1215
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "RemoteModule"
            Height          =   195
            Index           =   28
            Left            =   4710
            TabIndex        =   322
            Top             =   2280
            Width           =   1155
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Contact"
            Height          =   195
            Index           =   27
            Left            =   150
            TabIndex        =   321
            Top             =   1170
            Width           =   555
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Street"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   320
            Top             =   645
            Width           =   1215
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Suburb"
            Height          =   195
            Index           =   3
            Left            =   4710
            TabIndex        =   319
            Top             =   645
            Width           =   1215
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "STD"
            Height          =   195
            Index           =   4
            Left            =   150
            TabIndex        =   318
            Top             =   1725
            Width           =   360
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Modem"
            Height          =   195
            Index           =   6
            Left            =   750
            TabIndex        =   317
            Top             =   1725
            Width           =   690
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Node"
            Height          =   195
            Index           =   8
            Left            =   1830
            TabIndex        =   316
            Top             =   1725
            Width           =   930
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "RAS Username"
            Height          =   195
            Index           =   10
            Left            =   3285
            TabIndex        =   315
            Top             =   1725
            Width           =   1170
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "RAS Password"
            Height          =   195
            Index           =   11
            Left            =   3285
            TabIndex        =   314
            Top             =   2280
            Width           =   1170
         End
         Begin VB.Line Line 
            BorderWidth     =   3
            Index           =   0
            X1              =   600
            X2              =   695
            Y1              =   2040
            Y2              =   2040
         End
         Begin VB.Label lblFranID 
            BackStyle       =   0  'Transparent
            Caption         =   "Fran-ID"
            Height          =   195
            Index           =   7
            Left            =   4710
            TabIndex        =   313
            Top             =   1725
            Width           =   585
         End
         Begin VB.Label lblTSGFranchiseID 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   4710
            TabIndex        =   312
            Top             =   1920
            Width           =   555
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            Height          =   195
            Index           =   55
            Left            =   1830
            TabIndex        =   311
            Top             =   2280
            Width           =   930
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
            Height          =   195
            Index           =   5
            Left            =   750
            TabIndex        =   310
            Top             =   2280
            Width           =   615
         End
         Begin VB.Label lblState 
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   195
            Left            =   6330
            TabIndex        =   309
            Top             =   645
            Width           =   495
         End
         Begin VB.Label lblDCTabPromoGrade 
            BackStyle       =   0  'Transparent
            Caption         =   "Promo Grade"
            Height          =   195
            Left            =   6075
            TabIndex        =   308
            Top             =   1725
            Width           =   1095
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "To edit drop down list double click the caption"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   150
            Index           =   9
            Left            =   2460
            TabIndex        =   307
            Top             =   120
            Width           =   2520
         End
      End
      Begin VB.Frame frameNonCompliant 
         Caption         =   "Yesterday's Non-Compliant Franchises"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4785
         Left            =   -68355
         TabIndex        =   278
         Top             =   3510
         Width           =   4920
         Begin ComctlLib.ListView lvwNonCompliant 
            Height          =   3465
            Left            =   60
            TabIndex        =   120
            Top             =   195
            Width           =   4785
            _ExtentX        =   8440
            _ExtentY        =   6112
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            MultiSelect     =   -1  'True
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            _Version        =   327682
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Franchise"
               Object.Width           =   1023
            EndProperty
            BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   1
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Franchise ID"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   2
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Product"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   3
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Normal"
               Object.Width           =   706
            EndProperty
            BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   4
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Actual"
               Object.Width           =   706
            EndProperty
            BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   5
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Promo"
               Object.Width           =   706
            EndProperty
            BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
               SubItemIndex    =   6
               Key             =   ""
               Object.Tag             =   ""
               Text            =   "Date"
               Object.Width           =   706
            EndProperty
         End
         Begin VB.CommandButton cmdPromotion 
            Caption         =   "Print ALL"
            Height          =   330
            Index           =   5
            Left            =   615
            TabIndex        =   165
            Top             =   3765
            Width           =   1230
         End
         Begin VB.CommandButton cmdPromotion 
            Caption         =   "Print Selected"
            Enabled         =   0   'False
            Height          =   330
            Index           =   9
            Left            =   2370
            TabIndex        =   166
            Top             =   3765
            Width           =   1230
         End
         Begin VB.Frame Frame 
            Height          =   600
            Index           =   25
            Left            =   150
            TabIndex        =   279
            Top             =   4095
            Width           =   4680
            Begin VB.CommandButton cmdPromoTabSaveNonCompliantSelected 
               Caption         =   "Save to File (Selected)"
               Enabled         =   0   'False
               Height          =   315
               Left            =   1995
               TabIndex        =   172
               Top             =   195
               Width           =   1755
            End
            Begin VB.CommandButton cmdPromoTabSaveNonCompliantAll 
               Caption         =   "Save to File (ALL)"
               Height          =   315
               Left            =   105
               TabIndex        =   170
               Top             =   195
               Width           =   1755
            End
            Begin VB.OptionButton optNonCompliantSaveToFile 
               Caption         =   "TXT"
               Height          =   195
               Index           =   0
               Left            =   3870
               TabIndex        =   173
               Top             =   165
               Value           =   -1  'True
               Width           =   630
            End
            Begin VB.OptionButton optNonCompliantSaveToFile 
               Caption         =   "CSV"
               Height          =   195
               Index           =   1
               Left            =   3870
               TabIndex        =   174
               Top             =   360
               Width           =   630
            End
         End
      End
      Begin VB.Frame fraAddPromotion 
         Caption         =   "New Promotion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4785
         Left            =   -74865
         TabIndex        =   268
         Top             =   3510
         Width           =   6420
         Begin VB.Frame fraPromoSelFranchise 
            Height          =   2985
            Left            =   4710
            TabIndex        =   283
            Top             =   105
            Width           =   1710
            Begin VB.ListBox lstPromoFranchise 
               Height          =   2790
               Left            =   60
               MultiSelect     =   2  'Extended
               TabIndex        =   284
               Top             =   135
               Width           =   1590
            End
         End
         Begin VB.CheckBox chkPromoSelectFranchise 
            Caption         =   "Select Specific Franchisse(s)"
            Height          =   360
            Left            =   630
            TabIndex        =   282
            Top             =   1320
            Width           =   1485
         End
         Begin VB.CommandButton cmdPromotion 
            Caption         =   "Save Promotion"
            Height          =   525
            Index           =   0
            Left            =   5385
            TabIndex        =   162
            Top             =   3555
            Width           =   945
         End
         Begin VB.CommandButton cmdPromotion 
            Caption         =   "Clear"
            Height          =   315
            Index           =   1
            Left            =   5370
            TabIndex        =   169
            Top             =   4215
            Width           =   975
         End
         Begin VB.TextBox txtPromoName 
            Height          =   285
            Left            =   630
            TabIndex        =   105
            Top             =   210
            Width           =   1575
         End
         Begin VB.Frame Frame 
            Height          =   1515
            Index           =   28
            Left            =   555
            TabIndex        =   270
            Top             =   3165
            Width           =   1650
            Begin VB.ListBox lstPromoTabRegion 
               Enabled         =   0   'False
               Height          =   1035
               Left            =   60
               MultiSelect     =   2  'Extended
               TabIndex        =   159
               Top             =   420
               Width           =   1425
            End
            Begin VB.CheckBox chkPromoTabAllRegions 
               Caption         =   "All Regions"
               Height          =   195
               Left            =   60
               TabIndex        =   158
               Top             =   165
               Value           =   1  'Checked
               Width           =   1125
            End
         End
         Begin VB.Frame fraPromoSelectState 
            Height          =   1485
            Left            =   555
            TabIndex        =   269
            Top             =   1695
            Width           =   1650
            Begin VB.ListBox lstPromoTabState 
               Enabled         =   0   'False
               Height          =   1035
               Left            =   60
               MultiSelect     =   2  'Extended
               TabIndex        =   149
               Top             =   375
               Width           =   1425
            End
            Begin VB.CheckBox chkPromoTabAllStates 
               Caption         =   "All States"
               Height          =   195
               Left            =   60
               TabIndex        =   131
               Top             =   165
               Value           =   1  'Checked
               Width           =   990
            End
         End
         Begin VB.ListBox lstPromoProducts 
            BackColor       =   &H80000004&
            Height          =   2595
            Left            =   3435
            TabIndex        =   113
            Top             =   510
            Width           =   2895
         End
         Begin VB.ListBox lstPromoSubCat 
            Height          =   2595
            Left            =   2370
            MultiSelect     =   2  'Extended
            TabIndex        =   108
            Top             =   510
            Width           =   975
         End
         Begin MSComCtl2.DTPicker dtpPromoEnd 
            Height          =   285
            Left            =   630
            TabIndex        =   132
            Top             =   885
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   164954115
            CurrentDate     =   38477
         End
         Begin MSComCtl2.DTPicker dtpPromoStart 
            Height          =   285
            Left            =   630
            TabIndex        =   106
            Top             =   555
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dd MMM yyyy"
            Format          =   164954115
            CurrentDate     =   38477
         End
         Begin VSFlex8Ctl.VSFlexGrid grdPromoTabRebates 
            Height          =   1200
            Left            =   2340
            TabIndex        =   160
            Top             =   3450
            Width           =   2955
            _cx             =   5212
            _cy             =   2117
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   3
            GridLinesFixed  =   3
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"TSG Data Warehouse.frx":0487
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   5
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   1
            WordWrap        =   -1  'True
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   1
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   2
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   -2147483630
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblPromoFran 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Fran"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   4215
            TabIndex        =   285
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblAddPromotion 
            BackStyle       =   0  'Transparent
            Caption         =   "Promotion Rebates"
            Height          =   195
            Index           =   4
            Left            =   2340
            TabIndex        =   277
            Top             =   3255
            Width           =   1365
         End
         Begin VB.Label lblAddPromotion 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "End"
            Height          =   180
            Index           =   3
            Left            =   195
            TabIndex        =   276
            Top             =   930
            Width           =   330
         End
         Begin VB.Label lblAddPromotion 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Start"
            Height          =   180
            Index           =   2
            Left            =   210
            TabIndex        =   275
            Top             =   600
            Width           =   330
         End
         Begin VB.Label lblAddPromotion 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Name"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   274
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblPromoState 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "State"
            Height          =   195
            Left            =   90
            TabIndex        =   273
            Top             =   1875
            Width           =   405
         End
         Begin VB.Label lblPromoRegion 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Region"
            Height          =   195
            Left            =   0
            TabIndex        =   272
            Top             =   3345
            Width           =   525
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Sub Categories"
            Height          =   195
            Index           =   56
            Left            =   2385
            TabIndex        =   271
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.Frame fraAztec 
         Caption         =   "AZTEC Uploads"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4260
         Left            =   -74715
         TabIndex        =   265
         Top             =   3990
         Width           =   10875
         Begin VSFlex8Ctl.VSFlexGrid grdAztecUploads 
            Height          =   3930
            Left            =   255
            TabIndex        =   104
            Top             =   240
            Width           =   6270
            _cx             =   11060
            _cy             =   6932
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"TSG Data Warehouse.frx":053C
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.CommandButton cmdAztecUpload 
            Caption         =   "Upload Most Recent Data (IF NOT UPLOADED)     to AZTEC"
            Height          =   885
            Left            =   7725
            TabIndex        =   122
            Top             =   240
            Width           =   2010
         End
      End
      Begin VB.CheckBox chkStockTab_IncludeDeletedStock 
         Caption         =   "Show Deleted Stock"
         Height          =   195
         Left            =   -74640
         TabIndex        =   2
         Top             =   750
         Width           =   1920
      End
      Begin VB.CheckBox chkSalesRptTab_IncludeClosedFrans 
         Caption         =   "Show Closed Franchises"
         Height          =   195
         Left            =   -73815
         TabIndex        =   3
         Top             =   540
         Width           =   2190
      End
      Begin VB.Frame Frame 
         Caption         =   "Export or Delete Selected Stock Items"
         Height          =   3060
         Index           =   13
         Left            =   -74835
         TabIndex        =   264
         Top             =   5235
         Width           =   6330
         Begin VB.CommandButton cmdStockTabDelete 
            Caption         =   "Delete Selected Stock"
            Enabled         =   0   'False
            Height          =   330
            Left            =   4380
            TabIndex        =   161
            Top             =   1275
            Width           =   1755
         End
         Begin VB.CommandButton cmdStockTabExport 
            Caption         =   "Export &Selected Stock"
            Enabled         =   0   'False
            Height          =   330
            Left            =   4380
            TabIndex        =   153
            Top             =   345
            Width           =   1755
         End
         Begin VB.ListBox lstStcokTabSelectedSoctkExport 
            Height          =   2595
            Left            =   135
            MultiSelect     =   2  'Extended
            TabIndex        =   151
            Top             =   315
            Width           =   4095
         End
      End
      Begin VB.CommandButton cmdPromotionRecall 
         Caption         =   "Recall Sel'd"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -64125
         TabIndex        =   43
         Top             =   825
         Width           =   660
      End
      Begin VB.TextBox txtNewFranchiseBusinessName 
         Height          =   375
         Left            =   120
         TabIndex        =   44
         Top             =   780
         Visible         =   0   'False
         Width           =   2100
      End
      Begin ComctlLib.ListView lvwVersions 
         Height          =   7605
         Left            =   -74640
         TabIndex        =   9
         Top             =   600
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   13414
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Franchise"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Type"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "In Capture"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Remote Stat's"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Price Module"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Retail Mgr"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Operating Sys"
            Object.Width           =   1587
         EndProperty
      End
      Begin ComctlLib.ListView lvwUploadsPending 
         Height          =   2415
         Left            =   -74760
         TabIndex        =   156
         Top             =   5835
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   4260
         View            =   3
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Franchise"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Item to be Uploaded"
            Object.Width           =   6174
         EndProperty
      End
      Begin VB.Frame Frame 
         Caption         =   "State Specifics"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Index           =   29
         Left            =   -66915
         TabIndex        =   242
         Top             =   1410
         Width           =   2265
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   7
            Left            =   1005
            TabIndex        =   144
            Top             =   3120
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   6
            Left            =   1005
            TabIndex        =   124
            Top             =   2775
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   5
            Left            =   1005
            TabIndex        =   123
            Top             =   2415
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   4
            Left            =   1005
            TabIndex        =   101
            Top             =   2040
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   3
            Left            =   1005
            TabIndex        =   90
            Top             =   1695
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   2
            Left            =   1005
            TabIndex        =   79
            Top             =   1335
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   1
            Left            =   1005
            TabIndex        =   78
            Top             =   975
            Width           =   500
         End
         Begin VB.TextBox txtDialSequence 
            Height          =   350
            Index           =   0
            Left            =   1005
            TabIndex        =   50
            Top             =   615
            Width           =   500
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Dial Sequence"
            Height          =   300
            Index           =   49
            Left            =   915
            TabIndex        =   251
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "NSW"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   250
            Top             =   720
            Width           =   495
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "VIC"
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   249
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "QLD"
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   248
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "SA"
            Height          =   255
            Index           =   3
            Left            =   480
            TabIndex        =   247
            Top             =   1800
            Width           =   375
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "WA"
            Height          =   255
            Index           =   4
            Left            =   480
            TabIndex        =   246
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "TAS"
            Height          =   255
            Index           =   5
            Left            =   480
            TabIndex        =   245
            Top             =   2520
            Width           =   375
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "ACT"
            Height          =   255
            Index           =   6
            Left            =   480
            TabIndex        =   244
            Top             =   2880
            Width           =   375
         End
         Begin VB.Label lblSateName 
            BackStyle       =   0  'Transparent
            Caption         =   "NT"
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   243
            Top             =   3240
            Width           =   255
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Master - Slave Configuration"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   27
         Left            =   -74760
         TabIndex        =   240
         Top             =   3637
         Width           =   6570
         Begin VB.TextBox txtThisNodeName 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   107
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "This ComputerName"
            Height          =   255
            Index           =   37
            Left            =   120
            TabIndex        =   241
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.CommandButton cmdStickReport 
         Caption         =   "Start &Report"
         Height          =   285
         Left            =   -65205
         TabIndex        =   127
         Top             =   3930
         Width           =   1200
      End
      Begin VB.Frame Frame 
         Caption         =   "Step 4: Send now or later"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Index           =   35
         Left            =   -67680
         TabIndex        =   237
         Top             =   6540
         Width           =   4065
         Begin VB.CommandButton UploadToFranchises 
            Caption         =   "Upload ALL NOW"
            Height          =   510
            Index           =   0
            Left            =   1575
            TabIndex        =   167
            ToolTipText     =   "Upload the stuff you just nominated in this session as well as all the other stuff that was pending before you got in here."
            Top             =   270
            Width           =   915
         End
         Begin VB.CommandButton UploadToFranchises 
            Caption         =   "&Upload This NOW"
            Height          =   510
            Index           =   2
            Left            =   150
            TabIndex        =   164
            ToolTipText     =   $"TSG Data Warehouse.frx":05E9
            Top             =   270
            Width           =   915
         End
         Begin VB.CommandButton UploadToFranchises 
            Caption         =   "Upload ALL Later"
            Height          =   510
            Index           =   1
            Left            =   3000
            TabIndex        =   168
            Top             =   270
            Width           =   915
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "OR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   50
            Left            =   2595
            TabIndex        =   239
            Top             =   405
            Width           =   270
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "OR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   36
            Left            =   1185
            TabIndex        =   238
            Top             =   405
            Width           =   270
         End
      End
      Begin VB.ListBox lstProductReportsFranchiseBusinessName 
         Height          =   4350
         Left            =   -74880
         MultiSelect     =   2  'Extended
         TabIndex        =   5
         Top             =   765
         Width           =   2100
      End
      Begin VB.Frame Frame 
         Caption         =   "Step 2: Select items to upload"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Index           =   33
         Left            =   -74775
         TabIndex        =   236
         Top             =   3420
         Width           =   6660
         Begin VB.CommandButton btnClearUploadDir 
            Caption         =   "Clear Upload List"
            Height          =   795
            Left            =   5730
            TabIndex        =   138
            Top             =   1170
            Width           =   750
         End
         Begin VB.CommandButton cmdBrowseUploads 
            Caption         =   "Browse"
            Height          =   375
            Left            =   5715
            TabIndex        =   118
            Top             =   720
            Width           =   735
         End
         Begin VB.ListBox lstUploadItemList 
            Height          =   1815
            Left            =   180
            MultiSelect     =   2  'Extended
            TabIndex        =   103
            Top             =   285
            Width           =   5340
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Send report to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Index           =   0
         Left            =   -72660
         TabIndex        =   235
         Top             =   3120
         Width           =   3825
         Begin VB.CheckBox chkProductReportTabDelimited 
            Caption         =   "Tab deli&mited"
            Height          =   225
            Left            =   120
            TabIndex        =   109
            Top             =   840
            Width           =   1560
         End
         Begin VB.OptionButton optSendProductReportToFile 
            Caption         =   "&Temporary report file >>>>"
            Height          =   240
            Left            =   120
            TabIndex        =   110
            Top             =   592
            Width           =   2205
         End
         Begin VB.OptionButton optSendProductReportToPrinter 
            Caption         =   "&Printer"
            Height          =   255
            Left            =   2160
            TabIndex        =   100
            Top             =   240
            Width           =   810
         End
         Begin VB.OptionButton optSendProductReportToDisplay 
            Caption         =   "&Display(Below)"
            Height          =   300
            Left            =   120
            TabIndex        =   99
            Top             =   285
            Value           =   -1  'True
            Width           =   1440
         End
         Begin VB.CommandButton cmdViewProductReport 
            Caption         =   "&View"
            Height          =   285
            Left            =   2280
            TabIndex        =   116
            Top             =   600
            Width           =   735
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Report on:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   1
         Left            =   -72660
         TabIndex        =   234
         Top             =   720
         Width           =   2025
         Begin VB.OptionButton optProductReportOnSelectedFranchisesOnly 
            Caption         =   "&Selected franchise(s)"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   270
            Value           =   -1  'True
            Width           =   1780
         End
         Begin VB.OptionButton optProductReportOnSelectedFranchisesOnly 
            Caption         =   "All franchises"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   525
            Width           =   1335
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Present:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   2
         Left            =   -70455
         TabIndex        =   233
         Top             =   720
         Width           =   1290
         Begin VB.OptionButton optDescription 
            Caption         =   "D&escription"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   285
            Value           =   -1  'True
            Width           =   1110
         End
         Begin VB.OptionButton optDescription 
            Caption         =   "Barcode"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   525
            Width           =   900
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Transaction period:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   3
         Left            =   -72660
         TabIndex        =   228
         Top             =   4320
         Width           =   3705
         Begin VB.TextBox lblProductReportFinishDate 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2490
            TabIndex        =   136
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox lblProductReportStartDate 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   930
            TabIndex        =   133
            Top             =   240
            Width           =   855
         End
         Begin Spin.SpinButton spnProductReportFinishDate 
            Height          =   195
            Left            =   2490
            TabIndex        =   229
            Top             =   555
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Delay           =   100
            ShadowThickness =   1
            SpinOrientation =   1
            TdThickness     =   1
         End
         Begin Spin.SpinButton spnProductReportStartDate 
            Height          =   195
            Left            =   945
            TabIndex        =   230
            Top             =   555
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Delay           =   100
            ShadowThickness =   1
            SpinOrientation =   1
            TdThickness     =   1
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "to:"
            Height          =   195
            Index           =   17
            Left            =   2055
            TabIndex        =   232
            Top             =   240
            Width           =   165
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
            Height          =   195
            Index           =   13
            Left            =   405
            TabIndex        =   231
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Transaction period:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   5
         Left            =   -72750
         TabIndex        =   221
         Top             =   3675
         Width           =   3945
         Begin Spin.SpinButton spnStickReportStartDate 
            Height          =   195
            Left            =   720
            TabIndex        =   222
            Top             =   540
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Delay           =   100
            ShadowThickness =   1
            SpinOrientation =   1
            TdThickness     =   1
         End
         Begin Spin.SpinButton spnStickReportFinishDate 
            Height          =   195
            Left            =   2235
            TabIndex        =   223
            Top             =   525
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Delay           =   100
            ShadowThickness =   1
            SpinOrientation =   1
            TdThickness     =   1
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "From:"
            Height          =   195
            Index           =   14
            Left            =   150
            TabIndex        =   227
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblStickReportStartDate 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   720
            TabIndex        =   226
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lblStickReportFinishDate 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2235
            TabIndex        =   225
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "to:"
            Height          =   195
            Index           =   24
            Left            =   1800
            TabIndex        =   224
            Top             =   240
            Width           =   165
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Report on:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Index           =   6
         Left            =   -72780
         TabIndex        =   220
         Top             =   1335
         Width           =   3960
         Begin VB.OptionButton optStickReportOnSelectedFranchisesOnly 
            Caption         =   "&Selected franchise(s)"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton optStickReportOnSelectedFranchisesOnly 
            Caption         =   "All Franchises"
            Height          =   195
            Index           =   1
            Left            =   2175
            TabIndex        =   22
            Top             =   240
            Width           =   1290
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Send report to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   7
         Left            =   -72750
         TabIndex        =   219
         Top             =   2790
         Width           =   8970
         Begin VB.CommandButton cmdViewStickReport 
            Caption         =   "&View"
            Enabled         =   0   'False
            Height          =   285
            Left            =   5955
            TabIndex        =   91
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optSendStickReportToPrinter 
            Caption         =   "&Printer"
            Height          =   255
            Left            =   2000
            TabIndex        =   87
            Top             =   255
            Width           =   775
         End
         Begin VB.OptionButton optSendStickReportToFile 
            Caption         =   "&Temporary report file >>>>"
            Height          =   255
            Left            =   3675
            TabIndex        =   89
            Top             =   255
            Width           =   2175
         End
         Begin VB.OptionButton optSendStickReportToDisplay 
            Caption         =   "&Display (below)"
            Height          =   255
            Left            =   120
            TabIndex        =   85
            Top             =   255
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.CheckBox chkStickReportTabDelimited 
            Caption         =   "Tab deli&mited"
            Enabled         =   0   'False
            Height          =   255
            Left            =   7230
            TabIndex        =   94
            Top             =   255
            Width           =   1440
         End
      End
      Begin VB.ListBox lstStickReportsFranchiseBusinessName 
         Height          =   2985
         Left            =   -74820
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   6
         Top             =   1425
         Width           =   1905
      End
      Begin VB.Frame Frame 
         Caption         =   "Show:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   8
         Left            =   -68670
         TabIndex        =   218
         Top             =   2040
         Width           =   4875
         Begin VB.OptionButton optStickReportDescription 
            Caption         =   "D&escription"
            Height          =   255
            Index           =   0
            Left            =   150
            TabIndex        =   74
            Top             =   255
            Value           =   -1  'True
            Width           =   1780
         End
         Begin VB.OptionButton optStickReportDescription 
            Caption         =   "Barcode"
            Height          =   255
            Index           =   1
            Left            =   2160
            TabIndex        =   80
            Top             =   240
            Width           =   1780
         End
      End
      Begin VB.ListBox lstStickReportRecipient 
         Height          =   645
         Left            =   -68640
         Sorted          =   -1  'True
         TabIndex        =   119
         Top             =   3765
         Width           =   3090
      End
      Begin VB.Frame Frame 
         Caption         =   "Inclusions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   9
         Left            =   -70455
         TabIndex        =   217
         Top             =   2010
         Width           =   3075
         Begin VB.CheckBox chkIncludeTotalCustomerCount 
            Caption         =   "C&ustomer count"
            Height          =   240
            Left            =   120
            TabIndex        =   69
            Top             =   480
            Width           =   1440
         End
         Begin VB.CheckBox chkNonTobaccoBarcodesAreIncluded 
            Caption         =   "Non-Tobacco Items "
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   240
            Width           =   2325
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Style:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Index           =   10
         Left            =   -72660
         TabIndex        =   216
         Top             =   2010
         Width           =   2025
         Begin VB.OptionButton optProductReportNotSummarised 
            Caption         =   "Not summarised"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   64
            Top             =   255
            Value           =   -1  'True
            Width           =   1440
         End
         Begin VB.OptionButton optProductReportNotSummarised 
            Caption         =   "Summarised"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   63
            Top             =   495
            Width           =   1170
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Detail Type"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   11
         Left            =   -72765
         TabIndex        =   215
         Top             =   2025
         Width           =   3945
         Begin VB.OptionButton optStickReportSummaryType 
            Caption         =   "Full"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton optStickReportSummaryType 
            Caption         =   "Summarised by Fran"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   66
            Top             =   240
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton optStickReportSummaryType 
            Caption         =   "Totals"
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   71
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame 
         Height          =   4590
         Index           =   12
         Left            =   -74805
         TabIndex        =   204
         Top             =   555
         Width           =   11220
         Begin VB.TextBox txtStock_ID_DEV 
            BackColor       =   &H0000FF00&
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9600
            TabIndex        =   280
            Top             =   765
            Width           =   1275
         End
         Begin VB.TextBox txtCartonsPerPacket 
            Height          =   285
            Left            =   9360
            TabIndex        =   126
            Top             =   3435
            Width           =   1185
         End
         Begin VB.ComboBox cboCtnContainingPkt 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   315
            Left            =   4365
            Style           =   2  'Dropdown List
            TabIndex        =   114
            Top             =   3435
            Width           =   4890
         End
         Begin VB.Frame fra 
            Caption         =   "Export Stock"
            Height          =   585
            Left            =   4365
            TabIndex        =   263
            Top             =   3885
            Width           =   4890
            Begin VB.CommandButton btnExportStock 
               Caption         =   "Export Stock"
               Height          =   330
               Left            =   3630
               TabIndex        =   145
               ToolTipText     =   "Create the 6 monthly price update file"
               Top             =   165
               Width           =   1110
            End
            Begin VB.CheckBox chkExportStkCategory 
               Caption         =   "Tobacco"
               Height          =   195
               Index           =   3
               Left            =   2535
               TabIndex        =   142
               Top             =   240
               Width           =   960
            End
            Begin VB.CheckBox chkExportStkCategory 
               Caption         =   "Cigars"
               Height          =   195
               Index           =   2
               Left            =   1732
               TabIndex        =   140
               Top             =   240
               Width           =   750
            End
            Begin VB.CheckBox chkExportStkCategory 
               Caption         =   "Cigarette Cartons"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   150
               TabIndex        =   135
               Top             =   240
               Value           =   1  'Checked
               Width           =   1530
            End
         End
         Begin VB.CommandButton cmdMerger 
            Caption         =   "Merge Databases"
            Height          =   330
            Left            =   9615
            TabIndex        =   148
            Top             =   4050
            Width           =   1425
         End
         Begin VB.TextBox txtStkItemDescription 
            Height          =   345
            Left            =   7305
            TabIndex        =   48
            Top             =   1410
            Width           =   3735
         End
         Begin VB.ComboBox cboSupplier 
            Height          =   315
            Left            =   4365
            Style           =   2  'Dropdown List
            TabIndex        =   46
            Top             =   1410
            Width           =   2295
         End
         Begin VB.CommandButton cmdSaveStockDetails 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   330
            Left            =   8040
            TabIndex        =   37
            Top             =   765
            Width           =   1020
         End
         Begin VB.TextBox txtBarcode 
            Height          =   285
            Left            =   4365
            TabIndex        =   24
            Top             =   765
            Width           =   2295
         End
         Begin VB.TextBox txtSticks 
            Height          =   285
            Left            =   4365
            TabIndex        =   67
            Top             =   2175
            Width           =   1485
         End
         Begin VB.CommandButton cmdAddNewStockItem 
            Caption         =   "< &Add New"
            Height          =   330
            Left            =   6825
            TabIndex        =   33
            Top             =   765
            Width           =   1020
         End
         Begin VB.ComboBox cboCategory 
            Height          =   315
            Left            =   4365
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   2820
            Width           =   1725
         End
         Begin VB.ComboBox cboSubCategory 
            Height          =   315
            ItemData        =   "TSG Data Warehouse.frx":0684
            Left            =   6705
            List            =   "TSG Data Warehouse.frx":0686
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   2820
            Width           =   1350
         End
         Begin VB.TextBox txtRRP 
            Height          =   315
            Left            =   7665
            TabIndex        =   76
            Top             =   2175
            Width           =   1200
         End
         Begin VB.TextBox txtWholesaleListPrice 
            Height          =   300
            Left            =   6030
            TabIndex        =   73
            Top             =   2175
            Width           =   1395
         End
         Begin VB.CheckBox chkUploadWholesaleListPrice 
            Caption         =   "Include in Upload"
            Height          =   270
            Left            =   6825
            TabIndex        =   32
            Top             =   465
            Value           =   1  'Checked
            Width           =   1545
         End
         Begin VB.ListBox lstDescription 
            Height          =   3960
            ItemData        =   "TSG Data Warehouse.frx":0688
            Left            =   120
            List            =   "TSG Data Warehouse.frx":068A
            TabIndex        =   7
            Top             =   495
            Width           =   4095
         End
         Begin VB.ComboBox cboGoodsTax 
            Height          =   315
            Left            =   9120
            TabIndex        =   81
            Top             =   2175
            Width           =   855
         End
         Begin VB.ComboBox cboSalesTax 
            Height          =   315
            Left            =   10185
            TabIndex        =   83
            Top             =   2175
            Width           =   855
         End
         Begin VB.CheckBox chkPackage 
            Caption         =   "Package (Packet or single item as part of a carton or box)"
            Height          =   555
            Left            =   8700
            TabIndex        =   92
            Top             =   2670
            Width           =   2460
         End
         Begin VB.CommandButton btnAddSubcat 
            Height          =   255
            Index           =   0
            Left            =   8160
            TabIndex        =   98
            Top             =   2820
            Width           =   255
         End
         Begin VB.CommandButton btnAddSubcat 
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   96
            Top             =   2820
            Width           =   255
         End
         Begin VB.Label lblStock_ID_DEV 
            BackColor       =   &H0000FF00&
            Caption         =   "View Stock_ID value in DEV"
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   8790
            TabIndex        =   281
            Top             =   540
            Width           =   2085
         End
         Begin VB.Label lblCartonsPerPacket 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Qty Per Package"
            Height          =   195
            Left            =   9360
            TabIndex        =   267
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblCtnContainingPkt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Component"
            Height          =   195
            Left            =   4365
            TabIndex        =   266
            Top             =   3240
            Width           =   810
         End
         Begin VB.Label lblRRP 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RRP (ex tax)"
            Height          =   195
            Left            =   7680
            TabIndex        =   214
            Top             =   1950
            Width           =   900
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Supplier"
            Height          =   195
            Index           =   22
            Left            =   4365
            TabIndex        =   213
            Top             =   1185
            Width           =   570
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Barcode"
            Height          =   195
            Index           =   21
            Left            =   4365
            TabIndex        =   212
            Top             =   525
            Width           =   600
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sub-category"
            Height          =   195
            Index           =   18
            Left            =   6720
            TabIndex        =   211
            Top             =   2595
            Width           =   945
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Category"
            Height          =   195
            Index           =   16
            Left            =   4365
            TabIndex        =   210
            Top             =   2595
            Width           =   630
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sticks"
            Height          =   195
            Index           =   15
            Left            =   4365
            TabIndex        =   209
            Top             =   1950
            Width           =   435
         End
         Begin VB.Label lblWholesaleListPrice 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cost (exTax)"
            Height          =   195
            Left            =   6060
            TabIndex        =   208
            Top             =   1950
            Width           =   885
         End
         Begin VB.Label lblDescription 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   " Description"
            Height          =   195
            Left            =   7305
            TabIndex        =   207
            Top             =   1200
            Width           =   840
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Goods Tax"
            Height          =   195
            Index           =   43
            Left            =   9120
            TabIndex        =   206
            Top             =   1950
            Width           =   780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Tax"
            Height          =   195
            Index           =   42
            Left            =   10185
            TabIndex        =   205
            Top             =   1950
            Width           =   705
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Report 3 weeks up to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   14
         Left            =   -74760
         TabIndex        =   203
         Top             =   825
         Width           =   2565
         Begin MSComCtl2.UpDown updNielsenRptTxDate 
            Height          =   285
            Left            =   2145
            TabIndex        =   16
            Top             =   255
            Width           =   240
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            OrigLeft        =   480
            OrigTop         =   870
            OrigRight       =   720
            OrigBottom      =   1185
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker dtpNielsenRptTxDate 
            Height          =   285
            Left            =   180
            TabIndex        =   12
            Top             =   240
            Width           =   2235
            _ExtentX        =   3942
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "dddd, dd MMM yyyy"
            Format          =   164954115
            CurrentDate     =   38512
         End
      End
      Begin VB.CommandButton cmdPurgeNielsenReportList 
         Caption         =   "&Purge"
         Height          =   320
         Left            =   -72015
         TabIndex        =   20
         Top             =   1275
         Width           =   1230
      End
      Begin VB.ListBox lstNielsenReportDisplayDate 
         Height          =   2205
         Left            =   -74715
         Sorted          =   -1  'True
         TabIndex        =   8
         Top             =   1650
         Width           =   3930
      End
      Begin VB.CommandButton cmdCreateNielsenReports 
         Caption         =   "&Create"
         Height          =   320
         Left            =   -72015
         TabIndex        =   19
         Top             =   915
         Width           =   1230
      End
      Begin VB.Frame Frame 
         Caption         =   "Archive Options"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2040
         Index           =   24
         Left            =   -74760
         TabIndex        =   199
         Top             =   1410
         Width           =   6570
         Begin VB.TextBox txtIOMEGADrive 
            BackColor       =   &H0000FF00&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   240
            TabIndex        =   60
            Text            =   " "
            Top             =   1350
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtDaysKeepLiveData 
            BackColor       =   &H0000FF00&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   375
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.TextBox txtTxnStartDate 
            BackColor       =   &H80000004&
            Enabled         =   0   'False
            Height          =   285
            Left            =   240
            Locked          =   -1  'True
            TabIndex        =   61
            Top             =   862
            Width           =   1215
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "ZIP Drive where Database is copied to automatically each night - V386"
            Height          =   195
            Index           =   47
            Left            =   1200
            TabIndex        =   202
            Top             =   1425
            Visible         =   0   'False
            Width           =   5010
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackColor       =   &H0000FF00&
            Caption         =   "Keep this many days of LiveData - V386"
            Height          =   195
            Index           =   53
            Left            =   855
            TabIndex        =   201
            Top             =   420
            Visible         =   0   'False
            Width           =   2835
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Oldest sales in database"
            Height          =   285
            Index           =   54
            Left            =   1560
            TabIndex        =   200
            Top             =   922
            Width           =   2250
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Report type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   4
         Left            =   -68895
         TabIndex        =   197
         Top             =   4320
         Width           =   5280
         Begin VB.CommandButton cmdAllItems 
            Caption         =   "All &items"
            Height          =   285
            Left            =   120
            TabIndex        =   139
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton cmdTopSellers 
            Caption         =   "To&p"
            Height          =   285
            Left            =   2280
            TabIndex        =   146
            Top             =   360
            Width           =   780
         End
         Begin VB.CommandButton cmdMarketShare 
            Caption         =   "S&hare"
            Height          =   285
            Left            =   1200
            TabIndex        =   143
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnMissingDaysSales 
            Caption         =   "Missing Days"
            Height          =   285
            Left            =   3360
            TabIndex        =   147
            Top             =   360
            Width           =   1095
         End
         Begin Spin.SpinButton spnTopSellers 
            Height          =   285
            Left            =   3000
            TabIndex        =   198
            Top             =   360
            Width           =   195
            _Version        =   65536
            _ExtentX        =   344
            _ExtentY        =   503
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            ShadowThickness =   1
            TdThickness     =   1
         End
      End
      Begin VB.TextBox txtActiveDatabase 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   120
         TabIndex        =   175
         Top             =   8010
         Width           =   9495
      End
      Begin VB.Timer tmrAutoDataCapture 
         Interval        =   60000
         Left            =   11010
         Top             =   3780
      End
      Begin VB.CommandButton btnPurgeUploadsPending 
         Caption         =   "<--- Purge Uploads Pending "
         Height          =   420
         Left            =   -67665
         TabIndex        =   171
         Top             =   7590
         Width           =   2100
      End
      Begin VB.ListBox lstPRProductReportsFranchiseBusinessName 
         Height          =   3765
         Left            =   -74640
         MultiSelect     =   2  'Extended
         TabIndex        =   11
         Top             =   1005
         Width           =   2450
      End
      Begin VB.CommandButton cmdPRPrint 
         Caption         =   "Report"
         Height          =   435
         Left            =   -68160
         TabIndex        =   141
         Top             =   4350
         Width           =   1200
      End
      Begin VB.Frame Frame 
         Caption         =   "Send report to:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Index           =   23
         Left            =   -72090
         TabIndex        =   196
         Top             =   3045
         Width           =   5175
         Begin VB.OptionButton optPRSendProductReportToFile 
            Caption         =   "Temporary Report File >>>>"
            Height          =   285
            Left            =   1680
            TabIndex        =   115
            Top             =   210
            Width           =   2340
         End
         Begin VB.OptionButton optPRSendProductReportToDisplay 
            Caption         =   "Display(Below)"
            Height          =   285
            Left            =   150
            TabIndex        =   111
            Top             =   210
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.CheckBox chkPRProductReportTabDelimited 
            Caption         =   "Tab deli&mited"
            Height          =   240
            Left            =   1725
            TabIndex        =   117
            Top             =   555
            Width           =   2220
         End
         Begin VB.OptionButton optPRSendProductReportToPrinter 
            Caption         =   "&Printer"
            Height          =   240
            Left            =   150
            TabIndex        =   112
            Top             =   540
            Width           =   1335
         End
         Begin VB.CommandButton cmdPRView 
            Caption         =   "&View"
            Height          =   285
            Left            =   3975
            TabIndex        =   121
            Top             =   195
            Width           =   975
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Report on:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   22
         Left            =   -72120
         TabIndex        =   195
         Top             =   885
         Width           =   3720
         Begin VB.OptionButton optPRReportonSelectedFranchises 
            Caption         =   "Selected &franchise(s)"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   21
            Top             =   255
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.OptionButton optPRReportonSelectedFranchises 
            Caption         =   "All Franchises"
            Height          =   240
            Index           =   1
            Left            =   2160
            TabIndex        =   27
            Top             =   255
            Width           =   1335
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Present:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Index           =   21
         Left            =   -68280
         TabIndex        =   194
         Top             =   885
         Width           =   1335
         Begin VB.OptionButton optPRDescription 
            Caption         =   "D&escription"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optPRDescription 
            Caption         =   "Barcode"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Transaction period:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Index           =   20
         Left            =   -72090
         TabIndex        =   189
         Top             =   4005
         Width           =   3705
         Begin VB.TextBox lblPRProductReportFinishDate 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2520
            TabIndex        =   137
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox lblPRProductReportStartDate 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   960
            TabIndex        =   134
            Top             =   240
            Width           =   855
         End
         Begin Spin.SpinButton spnPRProductReportFinishDate 
            Height          =   195
            Left            =   2520
            TabIndex        =   190
            Top             =   555
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Delay           =   100
            ShadowThickness =   1
            SpinOrientation =   1
            TdThickness     =   1
         End
         Begin Spin.SpinButton spnPRProductReportStartDate 
            Height          =   195
            Left            =   975
            TabIndex        =   191
            Top             =   555
            Width           =   855
            _Version        =   65536
            _ExtentX        =   1508
            _ExtentY        =   344
            _StockProps     =   73
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Delay           =   100
            ShadowThickness =   1
            SpinOrientation =   1
            TdThickness     =   1
         End
         Begin VB.Label lbl 
            Caption         =   "From:"
            Height          =   195
            Index           =   32
            Left            =   405
            TabIndex        =   193
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl 
            Caption         =   "to:"
            Height          =   195
            Index           =   31
            Left            =   2055
            TabIndex        =   192
            Top             =   240
            Width           =   165
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Style:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   18
         Left            =   -72090
         TabIndex        =   188
         Top             =   2355
         Width           =   3705
         Begin VB.OptionButton optPRProductReportNotSummarised 
            Caption         =   "Not summarised"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   86
            Top             =   240
            Width           =   1560
         End
         Begin VB.OptionButton optPRProductReportNotSummarised 
            Caption         =   "Summarised"
            Height          =   195
            Index           =   1
            Left            =   2160
            TabIndex        =   88
            Top             =   240
            Value           =   -1  'True
            Width           =   1200
         End
      End
      Begin VB.ListBox lstPRProductList 
         Height          =   3765
         Left            =   -66675
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   38
         Top             =   990
         Width           =   3045
      End
      Begin VB.Frame Frame 
         Caption         =   "Report on:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Index           =   19
         Left            =   -72090
         TabIndex        =   187
         Top             =   1620
         Width           =   3720
         Begin VB.OptionButton optPRSelectedProducts 
            Caption         =   "S&elected Products"
            Height          =   270
            Index           =   0
            Left            =   150
            TabIndex        =   65
            Top             =   225
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton optPRSelectedProducts 
            Caption         =   "All &Products"
            Height          =   270
            Index           =   1
            Left            =   2160
            TabIndex        =   70
            Top             =   225
            Width           =   1245
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Step 1:  Create a new Message (optional) "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2820
         Index           =   30
         Left            =   -74760
         TabIndex        =   185
         Top             =   570
         Width           =   6645
         Begin VB.TextBox txtNewMessage 
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1785
            Left            =   180
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   930
            Width           =   5295
         End
         Begin VB.CommandButton btnSaveMessage 
            Caption         =   "Save"
            Height          =   345
            Left            =   5730
            TabIndex        =   72
            Top             =   2235
            Width           =   735
         End
         Begin VB.TextBox txtMessageTitle 
            Height          =   375
            Left            =   1680
            TabIndex        =   14
            Top             =   360
            Width           =   3015
         End
         Begin VB.Label lbl 
            BackStyle       =   0  'Transparent
            Caption         =   "Message Title --->"
            Height          =   375
            Index           =   35
            Left            =   360
            TabIndex        =   186
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdPromotion 
         Caption         =   "Delete ALL"
         Height          =   495
         Index           =   2
         Left            =   -64095
         TabIndex        =   84
         Top             =   2745
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.CommandButton cmdSettingsPassword 
         Caption         =   "Password"
         Height          =   315
         Index           =   0
         Left            =   -65625
         TabIndex        =   155
         Top             =   5520
         Width           =   1035
      End
      Begin VB.TextBox txtConfirmPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   -67035
         PasswordChar    =   "*"
         TabIndex        =   154
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton cmdPromotion 
         Caption         =   "Print List"
         Height          =   495
         Index           =   6
         Left            =   -64095
         TabIndex        =   59
         Top             =   2145
         Width           =   630
      End
      Begin VB.Frame Frame 
         Caption         =   "Retail/NCS split"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1050
         Index           =   15
         Left            =   -68985
         TabIndex        =   184
         Top             =   720
         Width           =   1590
         Begin VB.OptionButton optWRS 
            Caption         =   "All sales"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   255
            Value           =   -1  'True
            Width           =   885
         End
         Begin VB.OptionButton optWRS 
            Caption         =   "Retail Sales"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   495
            Width           =   1155
         End
         Begin VB.OptionButton optWRS 
            Caption         =   "NCS Sales"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   735
            Width           =   1065
         End
      End
      Begin VB.CommandButton cmdPromotion 
         Caption         =   "Show ALL"
         Height          =   495
         Index           =   8
         Left            =   -64125
         TabIndex        =   42
         ToolTipText     =   "Show all promotions including Expired promotions"
         Top             =   1545
         Width           =   660
      End
      Begin VB.Frame Frame 
         Caption         =   "Remote Settings"
         Height          =   540
         Index           =   34
         Left            =   -67680
         TabIndex        =   183
         Top             =   5850
         Width           =   4065
         Begin VB.CheckBox chkResetRemoteOpenedBy 
            Caption         =   "Reset OpenedBy"
            Height          =   255
            Left            =   300
            TabIndex        =   157
            Top             =   210
            Width           =   1575
         End
      End
      Begin VB.Frame Frame 
         Caption         =   "Step 3: Select Recipients"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5145
         Index           =   32
         Left            =   -67680
         TabIndex        =   180
         Top             =   570
         Width           =   4065
         Begin VB.OptionButton optUploadSelection 
            Caption         =   "Selected States"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   36
            Top             =   855
            Width           =   1440
         End
         Begin VB.OptionButton optUploadSelection 
            Caption         =   "Selected franchises"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   75
            Top             =   1755
            Width           =   1710
         End
         Begin VB.OptionButton optUploadSelection 
            Caption         =   "PM franchises"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   35
            Top             =   577
            Width           =   1680
         End
         Begin VB.OptionButton optUploadSelection 
            Caption         =   "ALL franchises"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   300
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.Frame Frame 
            Height          =   840
            Index           =   37
            Left            =   165
            TabIndex        =   182
            Top             =   840
            Width           =   3750
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "VIC"
               Enabled         =   0   'False
               Height          =   195
               Index           =   0
               Left            =   375
               TabIndex        =   49
               Top             =   270
               Width           =   585
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "NSW"
               Enabled         =   0   'False
               Height          =   195
               Index           =   1
               Left            =   1140
               TabIndex        =   51
               Top             =   270
               Width           =   690
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "QLD"
               Enabled         =   0   'False
               Height          =   195
               Index           =   3
               Left            =   2895
               TabIndex        =   54
               Top             =   270
               Width           =   645
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "NT"
               Enabled         =   0   'False
               Height          =   195
               Index           =   6
               Left            =   2010
               TabIndex        =   57
               Top             =   555
               Width           =   570
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "ACT"
               Enabled         =   0   'False
               Height          =   195
               Index           =   2
               Left            =   2010
               TabIndex        =   53
               Top             =   270
               Width           =   720
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "SA"
               Enabled         =   0   'False
               Height          =   195
               Index           =   5
               Left            =   1140
               TabIndex        =   56
               Top             =   555
               Width           =   675
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "TAS"
               Enabled         =   0   'False
               Height          =   195
               Index           =   4
               Left            =   375
               TabIndex        =   55
               Top             =   555
               Width           =   660
            End
            Begin VB.CheckBox chkUpload_State 
               Caption         =   "WA"
               Enabled         =   0   'False
               Height          =   195
               Index           =   7
               Left            =   2895
               TabIndex        =   58
               Top             =   555
               Width           =   645
            End
         End
         Begin VB.Frame Frame 
            Height          =   3285
            Index           =   38
            Left            =   165
            TabIndex        =   181
            Top             =   1755
            Width           =   3750
            Begin VB.ListBox lstUploadFranchiseList 
               Enabled         =   0   'False
               Height          =   2985
               ItemData        =   "TSG Data Warehouse.frx":068C
               Left            =   375
               List            =   "TSG Data Warehouse.frx":068E
               MultiSelect     =   2  'Extended
               Sorted          =   -1  'True
               TabIndex        =   77
               Top             =   210
               Width           =   3150
            End
         End
      End
      Begin VB.Frame fraBataTab 
         Caption         =   "BATA Reporting Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7710
         Left            =   -74895
         TabIndex        =   177
         Top             =   555
         Width           =   11385
         Begin VB.CommandButton cmdBataTabSaveSelected 
            Caption         =   "Save Selected"
            Height          =   345
            Left            =   9510
            TabIndex        =   344
            Top             =   2370
            Width           =   1545
         End
         Begin VB.ComboBox cboBataTabTxOrProcessedDate 
            Height          =   315
            ItemData        =   "TSG Data Warehouse.frx":0690
            Left            =   9375
            List            =   "TSG Data Warehouse.frx":069A
            Style           =   2  'Dropdown List
            TabIndex        =   343
            Top             =   150
            Width           =   1890
         End
         Begin VB.CommandButton cmdBataTabProcessSelected 
            Caption         =   "Process  [Select ...]"
            Height          =   345
            Left            =   9540
            TabIndex        =   342
            Top             =   2805
            Width           =   1545
         End
         Begin VB.CommandButton cmdBataTabProcessUnProcessed 
            Caption         =   "Process Un-processed"
            Height          =   480
            Left            =   9525
            TabIndex        =   341
            Top             =   3300
            Width           =   1545
         End
         Begin VB.CommandButton cmdBataTabUploadSelected 
            BackColor       =   &H0000FF00&
            Caption         =   "Upload  [Select ...]"
            Height          =   345
            Left            =   9525
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   4440
            Width           =   1545
         End
         Begin VB.CommandButton cmdBataTabPrintGrid 
            Caption         =   "Print"
            Enabled         =   0   'False
            Height          =   345
            Left            =   9525
            TabIndex        =   125
            ToolTipText     =   "Print Grid"
            Top             =   3990
            Width           =   660
         End
         Begin VB.CommandButton cmdBataTabExportGrid 
            Caption         =   "Export"
            Enabled         =   0   'False
            Height          =   345
            Left            =   10410
            TabIndex        =   128
            ToolTipText     =   "Export Grid"
            Top             =   3990
            Width           =   660
         End
         Begin VB.Frame fraBataTabProcessedStatus 
            Height          =   930
            Left            =   9360
            TabIndex        =   178
            Top             =   930
            Width           =   1905
            Begin VB.OptionButton optBataProcessed 
               Caption         =   "ALL Bata Franchises"
               Height          =   195
               HelpContextID   =   2
               Index           =   2
               Left            =   90
               TabIndex        =   52
               Top             =   660
               Width           =   1785
            End
            Begin VB.OptionButton optBataProcessed 
               Caption         =   "Processed"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   40
               Top             =   180
               Value           =   -1  'True
               Width           =   1065
            End
            Begin VB.OptionButton optBataProcessed 
               Caption         =   "NOT Processed"
               Height          =   195
               HelpContextID   =   1
               Index           =   1
               Left            =   90
               TabIndex        =   41
               Top             =   420
               Width           =   1455
            End
         End
         Begin VB.CommandButton cmdBataTabViewSelected 
            Caption         =   "View Selected"
            Height          =   345
            Left            =   9525
            TabIndex        =   82
            Top             =   1935
            Width           =   1545
         End
         Begin VB.CommandButton cmdBataTabUploadUnSent 
            BackColor       =   &H0000FF00&
            Caption         =   "Upload Unsent"
            Height          =   345
            Left            =   9525
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   4860
            Width           =   1545
         End
         Begin VB.CommandButton cmdImportBatscanFiles 
            Caption         =   "Import BATScan Files"
            Height          =   555
            Left            =   9525
            TabIndex        =   163
            Top             =   6000
            Width           =   1545
         End
         Begin VSFlex8Ctl.VSFlexGrid grdBataRpts 
            Height          =   7425
            Left            =   180
            TabIndex        =   10
            Top             =   195
            Width           =   9090
            _cx             =   16034
            _cy             =   13097
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   3
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"TSG Data Warehouse.frx":06C0
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   4
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   5
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            DataMode        =   0
            VirtualData     =   -1  'True
            DataMember      =   ""
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin MSComCtl2.DTPicker dtpBataTabTxDate 
            Height          =   285
            Left            =   9360
            TabIndex        =   39
            Top             =   495
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "ddd, dd MMM yyyy"
            Format          =   164954115
            CurrentDate     =   38512
         End
         Begin VB.Label lblBataTabFranCount 
            BackStyle       =   0  'Transparent
            Caption         =   "BATA Franchise Count"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   9330
            TabIndex        =   179
            Top             =   6900
            Width           =   1950
         End
      End
      Begin ComctlLib.ListView lvwPromo 
         Height          =   2865
         Left            =   -74865
         TabIndex        =   1
         Top             =   555
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5054
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Promotion Name"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Product"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "From"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "To"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Pkt Rebate"
            Object.Width           =   1252
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Ctn Rebate"
            Object.Width           =   1252
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "State"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Region"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Grade"
            Object.Width           =   617
         EndProperty
         BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   9
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "ID"
            Object.Width           =   706
         EndProperty
         BeginProperty ColumnHeader(11) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   10
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
      End
      Begin ComctlLib.ListView lvwProductReport 
         Height          =   3015
         Left            =   -74880
         TabIndex        =   150
         ToolTipText     =   "Sales report. Also shows 'missing' days sales."
         Top             =   5295
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   5318
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   10
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description, Barcode or Supplier"
            Object.Width           =   4674
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   441
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Normal Sell"
            Object.Width           =   1499
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Actual Sell"
            Object.Width           =   1517
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Total"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "NCS Qty"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   6
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "NCS Total"
            Object.Width           =   1305
         EndProperty
         BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   7
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "NCS Sell"
            Object.Width           =   1094
         EndProperty
         BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   8
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "$NCS %"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   9
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "NCS %"
            Object.Width           =   1058
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtxNielsenReportContents 
         Height          =   3000
         Left            =   -70530
         TabIndex        =   23
         Top             =   885
         Width           =   6645
         _ExtentX        =   11721
         _ExtentY        =   5292
         _Version        =   393217
         BorderStyle     =   0
         ScrollBars      =   3
         RightMargin     =   50000
         TextRTF         =   $"TSG Data Warehouse.frx":087B
      End
      Begin MSComDlg.CommonDialog cdlTSGDataWarehouse 
         Left            =   10485
         Top             =   3780
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
         FontBold        =   -1  'True
         FontName        =   "courier new"
         FontSize        =   12
      End
      Begin ComctlLib.ListView lvwStickReport 
         Height          =   3630
         Left            =   -74805
         TabIndex        =   130
         Top             =   4665
         Width           =   10980
         _ExtentX        =   19368
         _ExtentY        =   6403
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Product"
            Object.Width           =   4939
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Supplier"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Market %"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Sticks/1000"
            Object.Width           =   1367
         EndProperty
      End
      Begin ComctlLib.ListView lvwPRProductReport 
         Height          =   3195
         Left            =   -74640
         TabIndex        =   152
         Top             =   5070
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5636
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Description,Barcode or Supplier"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Qty"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Actual Sell Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Normal Sell Price"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   4
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Total"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   5
            Key             =   ""
            Object.Tag             =   ""
            Text            =   ""
            Object.Width           =   2540
         EndProperty
      End
      Begin ComctlLib.ListView lvwEventLog 
         Height          =   4230
         Left            =   120
         TabIndex        =   129
         Top             =   3735
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   7461
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Date/Time"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Franchise"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Event"
            Object.Width           =   13500
         EndProperty
      End
      Begin TSGDataWarehouse.TDatePicker tdpEventLogDate 
         Height          =   315
         Left            =   9705
         TabIndex        =   286
         Top             =   8010
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         EditMode        =   0
         Enabled         =   -1  'True
         MinDate         =   -657434
         MaxDate         =   2958465
         ToolTipUC       =   ""
         Value           =   42185
      End
      Begin VB.ListBox lstDataCaptureFranchiseBusinessName 
         Height          =   2790
         Left            =   120
         MultiSelect     =   2  'Extended
         Sorted          =   -1  'True
         TabIndex        =   339
         Top             =   750
         Width           =   2100
      End
      Begin VB.Label lblEventLog 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Event Log"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5190
         TabIndex        =   340
         Top             =   3525
         Width           =   1275
      End
      Begin VB.Label lblDisplayedFranchise 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Selected Franchise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1155
         TabIndex        =   262
         Top             =   525
         Width           =   2610
      End
      Begin VB.Label lblVersion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Version"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   9825
         TabIndex        =   253
         Top             =   525
         Width           =   1680
      End
      Begin VB.Label lbl 
         Caption         =   "Franchise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   259
         Top             =   525
         Width           =   855
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Franchise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   33
         Left            =   -74640
         TabIndex        =   261
         Top             =   630
         Width           =   975
      End
      Begin VB.Label lblProducts 
         BackStyle       =   0  'Transparent
         Caption         =   "Product List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -66000
         TabIndex        =   260
         Top             =   630
         Width           =   1395
      End
      Begin VB.Label lbl 
         Caption         =   "Franchise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   -74805
         TabIndex        =   258
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Franchise"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   25
         Left            =   -74880
         TabIndex        =   257
         Top             =   1050
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Recipient"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   26
         Left            =   -68610
         TabIndex        =   256
         Top             =   3525
         Width           =   1215
      End
      Begin VB.Label lbl 
         Caption         =   "Some of these settings are password protected and should only be modified by the System Administrator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   39
         Left            =   -74640
         TabIndex        =   255
         Top             =   1050
         Width           =   10815
      End
      Begin VB.Label lblUploadsPending 
         BackStyle       =   0  'Transparent
         Caption         =   "Caption Set in Code [lblUploadPendingCount]"
         Height          =   195
         Left            =   -74760
         TabIndex        =   254
         Top             =   5655
         Width           =   4095
      End
      Begin VB.Label lblLocalDriveFranFolder 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Using local Franchise folder (For ALL Franchises)"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   252
         Top             =   3540
         Width           =   4200
      End
   End
End
Attribute VB_Name = "frmTSGDataWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'$ PzROBHIDE ALL BEGIN
' Above relates to processing by Project.exe
'!!! ManualFix Clearing: Problem auto-fix by Project Analyzer 7.1.07 on 16/12/2005
Option Explicit
Private Enum TabEnum
    eDataCaptureTab = 0
    eSalesRptsTab = 1
    eStickRptsTab = 2
    eStockTab = 3
    eBataTab = 4
    eNielsenTab = 5
    eVersionsTab = 6
    eProductRptsTab = 7
    eSettingsTab = 8
    eUploadsTab = 9
    ePromotionsTab = 10
End Enum

Private Enum UploadSelModeEnum
    eUpldAllFrans = 0
'   eUpldVpnFrans = 1       ' All franchises are now on the VPN (cf dial-up)
    eUpldPmAndRmFrans = 2
    eUpldSelStates = 3
    eUpldSelFrans = 4
End Enum
    
Public Enum SelFranEnum
    eSelFran_CaptureCycleAuto = 1        ' Sorted by 1. Dial Seq, 2. Name - Franchises pending data collection or pending uploads
    eSelFran_CaptureCycleManual = 2      ' Sorted by 1. Dial Seq, 2. Name - Franchises pending data collection
    eSelFran_CaptureCycleExcluded = 3    ' Sorted By Name                 - Franchises excluded from Capture Cycle
End Enum

Private Type udt
    astrFranTypeTooltip() As String
    ablnTabRefreshed(TabEnum.eDataCaptureTab To TabEnum.ePromotionsTab) As Boolean
End Type
Private m As udt
Private WithEvents moBataRpts As clsBataRpts
Attribute moBataRpts.VB_VarHelpID = -1

' ----- These 3 must match the folders on the remote machine ---'
Private Const mkMessageFolderName As String = "NewMessages"     '
Private Const mkNewStockFolderName As String = "NewStock"       '
Private Const mkWLPUpgradesFolderName As String = "WLPUpdates"  '
' --------------------------------------------------------------'
Private Const mkPromoRegionsAll As Long = 0
Private Const mkPromoStatesAll As String = "ALL"
Private Const mkPromoRegionsNA As Long = -1
Private Const mkPromoStatesNA As String = vbNullString
Private Const mkPromoGradeIdNA As Long = -1

Private Sub AcceptSettings()
''' MySQL Review
''' PROCEDURE HAS BUG NOT BEING ADDRESSED IN TRANSLATION FROM Access Mdb to MySQL
''' PROCEDURE DOESN'T ACCOMMODATE UNIQUE INDEXING OF FranchiseDialSequence IN tlkpState
''' TRANSLATION WOULD TAKE FOREVER IF I FIXED EVERY BUG I ENCOUNTERED
Const kProcName As String = "AcceptSettings"
    Dim iIndex As Integer
Dim strSQL As String
Dim strErrMsg As String
Dim rsStateSpecific As ADODB.Recordset
Dim rsFranch As ADODB.Recordset
    
'-  Begin data transaction *************************************'*
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx '*
    On Error GoTo Procedure_Error_Rollback                      '*
'   **************************************************************
    
    For iIndex = 0 To 7
        strSQL = "SELECT * FROM tlkpState WHERE StateOfOz = " & SqlQ(lblSateName(iIndex))
        Set rsStateSpecific = GetRst(pCnn:=g.cnnDW, _
                                     pSource:=strSQL, _
                                     pSourceType:=adCmdText, _
                                     pRstType:=eEditableFwdOnly, _
                                     pErrMsg:=strErrMsg)
        If Not (rsStateSpecific.BOF And rsStateSpecific.EOF) Then
                rsStateSpecific!DialSequence = Val(txtDialSequence(iIndex))
            rsStateSpecific.Update
        End If
        
        ' Now set the dial sequence of all franchises in this stateofOz
        strSQL = "SELECT * FROM Franchises " & vbNewLine & _
                 "WHERE FranchiseStateOfOz = " & SqlQ(lblSateName(iIndex))
        Set rsFranch = GetRst(pCnn:=g.cnnDW, _
                              pSource:=strSQL, _
                              pSourceType:=adCmdText, _
                              pRstType:=eEditableFwdOnly, _
                              pErrMsg:=strErrMsg)
        Do Until rsFranch.EOF
                rsFranch!FranchiseDialSequence = Val(txtDialSequence(iIndex))
            rsFranch.Update     ' may be superfluous
            rsFranch.MoveNext
        Loop
    Next iIndex
    
    rsStateSpecific.Close
    Set rsStateSpecific = Nothing
    rsFranch.Close
    Set rsFranch = Nothing

    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx
    
    MsgBox "New settings accepted", vbInformation
    
Procedure_Exit:
    Exit Sub

Procedure_Error_Rollback:
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
    MsgBox "Changes not saved." & vbNewLine & Err.Description
    Resume Procedure_Exit

End Sub

Private Function AddNewPromo(ByVal pPromoName As String, _
                          ByVal pSubCat As String, _
                          ByVal pPromoStart As Date, _
                          ByVal pPromoEnd As Date, _
                          ByVal pCtnDiscount As Currency, _
                          ByVal pPktDiscount As Currency, _
                          ByVal pRegionID As Long, _
                          ByVal pState As String, _
                          ByVal pPromoGradeID As Long, _
                          ByRef pErrMsg As String) As Boolean

Const kProcName As String = "AddNewPromo"
Dim lngPromoID As Long
Dim dtmPromoUpdate As Date
Dim strWC As String
Dim strSQL As String
Dim strErrMsg As String
Dim strFranSelnWC As String
Dim colFranIDs As VBA.Collection
Dim rstFranIDs As ADODB.Recordset

'-> Get list of FranIDs promo applies to, then call AddPromoToFUandFP() with FranID list
'   Each Promo Name may have any number of promo records each with its own PromoID, State/Region, product combination
'   Stores are NOT excluded according to whether they are excluded from capture cycle (i.e. FranchiseIncludedInStatistics = False)
'   Stores may be temporarily excluded and later included in which case they will receive the uploads created for them
    If ChkBoxToBool(chkPromoSelectFranchise) Then
        Set colFranIDs = ListBoxGetCollection(pListBox:=lstPromoFranchise, pItemData:=True, pSelected:=True)
        If colFranIDs.Count = 0 Then
            strFranSelnWC = "(-1)"  ' Prevent selecting any franchises should we get to this code (selection validation fail)
        Else
            strFranSelnWC = GetWcValueListFromColn(colFranIDs)
        End If
    Else
        strWC = vbNullString
        If pState <> mkPromoStatesAll Then
            If Len(strWC) Then strWC = strWC & " AND "
            strWC = Bracket("FranchiseStateOfOz = " & SqlQ(pState))
        End If
        If pRegionID <> mkPromoRegionsAll Then
            If Len(strWC) Then strWC = strWC & " AND "
            strWC = strWC & Bracket("FranchiseRegionId = " & pRegionID)
        End If
    '   When adding a PROMO ensure franchise has approprite PromoGrade
    '   and is a FranType that promos are uploaded to
    '   [i.e. FranType < 60 see fOKToUploadItem() OR FranType = gkOPosFranType]
        If Len(strWC) Then strWC = strWC & " AND "
        strWC = strWC & Bracket("PromoGradeID = " & pPromoGradeID) & _
                " AND " & Bracket("(FranchiseType < 60) OR (FranchiseType = " & gkOPosFranType & ")")
''' Review:  Above WC could do with some comments about the addition of FranType seln
'''          and most significantly what these frantype values represent. Selections here
'''          are pretty much stolen from fOKToUploadItem() and perhaps everything could be simplified

    '   New promos are only added to Live frans. During the life of the program the number
    '   of dead frans is constantly increasing. New promotions are still added to frans suspended
    '   from the capture cycle as they may be un-suspended during the life of the promo
        strSQL = "SELECT FranchiseIDTSG " & vbNewLine & _
                 "FROM qryfranchiselive " & vbNewLine & _
                 "WHERE " & strWC
        
        Set rstFranIDs = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pRstType:=eReadOnlyFwdOnly, pErrMsg:=strErrMsg)
        Set colFranIDs = GetCollectionFromRst(pRst:=rstFranIDs, pFldName:="FranchiseIDTSG", pForceMoveFirst:=False)
        strFranSelnWC = GetWcValueListFromColn(colFranIDs)
        
        Set colFranIDs = Nothing
        Set rstFranIDs = Nothing
    End If
    
    If Len(strFranSelnWC) = 0 Then
        strErrMsg = "DOES NOT apply to any live franchises."
    Else
    '   SQL for adding promotion to Promotions table
    '?  Possibly write SQL procedure to add record and return auto-increment field value.
    '?  (might be nice to create my own helper procedures for creating common SQL statements.
    '?  (i guess it's what rsts do, but if I create my own stmts I can bundle a number together into same call)
        strSQL = "INSERT INTO Promotions " & vbNewLine & _
                 " (PromoName, PromoSubCat, PromoStart, " & _
                 "  PromoEnd, PromoCartonDiscount, PromoPacketDiscount, " & _
                 "  PromoState, PromoRegionID, PromoGradeID, PromoStatus) " & vbNewLine & _
                 "VALUES(" & _
                    SqlQ(pPromoName) & ", " & SqlQ(pSubCat) & ", " & MySqlDate(pPromoStart) & ", " & _
                    MySqlDate(pPromoEnd) & ", " & pCtnDiscount & ", " & pPktDiscount & ", " & _
                    SqlQ(pState) & ", " & pRegionID & ", " & pPromoGradeID & ", " & SqlQ(PROMO_NOT_SENT) & ")"
    
    '-  Begin data transaction *************************************'*
        Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx '*
        On Error GoTo Procedure_Error_Rollback                      '*
    '   **************************************************************
    
    '*  Add promotion to Promotions table
        CnnDwExecute pCommandText:=strSQL
        dtmPromoUpdate = Now
        
        If g.cnnDW.Errors.Count Then
            strErrMsg = "Failed to add promotion to Promotions table. " & g.cnnDW.Errors(0).Description
            Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
        Else
            strSQL = "Select Max(PromoID) FROM Promotions"
            lngPromoID = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL)
SetTableUpdateTime pTableName:="Promotions", pTimeStamp:=dtmPromoUpdate ''' Review
        
        '   AddPromoToFUandFP adds records to FranchiseUploads to add (action:='PROMO') or recall (action:='DELPROMO')
        '   Pending uploads are automatically purged as part of the capture cycle according to TsgDwMdb.Defaults!MonthsOfFranchiseUploads ''' Review
            AddPromoToFUandFP pAction:=gkPromoADD, _
                              pPromoID:=lngPromoID, _
                              pFranIdWcValueList:=strFranSelnWC, _
                              pErrMsg:=strErrMsg
            If Len(strErrMsg) Then
                Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
            Else
                Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx
            End If
        End If
    End If

Procedure_Exit:
    If Len(strErrMsg) Then pErrMsg = kProcName & "() -> " & strErrMsg ' prepend calling proc/stack
    AddNewPromo = (Len(strErrMsg) = 0)
    Exit Function

Procedure_Error_Rollback:
    If Err.Number Then strErrMsg = Trim$(strErrMsg & " " & Err.Source & " " & Err.Number & ": " & Err.Description)
    If g.cnnDW.Errors.Count Then strErrMsg = Trim$(strErrMsg & " " & g.cnnDW.Errors(0).Description)
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
    Resume Procedure_Exit
    
End Function

Private Sub AddPromoToFUandFP(ByVal pAction As String, _
                              ByVal pPromoID As Long, _
                              ByVal pFranIdWcValueList As String, _
                              ByRef pErrMsg As String)
'   pAction: gkPromoADD (adding a new promo) or gkPromoDELETE (recalling a promo)
'   This procedure is called in 2 places (V385):
'       1. cmdPromotion_Click -> SaveNewPromotion -> AddNewPromo -> AddPromoToFUandFP
'       2. cmdPromotionRecall_Click -> PromotionRecall -> AddPromoToFUandFP
Const kProcName As String = "AddPromoToFUandFP"
Dim bFranchiseUploads As Boolean
Dim strSQL As String
Dim strSQL_FranRst As String
Dim strFU_SQL As String
Dim strFU_Vals As String
Dim strFP_Vals As String
Dim strPromoUpload As String
Dim strErrMsg As String
Dim rstFran As ADODB.Recordset

'   Calling procedure determines fran selection and passes it via pFranIdWcValueList
'   Following SQL selects frans from Franchises table rather than qryFranchiseLive to allow
'   calling procedure to select any frans.[RecallPromotions should recall from frans regardless of status]
    If Len(pFranIdWcValueList) Then
''' V400 Start
'''     strSQL_FranRst = "SELECT FranchiseIDTSG, FranchiseType, FranchiseMessageFlag " & vbNewLine & _
'''                      "FROM qryFranchiseLive " & vbNewLine & _
'''                      "WHERE FranchiseIDTSG IN " & pFranIdWcValueList
        strSQL_FranRst = "SELECT FranchiseIDTSG, FranchiseType, FranchiseMessageFlag " & vbNewLine & _
                         "FROM Franchises " & vbNewLine & _
                         "WHERE FranchiseIDTSG IN " & pFranIdWcValueList
''' V400 End
        Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL_FranRst, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        If rstFran Is Nothing Then
            strErrMsg = kProcName & "() -> " & strErrMsg ' prepend calling proc/stack
        Else
            If Not (rstFran.BOF And rstFran.EOF) Then
'*** SHOULD ADD SOME TRIGGERS TO FORCE REFERENTIAL INTEGRITY BETWEEN PROMOUPLOAD TABLE AND Promotions Table *** (AMONG OTHER THINGS)
                strPromoUpload = pAction & pPromoID
                Do Until rstFran.EOF
                '   fOKToUploadItem() determines if TsgDw.exe should upload the item/fran combo
                '   If it should then it adds a row to FranchiseUploads table among other things
                '   TsgDw.exe doesn't upload to oPOS frans but uses tblFranhciseUploads and Promotions
                '   tables to communicate what Promos should be downloaded by oPOS and related S/W
                    bFranchiseUploads = fOKToUploadItem(pFranID:=rstFran!FranchiseIDTSG, _
                                                        pFranType:=Cn(rstFran!FranchiseType, gkOPosFranType), _
                                                        pFranMsgFlag:=CBool(rstFran!FranchiseMessageFlag), _
                                                        pUploadFile:=strPromoUpload)
                    If bFranchiseUploads Then
                    '   FranchiseUploads is populated for both Adding and Deleting PROMOs
                        strFU_Vals = strFU_Vals & "(" & rstFran!FranchiseIDTSG & ", " & SqlQ(strPromoUpload) & "), " & vbNewLine
                    End If
                    
                '   TsgDw does NOT connect/upload to oPos Frans therefore FU records NOT added
                '   BUT Promo & FranPromo records ARE ADDED for oPOS related s/w to manage promos
                    If bFranchiseUploads Or (rstFran!FranchiseType = gkOPosFranType) Then
                    '   We only only need to create strFP_Vals string when adding.
                        If pAction = gkPromoADD Then
                            strFP_Vals = strFP_Vals & "(" & rstFran!FranchiseIDTSG & ", " & pPromoID & ", " & FpTfrEnum.FpTfrRequested & "), " & vbNewLine
                        End If
                    End If
                    rstFran.MoveNext
                Loop
            End If ' If Not (rstFran.BOF And rstFran.EOF) Then
            rstFran.Close
            Set rstFran = Nothing
            
        '   FranchiseUpload table TREATED THE SAME FOR AddNewPromo & DeletePromo
            If Len(strFU_Vals) Then
                strFU_Vals = Left$(strFU_Vals, InStrRev(strFU_Vals, ",") - 1)
                strFU_SQL = "INSERT INTO FranchiseUploads (FranchiseID, UploadFile) " & vbNewLine & _
                            "VALUES " & strFU_Vals
            End If
    
         '  FranchiseUpload record added when promo created/added and status field maintained from then on
            Select Case pAction
                Case gkPromoDELETE
                '   Records edited when promo is deleted/recalled (Should I test that records I'm expecting to edit exist?)
                    strSQL = "UPDATE tblFranchisePromotions " & vbNewLine & _
                                 " SET TfrStatus = " & FpTfrEnum.FpRecallRequested & " " & vbNewLine & _
                                "WHERE (PromotionID = " & pPromoID & ") " & _
                                 " AND (FranchiseID IN " & pFranIdWcValueList & ")" & _
                                 " AND (TfrStatus = " & FpTfrEnum.FpTfrCompleted & ") "
                    '   Cater for cases where FU records are required and where FU records aren't required
                    '   i.e. At least some non oPOS frans and only oPOS frans
                    '   i.e. Records to add for tblFranchisePromotions but not for FranchiseUploads,
                    '           and records to add for both tables
                    If Len(strFU_SQL) Then
                        strSQL = strSQL & ";" & vbNewLine & strFU_SQL
                    End If
                
                Case gkPromoADD
                '   Cater to all cases. Applies to 1.no frans, 2. only oPOS 3. only non oPOS 4. combo of oPOS and non oPOS
                    If Len(strFP_Vals) Then
                    '   Records added when promo is created/added.
                        strFP_Vals = Left$(strFP_Vals, InStrRev(strFP_Vals, ",") - 1)
                        strSQL = "INSERT INTO tblFranchisePromotions (FranchiseID, PromotionID, TfrStatus)" & vbNewLine & _
                                 "VALUES " & strFP_Vals
                    '   Cater for cases where FU records are required and where FU records aren't required
                    '   i.e. At least some non oPOS frans and only oPOS frans
                    '   i.e. Records to add for tblFranchisePromotions but not for FranchiseUploads,
                    '           and records to add for both tables
                        If Len(strFU_SQL) Then
                            strSQL = strSQL & ";" & vbNewLine & strFU_SQL
                        End If
                    End If
                
                Case Else
                    Err.Raise Number:=-1, _
                              Source:="AddPromoToFUandFP", _
                              Description:="Invalid value for pAction parameter: " & DQ(CStr(pAction))
            End Select
        
            If Len(strSQL) = 0 Then
''' Review: Alternative would be to add an event log entry that promo combination does not currently appply
'''         rather than raise an error that percolates up the call tree. This approach would allow calling
'''         AddPromoToFUandFP() to add a Promotions record should any oPOS frans be created later
                strErrMsg = "DOES NOT apply to any franchises."
            Else
            '-  Begin data transaction *************************************'*
                Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx '*
                On Error GoTo Procedure_Error_Rollback                      '*
            '   **************************************************************
                CnnDwExecute pCommandText:=strSQL
                If g.cnnDW.Errors.Count Then
                    strErrMsg = g.cnnDW.Errors(0).Description
                    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
                Else
                    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx
                End If
            End If
        End If
    End If
    
Procedure_Exit:
    If Len(strErrMsg) Then pErrMsg = kProcName & "() -> " & strErrMsg ' prepend calling proc/stack
'   AddPromoToFUandFP = (Len(strErrMsg) = 0)    ' Can resinstate line if we use this proc as a function
    Exit Sub
    
Procedure_Error_Rollback:
    If Err.Number Then strErrMsg = Trim$(strErrMsg & " " & Err.Source & " " & Err.Number & ": " & Err.Description)
    If g.cnnDW.Errors.Count Then strErrMsg = Trim$(strErrMsg & " " & g.cnnDW.Errors(0).Description)
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
    Resume Procedure_Exit
    Resume  ' Not executed but assists when debugging in IDE
    
End Sub

Private Sub AddSelRptsFromGrid(ByRef pBataRpts As clsBataRpts)
Dim lngLoop As Long

    With grdBataRpts
        For lngLoop = 0 To .SelectedRows - 1
            If Len(.TextMatrix(Row:=.SelectedRow(lngLoop), col:=.ColIndex("RptType"))) Then
                pBataRpts.Add pFranName:=.TextMatrix(Row:=.SelectedRow(lngLoop), col:=.ColIndex("FranName")), _
                              pFranID:=.ValueMatrix(Row:=.SelectedRow(lngLoop), col:=.ColIndex("FranID")), _
                              pBataFranID:=.ValueMatrix(Row:=.SelectedRow(lngLoop), col:=.ColIndex("BataFranID")), _
                              pTxDate:=DateValue(.TextMatrix(Row:=.SelectedRow(lngLoop), col:=.ColIndex("TxDate"))), _
                              pBataRptType:=.ValueMatrix(Row:=.SelectedRow(lngLoop), col:=.ColIndex("RptType"))
            End If
        Next lngLoop
    End With

End Sub

'**********************************************************************************
'This prints a report to the Temp folder for perusal. Hard coded for modularity-can be
'placed in globals when installed in TSGHO1
Sub addToAnotherReport(ByRef rstStock As ADODB.Recordset)
Dim intFileNum As Long
Dim sFile As String
Dim fso As Scripting.FileSystemObject

On Error GoTo ErrorHandler

    sFile = g.strTsTemp & "\InHQButNotInAndy.txt"
    
    Set fso = New Scripting.FileSystemObject
    If Not fso.FileExists(sFile) Then
    '   Create an empty text file
        SaveTextFile sFile, vbNullString
    End If
    Set fso = Nothing
    
    intFileNum = FreeFile   ' Get unused file
    Open sFile For Append As #intFileNum
        Print #intFileNum, Chr(34) & rstStock(gconStockTableBarcodeField) & Chr(34) & ",", _
                    Chr(34) & rstStock(gconStockTableDescriptionField) & Chr(34) & ",", _
                    Chr(34) & rstStock(gconStockTableCategoryField) & Chr(34) & ",", _
                    Chr(34) & rstStock(gconStockTableSubCategoryField) & Chr(34) & ",", _
                    rstStock(gconStockTableSupplierIDField) & ",", _
                    rstStock(gconStockTableCostField) & ",", _
                    rstStock(gconStockTableSellField) & vbCrLf
                    
    Close #intFileNum
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description

End Sub

Private Sub addToDiffReport(ByRef rstStock As ADODB.Recordset, ByVal sDiffs As String)
Dim intFileNum As Long
Dim sFile As String
Dim fso As Scripting.FileSystemObject
On Error GoTo ErrorHandler

    sFile = g.strTsTemp & "\AndyDiffHQ.txt"

    Set fso = New Scripting.FileSystemObject
    If Not fso.FileExists(sFile) Then
        SaveTextFile sFile, vbNullString
    End If
    Set fso = Nothing
 
    intFileNum = FreeFile   ' Get unused file
    Open sFile For Append As #intFileNum
        Print #intFileNum, sDiffs & vbCrLf
        Print #intFileNum, Chr(34) & rstStock(gconStockTableBarcodeField) & Chr(34) & ",", _
                              Chr(34) & rstStock(gconStockTableDescriptionField) & Chr(34) & ",", _
                              Chr(34) & rstStock(gconStockTableCategoryField) & Chr(34) & ",", _
                              Chr(34) & rstStock(gconStockTableSubCategoryField) & Chr(34) & ",", _
                              rstStock(gconStockTableSupplierIDField) & ",", _
                              rstStock(gconStockTableCostField) & ",", _
                              rstStock(gconStockTableSellField) & vbCrLf
                    
    Close #intFileNum
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    Exit Sub
    Resume  ' Not executed but assists when debugging in IDE
    
End Sub

Private Sub AddToRemoteEventLog(ByVal pEvent As String, _
                                ByVal pFranName As String, _
                                ByRef pCnnRemote As ADODB.Connection)
Dim rst As ADODB.Recordset

On Error GoTo Procedure_Error
    
    Set rst = GetRst(pCnn:=pCnnRemote, _
                     pSource:="EventLog", _
                     pSourceType:=adCmdTable, _
                     pRstType:=eEditableFwdOnly)
    
    rst.AddNew
        rst!DateTime = Now
        rst!Event = Left$(Trim$(pEvent), rst!Event.DefinedSize)
    rst.Update
    rst.Close
    
    Set rst = Nothing
    
Procedure_Exit:
    Exit Sub
    
Procedure_Error:
    StatusBar "Remote EventLog Error " & Err.Number & ": " & Err.Description & " EventToLog: " & pEvent, _
              pFranName, _
              pRefreshEventLogDisplay:=False
    Resume Procedure_Exit

End Sub

Sub AddToUpdateFile(ByRef pRstStock As ADODB.Recordset, _
                    ByVal pRstDbType As DbTypeEnum, _
                    ByVal bAddNewItem As Boolean, _
                    ByVal pUseRecordValues As Boolean) ' asdf is pRstDbType flexiblity required
'   A counter is kept for the file name of the update file
'   The current file is checked to see whether it is read-only. If it is, then it has
'   already been sent to at least one franchise (at which point it is made read-only)
'   and a new file is started for the new updates

Dim strSQL As String
Dim strErrMsg As String

    Dim sFile As String
    Dim sType As String
    Dim fNew As Boolean
    Dim iNextWLPFileNum As Integer
    Dim iNextSTKFileNum As Integer
Dim bPackage As Boolean
Dim rsStkFileNums As ADODB.Recordset
Dim rsWLPFileNums As ADODB.Recordset
       
    If bAddNewItem Then
        sType = "New"
        strSQL = "SELECT * FROM STKFileNums ORDER BY FileNum DESC"
        Set rsStkFileNums = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, _
                                   pSourceType:=adCmdText, _
                                   pRstType:=eEditableDynamic, _
                                   pErrMsg:=strErrMsg)
        With rsStkFileNums
            If Not (.BOF And .EOF) Then
                .MoveLast
                iNextSTKFileNum = .Fields!FileNum.Value
            Else
                iNextSTKFileNum = 1
                .AddNew
                    .Fields!FileNum.Value = iNextSTKFileNum
                .Update
            End If
            
            sFile = g.strUploadsFolder & "\" & gconNewStockFilePrefix & Format(iNextSTKFileNum, "000") & gconTextFileSuffix
            Do While Dir(sFile) <> ""
            '   Bitwise operation
                If (GetAttr(sFile) And vbReadOnly) <> vbReadOnly Then
                '   File is writeable -> append to file
                    Exit Do
                Else
                '   File is Read-only -> increment file count
                    iNextSTKFileNum = iNextSTKFileNum + 1
                    sFile = g.strUploadsFolder & "\" & gconNewStockFilePrefix & Format(iNextSTKFileNum, "000") & gconTextFileSuffix
                    .Fields!FileNum.Value = iNextSTKFileNum
                    .Update
                End If
            Loop
            .Close
        End With
        Set rsStkFileNums = Nothing
        
    Else
        sType = "Updated"
        strSQL = "SELECT * FROM " & "WLPFileNums" & " ORDER BY " & "FileNum" & " DESC"
        Set rsWLPFileNums = GetRst(pCnn:=g.cnnDW, _
                                   pSource:=strSQL, _
                                   pSourceType:=adCmdText, _
                                   pRstType:=eEditableFwdOnly, _
                                   pErrMsg:=strErrMsg)
        If Not (rsWLPFileNums.BOF And rsWLPFileNums.EOF) Then
            iNextWLPFileNum = rsWLPFileNums!FileNum
        Else
            iNextWLPFileNum = 1
        End If
        sFile = g.strUploadsFolder & "\" & gconWLPUpgradePrefix & Format(iNextWLPFileNum, "000") & gconTextFileSuffix
        
        If Dir(sFile) <> "" Then
            If GetAttr(sFile) = vbReadOnly Then
                iNextWLPFileNum = iNextWLPFileNum + 1
                sFile = g.strUploadsFolder & "\" & gconWLPUpgradePrefix & Format(iNextWLPFileNum, "000") & gconTextFileSuffix
                rsWLPFileNums.AddNew
                    rsWLPFileNums("FileNum") = iNextWLPFileNum
                rsWLPFileNums.Update
            End If
        End If
        rsWLPFileNums.Close
        Set rsWLPFileNums = Nothing
    End If
    
    If Dir(sFile) = "" Then
        fNew = True
    End If
    
    If Not pUseRecordValues Then
        pRstStock(gconStockTableBarcodeField) = Left(txtBarcode, gconStockTableBarcodeFieldWidth)
        pRstStock(gconStockTableSupplierIDField) = flSupplierIDFrom(cboSupplier)
        pRstStock(gconStockTableSticksField) = Val(txtSticks)
        pRstStock(gconStockTableCategoryField) = cboCategory
        pRstStock(gconStockTableSubCategoryField) = cboSubCategory
        pRstStock(gconStockTableSellField) = Val(Format(txtRRP, "####0.00"))            'PAL
        pRstStock(gconStockTableCostField) = Val(Format(txtWholesaleListPrice, "####0.00"))
        pRstStock(gconStockTableSalesTaxCodeField) = cboSalesTax
        pRstStock(gconStockTableGoodsTaxCodeField) = cboGoodsTax
        bPackage = ChkBoxToBool(chkPackage)
        pRstStock(gconStockTablePackageField) = CBoolDb(bPackage, pDbType:=pRstDbType)    'PAL
        pRstStock(gconStockTableTaxComponentsField) = CBoolDb(bPackage, pDbType:=pRstDbType)
        pRstStock(gconStockTableAllowFractionsField) = CBoolDb(Not bPackage, pDbType:=pRstDbType)
    End If
    
    WriteStockToTextFile pRstStock:=pRstStock, pFullFilename:=sFile
    
    If fNew Then
        lstUploadItemList.AddItem sFile
        lstUploadItemList.Refresh
    End If
    
    If Not pUseRecordValues Then
        MsgBox sType & " stock item " & pRstStock(gconStockTableBarcodeField) & " has been added to " & sFile & vbCrLf
    End If
    
End Sub

Private Sub btnAddSubcat_Click(Index As Integer)
    Dim str As String
    
    If Index = 0 Then
        str = InputBox("Enter new SubCategory")
        If str <> "" Then
            cboSubCategory.AddItem str
        End If
    ElseIf Index = 1 Then
        str = InputBox("Enter new Category")
         If str <> "" Then
            cboCategory.AddItem str
        End If
    End If
End Sub

Private Sub btnClearUploadDir_Click()
    Dim sTemporaryFileList As String
  
    ' delete each file in the 'uploads' directory if there are no uploads pending on the file
    sTemporaryFileList = Dir(g.strUploadsFolder & "\*.*", vbDirectory)
    If Len(sTemporaryFileList) > gconZeroValue Then
        Do Until sTemporaryFileList = ""
            If sTemporaryFileList <> "." And sTemporaryFileList <> ".." Then
                If fOKToUploadItem(pFranID:=0, _
                                   pFranType:=0, _
                                   pFranMsgFlag:=True, _
                                   pUploadFile:=g.strUploadsFolder & "\" & sTemporaryFileList) Then
                    SetAttr g.strUploadsFolder & "\" & sTemporaryFileList, vbNormal
                    Kill g.strUploadsFolder & "\" & sTemporaryFileList
                Else
                    MsgBox "Cannot delete " & g.strUploadsFolder & "\" & sTemporaryFileList & vbNewLine & _
                           " because there are uploads pending for this file."
                End If
                ' dont use this because it does a 'dir' which buggers up our next dir
                ' Call deleteFile(gsUploadsFolder & sTemporaryFileList)
            End If
            sTemporaryFileList = Dir
        Loop
    End If
    Call LoadUploadTab
End Sub

Private Sub btnExportStock_Click()
Dim rstStock As ADODB.Recordset
Dim rsWLPFileNums As ADODB.Recordset
    Dim cCtns As Integer
    Dim sFile As String
    Dim iNextWLPFileNum As Integer
    Dim strConsultantDBPath As String
Dim cnnRecent As ADODB.Connection   '!!! ManualFix Clearing: Object variable not cleared: CNNRecent
Dim bContinue As Boolean
Dim strSQL As String
Dim strWC As String
Dim strSELECT As String
Dim strErrMsg As String
Dim eRstDbType As DbTypeEnum    ' asdf is pRstDbType flexiblity required

    If MsgBox("This will create a text file containing (according to your selection) the current prices of" & vbNewLine & _
               "all cigarette cartons, cigars & tobacco. This is the file that is sent to the stores every" & vbNewLine & _
               "6 months when the price upgrades occur." & vbNewLine & vbNewLine & _
               "You can choose to use the main TSG database for the export" & vbNewLine & _
               "or a separate stock table (for example, in a recent.mdb from " & _
               "another machine.)" & vbNewLine & vbNewLine & _
               "If the current prices are not yet in the TSG database," & vbNewLine & _
               "but are in a recent.mdb file (like on Andy's laptop), then " & vbNewLine & _
               "that recent.mdb file can be merged into the TSG database using" & vbNewLine & _
               "the 'Merge Databases' button on this tab before creating the text file." & vbNewLine & vbNewLine & _
               "Press OK to continue.", vbOKCancel + vbInformation) = vbOK Then
    
        If MsgBox("Do you want to export from an external recent.mdb or the main (recommended) TSG database." & vbNewLine & vbNewLine & _
                  "Click Yes to select an external recent.mdb, no to use the TSG Database", vbYesNo + vbQuestion) <> vbYes Then
            bContinue = True
        Else
            strConsultantDBPath = fdlgCommon.GetFullFileName(pMethod:=ShowOpenOrSaveEnum.eShowOpen, _
                                                             pFilename:="*.mdb", _
                                                             pFilter:="mdb", _
                                                             pFilterDescription:="Microsoft Access Database (*.mdb)", _
                                                             pDefaultExtension:="mdb")
            bContinue = Len(strConsultantDBPath) <> 0
        End If
    
        If bContinue Then
            strSQL = "SELECT * FROM WLPFileNums ORDER BY FileNum DESC"
            Set rsWLPFileNums = GetRst(pCnn:=g.cnnDW, _
                                       pSource:=strSQL, _
                                       pSourceType:=adCmdText, _
                                       pRstType:=eEditableFwdOnly, _
                                       pErrMsg:=strErrMsg)

            If Not (rsWLPFileNums.BOF And rsWLPFileNums.EOF) Then
                iNextWLPFileNum = rsWLPFileNums!FileNum
            Else
                iNextWLPFileNum = 1
            End If
            
            sFile = g.strUploadsFolder & "\" & gconWLPUpgradePrefix & Format(iNextWLPFileNum, "000") & gconTextFileSuffix
            
            ' We want to create a separate new file. If one already exists, close it off and get the
            ' next available number.
            If Dir(sFile) <> "" Then
                SetAttr sFile, vbReadOnly
                iNextWLPFileNum = iNextWLPFileNum + 1
                sFile = g.strUploadsFolder & "\" & gconWLPUpgradePrefix & Format(iNextWLPFileNum, "000") & gconTextFileSuffix
                rsWLPFileNums.AddNew
                    rsWLPFileNums!FileNum = iNextWLPFileNum
                rsWLPFileNums.Update
            End If
            rsWLPFileNums.Close
            Set rsWLPFileNums = Nothing
        ''' Review SHOULD PROBABLY EXPLICITLY CLOSE THE CONSULTANT DB IF THAT IS WHAT WE ARE OPENING
            If Len(strConsultantDBPath) Then
                Set cnnRecent = GetCnn(pDataSource:=strConsultantDBPath, _
                                       pCnnMode:=adModeReadWrite, _
                                       pDataSourceType:=DataSourceTypeEnum.eMdb, _
                                       pErrMsg:=strErrMsg)
                strSELECT = "SELECT * FROM Stock"
                eRstDbType = eJetDb
            Else
                Set cnnRecent = g.cnnDW
                strSELECT = "SELECT * FROM qryStock"
                eRstDbType = eMySqlDb
            End If
            
            strWC = gconStockTableCategoryField & " = " & SqlQ(gkCAT_CigCtn)
            If chkExportStkCategory(2).Value = 1 Then
                strWC = strWC & " OR " & gconStockTableCategoryField & " = " & SqlQ(gkCAT_Cigar)
            End If
            If Me.chkExportStkCategory(3).Value = 1 Then
                strWC = strWC & " OR " & gconStockTableCategoryField & " = " & SqlQ(gkCAT_Tobac)
            End If
        
        '   Use qryStock to exclude stock flagged as deleted
            strSQL = strSELECT & vbNewLine & _
                     "WHERE " & strWC & vbNewLine & _
                     "ORDER BY " & gconStockTableCategoryField & ", " & gconStockTableDescriptionField
                     
            Set rstStock = GetRst(pCnn:=cnnRecent, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
            Do Until rstStock.EOF
                AddToUpdateFile pRstStock:=rstStock, _
                                pRstDbType:=eRstDbType, _
                                bAddNewItem:=False, _
                                pUseRecordValues:=True
                cCtns = cCtns + 1
                StatusBar "Processing " & cCtns, pLog:=False
                rstStock.MoveNext
            Loop
            
            rstStock.Close
            Set rstStock = Nothing
            SetAttr sFile, vbReadOnly
            
            MsgBox "Text file " & sFile & " containing the latest prices (" & cCtns & " of)" & vbNewLine & _
                    "has been created. This can be uploaded to all franchises now.", vbInformation
        End If
    End If
    
End Sub

Private Sub btnMissingDaysSales_Click()
'   Button on Sales Rports tab
    Dim lArrayRowIndex As Long
    Dim iNumberOfFranchisesIncluded As Integer
    
    Dim sIncludedFranchiseNames As String
    Dim sPlural As String
    Dim sReportingPeriod As String
    Dim sSQLQuery As String
    
    Dim cDays, cDaysInPeriod As Long
    Dim sFranchName As String
    Dim fMissingDays As Boolean
    Dim cFranchises As Integer
    
    Dim intFileNum As Long
    
    Const conReportType = "Missing Sales Days report for "
    
Dim datThisDay As Date
Dim strSQL As String
Dim strErrMsg As String
Dim strFranID As String
Dim rstDistinctDayForFranch As ADODB.Recordset
Dim rstFranIncluded As ADODB.Recordset

    With lvwProductReport
        .ListItems.Clear
        .Refresh
    End With
    
    If Not IsDateFmtOk() Then   ''' Review Fix Reliance on date format when time permits
        MsgBox "incorrect system date format"
        Exit Sub
    End If
    
    btnMissingDaysSales.Enabled = False
    Call subWriteSearchingMessageToStatusBar
    
    'build a query spec for all dates within the range
    If lblProductReportStartDate = lblProductReportFinishDate Then
        sReportingPeriod = lblProductReportStartDate
    Else
        sReportingPeriod = lblProductReportStartDate & " to " & lblProductReportFinishDate
    End If
    
    ReDim lFranchiseID(gconZeroValue) As Long
    
    ' Build an array containing the ID for each selected franchise,
    ' but only if this franchise is included in the daily capture
    For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
        If (lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) And optProductReportOnSelectedFranchisesOnly(0)) Or _
            (Not optProductReportOnSelectedFranchisesOnly(0)) Then
            ' check if this franchise is included in the daily capture
            strFranID = fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex))
            strSQL = "SELECT " & gconFranchiseTableTSGFranchiseIDField & vbNewLine & _
                     "FROM qryFranchiseLive" & vbNewLine & _
                     "WHERE " & gconFranchiseTableTSGFranchiseIDField & " = " & strFranID
            Set rstFranIncluded = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            If Not rstFranIncluded Is Nothing Then
                If Not (rstFranIncluded.BOF And rstFranIncluded.EOF) Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)
                    lFranchiseID(iNumberOfFranchisesIncluded) = strFranID
                    sIncludedFranchiseNames = sIncludedFranchiseNames & lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
                End If
                rstFranIncluded.Close
            End If
        End If
    Next lArrayRowIndex
    Set rstFranIncluded = Nothing
    
    If iNumberOfFranchisesIncluded = 0 Then
        MsgBox "No live franchise selected", vbExclamation, gconReportManager
        btnMissingDaysSales.Enabled = True
        Exit Sub
    End If
    
    If optProductReportOnSelectedFranchisesOnly(0) Then
        'get rid of the last delimiters
        sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
        'sFranchiseMessageBox = " for " & sIncludedFranchiseNames
    Else
        sIncludedFranchiseNames = gconAllFranchises
        'sFranchiseMessageBox = ""
    End If
    
    If iNumberOfFranchisesIncluded > 1 Then
        sPlural = "s"
    End If
    
    Dim iCurrentFranchise As Integer
                                
    If optSendProductReportToPrinter Then
        On Error GoTo noRprinterMS
        cdlTSGDataWarehouse.ShowPrinter
        Me.Refresh
        
        Printer.Print "Tobacco Station" & sPlural & " - " & sIncludedFranchiseNames
        'leave a dual gap
        Printer.Print vbCrLf

        Printer.Print conReportType & sReportingPeriod
        Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
        'leave a dual gap
        Printer.Print vbCrLf
    ElseIf optSendProductReportToFile Then
        If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
            intFileNum = FreeFile   ' Get unused file
            Open gsProductReportPathAndFilename For Output As #intFileNum
            Print #intFileNum, "Tobacco Station" & sPlural & _
                      " - " & sIncludedFranchiseNames
            'leave a dual gap
            Print #intFileNum, vbCrLf
        
            Print #intFileNum, conReportType & sReportingPeriod
            Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
            'leave a dual gap
            Print #intFileNum, vbCrLf
            Close #intFileNum   ' No Close statement prior to V 3.0.9027
        Else 'environment was not created
            MsgBox "Report was aborted", vbExclamation
        End If 'environement created ?
    End If 'sent prod report to file
    cDaysInPeriod = DateDiff("d", gfsSplitDate(lblProductReportStartDate), gfsSplitDate(lblProductReportFinishDate)) + 1
    
    For iCurrentFranchise = 1 To iNumberOfFranchisesIncluded
        sFranchName = GetFranName(lFranchiseID(iCurrentFranchise))
        datThisDay = GetDateFrom_ddmmmyy(lblProductReportStartDate)
        For cDays = 1 To cDaysInPeriod
            StatusBar sFranchName & " " & cDays, pLog:=False
            DoEvents
            sSQLQuery = "SELECT TransactionDate FROM LiveData " & _
                        " WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & _
                        "  AND Barcode <> " & SqlQ("TOTALCUSTOMERS") & _
                        "  AND (TransactionDate = " & MySqlDate(datThisDay) & ")"
            Set rstDistinctDayForFranch = GetRst(g.cnnDW, pSource:=sSQLQuery, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            If (rstDistinctDayForFranch.BOF And rstDistinctDayForFranch.EOF) Then
                Set gvListItem = frmTSGDataWarehouse.lvwProductReport.ListItems.Add()
                gvListItem.Text = sFranchName
                gsubAddSubItemToListview Format$(datThisDay, gkFmtDateUnambiguous), 1
                fMissingDays = True
            End If
            rstDistinctDayForFranch.Close
            Set rstDistinctDayForFranch = Nothing
            
            ' get the next day
            datThisDay = DateAdd("d", 1, datThisDay)
            
        Next cDays
        If fMissingDays Then
            cFranchises = cFranchises + 1
        End If
        fMissingDays = False
    Next iCurrentFranchise
    
    If cFranchises = 0 Then
        MsgBox "No missing days data for " & sIncludedFranchiseNames
    End If
    
    StatusBar cFranchises & " out of " & iNumberOfFranchisesIncluded & " had missing sales", pLog:=False

    btnMissingDaysSales.Enabled = True
    Exit Sub
    
noRprinterMS:
    btnMissingDaysSales.Enabled = True

End Sub

Private Sub btnPurgeUploadsPending_Click()
    PurgeUploadsPending fAll:=True, pFranID:=0      ' true means all of them
End Sub

Private Sub btnSaveMessage_Click()
' Ascertain the next available message num
Dim strSQL As String
Dim strErrMsg As String
    Dim rsMessages As ADODB.Recordset
    Dim iNextMessageNum As Integer
    Dim sMessageFileName As String
    Dim sUploadMessage As String
    Dim sMessageText As String
    Dim intFileNum As Long
    
    If txtMessageTitle.Text = "" Then
        MsgBox "Please enter a title for this message"
        txtMessageTitle.SetFocus
        Exit Sub
    End If
    
    strSQL = "SELECT Max(MessageNum) FROM Messages"
    iNextMessageNum = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=0) + 1
    
    sMessageFileName = g.strLocalMessageFolder & "\" & gconNewMessageFilePrefix & Format(iNextMessageNum, "000") & gconTextFileSuffix
    DeleteFile (sMessageFileName)
    
    ' break the text into 40-char max lines by inserting CR so it fits on docket printer
    sMessageText = breaktext(txtNewMessage.Text)
    
    intFileNum = FreeFile   ' Get unused file
    Open sMessageFileName For Output As #intFileNum
        Print #intFileNum, txtMessageTitle.Text
        Print #intFileNum, "---------------------------------------------------"
        Print #intFileNum, sMessageText
        Print #intFileNum, "---------------------------------------------------"
    Close #intFileNum
    
    Set rsMessages = GetRst(pCnn:=g.cnnDW, pSource:="Messages", pSourceType:=adCmdTable, pRstType:=eEditableFwdOnly, pErrMsg:=strErrMsg)
    rsMessages.AddNew
        rsMessages("MessageNum") = iNextMessageNum
        rsMessages("MessageTitle") = txtMessageTitle.Text
    rsMessages.Update
    rsMessages.Close
    Set rsMessages = Nothing
    
    ' Copy the message file into the 'uploads' folder ready to be uploaded.
    sUploadMessage = g.strUploadsFolder & "\" & gconNewMessageFilePrefix & Format(iNextMessageNum, "000") & gconTextFileSuffix
    DeleteFile (sUploadMessage)
    FileCopy sMessageFileName, sUploadMessage
    lstUploadItemList.AddItem sUploadMessage
    lstUploadItemList.Refresh
    
    MsgBox "Message " & SQ(txtMessageTitle.Text) & " saved as " & sUploadMessage & vbCrLf
    txtMessageTitle.Text = ""
    txtMessageTitle.Refresh
    txtNewMessage.Text = ""
    txtNewMessage.Refresh
    
End Sub

Private Sub cboBataTabTxOrProcessedDate_Click()
Dim bEnableProcessedStatusSelection As Boolean
Dim opt As OptionButton

    bEnableProcessedStatusSelection = InStr(cboBataTabTxOrProcessedDate, "Transaction")
    
    fraBataTabProcessedStatus.Enabled = bEnableProcessedStatusSelection
    For Each opt In Me.optBataProcessed
        opt.Enabled = bEnableProcessedStatusSelection
    Next opt
    
    'dtpBataTabTxDate is used as either TxDate or ProcessingDate according to
    'selection in cboBataTabTxOrProcessedDate combo. dtpBataTabTxDate should
    'really be renamed to something like dtpBataTabSelnDate
    If bEnableProcessedStatusSelection Then
    '   Enabling selection by process status AND therefore also Tx Date
        If dtpBataTabTxDate.Value = Date Then
            dtpBataTabTxDate.Value = fdtmYesterday()
        End If
        dtpBataTabTxDate.MaxDate = fdtmYesterday()
    Else
   '    Not enabling selection by process status BUT by Processing Date
        dtpBataTabTxDate.MaxDate = Date
    End If
    
    RefreshBataTabGrid
End Sub

Private Sub cboCategory_Click()
    ConfigureStkTabCtls
End Sub

Private Sub cboDCTabPromoGrade_Click()
    If cboDCTabPromoGrade.ListIndex > -1 Then
        cmdSaveFranchiseDetails.Enabled = True
    End If
End Sub

Private Sub cboDCTabRegion_Click()
    If cboDCTabRegion.ListIndex > -1 Then
    '   Cbo has a selection
        cmdSaveFranchiseDetails.Enabled = True
    End If
End Sub

Private Sub cboFranchiseType_Click()
    If cboFranchiseType.ListIndex > -1 Then
    '   When adding a new franchise combo is deselected by setting ListIndex = -1
        cboFranchiseType.ToolTipText = m.astrFranTypeTooltip(cboFranchiseType.ListIndex)
        cmdSaveFranchiseDetails.Enabled = True
    End If
End Sub

Private Sub cboGoodsTax_Change()
    cboSalesTax = cboGoodsTax
End Sub

Private Sub cboState_Click()
    If cboState.ListIndex > -1 Then
        cmdSaveFranchiseDetails.Enabled = True
    End If
End Sub

Private Sub cboSupplier_Change()
    If Not gbClickEventIsSuppressed Then
        cmdSaveStockDetails.Enabled = True
    End If
End Sub

Private Sub checkMissingDaysSales()
Const kDaysAgo As Long = 5
    
    Dim rstSalesForFranch As ADODB.Recordset
    Dim rsFranchIncluded As ADODB.Recordset
    Dim sFranchName As String
    Dim sSQLQuery As String
Dim strErrMsg As String
Dim strMsg As String

    
    Set rsFranchIncluded = GetRst(pCnn:=g.cnnDW, _
                                  pSource:="qryFranchiseLive", _
                                  pSourceType:=adCmdTable, _
                                  pErrMsg:=strErrMsg)
    
    Do Until rsFranchIncluded.EOF
        sSQLQuery = "SELECT FranchiseIDTSG FROM LiveData " & _
                    " WHERE FranchiseIDTSG = " & rsFranchIncluded!FranchiseIDTSG & _
                    "  AND Barcode <> " & SqlQ("TOTALCUSTOMERS") & _
                    "  AND TransactionDate >= " & MySqlDate(DateAdd("d", kDaysAgo * (-1), Date))
    
        Set rstSalesForFranch = GetRst(pCnn:=g.cnnDW, _
                                       pSource:=sSQLQuery, _
                                       pSourceType:=adCmdText, _
                                       pErrMsg:=strErrMsg)
                                       
        sFranchName = rsFranchIncluded!FranchiseBusinessName
        If (rstSalesForFranch.BOF And rstSalesForFranch.EOF) Then
            strMsg = "*** No Sales data collected for at least " & kDaysAgo - 1 & " days! ***"
            StatusBar strMsg, sFranchName
        End If
        rstSalesForFranch.Close
        Set rstSalesForFranch = Nothing
        rsFranchIncluded.MoveNext
        DoEvents
    Loop
    rsFranchIncluded.Close
    Set rsFranchIncluded = Nothing
    
End Sub

Private Sub chkIncludeInDataCaptureCycle_Click()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub chkPromoSelectFranchise_Click()
Dim bSelectFranchise As Boolean
    
    bSelectFranchise = ChkBoxToBool(chkPromoSelectFranchise)
    
    LoadGrdPromoTabRebates pUsePromoGrades:=Not bSelectFranchise
    
    fraPromoSelFranchise.Visible = bSelectFranchise
    lblPromoFran.Visible = bSelectFranchise
    
    lblPromoState.Visible = Not bSelectFranchise
    lblPromoRegion.Visible = Not bSelectFranchise
    
End Sub

Private Sub chkPromoTabAllRegions_Click()
Dim lngLoop As Long

    With lstPromoTabRegion
        .Enabled = Not (chkPromoTabAllRegions.Value = vbChecked)
    '   Clear Selections
        For lngLoop = 0 To .ListCount - 1
            .Selected(lngLoop) = False
        Next lngLoop
    End With
    
End Sub

Private Sub chkPromoTabAllStates_Click()
Dim lngLoop As Long

    With lstPromoTabState
        .Enabled = Not (chkPromoTabAllStates.Value = vbChecked)
    '   Clear Selections
        For lngLoop = 0 To .ListCount - 1
            .Selected(lngLoop) = False
        Next lngLoop
    End With

End Sub

Private Sub chkSalesRptTab_IncludeClosedFrans_Click()
    PopulateLstProductReportsFranchiseBusinessName pIncludeClosedFrans:=ChkBoxToBool(chkSalesRptTab_IncludeClosedFrans)
End Sub

Private Sub chkStockTab_IncludeDeletedStock_Click()
    subPopulateStockListboxes pExclDeletedStk:=ChkBoxToBool(chkStockTab_IncludeDeletedStock)
End Sub

Private Sub ClearCreatePromoCtls()
Dim lngLoop As Long

    txtPromoName.Text = vbNullString
    lstPromoProducts.Clear
    ListBoxClearSelections lstPromoFranchise
    ListBoxClearSelections lstPromoSubCat
    ListBoxClearSelections lstPromoTabState
    ListBoxClearSelections lstPromoTabRegion
    chkPromoTabAllStates.Value = vbChecked  ' Checked is default
    chkPromoTabAllRegions.Value = vbChecked ' Checked is default
    With grdPromoTabRebates
        For lngLoop = .FixedRows To .Rows - 1
            .TextMatrix(Row:=lngLoop, col:=2) = 0   ' Carton Rebate
            .TextMatrix(Row:=lngLoop, col:=3) = 0   ' Packet Rebate
        Next lngLoop
    End With
    txtPromoName.SetFocus

End Sub

Private Sub cmdAddNewFranchise_Click()
Dim strBusinessName As String

    Dim rstFranBusinessName As ADODB.Recordset  '!!! ManualFix Clearing: Object variable not cleared: rstFranBusinessName
    Dim strMsg As String
Dim strSQL As String
Dim strErrMsg As String


    If Right(cmdAddNewFranchise.Caption, 1) = "w" Then 'is add new
        cmdAddNewFranchise.Caption = "&OK"
        cmdSaveFranchiseDetails.Enabled = False
        cmdSaveFranchiseDetails.Visible = False
        cmdCaptureData.Enabled = False
        cmdCaptureData.Visible = False
        cmdCaptureSelected.Enabled = False
        cmdCaptureSelected.Visible = False
        cmdCloseSelectedFranchises.Enabled = False
        cmdCloseSelectedFranchises.Visible = False
        lstDataCaptureFranchiseBusinessName.Enabled = False
        
        With txtNewFranchiseBusinessName
            .Text = ""
            .Enabled = True
            .Visible = True
            .SetFocus
        End With
        LockDCTabFranchiseCtls pLocked:=False
        subClearFranchiseDetails
        txtRASUsername = gCompanyIdentifier
        cboDCTabPromoGrade.ListIndex = -1 ' No selection
        txtFaxNum = "unknown"
        txtBATAFranchiseID = "0"
    Else 'must have been editing the fields and is now accepted "OK" by the user
        If Trim(txtNewFranchiseBusinessName) <> "" Then
            txtNewFranchiseBusinessName = fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks(txtNewFranchiseBusinessName)
            strSQL = "SELECT FranchiseBusinessName, Live FROM Franchises " & _
                     "WHERE FranchiseBusinessName = " & SqlQ(txtNewFranchiseBusinessName)
            Set rstFranBusinessName = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
''          If Len(Trim$(txtNewFranchiseBusinessName.Text)) = 0 Then
''              MsgBox "Franchise business name is mandatory." & vbNewLine & "Please enter a business name.", vbCritical
''              With txtNewFranchiseBusinessName
''                  .Text = ""
''                  .SetFocus
''              End With
''              Exit Sub
            If Not (rstFranBusinessName.BOF And rstFranBusinessName.EOF) Then
                strBusinessName = rstFranBusinessName(gconFranchiseTableBusinessNameField)
                strMsg = "Franchise name cannot be added." & vbNewLine & _
                          SQ(strBusinessName) & " already exists in database."
                If Not CBool(rstFranBusinessName!Live) Then
                    strMsg = strMsg & vbNewLine & "This name was used for a franchise that has been closed but not permanently deleted."
                End If
                MsgBox strMsg, vbExclamation
                txtNewFranchiseBusinessName.Text = ""
                txtNewFranchiseBusinessName.SetFocus
                Exit Sub
            ElseIf cboDCTabRegion.ListIndex = -1 Then
                strMsg = "Region is mandatory." & vbNewLine & "Please select a region."
                MsgBox strMsg, vbInformation
                cboDCTabRegion.SetFocus
                Exit Sub
            ElseIf cboState.ListIndex = -1 Then
                strMsg = "State is mandatory." & vbNewLine & "Please select a state."
                MsgBox strMsg, vbInformation
                cboState.SetFocus
                Exit Sub
            ElseIf cboDCTabPromoGrade.ListIndex = -1 Then
                strMsg = "Promo Grade is mandatory." & vbNewLine & "Please select a promo grade."
                MsgBox strMsg, vbInformation
                cboDCTabPromoGrade.SetFocus
                Exit Sub
            Else
                'name was not already found, so add it
                'windows limitation for nodename is actually 15, but due to pricemodule
                'backoffice we only allow 14 (for the addition of backoffice nodes
                'nodename1, nodename2, nodename3 etc., which adds an extra character)
                If Len(Trim(txtNodename)) < 1 Then
                    txtNodename = Left(gCompanyIdentifier & fsSpacesRemovedFrom(txtNewFranchiseBusinessName), 14)
                End If
                txtRASPassword.Text = "headoffice"
                If Len(Trim(txtModem)) < 1 Then
                    txtModem = "99999999"
                End If
                subSaveFranchiseDetails bAddNewFranchise:=True
                subPopulateFranchiseBusinessNameListBoxes
                gsubRefreshEventLogDisplay
            End If
        End If
        'reset the form
        With txtNewFranchiseBusinessName
            .Text = ""
            .Visible = False
            .Enabled = False
        End With
        cmdAddNewFranchise.Caption = "< &Add New"
        subClearFranchiseDetails
        
        cmdSaveFranchiseDetails.Enabled = False
        cmdSaveFranchiseDetails.Visible = True
        cmdCaptureData.Enabled = True
        cmdCaptureData.Visible = True
        cmdCaptureSelected.Enabled = g.bMaster And (lstDataCaptureFranchiseBusinessName.SelCount > 0)
        cmdCaptureSelected.Visible = True
        cmdCloseSelectedFranchises.Enabled = g.bMaster And (lstDataCaptureFranchiseBusinessName.SelCount > 0)
        cmdCloseSelectedFranchises.Visible = True
        gbClickEventIsSuppressed = True
        With lstDataCaptureFranchiseBusinessName
            .Enabled = True
            .ListIndex = gconDoNotDisplayAnyItems
            .SetFocus
        End With
        gbClickEventIsSuppressed = False
    End If

End Sub

Private Sub cmdAddNewStockItem_Click()
Dim bCaptionIsAddNew As Boolean
Dim bResetStockTab As Boolean
Dim lngStkID As Long
Dim strMsg As String
Dim strErrMsg As String
    
    bCaptionIsAddNew = Right(cmdAddNewStockItem.Caption, 1) = "w"
    If bCaptionIsAddNew Then
        cmdAddNewStockItem.Caption = "&OK"
        With cmdSaveStockDetails
            .Enabled = False
            .Visible = False
        End With
        lstDescription.Enabled = False
        Call subClearStockFields
        txtStkItemDescription.SetFocus
        subPopulateStockListboxes pExclDeletedStk:=True
    
    Else 'must have been editing the fields and is now accepted "OK" by the user
        txtStkItemDescription = fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks(txtStkItemDescription)
        If Not IsStkCtlsValid(strErrMsg) Then
            MsgBox strErrMsg, vbExclamation
        Else
            If cboCategory = gkCAT_CigPkt And (Not ChkBoxToBool(chkPackage)) Then
                If MsgBox("Should the 'Package' box be ticked?", vbYesNo) = vbYes Then
                    chkPackage.Value = 1
                    chkPackage.Refresh
                End If
            End If
            lngStkID = subSaveStockDetails(pAddNewItem:=True)
            If Not ((cboCategory = gkCAT_CigCtn) And (chkPackage.Value = 0)) Then
                bResetStockTab = True
                cboSubCategory.Enabled = True
                cboCategory.Enabled = True
                chkPackage.Enabled = True
            Else
            '   Prompt to add the packet linked to the carton
                strMsg = "Add packet linked to " & SQ(txtStkItemDescription) & "?"
                If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                '   Keep as many of the appropriate prefilled values as possible and clear the inaappropriate values
                    txtBarcode.Text = vbNullString
                    txtSticks.Text = vbNullString
                    txtWholesaleListPrice.Text = 0
                    txtRRP.Text = 0
                    cboGoodsTax = "N/A"
                    cboSalesTax = "N/A"
                    chkPackage.Value = 1
                    cboCategory.Text = gkCAT_CigPkt
                    LoadCboCtnContainingPkt pRecordSource:="qryStock"  ' ie Query on Stock table excluding deleted stock items
                    SetCboCtnContainingPkt pStkId:=lngStkID
                    cboCtnContainingPkt.Visible = True
                    lblCartonsPerPacket.Visible = True
                    txtCartonsPerPacket.Visible = True
                    txtCartonsPerPacket.Enabled = True
                    cboSubCategory.Enabled = False
                    cboCategory.Enabled = False
                    chkPackage.Enabled = False
                End If
            End If
            If bResetStockTab Then
                subPopulateStockListboxes pExclDeletedStk:=ChkBoxToBool(chkStockTab_IncludeDeletedStock)
                subResetStockForm
                gbClickEventIsSuppressed = False
            End If
        End If
    End If
End Sub

Private Sub cmdAllItems_Click()
'--------------------------------------------------------------------------------------------------------------
'  AUrban Procedure is a candidate for splitting into two procedures (Summarised Rpt and Not Summarised Rpt
'  AUrban (cmdAllItems_Click & cmdMarketShare_Click are v. similar. Prob cut & pasted and modified)
'--------------------------------------------------------------------------------------------------------------
Dim strErrMsg As String
Dim rstAllSameBarcodesForTheReportingPeriod As ADODB.Recordset

    Dim bIndexSwapped As Boolean
    
    Dim cTotalSalesIncludingTaxForThePeriod As Currency
        
    Dim iArrayColumnIndex As Integer, _
        iNumberOfCopies As Integer, _
        iNumberOfFranchisesIncluded As Integer, _
        iPageNumber As Integer
    
    Dim lArrayRowIndex As Long
    Dim BarcodeArrayIndex As Long

    Dim lTotalCustomersForThePeriod As Long, _
        lTotalNumberOfBarcodesForTheReportingPeriod As Long, _
        lTotalItemsSoldForThePeriod As Long
    
Dim rstDistinctBarcodesForTheReportingPeriod As ADODB.Recordset

    Dim sFranchiseMessageBox As String, _
        sIncludedDates As String, _
        sIncludedFranchiseIDs As String, _
        sIncludedFranchiseNames As String, _
        sPlural As String
    Dim sReportingPeriod As String
    Dim sSQLQuery As String

    Dim vPlaceHolder As Variant
    ReDim sArrBarcode(0) As String
    Dim sWholesaleQuery As String
    Dim sRetailSell As String
    Dim sWHSSell As String
    Dim sWHSpercent As String
    Static bRptlnProcess As Boolean
    Dim intFileNum As Integer

    'data array
    '
    Const conSortIndex = 1, _
          conProduct = 2, _
          conQuantity = 3, _
          conNormalSell = 4, _
          conNormalSellcount = 5, _
          conTotalSalesInc = 6, _
          conWHSQty = 7, _
          conWHSTotalSell = 8, _
          conWHSActualSell = 9, _
          conWHSAmntPercent = 10, _
          conWHSQtyPercent = 11
    
    'tabstop array uses same as data array except for this extra
    Const conDisplayAverageSalesInc = 2
    
    Const conReportType = "Sales report for "
    
Dim datReportStart As Date
Dim lngRecCount As Long

    If bRptlnProcess Then
      ' If processing is in progress, cancel it.
      cmdAllItems.Caption = "All items"
      bRptlnProcess = False
      Exit Sub
    End If
    
    With lvwProductReport
        .ListItems.Clear
        .Refresh
    End With

    If Not IsDateFmtOk() Then   ''' Review Fix Reliance on date format when time permits
        MsgBox "incorrect system date format"
        Exit Sub
    End If
    
    cmdAllItems.Caption = "Cancel"
    bRptlnProcess = True
    
    Call subWriteSearchingMessageToStatusBar
    
    'build a query spec for all dates within the range
    datReportStart = GetDateFrom_ddmmmyy(lblProductReportStartDate)

    If lblProductReportStartDate = lblProductReportFinishDate Then
        sIncludedDates = "TransactionDate = " & MySqlDate(datReportStart)
        sReportingPeriod = lblProductReportStartDate
    Else
        sIncludedDates = "TransactionDate BETWEEN " & _
                         MySqlDate(datReportStart) & " AND " & MySqlDate(GetDateFrom_ddmmmyy(lblProductReportFinishDate))
        sReportingPeriod = lblProductReportStartDate & " to " & lblProductReportFinishDate
    End If
    
    If optWRS(0) Then ' show all sales, so ignore the wholesale qty
            sWholesaleQuery = ""
    ElseIf optWRS(1) Then ' show sales that were totally retail
            sWholesaleQuery = " AND " & gconLiveDataTableWholesaleQty & " = 0 "
    ElseIf optWRS(2) Then ' show sales that had some or all Wholesale
            sWholesaleQuery = " AND " & gconLiveDataTableWholesaleQty & " <> 0 "
    End If
    
    If optProductReportNotSummarised(0) Then
        ReDim lFranchiseID(gconZeroValue) As Long
        
            'build an array containing the ID for each selected franchise
        For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
            If (lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) And optProductReportOnSelectedFranchisesOnly(0)) Or _
                (Not optProductReportOnSelectedFranchisesOnly(0)) Then
                iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)
                lFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex))
                sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                          lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
            End If
        Next lArrayRowIndex
        
        If iNumberOfFranchisesIncluded = 0 Then
            MsgBox "No franchise selected", vbExclamation, gconReportManager
            cmdAllItems.Caption = "All items"
            bRptlnProcess = False
            Exit Sub
        End If
        
        If optProductReportOnSelectedFranchisesOnly(0) Then
            'get rid of the last delimiters
            sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
            sFranchiseMessageBox = " for " & sIncludedFranchiseNames
        Else
            sIncludedFranchiseNames = gconAllFranchises
            sFranchiseMessageBox = ""
        End If
        
        If iNumberOfFranchisesIncluded > 1 Then
            sPlural = "s"
        End If
        
        Dim iCurrentFranchise As Integer
                                    
        If optSendProductReportToPrinter Then
                'On Error GoTo NotSummarisedPrinterErrorHandler
            cdlTSGDataWarehouse.ShowPrinter
            Me.Refresh
            
            Printer.Print "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
            'leave a dual gap
            Printer.Print vbCrLf

            Printer.Print conReportType & sReportingPeriod
            Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
            'leave a dual gap
            Printer.Print vbCrLf
        ElseIf optSendProductReportToFile Then
            If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                intFileNum = FreeFile   ' Get unused file
                Open gsProductReportPathAndFilename For Output As #intFileNum
                Print #intFileNum, "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
                'leave a dual gap
                Print #intFileNum, vbCrLf
            
                Print #intFileNum, conReportType & sReportingPeriod
                Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                'leave a dual gap
                Print #intFileNum, vbCrLf
            Else 'environment was not created
                MsgBox "Report was aborted", vbExclamation
                GoTo TidyUpnotSummarised
            End If 'environement created ?
        End If 'sent prod report to file
        
        
        For iCurrentFranchise = 1 To iNumberOfFranchisesIncluded
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & _
                        "WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & " " & _
                         " AND (" & sIncludedDates & ") " & _
                         " AND (Quantity <> 0) " & _
                          sWholesaleQuery
        
            Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                  pSource:=sSQLQuery, _
                                                                  pSourceType:=adCmdText, _
                                                                  pErrMsg:=strErrMsg)

            If Not (rstDistinctBarcodesForTheReportingPeriod.BOF And _
                    rstDistinctBarcodesForTheReportingPeriod.EOF) Then
                
                Call subWriteSizingArraysMessageToStatusBar
                
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conDisplayAverageSalesInc To conWHSQtyPercent) As Integer
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + _
                                               Len(gconTruncateCharacter) + _
                                               gconTruncateExtensionWidth + _
                                               Len(gconSpace) + _
                                               Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conNormalSell) = iArrTabStop(conQuantity) + _
                                                             Len(gcon5DigitDollarFormat) + _
                                                             Len(gconSpace) '
                    
                    iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conNormalSell) + _
                                                             Len(gcon5DigitDollarFormat) + _
                                                             Len(gconSpace) '
                    
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conDisplayAverageSalesInc) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                    iArrTabStop(conWHSQty) = iArrTabStop(conTotalSalesInc) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                    iArrTabStop(conWHSTotalSell) = iArrTabStop(conWHSQty) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                Else
                    iArrTabStop(conQuantity) = 43
                    iArrTabStop(conNormalSell) = 55 '
                    iArrTabStop(conDisplayAverageSalesInc) = 69
                    iArrTabStop(conTotalSalesInc) = 80 ' 69
                    iArrTabStop(conWHSQty) = 90
                    iArrTabStop(conWHSTotalSell) = 104
                    iArrTabStop(conWHSActualSell) = 116
                    iArrTabStop(conWHSAmntPercent) = 128
                    iArrTabStop(conWHSQtyPercent) = 138
                End If
                
                'size the array
                'do not remove this
                lTotalNumberOfBarcodesForTheReportingPeriod = gconZeroValue
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    If fbProductIsIncludedInThisProductReport(rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                        lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod + 1
                        ReDim Preserve sArrBarcode(lTotalNumberOfBarcodesForTheReportingPeriod) 'PAL
                        sArrBarcode(lTotalNumberOfBarcodesForTheReportingPeriod) = rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) 'PAL
                    End If
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
                rstDistinctBarcodesForTheReportingPeriod.Close
                Set rstDistinctBarcodesForTheReportingPeriod = Nothing
                ReDim Varrsalesdata(conSortIndex To conWHSTotalSell, _
                                    1 To lTotalNumberOfBarcodesForTheReportingPeriod) As Variant
                
                Call subWriteCollatingMessageToStatusBar
                
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                lTotalItemsSoldForThePeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                lTotalCustomersForThePeriod = gconZeroValue
                
                For BarcodeArrayIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod ' PALb50
                    sSQLQuery = "SELECT * FROM LiveData " & _
                                " WHERE FranchiseIDTSG  = " & lFranchiseID(iCurrentFranchise) & _
                                 " AND (" & sIncludedDates & ")" & _
                                 " AND (Barcode = " & SqlQ(sArrBarcode(BarcodeArrayIndex)) & ")"
                    DoEvents
                    If Not bRptlnProcess Then
                        StatusBar "Report cancelled.", pLog:=False
                        Exit Sub
                    End If
                    lngRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
                    Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                         pSource:=sSQLQuery, _
                                                                         pSourceType:=adCmdText, _
                                                                         pErrMsg:=strErrMsg)
                    
                    'has to be more than zero records, so don't waste time testing for it
                    
                    If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) = "TOTALCUSTOMERS" Then
                        Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                            lTotalCustomersForThePeriod = lTotalCustomersForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                            rstAllSameBarcodesForTheReportingPeriod.MoveNext
                        Loop
                    Else
                        lArrayRowIndex = lArrayRowIndex + 1
                        Call subDisplayCurrentRecordToUser( _
                             lArrayRowIndex, _
                             lTotalNumberOfBarcodesForTheReportingPeriod)
                        
                        If optDescription(0) Then
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                fsDescriptionFrom(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))
                        Else
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)
                        End If
                    
                        Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                            DoEvents
                            If Not bRptlnProcess Then
                                StatusBar "Report cancelled.", pLog:=False
                                Exit Sub
                            End If
                            'avert an overflow divide by zero
                            If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField) <> 0 Then
                                Varrsalesdata(conQuantity, lArrayRowIndex) = _
                                    Varrsalesdata(conQuantity, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                
                                lTotalItemsSoldForThePeriod = _
                                    lTotalItemsSoldForThePeriod + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                
                                Varrsalesdata(conNormalSell, lArrayRowIndex) = _
                                    Varrsalesdata(conNormalSell, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod!NormalSellInc
                                 
                                 Varrsalesdata(conTotalSalesInc, lArrayRowIndex) = _
                                    Varrsalesdata(conTotalSalesInc, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                 
                                 Varrsalesdata(conWHSQty, lArrayRowIndex) = _
                                    Varrsalesdata(conWHSQty, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableWholesaleQty)
                                
                                 Varrsalesdata(conWHSTotalSell, lArrayRowIndex) = _
                                    Varrsalesdata(conWHSTotalSell, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableWholesaleActualSell)
                                
                                cTotalSalesIncludingTaxForThePeriod = _
                                    cTotalSalesIncludingTaxForThePeriod + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                            End If
                            Varrsalesdata(conNormalSellcount, lArrayRowIndex) = lngRecCount
                            rstAllSameBarcodesForTheReportingPeriod.MoveNext
                        Loop
                    End If
                    rstAllSameBarcodesForTheReportingPeriod.Close
                Next BarcodeArrayIndex
                Set rstAllSameBarcodesForTheReportingPeriod = Nothing
                
                If lTotalCustomersForThePeriod <> gconZeroValue Then
                    'don't want to sort this as an array component
                    lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod - 1
                End If
                
                Call subWriteSortingMessageToStatusBar
                
                Do 'at least one sort by description pass
                    bIndexSwapped = False
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod - 1
                        If (Varrsalesdata(conProduct, lArrayRowIndex) > _
                            Varrsalesdata(conProduct, lArrayRowIndex + 1)) Then 'swap
                            For iArrayColumnIndex = conProduct To conWHSTotalSell
                                vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                bIndexSwapped = True
                            Next iArrayColumnIndex
                        End If
                    Next lArrayRowIndex
                Loop While bIndexSwapped
                
                If optSendProductReportToDisplay Then
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise))
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            Set gvListItem = lvwProductReport.ListItems.Add()
                            gvListItem.Text = Varrsalesdata(conProduct, lArrayRowIndex)
                            Call gsubAddSubItemToListview( _
                                 Varrsalesdata(conQuantity, lArrayRowIndex), 1)
                            Call gsubAddSubItemToListview( _
                                 Format(Varrsalesdata(conNormalSell, lArrayRowIndex) / _
                                        Varrsalesdata(conNormalSellcount, lArrayRowIndex), gcon5DigitDollarFormat), 2)
                            ' only calculate avg. actual retail sell if retail qty (which is total - whs) is not zero
                            If Varrsalesdata(conQuantity, lArrayRowIndex) <> Varrsalesdata(conWHSQty, lArrayRowIndex) Then
                                Call gsubAddSubItemToListview( _
                                 Format((Varrsalesdata(conTotalSalesInc, lArrayRowIndex) - Varrsalesdata(conWHSTotalSell, lArrayRowIndex)) / _
                                        (Varrsalesdata(conQuantity, lArrayRowIndex) - Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat), 3)
                            End If
                            Call gsubAddSubItemToListview( _
                                 Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 4)
                            On Error Resume Next
                            Call gsubAddSubItemToListview( _
                                 Varrsalesdata(conWHSQty, lArrayRowIndex), 5)
                            Call gsubAddSubItemToListview( _
                                 Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex), gcon6DigitDollarFormat), 6)
                            If Varrsalesdata(conWHSQty, lArrayRowIndex) <> 0 Then
                                Call gsubAddSubItemToListview( _
                                 Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) / Varrsalesdata(conWHSQty, lArrayRowIndex), gcon5DigitDollarFormat), 7)
                                Call gsubAddSubItemToListview( _
                                 Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) * 100 / Varrsalesdata(conTotalSalesInc, lArrayRowIndex), "###"), 8)
                                Call gsubAddSubItemToListview( _
                                  Format(Varrsalesdata(conWHSQty, lArrayRowIndex) * 100 / Varrsalesdata(conQuantity, lArrayRowIndex), "###"), 9)
                            End If
                    
                        End If
                    Next lArrayRowIndex
                    
                    'leave a gap then totals
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                    
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = "Total items"
                    Call gsubAddSubItemToListview(lTotalItemsSoldForThePeriod, 1)
                    Call gsubAddSubItemToListview(Format(cTotalSalesIncludingTaxForThePeriod, gcon6DigitDollarFormat), 4) ' 3)
                    
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave another gap then total customers
                            Set gvListItem = lvwProductReport.ListItems.Add()
                            gvListItem.Text = gconSpace
                            
                            Set gvListItem = lvwProductReport.ListItems.Add()
                            gvListItem.Text = "Total customers"
                            Call gsubAddSubItemToListview(lTotalCustomersForThePeriod, 1)
                        End If
                    End If
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                    Me.Refresh
                ElseIf optSendProductReportToPrinter Then
                    Printer.Print GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    'leave a gap
                    Printer.Print gconSpace
                    
                    'headings
                    '  added 2 lines below: Tab & "Normal Sell"
                    Printer.Print "Product"; _
                                   Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "  Qty"; _
                                   Tab(iArrTabStop(conNormalSell) - Len("Normal Sell")); _
                                  "Normal Sell"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Actual Sell")); _
                                  " Actual Sell"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "Tot (inc)"; _
                                  Tab(iArrTabStop(conWHSQty) - Len("NCS Qty")); _
                                  "NCS Qty"; _
                                  Tab(iArrTabStop(conWHSTotalSell) - Len("NCS Total)")); _
                                  "NCS Total"; _
                                  Tab(iArrTabStop(conWHSActualSell) - Len("NCS Sell)")); _
                                  "NCS Sell"; _
                                  Tab(iArrTabStop(conWHSAmntPercent) - Len("$NCS %)")); _
                                  "$NCS %"; _
                                  Tab(iArrTabStop(conWHSQtyPercent) - Len("NCS %)")); _
                                  "NCS %"
                    'leave a gap
                    Printer.Print gconSpace
                    
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            '  added 2 lines below: Tab & Format with Normalsell
                            If Varrsalesdata(conQuantity, lArrayRowIndex) <> Varrsalesdata(conWHSQty, lArrayRowIndex) Then
                                sRetailSell = Format((Varrsalesdata(conTotalSalesInc, lArrayRowIndex) - Varrsalesdata(conWHSTotalSell, lArrayRowIndex)) / _
                                        (Varrsalesdata(conQuantity, lArrayRowIndex) - Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat)
                            Else
                                sRetailSell = ""
                            End If
                            If Varrsalesdata(conWHSQty, lArrayRowIndex) <> 0 Then
                                sWHSSell = Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) / Varrsalesdata(conWHSQty, lArrayRowIndex), gcon5DigitDollarFormat)
                                sWHSpercent = Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex) * 100 / Varrsalesdata(conQuantity, lArrayRowIndex)), "###")
                            Else
                                sWHSSell = ""
                                sWHSpercent = ""
                            End If
                            
                            Printer.Print _
                                Varrsalesdata(conProduct, lArrayRowIndex); _
                                Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                sRetailSell; _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conWHSQty) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conWHSTotalSell) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conWHSActualSell) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                sWHSSell; _
                                Tab(iArrTabStop(conWHSAmntPercent) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) * 100 / Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), "###"); _
                                Tab(iArrTabStop(conWHSQtyPercent) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                sWHSpercent
                        End If
                    Next lArrayRowIndex
                    
                    'leave a gap
                    Printer.Print gconSpace
                    
                    'expose the totals
                    Printer.Print "Total"; _
                                Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); _
                                Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); _
                                Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave a dual gap
                            Printer.Print vbCrLf
                    
                            'expose number of customers
                            Printer.Print "Total customers"; _
                                           Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); _
                                           Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                        End If
                    End If
                    'leave a dual gap
                    Printer.Print vbCrLf
                Else 'must be to file
                    Print #intFileNum, GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    'leave a gap
                    Print #intFileNum, gconSpace
                
                    'headings
                    If chkProductReportTabDelimited Then
                        '  added 2 lines "Normal" & vbtab
                        Print #intFileNum, "Product"; _
                                   vbTab; _
                                  "Qty"; _
                                   vbTab; _
                                  "Normal Sell"; _
                                   vbTab; _
                                  "Actual Sell"; _
                                   vbTab; _
                                  "Tot (inc)"; _
                                   vbTab; _
                                  "NCS Qty"; _
                                   vbTab; _
                                  "NCS Total"; _
                                  vbTab; _
                                  "NCS Sell"; _
                                  vbTab; _
                                  "$NCS %"; _
                                  vbTab; _
                                  "NCS %"
                    Else 'normal tabs
                        '  added 2 lines "Normal" & tab
                    
                        Print #intFileNum, "Product"; _
                                   Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "  Qty"; _
                                   Tab(iArrTabStop(conNormalSell) - Len("Normal")); _
                                  "Normal"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Promo")); _
                                  "Promo"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "  Tot (inc)"; _
                                   Tab(iArrTabStop(conWHSQty) - Len("Tot (inc)")); _
                                  "  NCS Qty"; _
                                   Tab(iArrTabStop(conWHSTotalSell) - Len("Tot (inc)")); _
                                  "  NCS Total"; _
                                  Tab(iArrTabStop(conWHSActualSell) - Len("Tot (inc)")); _
                                  "  NCS Sell"; _
                                  Tab(iArrTabStop(conWHSAmntPercent) - Len("Tot (inc)")); _
                                  "  $NCS %"; _
                                  Tab(iArrTabStop(conWHSQtyPercent) - Len("Tot (inc)")); _
                                  "  NCS %";
                        
                        Print #intFileNum, Tab(iArrTabStop(conNormalSell) - Len("Sell")); _
                                  "Sell"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Sell")); _
                                  "Sell"
                    End If
                    
                    'leave a gap
                    Print #intFileNum, gconSpace
                    
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            If Varrsalesdata(conQuantity, lArrayRowIndex) <> Varrsalesdata(conWHSQty, lArrayRowIndex) Then
                                sRetailSell = Format((Varrsalesdata(conTotalSalesInc, lArrayRowIndex) - Varrsalesdata(conWHSTotalSell, lArrayRowIndex)) / _
                                        (Varrsalesdata(conQuantity, lArrayRowIndex) - Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat)
                            Else
                                sRetailSell = ""
                            End If
                            If Varrsalesdata(conWHSQty, lArrayRowIndex) <> 0 Then
                                sWHSSell = Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) / Varrsalesdata(conWHSQty, lArrayRowIndex), gcon5DigitDollarFormat)
                                sWHSpercent = Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex) * 100 / Varrsalesdata(conQuantity, lArrayRowIndex)), "###")
                            Else
                                sWHSSell = ""
                                sWHSpercent = ""
                            End If
                            If chkProductReportTabDelimited Then
                                Print #intFileNum, _
                                    Varrsalesdata(conProduct, lArrayRowIndex); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    vbTab; _
                                    sRetailSell; _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    vbTab; _
                                    sWHSSell; _
                                    vbTab; _
                                    Format(Val(SafeDivide(pNumerator:=Varrsalesdata(conWHSTotalSell, lArrayRowIndex) * 100, _
                                                          pDenominator:=Varrsalesdata(conTotalSalesInc, lArrayRowIndex), _
                                                          pAnswerForZeroDividedByZero:="")), _
                                           "###"); _
                                    vbTab; _
                                    sWHSpercent
                                    
                            Else 'normal report
                            '
                                On Error Resume Next
                                Print #intFileNum, _
                                    Varrsalesdata(conProduct, lArrayRowIndex); _
                                    Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                    Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    sRetailSell; _
                                    Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conWHSQty) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                    Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    Tab(iArrTabStop(conWHSTotalSell) - Len(Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conWHSActualSell) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    sWHSSell; _
                                    Tab(iArrTabStop(conWHSAmntPercent) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) * 100 / Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), "###"); _
                                    Tab(iArrTabStop(conWHSQtyPercent) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    sWHSpercent
                            End If
                        End If
                    Next lArrayRowIndex
                    
                    'leave a gap
                    Print #intFileNum, gconSpace
                    
                    'expose the totals
                    If chkProductReportTabDelimited Then
                        Print #intFileNum, "Total"; _
                                   vbTab; _
                                   Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); _
                                   vbTab; _
                                   vbTab; _
                                   vbTab; _
                                   Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    Else 'normal report
                        Print #intFileNum, "Total"; _
                                   Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); _
                                   Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); _
                                   Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    End If
                    
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave a dual gap
                            Print #intFileNum, gconSpace
                            Print #intFileNum, gconSpace
                            'expose number of customers
                            If chkProductReportTabDelimited Then
                                Print #intFileNum, "Total customers"; _
                                           vbTab; _
                                           Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            Else 'normal report
                                Print #intFileNum, "Total customers"; _
                                           Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); _
                                           Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            End If
                        End If
                    End If
                    
                    Print #intFileNum, vbCrLf
                End If 'report destination
                
                On Error GoTo 0
                'conserve memory
                Erase iArrTabStop
                Erase Varrsalesdata
                Erase sArrBarcode
                
            Else 'no transactions for the date
                If optSendProductReportToDisplay Then
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                ElseIf optSendProductReportToPrinter Then
                    Printer.Print "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Printer.Print vbCrLf
                Else 'is destined for the file
                    Print #intFileNum, "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Print #intFileNum, vbCrLf
                End If
            End If 'any transactions for the report date ?
            
            With stb
                .SimpleText = ""
                .Refresh
            End With

        Next iCurrentFranchise
        
TidyUpnotSummarised:
        On Error GoTo 0
        
        If optSendProductReportToDisplay Then
            'do nothing
        ElseIf optSendProductReportToPrinter Then
            Printer.EndDoc
            MsgBox "Report was successfully submitted to the selected printer", _
                    vbInformation, gconReportManager
        Else 'was to file
            Close #intFileNum
            Call subSetProductReportViewButton
            MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & _
                   ". Use the 'View' button to display", _
                    vbInformation, gconReportManager
        End If
'--------------------------------------------------------------------------------------------------------------
'  AUrban SUMMARISED REPORT: Procedure is a candidate for splitting above and below here into two procedures
'--------------------------------------------------------------------------------------------------------------
    Else 'summarised
        If optProductReportOnSelectedFranchisesOnly(0) Then
            'build a query spec for all selected franchises
            For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
                If lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    
                    sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                              lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
                    
                    sIncludedFranchiseIDs = sIncludedFranchiseIDs & _
                                            gconLiveDataTableTSGFranchiseIDField & " = " & _
                                            fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex)) & " OR "
                End If
            Next lArrayRowIndex
            
            If iNumberOfFranchisesIncluded Then
                'get rid of the last delimiters
                sIncludedFranchiseNames = Left(sIncludedFranchiseNames, _
                                          Len(sIncludedFranchiseNames) - Len(", "))
                
                sIncludedFranchiseIDs = Left(sIncludedFranchiseIDs, _
                                        Len(sIncludedFranchiseIDs) - Len(" OR "))
            
                sFranchiseMessageBox = " for " & sIncludedFranchiseNames
            End If
            
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & _
                        " WHERE (" & sIncludedFranchiseIDs & ") " & _
                        "  AND (" & sIncludedDates & ") " & _
                        "  AND (Quantity <> 0) " & _
                          sWholesaleQuery
        
        Else 'all franchises option was selected
            iNumberOfFranchisesIncluded = lstProductReportsFranchiseBusinessName.ListCount
            
            sFranchiseMessageBox = ""
            
            sIncludedFranchiseNames = gconAllFranchises
            
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & _
                        " WHERE (" & sIncludedDates & ") " & _
                        "  AND (Quantity <> 0)"
        End If
        
        If iNumberOfFranchisesIncluded = 0 Then
            MsgBox "No franchise selected", vbExclamation, gconReportManager
            cmdAllItems.Caption = "All items"
            bRptlnProcess = False
            Exit Sub
        End If

        If iNumberOfFranchisesIncluded > 1 Then
            sPlural = "s"
        End If
        
        Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                              pSource:=sSQLQuery, _
                                                              pSourceType:=adCmdText, _
                                                              pErrMsg:=strErrMsg)
        
        If Not (rstDistinctBarcodesForTheReportingPeriod.BOF And _
                rstDistinctBarcodesForTheReportingPeriod.EOF) Then
            
            Call subWriteSizingArraysMessageToStatusBar
            
            'use the tabstop array to store the right justified position
            ReDim iArrTabStop(conDisplayAverageSalesInc To conWHSQtyPercent) As Integer
            'truncate if the docket printer is enabled in the defaults
            If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + _
                                           Len(gconTruncateCharacter) + _
                                           gconTruncateExtensionWidth + _
                                           Len(gconSpace) + _
                                           Len(gconStandardQuantityFormat)
                
                iArrTabStop(conNormalSell) = iArrTabStop(conQuantity) + _
                                                             Len(gcon5DigitDollarFormat) + _
                                                             Len(gconSpace) '
                iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conNormalSell) + _
                                                         Len(gcon5DigitDollarFormat) + _
                                                         Len(gconSpace)
                
                iArrTabStop(conTotalSalesInc) = iArrTabStop(conDisplayAverageSalesInc) + _
                                                Len(gcon5DigitDollarFormat) + _
                                                Len(gconSpace)
                iArrTabStop(conWHSQty) = iArrTabStop(conTotalSalesInc) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                iArrTabStop(conWHSTotalSell) = iArrTabStop(conWHSQty) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
            Else
                    iArrTabStop(conQuantity) = 43
                    iArrTabStop(conNormalSell) = 55 '
                    iArrTabStop(conDisplayAverageSalesInc) = 69
                    iArrTabStop(conTotalSalesInc) = 80 ' 69
                    iArrTabStop(conWHSQty) = 90
                    iArrTabStop(conWHSTotalSell) = 104
                    iArrTabStop(conWHSActualSell) = 116
                    iArrTabStop(conWHSAmntPercent) = 128
                    iArrTabStop(conWHSQtyPercent) = 138
            End If
            
            'size the array
            Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                If fbProductIsIncludedInThisProductReport(rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                    lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod + 1
                    ReDim Preserve sArrBarcode(lTotalNumberOfBarcodesForTheReportingPeriod) 'PAL
                    sArrBarcode(lTotalNumberOfBarcodesForTheReportingPeriod) = rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) 'PAL
                End If
                rstDistinctBarcodesForTheReportingPeriod.MoveNext
            Loop
            rstDistinctBarcodesForTheReportingPeriod.Close
            'ReDim udtSalesData(1 To lTotalNumberOfBarcodesForTheReportingPeriod) As gudtSalesData
            
            ReDim Varrsalesdata(conSortIndex To conWHSTotalSell, _
                                1 To lTotalNumberOfBarcodesForTheReportingPeriod) As Variant
            
            Call subWriteCollatingMessageToStatusBar
            
            cTotalSalesIncludingTaxForThePeriod = gconZeroValue
            lTotalItemsSoldForThePeriod = gconZeroValue
            lArrayRowIndex = gconZeroValue
            lTotalCustomersForThePeriod = gconZeroValue
            
            For BarcodeArrayIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod ' PALb50
                If optProductReportOnSelectedFranchisesOnly(0) Then
                    sSQLQuery = "SELECT * FROM LiveData " & vbNewLine & _
                                "WHERE (" & sIncludedFranchiseIDs & ") " & _
                                  "AND (" & sIncludedDates & ") " & _
                                  "AND (Barcode = " & SqlQ(sArrBarcode(BarcodeArrayIndex)) & ")" ' PALb50
                Else 'all franchises, no requirement to discriminate (performance reasons)
                    sSQLQuery = "SELECT * FROM LiveData " & _
                                "WHERE (" & sIncludedDates & ")" & _
                                " AND (Barcode = " & SqlQ(sArrBarcode(BarcodeArrayIndex)) & ")"
                End If
                
                lngRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
                Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                     pSource:=sSQLQuery, _
                                                                     pSourceType:=adCmdText, _
                                                                     pErrMsg:=strErrMsg)
                
                'has to be more than zero records, so don't waste time testing for it
                If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) = "TOTALCUSTOMERS" Then
                    Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                        lTotalCustomersForThePeriod = lTotalCustomersForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                        rstAllSameBarcodesForTheReportingPeriod.MoveNext
                    Loop
                Else
                    'If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                        lArrayRowIndex = lArrayRowIndex + 1
                        Call subDisplayCurrentRecordToUser( _
                             lArrayRowIndex, _
                             lTotalNumberOfBarcodesForTheReportingPeriod)
                        
                        If optDescription(0) Then
                            
                            'udtSalesData(lArrayRowIndex).Description = ""
                            
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                fsDescriptionFrom(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))
                        Else
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)
                        End If
                    
                        Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                            'avert an overflow divide by zero
                            If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField) <> 0 Then
                                Varrsalesdata(conQuantity, lArrayRowIndex) = _
                                    Varrsalesdata(conQuantity, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                
                                lTotalItemsSoldForThePeriod = _
                                    lTotalItemsSoldForThePeriod + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                
                                Varrsalesdata(conNormalSell, lArrayRowIndex) = _
                                    Varrsalesdata(conNormalSell, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod!NormalSellInc
                                    
                                 Varrsalesdata(conTotalSalesInc, lArrayRowIndex) = _
                                    Varrsalesdata(conTotalSalesInc, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                 
                                 Varrsalesdata(conWHSQty, lArrayRowIndex) = _
                                    Varrsalesdata(conWHSQty, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableWholesaleQty)
                                
                                 Varrsalesdata(conWHSTotalSell, lArrayRowIndex) = _
                                    Varrsalesdata(conWHSTotalSell, lArrayRowIndex) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableWholesaleActualSell)
                                
                                cTotalSalesIncludingTaxForThePeriod = _
                                    cTotalSalesIncludingTaxForThePeriod + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                            End If
                            
                            Varrsalesdata(conNormalSellcount, lArrayRowIndex) = lngRecCount
                            rstAllSameBarcodesForTheReportingPeriod.MoveNext
                        Loop
                    'End If
                End If
                rstAllSameBarcodesForTheReportingPeriod.Close
            Next BarcodeArrayIndex
            Set rstAllSameBarcodesForTheReportingPeriod = Nothing
            
            If lTotalCustomersForThePeriod <> gconZeroValue Then
                'don't want to sort this as an array component
                lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod - 1
            End If
            
            Call subWriteSortingMessageToStatusBar
            
            'sort by description
            Do
                bIndexSwapped = False
                For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod - 1
                    If (Varrsalesdata(conProduct, lArrayRowIndex) > _
                        Varrsalesdata(conProduct, lArrayRowIndex + 1)) Then 'swap
                        For iArrayColumnIndex = conProduct To conWHSTotalSell
                            vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                            Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                            Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                            bIndexSwapped = True
                        Next iArrayColumnIndex
                    End If
                Next lArrayRowIndex
            Loop While bIndexSwapped
            
            If optSendProductReportToDisplay Then
                For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                    If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                        Set gvListItem = lvwProductReport.ListItems.Add()
                        gvListItem.Text = Varrsalesdata(conProduct, lArrayRowIndex)
                        Call gsubAddSubItemToListview( _
                             Varrsalesdata(conQuantity, lArrayRowIndex), 1)
                        Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conNormalSell, lArrayRowIndex) / _
                                    Varrsalesdata(conNormalSellcount, lArrayRowIndex), gcon5DigitDollarFormat), 2)
                        ' only calculate avg. actual retail sell if retail qty (which is total - whs) is not zero
                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> Varrsalesdata(conWHSQty, lArrayRowIndex) Then
                            Call gsubAddSubItemToListview( _
                             Format((Varrsalesdata(conTotalSalesInc, lArrayRowIndex) - Varrsalesdata(conWHSTotalSell, lArrayRowIndex)) / _
                                    (Varrsalesdata(conQuantity, lArrayRowIndex) - Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat), 3)
                        End If
                        Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 4)
                        Call gsubAddSubItemToListview( _
                             Varrsalesdata(conWHSQty, lArrayRowIndex), 5)
                        Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex), gcon6DigitDollarFormat), 6)
                        If Varrsalesdata(conWHSQty, lArrayRowIndex) <> 0 Then
                            Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) / Varrsalesdata(conWHSQty, lArrayRowIndex), gcon5DigitDollarFormat), 7)
                            Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) * 100 / Varrsalesdata(conTotalSalesInc, lArrayRowIndex), "###"), 8)
                            Call gsubAddSubItemToListview( _
                              Format(Varrsalesdata(conWHSQty, lArrayRowIndex) * 100 / Varrsalesdata(conQuantity, lArrayRowIndex), "###"), 9)
                        End If
                    End If
                Next lArrayRowIndex
                
                'leave a gap then totals
                Set gvListItem = lvwProductReport.ListItems.Add()
                gvListItem.Text = gconSpace
                
                Set gvListItem = lvwProductReport.ListItems.Add()
                gvListItem.Text = "Total items"
                Call gsubAddSubItemToListview(lTotalItemsSoldForThePeriod, 1)
                Call gsubAddSubItemToListview(Format(cTotalSalesIncludingTaxForThePeriod, gcon6DigitDollarFormat), 4)
                
                If chkIncludeTotalCustomerCount Then
                    If lTotalCustomersForThePeriod <> gconZeroValue Then
                        'leave another gap then total customers
                        Set gvListItem = lvwProductReport.ListItems.Add()
                        gvListItem.Text = gconSpace
                        
                        Set gvListItem = lvwProductReport.ListItems.Add()
                        gvListItem.Text = "Total customers"
                        Call gsubAddSubItemToListview(lTotalCustomersForThePeriod, 1)
                    End If
                End If
            ElseIf optSendProductReportToPrinter Then
                ' PAL TODO - make changes fro Promo stuff for printer option
                On Error GoTo SummarisedPrinterErrorHandler
                cdlTSGDataWarehouse.ShowPrinter
                iNumberOfCopies = cdlTSGDataWarehouse.Copies
                Me.Refresh
                
                For iPageNumber = 1 To iNumberOfCopies
                    Printer.Print "Tobacco Station" & sPlural & _
                                  " - " & sIncludedFranchiseNames
                    'leave a dual gap
                    Printer.Print vbCrLf
        
                    Printer.Print conReportType & sReportingPeriod
                    Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                    'leave a dual gap
                    Printer.Print vbCrLf
                    
                    'headings
                    Printer.Print "Product"; _
                                    Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "  Qty"; _
                                   Tab(iArrTabStop(conNormalSell) - Len("Normal Sell")); _
                                  "Normal Sell"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Actual Sell")); _
                                  " Actual Sell"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "Tot (inc)"; _
                                  Tab(iArrTabStop(conWHSQty) - Len("NCS Qty")); _
                                  "NCS Qty"; _
                                  Tab(iArrTabStop(conWHSTotalSell) - Len("NCS Total)")); _
                                  "NCS Total"; _
                                  Tab(iArrTabStop(conWHSActualSell) - Len("NCS Sell)")); _
                                  "NCS Sell"; _
                                  Tab(iArrTabStop(conWHSAmntPercent) - Len("$NCS %)")); _
                                  "$NCS %"; _
                                  Tab(iArrTabStop(conWHSQtyPercent) - Len("NCS %)")); _
                                  "NCS %"
                    'leave a gap
                    Printer.Print gconSpace
                    
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            If Varrsalesdata(conQuantity, lArrayRowIndex) <> Varrsalesdata(conWHSQty, lArrayRowIndex) Then
                                sRetailSell = Format((Varrsalesdata(conTotalSalesInc, lArrayRowIndex) - Varrsalesdata(conWHSTotalSell, lArrayRowIndex)) / _
                                        (Varrsalesdata(conQuantity, lArrayRowIndex) - Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat)
                            Else
                                sRetailSell = ""
                            End If
                            If Varrsalesdata(conWHSQty, lArrayRowIndex) <> 0 Then
                                sWHSSell = Format(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) / Varrsalesdata(conWHSQty, lArrayRowIndex), gcon5DigitDollarFormat)
                            Else
                                sWHSSell = ""
                            End If
                            
                            Printer.Print _
                                Varrsalesdata(conProduct, lArrayRowIndex); _
                                Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                sRetailSell; _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conWHSQty) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conWHSTotalSell) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conWHSActualSell) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                sWHSSell; _
                                Tab(iArrTabStop(conWHSAmntPercent) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conWHSTotalSell, lArrayRowIndex) * 100 / Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), "###"); _
                                Tab(iArrTabStop(conWHSQtyPercent) - Len(Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conWHSQty, lArrayRowIndex) * 100 / Varrsalesdata(conQuantity, lArrayRowIndex)), "###")

                        End If
                    Next lArrayRowIndex
                    
                    'leave a gap
                    Printer.Print gconSpace
                    
                    'expose the totals
                    Printer.Print "Total"; _
                                Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); _
                                Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); _
                                Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave a dual gap
                            Printer.Print vbCrLf
                    
                            'expose number of customers
                            Printer.Print "Total customers"; _
                                           Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); _
                                           Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                        End If
                    End If
                Next iPageNumber
                Printer.EndDoc
                
                MsgBox "Report was successfully submitted to the selected printer", _
                        vbInformation, gconReportManager
            Else 'must be to file
                If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                    intFileNum = FreeFile   ' Get unused file
                    Open gsProductReportPathAndFilename For Output As #intFileNum
                    Print #intFileNum, "Tobacco Station" & sPlural & _
                              " - " & sIncludedFranchiseNames
                    'leave a dual gap
                    Print #intFileNum, vbCrLf
                    
                    Print #intFileNum, conReportType & sReportingPeriod
                    Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                    'leave a dual gap
                    Print #intFileNum, vbCrLf
                    
                    'headings
                    If chkProductReportTabDelimited Then
                        Print #intFileNum, "Product"; _
                                   vbTab; _
                                  "Qty"; _
                                   vbTab; _
                                  "Normal"; _
                                   vbTab; _
                                  "Avg unit"; _
                                   vbTab; _
                                  "Tot (inc)"
                    Else 'normal tabs
                        Print #intFileNum, "Product"; _
                                   Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "  Qty"; _
                                   Tab(iArrTabStop(conNormalSell) - Len("Normal")); _
                                  "Normal"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Promo")); _
                                  "Promo"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "  Tot (inc)"
                        Print #intFileNum, Tab(iArrTabStop(conNormalSell) - Len("Sell")); _
                                  "Sell"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Sell")); _
                                  "Sell"
                    End If
                    
                    'leave a gap
                    Print #intFileNum, gconSpace
                    
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            If chkProductReportTabDelimited Then
                            '
                                Print #intFileNum, _
                                    Varrsalesdata(conProduct, lArrayRowIndex); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            Else 'normal report
                            '
                                Print #intFileNum, _
                                    Varrsalesdata(conProduct, lArrayRowIndex); _
                                    Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                    Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            End If
                        End If
                    Next lArrayRowIndex
                    
                    'leave a gap
                    Print #intFileNum, gconSpace
                        
                    'expose the totals
                    If chkProductReportTabDelimited Then
                        Print #intFileNum, "Total"; _
                                   vbTab; _
                                   Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); _
                                   vbTab; _
                                   vbTab; _
                                   vbTab; _
                                   Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    Else 'normal report
                        Print #intFileNum, "Total"; _
                                   Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); _
                                   Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); _
                                   Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    End If
                    
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave a dual gap
                            Print #intFileNum, gconSpace
                            Print #intFileNum, gconSpace
                            'expose number of customers
                            If chkProductReportTabDelimited Then
                                Print #intFileNum, "Total customers"; _
                                           vbTab; _
                                           Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            Else 'normal report
                                Print #intFileNum, "Total customers"; _
                                           Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); _
                                           Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            End If
                        End If
                    End If
                    Close #intFileNum
                    
                    Call subSetProductReportViewButton
                    
                    MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & _
                           ". Use the 'View' button to display", _
                            vbInformation, gconReportManager
                Else 'environment was not created
                    MsgBox "Report was aborted", vbExclamation
                End If 'environement created ?
            End If 'report destination

TidyUpSummarised:
            On Error GoTo 0
            
            'conserve memory
            Erase iArrTabStop
            Erase Varrsalesdata
            Erase sArrBarcode
            
        Else 'no transactions for the date
            MsgBox "No sales transactions" & sFranchiseMessageBox & gsReportPeriodWording & sReportingPeriod, _
                    vbInformation, gconReportManager
        End If 'any transactions for the report date ?
    
        With stb
            .SimpleText = ""
            .Refresh
        End With
   End If 'summarised or notsummarised ?
      
    cmdAllItems.Caption = "All items"
    bRptlnProcess = False
    Exit Sub

SummarisedPrinterErrorHandler:
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    Resume TidyUpSummarised

End Sub


Private Sub cmdAztecUpload_Click()
    Me.Enabled = False
    UploadLatestAztecRpts
    subRefreshAztecUploadsGrid
    Me.Enabled = True
End Sub

Private Sub cmdBataTabExportGrid_Click()
Dim strFileName As String
Dim strTimeStamp As String

    Me.Enabled = False

    strTimeStamp = "TxDate " & Format$(dtpBataTabTxDate.Value, "ddmmmyy")
    Select Case True
        Case optBataProcessed(0): strFileName = strTimeStamp & " Processed" ' Processed (Uploaded or Disk File Created)
        Case optBataProcessed(1): strFileName = strTimeStamp & " Missing"   ' Missing
        Case optBataProcessed(2): strFileName = strTimeStamp & " ALL Frans" ' All Bata Franchises
    End Select
    
    strFileName = fdlgCommon.GetFullFileName(pMethod:=eShowSave, _
                                             pFilename:=strFileName & ".xls", _
                                             pFilter:="xls", _
                                             pFilterDescription:="Excel", _
                                             pDefaultExtension:="xls")
    
    If Len(strFileName) Then
        grdBataRpts.SaveGrid FileName:=strFileName, SaveWhat:=flexFileExcel, Options:=flexXLSaveFixedCells
'       If MsgBox("Open saved file?", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
'           OpenFile strFileName ' Can't find generic way of opening diff versions of Excel across diff versions of Windows
'       End If
    End If

    Me.Enabled = True

End Sub

Private Sub cmdBataTabPrintGrid_Click()
Dim strDocName As String
Dim strTimeStamp As String

    Me.Enabled = False

    strTimeStamp = "Transaction Date " & Format$(dtpBataTabTxDate.Value, "d mmmm") & _
                   " (Printed " & Format$(Now, "d mmmm hh:nn am/pm") & ")"
    Select Case True
        Case optBataProcessed(0): strDocName = "Franchises Processed - " & strTimeStamp     ' Processed (Uploaded or Disk File Created)
        Case optBataProcessed(1): strDocName = "Franchises NOT Processed - " & strTimeStamp ' NOT Processed
        Case optBataProcessed(2): strDocName = "ALL Franchises - " & strTimeStamp           ' All Bata Franchises
    End Select
                   
    With grdBataRpts
        If .Rows = .FixedRows Then
            MsgBox "There is nothing to print.", vbInformation
        Else
            On Error Resume Next
            '   Handle Run-time error when you call PrintGrid method, and user selects a printer
            '   configured as a PrintFile but cancels the input form for the Output File Name. (Win 2000)
                .PrintGrid DocName:=(strDocName), ShowDialog:=True  ' NB strDocName MUST be passed in parenthesis
                                                                    ' otherwise its contents are not printed
                                                                    ' (ie must be passed as a literal, not a variable)
            On Error GoTo 0
        End If
    End With
    
    Me.Enabled = True
    
End Sub

Private Sub cmdBataTabProcessSelected_Click()
   frmProcessBataRpts.Show Modal:=vbModal, OwnerForm:=Me    ' should perhaps have an OpenF...() fn which
                                                            ' returns if rpeorts were procesed so grdBataRpts
                                                            ' could be conditionally refreshed
   RefreshBataTabGrid
   Me.SetFocus
End Sub

Private Sub cmdBataTabProcessUnProcessed_Click()
Dim intPrevMousePointer As Integer
Dim strMsg As String

    strMsg = "Process all un-processed data?"
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Me.Enabled = False
        intPrevMousePointer = SetMousePointer(vbHourglass)
        g.bCaptureCycleRunning = True   ''' Review PERHAPS A GOOD IDEA TO RENAME THIS GLOBAL VARIABLE TO SOMETHING LIKE SuspendAutoCaptureCycle OR SOMETHING MORE INSTRUCTIVE GIVEN THE USE OF THE VARIABLE HERE
        
        StatusBar pMsg:="Manually processing un-processed BATA reports", pLog:=True, pRefreshEventLogDisplay:=True
        
        Set moBataRpts = New clsBataRpts
        moBataRpts.Process pAddUnProcessed:=True
        Set moBataRpts = Nothing
        
        gsubRefreshEventLogDisplay
        
        g.bCaptureCycleRunning = False  ''' Review PERHAPS A GOOD IDEA TO RENAME THIS GLOBAL VARIABLE TO SOMETHING LIKE SuspendAutoCaptureCycle OR SOMETHING MORE INSTRUCTIVE GIVEN THE USE OF THE VARIABLE HERE
        SetMousePointer intPrevMousePointer
        Me.Enabled = True
    End If

 End Sub

Private Sub cmdBataTabSaveSelected_Click()
Dim intPrevMousePointer As Integer
Dim strLocalPath As String
Dim fso As Scripting.FileSystemObject
Dim oRpt As clsBataRpt
    
    Me.Enabled = False
    intPrevMousePointer = SetMousePointer(vbHourglass)
    
    MsgBox "Reports will be saved in a sub folder of " & g.strBataRptsFolder
    
    Set fso = New Scripting.FileSystemObject
    strLocalPath = g.strBataRptsFolder & "\" & Format$(Now, "yyyy-mm-dd hh nn ss am/pm")
    If Not fso.FolderExists(strLocalPath) Then
        fso.CreateFolder (strLocalPath)
    End If
    
    Set moBataRpts = New clsBataRpts
    AddSelRptsFromGrid pBataRpts:=moBataRpts
    StatusBar pMsg:="", pLog:=False
    For Each oRpt In moBataRpts
    '   Use Shell directly as subOpenFile somehow denies permission to delete
    '   parent folder once all reports are deleted in clsRpts terminate event
        fso.CopyFile Source:=oRpt.FullName, Destination:=strLocalPath & "\" & oRpt.Name, OverWriteFiles:=False
    Next oRpt
    
    Set oRpt = Nothing
    Set moBataRpts = Nothing
    Set fso = Nothing
    
    gsubRefreshEventLogDisplay
    
    GridClearSelections grdBataRpts
    
    SetMousePointer intPrevMousePointer
    Me.Enabled = True
    
End Sub

Private Sub cmdBataTabViewSelected_Click()
Dim intPrevMousePointer As Integer
Dim oRpt As clsBataRpt

    Me.Enabled = False
    intPrevMousePointer = SetMousePointer(vbHourglass)
    
    Set moBataRpts = New clsBataRpts
    AddSelRptsFromGrid pBataRpts:=moBataRpts
    StatusBar pMsg:="", pLog:=False
    For Each oRpt In moBataRpts
    '   Use Shell directly as subOpenFile somehow denies permission to delete
    '   parent folder once all reports are deleted in clsRpts terminate event
        Shell "Notepad " & DQ(oRpt.FullName), vbNormalNoFocus
    Next oRpt
    Set oRpt = Nothing
    Set moBataRpts = Nothing
    
    gsubRefreshEventLogDisplay
    
    GridClearSelections grdBataRpts
    
    SetMousePointer intPrevMousePointer
    Me.Enabled = True
    
End Sub

Private Sub cmdBrowseUploads_Click()
    Dim sNewUploadFile As String
    
    On Error GoTo uploadsCancel
    
    cdlTSGDataWarehouse.ShowOpen

    If Dir(cdlTSGDataWarehouse.FileName) <> "" Then
        sNewUploadFile = fGetLastWord(cdlTSGDataWarehouse.FileName, "\")
        ' delete the existing file in the 'uploads' holding area if it exists
        DeleteFile g.strUploadsFolder & "\" & sNewUploadFile
        FileCopy cdlTSGDataWarehouse.FileName, g.strUploadsFolder & "\" & sNewUploadFile
        lstUploadItemList.AddItem g.strUploadsFolder & "\" & sNewUploadFile
    End If
uploadsCancel:

End Sub

Private Sub cmdCaptureData_Click()
Dim strMsg As String

    strMsg = "Capture data for all franchises?"
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        If Not g.bCaptureCycleRunning Then
            subCaptureData pAutoCaptureCycle:=False
        End If
    End If
    
End Sub

Private Sub cmdCaptureSelected_Click()
Dim colFranIDs As VBA.Collection
Dim colFranNames As VBA.Collection
Dim oCaptureOptions As clsDataCaptureOptions

    Set colFranIDs = ListBoxGetCollection(pListBox:=lstDataCaptureFranchiseBusinessName, _
                                          pItemData:=True, _
                                          pSelected:=True)
    If colFranIDs.Count Then
        Set colFranNames = ListBoxGetCollection(pListBox:=lstDataCaptureFranchiseBusinessName, _
                                                pItemData:=False, _
                                                pSelected:=True)
        
        Set oCaptureOptions = fdlgGetCaptureOptions.GetOptions(pDataCapture:=eCaptureSelected, _
                                                               pColFranNames:=colFranNames)
                                                               
    '   Debug.Print "CancelCapture: " & oCaptureOptions.CancelCapture, _
                    "UpdateNonCompliants: " & oCaptureOptions.UpdateNonCompliants

        Set fdlgGetCaptureOptions = Nothing
        
        If Not oCaptureOptions.CancelCapture Then
            subCaptureData pAutoCaptureCycle:=False, _
                           pColSelFranIDs:=colFranIDs, _
                           pCaptureOptions:=oCaptureOptions
        End If
        
        Set oCaptureOptions = Nothing
            
    End If
    
    Set colFranIDs = Nothing

End Sub

Private Sub cmdCloseSelectedFranchises_Click()
Dim intPrevMousePointer As Integer
Dim strMsg As String
Dim strErrMsg As String
Dim vntFranID As Variant
Dim vntFranName As Variant
Dim colFranIDs As VBA.Collection
Dim colFranNames As VBA.Collection
Dim rst As ADODB.Recordset

    Set colFranIDs = ListBoxGetCollection(pListBox:=lstDataCaptureFranchiseBusinessName, _
                                          pItemData:=True, _
                                          pSelected:=True)
    If colFranIDs.Count Then
        Set colFranNames = ListBoxGetCollection(pListBox:=lstDataCaptureFranchiseBusinessName, _
                                                pItemData:=False, _
                                                pSelected:=True)
        For Each vntFranName In colFranNames
            strMsg = strMsg & vbNewLine & vntFranName
        Next vntFranName
        Set colFranNames = Nothing
        
        strMsg = "Close the following franchise(s)" & "?" & strMsg
        If MsgBox(Prompt:=strMsg, Buttons:=vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Me.Enabled = False
            intPrevMousePointer = SetMousePointer(pMousePointer:=vbHourglass)
            
            Set rst = GetRst(pCnn:=g.cnnDW, _
                             pSource:="Franchises", _
                             pSourceType:=adCmdTable, _
                             pRstType:=eEditableDynamic, _
                             pErrMsg:=strErrMsg)
            If Not rst Is Nothing Then
                If Not (rst.BOF And rst.EOF) Then
                    With rst
                        For Each vntFranID In colFranIDs
                            .MoveFirst
                            .Find "FranchiseIDTSG = " & vntFranID
                            If Not .EOF Then
                                .Fields!Live.Value = CBoolMySql(False)
                                .Update
                            End If
                        Next vntFranID
                        .Close
                    End With
                    Set rst = Nothing
                End If
            End If
            
            subPopulateFranchiseBusinessNameListBoxes
            
            SetMousePointer pMousePointer:=intPrevMousePointer
            Me.Enabled = True
        End If
    
    End If
    
    Set colFranIDs = Nothing

End Sub

Private Sub cmdCreateNielsenReports_Click()
    
    CreateNielsenReports pLastReportEndDate:=dtpNielsenRptTxDate.Value, pCalledAutomatically:=False

    MsgBox stb.SimpleText, vbInformation
    
    subRefreshNielsenReportListBox dtpNielsenRptTxDate.Value

    cmdCreateNielsenReports.Enabled = True

End Sub

Private Sub cmdDisplayDialupResults_Click()
Static strOriginalCaption As String
Dim intPrevMousePointer As Integer

    intPrevMousePointer = SetMousePointer(vbHourglass)
    
    If Len(strOriginalCaption) = 0 Then
        strOriginalCaption = cmdDisplayDialupResults.Caption
    End If
    
    If InStr(UCase$(cmdDisplayDialupResults.Caption), "DIAL") Then
    '   User wants to see dialup results -> display dialup results and toggle button
        DialupResults fPrint:=False ' false means display in eventlog window
        optDialupResults(0).Enabled = True
        optDialupResults(1).Enabled = True
        lblEventLog.Caption = "Dialup Results"
        cmdDisplayDialupResults.Caption = "Show Event Log"
    Else
    '   User wants to revert to event log -> display dialup results and toggle button
        gsubRefreshEventLogDisplay
        optDialupResults(0).Enabled = False
        optDialupResults(1).Enabled = False
        lblEventLog.Caption = "Event Log"
        cmdDisplayDialupResults.Caption = strOriginalCaption
    End If
    
    SetMousePointer intPrevMousePointer
    
End Sub

Private Sub cmdImportBatscanFiles_Click()
    
    Me.Enabled = False

    ImportBatScanFiles pManualImport:=True
'   AUrban Version 3.0.9009
'   Don't transfer records until Capture All Franchises because you may import a number of
'   different files and be able to update the records. Once a barcode/franchise combo is in
'   livedata it will never be updated by TfrAllPreLiveDataToLiveData ... unless I modify it
'   TfrAllPreLiveDataToLiveData

    Me.Enabled = True

End Sub

Private Sub cmdMarketShare_Click()
'--------------------------------------------------------------------------------------------------------------
'  AUrban Procedure is a candidate for splitting into two procedures (Summarised Rpt and Not Summarised Rpt)
'  AUrban (cmdAllItems_Click, cmdMarketShare_Click & cmdStickReport_Click are similar. Prob cut & pasted and modified)
'--------------------------------------------------------------------------------------------------------------

    Dim cTotalSalesIncludingTaxForThePeriod As Currency
    
    Dim iMaximumSuppliers As Integer, _
        iNumberOfCopies As Integer, _
        iNumberOfFranchisesIncluded As Integer, _
        iPageNumber As Integer
    
    Dim lArrayRowIndex As Long
    Dim lCurrentProductDisplayIndex As Long
    Dim lTotalNumberOfBarcodesForTheReportingPeriod As Long

Const kSqlSupplier As String = "SELECT MAX(supplier_id) FROM Supplier WHERE supplier_id <> 0"
Dim strErrMsg As String
Dim rstDistinctBarcodesForTheReportingPeriod As ADODB.Recordset '!!! ManualFix Clearing: Object variable not cleared: rstDistinctBarcodesForTheReportingPeriod
Dim rstAllSameBarcodesForTheReportingPeriod As ADODB.Recordset  '!!! ManualFix Clearing: Object variable not cleared: rstAllSameBarcodesForTheReportingPeriod
    
    Dim sFranchiseMessageBox As String, _
        sIncludedDates As String, _
        sIncludedFranchiseIDs As String, _
        sIncludedFranchiseNames As String, _
        sPlural As String
    Dim sReportingPeriod As String
    Dim sSQLQuery As String

    Dim intFileNum As Integer

    'data and tabstop arrays
    Const conPercentageMarketShare = 2, conTotalSalesInc = 3
    
    Const conReportType = "Sales by market share for "
    
Dim datReportStart As Date

    With lvwProductReport
        .ListItems.Clear
        .Refresh
    End With
    
    If Not IsDateFmtOk() Then   ''' Review Fix Reliance on date format when time permits
        MsgBox "incorrect system date format"
        Exit Sub
    End If
    
    cmdMarketShare.Enabled = False

    Call subWriteSearchingMessageToStatusBar
        
    'build a query spec for all dates within the range
    datReportStart = GetDateFrom_ddmmmyy(lblProductReportStartDate)
    If lblProductReportStartDate = lblProductReportFinishDate Then
        sIncludedDates = "TransactionDate = " & MySqlDate(datReportStart)
        sReportingPeriod = lblProductReportStartDate
    Else
        sIncludedDates = "TransactionDate BETWEEN " & _
                         MySqlDate(datReportStart) & " AND " & MySqlDate(GetDateFrom_ddmmmyy(lblProductReportFinishDate))
        sReportingPeriod = lblProductReportStartDate & " to " & lblProductReportFinishDate
    End If
    
    If optProductReportNotSummarised(0) Then
        ReDim lFranchiseID(gconZeroValue) As Long
            'build an array containing the ID for each selected franchise
        For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
            If (lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) And optProductReportOnSelectedFranchisesOnly(0)) Or _
                (Not optProductReportOnSelectedFranchisesOnly(0)) Then
                iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)
                lFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex))
                sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                          lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
            End If
        Next lArrayRowIndex
        
        If iNumberOfFranchisesIncluded = 0 Then
            MsgBox "No franchise selected", vbExclamation, gconReportManager
            cmdMarketShare.Enabled = True
            Exit Sub
        End If
        
        If optProductReportOnSelectedFranchisesOnly(0) Then
            'get rid of the last delimiters
            sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
            sFranchiseMessageBox = " for " & sIncludedFranchiseNames
        Else
            sIncludedFranchiseNames = gconAllFranchises
            sFranchiseMessageBox = ""
        End If
        
        If iNumberOfFranchisesIncluded > 1 Then
            sPlural = "s"
        End If
                
            
        Dim iCurrentFranchise As Integer
                                    
        If optSendProductReportToPrinter Then
            On Error GoTo NotSummarisedPrinterErrorHandler
            cdlTSGDataWarehouse.ShowPrinter
            Me.Refresh
            
            Printer.Print "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
            'leave a dual gap
            Printer.Print vbCrLf

            Printer.Print conReportType & sReportingPeriod
            Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
            'leave a dual gap
            Printer.Print vbCrLf
        ElseIf optSendProductReportToFile Then
            If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                intFileNum = FreeFile   ' Get unused file
                Open gsProductReportPathAndFilename For Output As #intFileNum
                Print #intFileNum, "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
                'leave a dual gap
                Print #intFileNum, vbCrLf
            
                Print #intFileNum, conReportType & sReportingPeriod
                Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                'leave a dual gap
                Print #intFileNum, vbCrLf
            Else 'environment was not created
                MsgBox "Report was aborted", vbExclamation
                GoTo TidyUpnotSummarised
            End If 'environement created ?
        End If 'sent prod report to file
                                    
        For iCurrentFranchise = 1 To iNumberOfFranchisesIncluded
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & vbNewLine & _
                        "WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & _
                         " AND (" & sIncludedDates & ")"
                                                         
            lTotalNumberOfBarcodesForTheReportingPeriod = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
            If lTotalNumberOfBarcodesForTheReportingPeriod Then
                Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                      pSource:=sSQLQuery, _
                                                                      pSourceType:=adCmdText, _
                                                                      pErrMsg:=strErrMsg) 'required for movelast etc...
                Call subWriteSizingArraysMessageToStatusBar
                
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conPercentageMarketShare To conTotalSalesInc) As Integer
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conPercentageMarketShare) = gconTruncateDescriptionBriefAt + _
                                                            Len(gconTruncateCharacter) + _
                                                            gconTruncateExtensionWidth + _
                                                            Len(gconSpace) + _
                                                            Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conPercentageMarketShare) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                Else
                    iArrTabStop(conPercentageMarketShare) = 48
                    iArrTabStop(conTotalSalesInc) = 69
                End If
                
                iMaximumSuppliers = GetRstVal(pCnn:=g.cnnDW, pSource:=kSqlSupplier)

                ReDim cArrSalesData(conTotalSalesInc, 1 To iMaximumSuppliers) As Currency
                
                Call subWriteCollatingMessageToStatusBar
                
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                lCurrentProductDisplayIndex = gconZeroValue
                
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    'If optProductReportOnSelectedFranchisesOnly(0) Then
                        sSQLQuery = "SELECT * FROM LiveData " & _
                                    " WHERE (FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & ") " & _
                                     " AND (" & sIncludedDates & ") " & _
                                     " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod!Barcode) & ")"
                    Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                         pSource:=sSQLQuery, _
                                                                         pSourceType:=adCmdText, _
                                                                         pErrMsg:=strErrMsg)
                    
                    'has to be more than zero records, so don't waste time testing for it
                    If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) <> "TOTALCUSTOMERS" Then
                        
                        lCurrentProductDisplayIndex = lCurrentProductDisplayIndex + 1
                        Call subDisplayCurrentRecordToUser( _
                             lCurrentProductDisplayIndex, _
                             lTotalNumberOfBarcodesForTheReportingPeriod)
                        
                        If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                            Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                                'add the value sold to the existing value, using the suplier id to identify the row
                                cArrSalesData(conTotalSalesInc, fiSupplierIDForBarcode(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))) = _
                                    cArrSalesData(conTotalSalesInc, fiSupplierIDForBarcode(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                    
                                cTotalSalesIncludingTaxForThePeriod = cTotalSalesIncludingTaxForThePeriod + _
                                                                      rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                rstAllSameBarcodesForTheReportingPeriod.MoveNext
                            Loop
                        End If
                    End If
                    rstAllSameBarcodesForTheReportingPeriod.Close
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
                
                If optSendProductReportToDisplay Then
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise))
                    For lArrayRowIndex = 1 To iMaximumSuppliers
                        If cArrSalesData(conTotalSalesInc, lArrayRowIndex) > gconZeroValue Then
                            Set gvListItem = lvwProductReport.ListItems.Add()
                            gvListItem.Text = fsSupplierNameFrom(lArrayRowIndex)
                            Call gsubAddSubItemToListview( _
                                 Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"), 2)
                            
                            Call gsubAddSubItemToListview( _
                                 Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 3)
                        End If
                    Next lArrayRowIndex
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                    Me.Refresh
                ElseIf optSendProductReportToPrinter Then
                    Printer.Print GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    'leave a gap
                    Printer.Print gconSpace
                    
                    'headings
                    Printer.Print "Supplier"; _
                                   Tab(iArrTabStop(conPercentageMarketShare) - Len("  %")); _
                                  "  %"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "Tot (inc)"
                    'leave a gap
                    Printer.Print gconSpace
                    
                    For lArrayRowIndex = 1 To iMaximumSuppliers
                        If cArrSalesData(conTotalSalesInc, lArrayRowIndex) > gconZeroValue Then
                            Printer.Print _
                                fsSupplierNameFrom(lArrayRowIndex); _
                                Tab(iArrTabStop(conPercentageMarketShare) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"))); _
                                Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"); _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat))); _
                                Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat)
                        End If
                    Next lArrayRowIndex
                    'leave a dual gap
                    Printer.Print vbCrLf
                Else 'must be to file
                    Print #intFileNum, GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    'leave a gap
                    Print #intFileNum, gconSpace
                
                    'headings
                    If chkProductReportTabDelimited Then
                        Print #intFileNum, "Supplier"; _
                                   vbTab; _
                                  "%"; _
                                   vbTab; _
                                  "Tot (inc)"
                    Else 'normal report
                        Print #intFileNum, "Supplier"; _
                                   Tab(iArrTabStop(conPercentageMarketShare) - Len("  %")); _
                                  "  %"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "Tot (inc)"
                    End If 'tab delimited ?
                    
                    'leave a gap
                    Print #intFileNum, gconSpace
                    
                    For lArrayRowIndex = 1 To iMaximumSuppliers
                        If cArrSalesData(conTotalSalesInc, lArrayRowIndex) > gconZeroValue Then
                            If chkProductReportTabDelimited Then
                                Print #intFileNum, _
                                    fsSupplierNameFrom(lArrayRowIndex); _
                                    vbTab; _
                                    Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"); _
                                    vbTab; _
                                    Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat)
                            Else 'normal report
                                Print #intFileNum, _
                                    fsSupplierNameFrom(lArrayRowIndex); _
                                    Tab(iArrTabStop(conPercentageMarketShare) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"))); _
                                    Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"); _
                                    Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat))); _
                                    Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat)
                            End If 'tab delimited
                        End If '>0 sales
                    Next lArrayRowIndex
                    Print #intFileNum, vbCrLf
                End If 'report destination
                
                On Error GoTo 0
                'conserve memory
                Erase iArrTabStop
                Erase cArrSalesData
                
                rstDistinctBarcodesForTheReportingPeriod.Close
                
            Else 'no transactions for the date
                If optSendProductReportToDisplay Then
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                ElseIf optSendProductReportToPrinter Then
                    Printer.Print "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Printer.Print vbCrLf
                Else 'is destined for the file
                    Print #intFileNum, "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Print #intFileNum, vbCrLf
                End If
            End If 'any transactions for the report date ?
            
            With stb
                .SimpleText = ""
                .Refresh
            End With

        Next iCurrentFranchise
        
TidyUpnotSummarised:
        On Error GoTo 0
        
        If optSendProductReportToDisplay Then
            'do nothing
        ElseIf optSendProductReportToPrinter Then
            Printer.EndDoc
            MsgBox "Report was successfully submitted to the selected printer", _
                    vbInformation, gconReportManager
        Else 'was to file
            Close #intFileNum
            Call subSetProductReportViewButton
            MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & _
                   ". Use the 'View' button to display", _
                    vbInformation, gconReportManager
        End If
'--------------------------------------------------------------------------------------------------------------
'  AUrban SUMMARISED REPORT: Procedure is a candidate for splitting above and below here into two procedures
'--------------------------------------------------------------------------------------------------------------
    Else 'summarised
        If optProductReportOnSelectedFranchisesOnly(0) Then
            'build a query spec for all selected franchises
            For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
                If lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    
                    sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                              lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
                    
                    sIncludedFranchiseIDs = sIncludedFranchiseIDs & _
                                            gconLiveDataTableTSGFranchiseIDField & " = " & _
                                            fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex)) & " OR "
                End If
            Next lArrayRowIndex
            
            If iNumberOfFranchisesIncluded Then
                'get rid of the last delimiters
                sIncludedFranchiseNames = Left(sIncludedFranchiseNames, _
                                          Len(sIncludedFranchiseNames) - Len(", "))
                
                sIncludedFranchiseIDs = Left(sIncludedFranchiseIDs, _
                                        Len(sIncludedFranchiseIDs) - Len(" OR "))
            
                sFranchiseMessageBox = " for " & sIncludedFranchiseNames
            End If
            
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & _
                        " WHERE (" & sIncludedFranchiseIDs & ") AND (" & sIncludedDates & ")"
        Else 'all franchises option was selected
            iNumberOfFranchisesIncluded = lstProductReportsFranchiseBusinessName.ListCount
            
            sFranchiseMessageBox = ""
            
            sIncludedFranchiseNames = gconAllFranchises
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData WHERE " & sIncludedDates
        End If
        
        If iNumberOfFranchisesIncluded Then 'franchises are included
            If iNumberOfFranchisesIncluded > 1 Then
                sPlural = "s"
            End If
            
            lTotalNumberOfBarcodesForTheReportingPeriod = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
            If lTotalNumberOfBarcodesForTheReportingPeriod Then
                Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                      pSource:=sSQLQuery, _
                                                                      pSourceType:=adCmdText, _
                                                                      pErrMsg:=strErrMsg)
                
                Call subWriteSizingArraysMessageToStatusBar
                
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conPercentageMarketShare To conTotalSalesInc) As Integer
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conPercentageMarketShare) = gconTruncateDescriptionBriefAt + _
                                                            Len(gconTruncateCharacter) + _
                                                            gconTruncateExtensionWidth + _
                                                            Len(gconSpace) + _
                                                            Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conPercentageMarketShare) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                Else
                    iArrTabStop(conPercentageMarketShare) = 48
                    iArrTabStop(conTotalSalesInc) = 69
                End If
                
                iMaximumSuppliers = GetRstVal(pCnn:=g.cnnDW, pSource:=kSqlSupplier)
                
                ReDim cArrSalesData(conTotalSalesInc, _
                                    1 To iMaximumSuppliers) As Currency
                
                Call subWriteCollatingMessageToStatusBar
                
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                lCurrentProductDisplayIndex = gconZeroValue
                
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    If optProductReportOnSelectedFranchisesOnly(0) Then
                        sSQLQuery = "SELECT * FROM LiveData " & _
                                    "WHERE (" & sIncludedFranchiseIDs & ") " & _
                                     " AND (" & sIncludedDates & ") " & _
                                     " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod!Barcode) & ")"
                    Else 'all franchises, no requirement to discriminate (performance reasons)
                        sSQLQuery = "SELECT * FROM LiveData " & _
                                    "WHERE (" & sIncludedDates & ") " & _
                                     " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod!Barcode) & ")"
                    End If
 
                    Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                         pSource:=sSQLQuery, _
                                                                         pSourceType:=adCmdText, _
                                                                         pErrMsg:=strErrMsg)
                    
                    'has to be more than zero records, so don't waste time testing for it
                    If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) <> "TOTALCUSTOMERS" Then
                        
                        lCurrentProductDisplayIndex = lCurrentProductDisplayIndex + 1
                        Call subDisplayCurrentRecordToUser( _
                             lCurrentProductDisplayIndex, _
                             lTotalNumberOfBarcodesForTheReportingPeriod)

                        
                        If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                            Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                                'add the value sold to the existing value, using the suplier id to identify the row
                                cArrSalesData(conTotalSalesInc, fiSupplierIDForBarcode(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))) = _
                                    cArrSalesData(conTotalSalesInc, fiSupplierIDForBarcode(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))) + _
                                    rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                    
                                cTotalSalesIncludingTaxForThePeriod = cTotalSalesIncludingTaxForThePeriod + _
                                                                      rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                                                
                                rstAllSameBarcodesForTheReportingPeriod.MoveNext
                            Loop
                        End If
                    End If
                    rstAllSameBarcodesForTheReportingPeriod.Close
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
                
                If optSendProductReportToDisplay Then
                    For lArrayRowIndex = 1 To iMaximumSuppliers
                        If cArrSalesData(conTotalSalesInc, lArrayRowIndex) > gconZeroValue Then
                            Set gvListItem = lvwProductReport.ListItems.Add()
                            gvListItem.Text = fsSupplierNameFrom(lArrayRowIndex)
                            Call gsubAddSubItemToListview( _
                                 Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"), 2)
                            
                            Call gsubAddSubItemToListview( _
                                 Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 3)
                        End If
                    Next lArrayRowIndex
                ElseIf optSendProductReportToPrinter Then
                    On Error GoTo NotSummarisedPrinterErrorHandler
                    cdlTSGDataWarehouse.ShowPrinter
                    iNumberOfCopies = cdlTSGDataWarehouse.Copies
                    Me.Refresh
                    
                    For iPageNumber = 1 To iNumberOfCopies
                        Printer.Print "Tobacco Station" & sPlural & _
                                      " - " & sIncludedFranchiseNames
                        'leave a dual gap
                        Printer.Print vbCrLf
            
                        Printer.Print conReportType & sReportingPeriod
                        Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                        'leave a dual gap
                        Printer.Print vbCrLf
                        
                        'headings
                        Printer.Print "Supplier"; _
                                       Tab(iArrTabStop(conPercentageMarketShare) - Len("  %")); _
                                      "  %"; _
                                       Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                      "Tot (inc)"
                        'leave a gap
                        Printer.Print gconSpace
                        
                        For lArrayRowIndex = 1 To iMaximumSuppliers
                            If cArrSalesData(conTotalSalesInc, lArrayRowIndex) > gconZeroValue Then
                                Printer.Print _
                                    fsSupplierNameFrom(lArrayRowIndex); _
                                    Tab(iArrTabStop(conPercentageMarketShare) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"))); _
                                    Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"); _
                                    Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat))); _
                                    Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat)
                            End If
                        Next lArrayRowIndex
                    Next iPageNumber
                    Printer.EndDoc
                    
                    MsgBox "Report was successfully submitted to the selected printer", _
                            vbInformation, gconReportManager
                Else 'must be to file
                    If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                        intFileNum = FreeFile   ' Get unused file
                        Open gsProductReportPathAndFilename For Output As #intFileNum
                        Print #intFileNum, "Tobacco Station" & sPlural & _
                                  " - " & sIncludedFranchiseNames
                        'leave a dual gap
                        Print #intFileNum, vbCrLf
                        
                        Print #intFileNum, conReportType & sReportingPeriod
                        Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                        'leave a dual gap
                        Print #intFileNum, vbCrLf
                        
                        
                        'headings
                        If chkProductReportTabDelimited Then
                            Print #intFileNum, "Supplier"; _
                                       vbTab; _
                                      "%"; _
                                       vbTab; _
                                      "Tot (inc)"
                        Else 'normal report
                            Print #intFileNum, "Supplier"; _
                                       Tab(iArrTabStop(conPercentageMarketShare) - Len("  %")); _
                                      "  %"; _
                                       Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                      "Tot (inc)"
                        End If
                        
                        'leave a gap
                        Print #intFileNum, gconSpace
                        
                        For lArrayRowIndex = 1 To iMaximumSuppliers
                            If cArrSalesData(conTotalSalesInc, lArrayRowIndex) > gconZeroValue Then
                                If chkProductReportTabDelimited Then
                                    Print #intFileNum, _
                                        fsSupplierNameFrom(lArrayRowIndex); _
                                        vbTab; _
                                        Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"); _
                                        vbTab; _
                                        Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat)
                                Else 'normal report
                                    Print #intFileNum, _
                                        fsSupplierNameFrom(lArrayRowIndex); _
                                        Tab(iArrTabStop(conPercentageMarketShare) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"))); _
                                        Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex) / cTotalSalesIncludingTaxForThePeriod * 100, "#0.#0"); _
                                        Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat))); _
                                        Format(cArrSalesData(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat)
                                End If
                            
                            End If
                        Next lArrayRowIndex
                        
                        Close #intFileNum
                        
                        Call subSetProductReportViewButton
                        
                        MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & _
                               ". Use the 'View' button to display", _
                                vbInformation, gconReportManager
                    Else 'environment was not created
                        MsgBox "Report was aborted", vbExclamation
                    End If 'environement created ?
                End If 'report destination

TidyUpSummarised:
                On Error GoTo 0
                
                'conserve memory
                Erase iArrTabStop
                Erase cArrSalesData
                
            Else 'no transactions for the date
                MsgBox "No sales transactions" & sFranchiseMessageBox & gsReportPeriodWording & sReportingPeriod, _
                        vbInformation, gconReportManager
            End If 'any transactions for the report date ?
        
            With stb
                .SimpleText = ""
                .Refresh
            End With
            
            rstDistinctBarcodesForTheReportingPeriod.Close
        Else
            MsgBox "No franchise selected", _
                    vbExclamation, gconReportManager
        End If
    End If 'summarised or notsummarised ?
    
    cmdMarketShare.Enabled = True
    Exit Sub
    
NotSummarisedPrinterErrorHandler:
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    
    Resume TidyUpnotSummarised

SummarisedPrinterErrorHandler:
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    
    Resume TidyUpSummarised

End Sub

Private Sub cmdMerger_Click()
'*****************************************************************************************
'This merges an offsite DB(Say Andys DB) and TSGHO1 database. It makes a call to one existing
'function "addToUpdateFile" and a new function below "addToAnotherReport" Copy these two functions
'and copy the STOCK tabs controls to get it to work(there are 3 new controls on the STOCK report)
'Text files are produced either for upload or information
'******************************************************************************************
Dim cnnConsultant As ADODB.Connection   '!!! ManualFix Clearing: Object variable not cleared: cnnConsultant
Dim rstSnpStockID As ADODB.Recordset    '!!! ManualFix Clearing: Object variable not cleared: rstSnpStockID
Dim rstHQDbNew As ADODB.Recordset       '!!! ManualFix Clearing: Object variable not cleared: rstHQDBNew
Dim rstConsultant As ADODB.Recordset    '!!! ManualFix Clearing: Object variable not cleared: rstConsultant
Dim rstStockHO As ADODB.Recordset       '!!! ManualFix Clearing: Object variable not cleared: rstStockHO
Dim rstStockCR As ADODB.Recordset       '!!! ManualFix Clearing: Object variable not cleared: rstStockCR
    
    Dim strConsultantDBPath As String
    Dim strAnswer As String
    Dim dNextAvailableStockID As Double
    Dim iTicker, iNewStock, iUpdatedStock, iExtras As Long
    Dim fDiffer As Boolean
    Dim sDiff As String
Dim strSQL As String
Dim strErrMsg As String
Dim rstHQExisting As ADODB.Recordset    '!!! ManualFix Clearing: Object variable not cleared: rstHQExisting
Dim lngHQExistingRecCount As Long

    If MsgBox("This option allows an external 'stock' table in a database" & vbCrLf & _
       "to be merged into the stock table in this database." & vbCrLf & _
       "Products in the external 'stock' table that are not in the local table" & vbCrLf & _
       "will be added to the local table." & vbCrLf & _
       "Products in the external 'stock' table whose fields differ from those in" & vbCrLf & _
       "a matching barcode in the local table will be added to the local table." & vbCrLf & vbCrLf & _
       "These updates are saved in text files that can be uploaded to the franchises, where" & vbCrLf & _
       "they can be merged in to the Retail Manager databases using Price Module.", vbOKCancel) = _
       vbCancel Then
       Exit Sub
    End If
    
    On Error GoTo DatabaseNotFound

    With stb
        .SimpleText = ""
        .Refresh
    End With
    
    strConsultantDBPath = InputBox("Enter the path to the Consultant Database", "TSG Database Updater")
    If strConsultantDBPath = "" Then
        strAnswer = MsgBox("Do you wish to continue with the Data Merge", vbOKCancel + vbExclamation, "Continue")
        If strAnswer <> 2 Then
        Do Until strConsultantDBPath <> ""
            strConsultantDBPath = InputBox("Enter the path to the Consultant Database", "TSG Database Updater")
            If strConsultantDBPath = "" Then
                strAnswer = MsgBox("Do you wish to continue with the Data Merge", vbOKCancel + vbExclamation, "Continue")
                If strAnswer = 2 Then
                    Exit Do
                End If
            End If
        Loop
        Else
            Exit Sub
        End If
        Exit Sub
    End If
      
    Set cnnConsultant = GetCnn(pDataSource:=strConsultantDBPath, _
                               pCnnMode:=adModeRead, _
                               pDataSourceType:=eMdb, _
                               pErrMsg:=strErrMsg)
    Set rstHQDbNew = GetRst(pCnn:=g.cnnDW, _
                            pSource:="Stock", _
                            pSourceType:=adCmdTable, _
                            pRstType:=eEditableFwdOnly, _
                            pErrMsg:=strErrMsg)

    'Overwrite TSGDatabase with Consultants data and write it to the
    'WLP update file and the Updated stock report

    Set rstConsultant = GetRst(pCnn:=cnnConsultant, pSource:="Stock", pSourceType:=adCmdTable, pErrMsg:=strErrMsg)

    iTicker = 1
    rstConsultant.MoveFirst
        
    Do Until rstConsultant.EOF
        strSQL = "SELECT * FROM Stock WHERE Barcode = " & SqlQ(rstConsultant!Barcode)
        lngHQExistingRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=strSQL)
        Set rstHQExisting = GetRst(pCnn:=g.cnnDW, _
                                   pSource:=strSQL, _
                                   pSourceType:=adCmdText, _
                                   pRstType:=eEditableDynamic, _
                                   pErrMsg:=strErrMsg)
        StatusBar "Processing record " & iTicker, pLog:=False
        iTicker = iTicker + 1

        If lngHQExistingRecCount <> 0 Then
            rstHQExisting.MoveLast
        End If
        If lngHQExistingRecCount = 0 Then
            'it must be a new item
            strSQL = "SELECT stock_id FROM Stock ORDER BY stock_id DESC"
            Set rstSnpStockID = GetRst(pCnn:=g.cnnDW, _
                                       pSource:=strSQL, _
                                       pSourceType:=adCmdText, _
                                       pErrMsg:=strErrMsg)
            If Not (rstSnpStockID.BOF And rstSnpStockID.EOF) Then
                dNextAvailableStockID = rstSnpStockID(gconStockTableStockIDField) + 1
            Else
                dNextAvailableStockID = 1
            End If
            iNewStock = iNewStock + 1
            rstSnpStockID.Close
            rstHQDbNew.AddNew
                rstHQDbNew("stock_id") = dNextAvailableStockID
                rstHQDbNew("barcode") = rstConsultant("barcode")
                rstHQDbNew("custom1") = rstConsultant("custom1")
                rstHQDbNew("custom2") = rstConsultant("custom2")
                rstHQDbNew("Sales_prompt") = rstConsultant("Sales_prompt")
                rstHQDbNew("inactive") = CBoolMySql(rstConsultant("inactive"))
                rstHQDbNew("allow_renaming") = CBoolMySql(rstConsultant("allow_renaming"))
                rstHQDbNew("allow_fractions") = CBoolMySql(rstConsultant("allow_fractions"))
                rstHQDbNew("package") = CBoolMySql(rstConsultant("package"))
                rstHQDbNew("tax_components") = CBoolMySql(rstConsultant("tax_components"))
                rstHQDbNew("print_components") = CBoolMySql(rstConsultant("print_components"))
                rstHQDbNew("description") = rstConsultant("description")
                rstHQDbNew("longdesc") = rstConsultant("longdesc")
                rstHQDbNew("cat1") = rstConsultant("cat1")
                rstHQDbNew("cat2") = rstConsultant("cat2")
                rstHQDbNew("goods_tax") = rstConsultant("goods_tax")
                rstHQDbNew("cost") = rstConsultant("cost")
                rstHQDbNew("sales_tax") = rstConsultant("sales_tax")
                rstHQDbNew("sell") = rstConsultant("sell")
                rstHQDbNew("quantity") = rstConsultant("quantity")
                rstHQDbNew("layby_qty") = rstConsultant("layby_qty")
                rstHQDbNew("date_created") = rstConsultant("date_created")
                rstHQDbNew("track_serial") = CBoolMySql(rstConsultant("track_serial"))
                rstHQDbNew("static_quantity") = CBoolMySql(rstConsultant("static_quantity"))
                rstHQDbNew("bonus") = rstConsultant("bonus")
                rstHQDbNew("order_threshold") = rstConsultant("order_threshold")
                rstHQDbNew("order_quantity") = rstConsultant("order_quantity")
                rstHQDbNew("supplier_id") = rstConsultant("supplier_id")
                rstHQDbNew("date_modified") = rstConsultant("date_modified")
                rstHQDbNew("freight") = CBoolMySql(rstConsultant("freight"))
                rstHQDbNew("sticks") = 0
            rstHQDbNew.Update

            'Append to the existing TSGDatabase and send to 'New Stock Report'
            'and New Stock upload file
            AddToUpdateFile pRstStock:=rstConsultant, _
                            pRstDbType:=eJetDb, _
                            pUseRecordValues:=True, _
                            bAddNewItem:=True
        ElseIf lngHQExistingRecCount = 1 Then
            ' Here if the product already exists in the HQ database, in which case we
            ' see if it differs from the Consukltants product in the important fields.
            ' If there is a difference, update the HQ database with the values from the
            ' consultant database.
              fDiffer = False
              sDiff = ""
              If CBool(rstHQExisting("tax_components")) <> CBool(rstConsultant("tax_components")) Then
                  fDiffer = True
                  sDiff = sDiff & "tax_components "
              End If
              If rstHQExisting("description") <> rstConsultant("description") Then
                  fDiffer = True
                  sDiff = sDiff & "description "
              End If
              If rstHQExisting("cat1") <> rstConsultant("cat1") Then
                  fDiffer = True
                  sDiff = sDiff & "cat "
              End If
              If rstHQExisting("cat2") <> rstConsultant("cat2") Then
                  fDiffer = True
                  sDiff = sDiff & "sub-cat "
              End If
              If rstHQExisting("goods_tax") <> rstConsultant("goods_tax") Then
                  fDiffer = True
                  sDiff = sDiff & "goods_tax "
              End If
              If rstHQExisting("cost") <> rstConsultant("cost") Then
                  fDiffer = True
                  sDiff = sDiff & "cost "
              End If
              If rstHQExisting("sales_tax") <> rstConsultant("sales_tax") Then
                  fDiffer = True
                  sDiff = sDiff & "sales_tax "
              End If
              If rstHQExisting("supplier_id") <> rstConsultant("supplier_id") Then
                  fDiffer = True
                  sDiff = sDiff & "supplier_id "
              End If
              If fDiffer Then
                  iUpdatedStock = iUpdatedStock + 1
                  rstHQExisting("tax_components") = CBoolMySql(rstConsultant("tax_components"))
                  rstHQExisting("cat1") = rstConsultant("cat1")
                  rstHQExisting("cat2") = rstConsultant("cat2")
                  rstHQExisting("goods_tax") = rstConsultant("goods_tax")
                  rstHQExisting("cost") = rstConsultant("cost")
                  rstHQExisting("sales_tax") = rstConsultant("sales_tax")
                  rstHQExisting("supplier_id") = rstConsultant("supplier_id")
                  rstHQExisting("date_modified") = rstConsultant("date_modified")
                  rstHQExisting.Update
                  AddToUpdateFile pRstStock:=rstHQExisting, _
                                  pRstDbType:=eMySqlDb, _
                                  bAddNewItem:=False, _
                                  pUseRecordValues:=True
                  Call addToDiffReport(rstHQExisting, sDiff)
              End If

        ElseIf lngHQExistingRecCount > 1 Then
        ' if more than one, biff all except one.
        '   Noting existing previous call to MoveLast
            Do Until rstHQExisting.BOF
                rstHQExisting.MovePrevious
                If Not rstHQExisting.BOF Then
                    rstHQExisting.Delete
                End If
            Loop
            rstHQExisting.MoveFirst '   ???
        End If
        rstConsultant.MoveNext
    Loop
        
       ' Now go through he HO DB and find the extra ones that are not in consultant DB
        Set rstStockHO = GetRst(pCnn:=g.cnnDW, pSource:="Stock", pSourceType:=adCmdTable, pErrMsg:=strErrMsg)
       'loop below generates a report on what the Head office DB has as compared to the Consultant
        Do Until rstStockHO.EOF
            ' find all those in HQ DB that are NOT in CR DB
            strSQL = "SELECT * FROM Stock WHERE barcode = " & SqlQ(rstStockHO("barcode"))
            Set rstStockCR = GetRst(pCnn:=cnnConsultant, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            If (rstStockCR.BOF And rstStockCR.EOF) Then
                Call addToAnotherReport(rstStockHO)
                iExtras = iExtras + 1
            End If
            rstStockHO.MoveNext
        Loop
    With stb
        .SimpleText = "Finished processing..."
        .Refresh
    End With
    rstStockCR.Close
    rstStockHO.Close
    rstHQDbNew.Close
    rstConsultant.Close
    rstHQExisting.Close
    cnnConsultant.Close
    Set cnnConsultant = Nothing
    
    subPopulateStockListboxes pExclDeletedStk:=ChkBoxToBool(chkStockTab_IncludeDeletedStock)
    
    MsgBox "Number of new stock added to HO Db:  " & iNewStock & vbCrLf & _
           "Number of updated stock items in HO: " & iUpdatedStock & vbCrLf & _
           "Number of stock in HO not in Andys DB: " & iExtras & vbCrLf & vbCrLf & _
           "Text files for upload can be found in, c:\ts\data\Upload folder", vbInformation, "Finished Processing"
    Exit Sub
    
DatabaseNotFound:
    MsgBox Err.Description, vbExclamation
    Exit Sub
    Resume  ' Not executed but assists when debugging in IDE
    
End Sub

Private Sub cmdPrintRejectedData_Click(Index As Integer)
    PrintRejectedData
End Sub

Private Sub cmdPromoTabSaveNonCompliantAll_Click()
Dim intFileNum As Integer
Dim strFileName As String
Dim fdlg As fdlgCommon

    Select Case True
        
        Case optNonCompliantSaveToFile.Item(1).Value ' CSV
            SaveNonCompliantRptToCsvFile
        
        Case optNonCompliantSaveToFile.Item(0).Value ' TXT
            
            Set fdlg = New fdlgCommon
            strFileName = fdlg.GetFullFileName(pMethod:=eShowSave, _
                                               pFilename:="NonCompliantRpt_ALL_" & Format$(Date, gkFmtDateUnambiguous) & ".txt", _
                                               pFilter:="*.txt", _
                                               pFilterDescription:="Text File (*.txt)", _
                                               pDefaultExtension:="txt")
            Set fdlg = Nothing
            If Len(strFileName) Then
                intFileNum = FreeFile   ' Get unused file
                Open strFileName For Output As #intFileNum
                printNonCompliantsReport pFileNum:=intFileNum
                Close #intFileNum
            End If
        
    End Select

End Sub

Private Sub cmdPromoTabSaveNonCompliantSelected_Click()
Dim intPrevMousePointer As Integer
Dim intFileNum As Integer
Dim strFileName As String
Dim colFranIDs As VBA.Collection
Dim fdlg As fdlgCommon

    
    Me.Enabled = False
    intPrevMousePointer = SetMousePointer(vbHourglass)
    
    Set colFranIDs = LvwGetSubItemCollection(pListView:=lvwNonCompliant, pSubItemIdx:=1, pSelected:=True)
    If Not colFranIDs Is Nothing Then
        Select Case True
            Case optNonCompliantSaveToFile.Item(1).Value ' CSV
                SaveNonCompliantRptToCsvFile pColSelFranIDs:=colFranIDs
            
            Case optNonCompliantSaveToFile.Item(0).Value ' TXT
                Set fdlg = New fdlgCommon
                strFileName = fdlg.GetFullFileName(pMethod:=eShowSave, _
                                                   pFilename:="NonCompliantRpt_" & Format$(Date, gkFmtDateUnambiguous) & ".txt", _
                                                   pFilter:="*.txt", _
                                                   pFilterDescription:="Text File (*.txt)", _
                                                   pDefaultExtension:="txt")
                Set fdlg = Nothing
                If Len(strFileName) Then
                    Set colFranIDs = LvwGetSubItemCollection(pListView:=lvwNonCompliant, pSubItemIdx:=1, pSelected:=True)
                    If Not colFranIDs Is Nothing Then
                        intFileNum = FreeFile   ' Get unused file
                        Open strFileName For Output As #intFileNum
                        printNonCompliantsReport pColSelFranIDs:=colFranIDs, pFileNum:=intFileNum
                        Close #intFileNum
                    End If
                End If
            
        
        End Select
    End If
                
    SetMousePointer intPrevMousePointer
    Me.Enabled = True

End Sub

Private Sub cmdPromotion_Click(Index As Integer)
Dim colFranIDs As VBA.Collection
    
    Select Case Index
        Case 0
            SaveNewPromotion
        Case 1
            ClearCreatePromoCtls
'   AUrban: Note that DELETE button has always been invisible while I have been at TSG
'        Case 2      ' delete all button
'            If MsgBox("Are you sure you want to delete all the existing Promotions listed?", vbYesNo) = vbYes Then
'                Set rsPromo = g.dbDW.OpenRecordset("SELECT * FROM Promotions;", dbOpenDynaset)
'                Do Until rsPromo.EOF
'                    Call deletePromoFromUploadsPending(rsPromo!promoID)
'                    rsPromo.Delete
'                    rsPromo.MoveNext
'                Loop
'                rsPromo.Close
'                Set rsPromo = Nothing
'
'                lvwPromo.ListItems.Clear
'                lvwPromo.Refresh
'                clearNonCompliant fPrompt:=False
'            End If
        Case 5  ' print non-compliants report
            printNonCompliantsReport
        Case 6  ' print list of promotions
            PrintPromoList
        Case 8  ' Show ALL/Show CURRENT
            With cmdPromotion(Index)
                If cmdPromotion(Index).Caption = "Show ALL" Then
                    LoadPromoListview pShowALL:=True, pForce:=True
                    .Caption = "Show Current"
                    .ToolTipText = "Show ONLY current promotions"
                Else 'Button was "Show Current"
                    LoadPromoListview pShowALL:=False, pForce:=True
                    .Caption = "Show ALL"
                    .ToolTipText = "Show ALL promotions including Expired promotions"
                End If
            End With
        Case 9  ' Print selected non-compliant promos
            Set colFranIDs = LvwGetSubItemCollection(pListView:=lvwNonCompliant, pSubItemIdx:=1, pSelected:=True)
            If Not colFranIDs Is Nothing Then
                printNonCompliantsReport pColSelFranIDs:=colFranIDs
                Set colFranIDs = Nothing
            End If
    End Select
    
End Sub

Private Sub cmdPromotionRecall_Click()
''' Review Can fix up message for single selected franchises later - z Test Msgs as part of testing
Const kProduct As Long = 1
Const kFrom As Long = 2
Const kTo As Long = 3
Const kState As Long = 6
Const kRegion As Long = 7
Const kPromoGrade As Long = 8
Const kPromoID As Long = 9
Const kSeparator As String = ", "
Dim intPrevMousePointer As Integer
Dim lngLoop As Long
Dim lngPromoID As Long
Dim strErrMsg As String
Dim strMsg As String
Dim strFailed As String
Dim strRecalled As String
Dim astrFailed() As String
Dim astrRecalled() As String
Dim vntPromo As Variant
Dim colPromos As VBA.Collection

    Me.Enabled = False
    
    Set colPromos = LvwGetItemCollection(pListView:=lvwPromo, pSelected:=True)
    If colPromos.Count Then
        For Each vntPromo In colPromos
            strMsg = strMsg & vbNewLine & _
                     vntPromo.Text & " - " & _
                     vntPromo.SubItems(kProduct) & kSeparator & _
                     "From: " & vntPromo.SubItems(kFrom) & kSeparator & _
                     "To: " & vntPromo.SubItems(kTo) & kSeparator & _
                     "State: " & vntPromo.SubItems(kState) & kSeparator & _
                     "Region: " & vntPromo.SubItems(kRegion) & kSeparator & _
                     "Grade: " & vntPromo.SubItems(kPromoGrade) & kSeparator & _
                     "ID: " & vntPromo.SubItems(kPromoID)
        Next vntPromo
        
    '   NB First iteration of above loop precedes 1st item with vbNewline
        strMsg = "Recall the following promotion(s)? " & vbNewLine & colPromos.Count & " selected" & strMsg
        If MsgBox(Prompt:=strMsg, Buttons:=vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            intPrevMousePointer = SetMousePointer(vbHourglass)
            
            For Each vntPromo In colPromos
                lngPromoID = CLng(vntPromo.SubItems(kPromoID))
                If Not PromotionRecall(pPromotionID:=lngPromoID, pErrMsg:=strErrMsg) Then
                    strFailed = strFailed & "FAILED Recall Promotion ID-" & vntPromo.SubItems(kPromoID) & " " & strErrMsg & vbNewLine
                    strMsg = "FAILED Recall Promotion " & "ID-" & vntPromo.SubItems(kPromoID) & vbNewLine & strErrMsg
                    MsgBox strMsg, vbExclamation
                Else
                    strRecalled = strRecalled & _
                                  vntPromo.Text & " - " & _
                                  vntPromo.SubItems(kProduct) & kSeparator & _
                                  "From: " & vntPromo.SubItems(kFrom) & kSeparator & _
                                  "To: " & vntPromo.SubItems(kTo) & kSeparator & _
                                  "State: " & vntPromo.SubItems(kState) & kSeparator & _
                                  "Region: " & vntPromo.SubItems(kRegion) & kSeparator & _
                                  "Grade: " & vntPromo.SubItems(kPromoGrade) & kSeparator & _
                                  "ID: " & vntPromo.SubItems(kPromoID) & _
                                  vbNewLine
                End If
            Next vntPromo
        
        '   Log successful promotion recalls (& recall requests) to Event Log
            astrRecalled() = Split(strRecalled, vbNewLine)
            strMsg = "Recalled Promotion(s): " & vbNewLine
            For lngLoop = LBound(astrRecalled) To UBound(astrRecalled) - 1
                strMsg = strMsg & astrRecalled(lngLoop) & vbNewLine
                StatusBar pMsg:="Recalled Promotion " & astrRecalled(lngLoop)
            Next lngLoop
            
            astrFailed() = Split(strFailed, vbNewLine)
            For lngLoop = LBound(astrFailed) To UBound(astrFailed) - 1
                StatusBar pMsg:=astrFailed(lngLoop)
            Next lngLoop
            
            SetMousePointer intPrevMousePointer
            LoadPromoListview pShowALL:=False
        
        End If
    
        Set colPromos = Nothing
        
        With Me.stb
            .SimpleText = vbNullString
            .Refresh
        End With

    End If
    
    Me.Enabled = True

End Sub

Private Sub cmdPRPrint_Click()
'--------------------------------------------------------------------------------------------------
'  AUrban Candidate for splitting into two procedures (Summarised Rpt and Not Summarised Rpt)
'         (Two procedures would perhaps call a common procedure - haven't fully analysed yet)
'--------------------------------------------------------------------------------------------------
'******************************************************************************************
'this now reports on not only franchise but also can generate reports
'on given stock for given franchises  - TCS
'******************************************************************************************
Dim rstAllSameBarcodesForTheReportingPeriod As ADODB.Recordset

    Dim bIndexSwapped As Boolean
    
    Dim cTotalSalesIncludingTaxForThePeriod As Currency
        
    Dim iArrayColumnIndex As Integer
    Dim iNumberOfCopies As Integer
    Dim iNumberOfProductSelected As Integer
    Dim iNumberOfFranchisesIncluded As Integer
    Dim iBarcodeIDCount As Integer
    Dim iCurrentProduct As Integer
    Dim iCountOfFranchises As Integer
    Dim iPageNumber As Integer
    
    Dim lArrayRowIndex As Long
    Dim lTotalCustomersForThePeriod As Long
    Dim lTotalNumberOfBarcodesForTheReportingPeriod As Long
    Dim lTotalItemsSoldForThePeriod As Long
    
    
Dim rstDistinctBarcodesForTheReportingPeriod As ADODB.Recordset
    
    Dim sFranchiseMessageBox As String
    Dim sIncludedDates As String
    Dim sIncludedFranchiseIDs As String
    Dim sIncludedFranchiseNames As String
    Dim sPlural As String
    Dim sReportingPeriod As String
    Dim stempName As String
    Dim sSQLQuery As String

    Dim vPlaceHolder As Variant
    
    'data array
    Const conSortIndex = 1, conProduct = 2, conQuantity = 3, conTotalSalesInc = 4, conNormalSell = 5, conCountSameBarcodes = 6  'for count of sameBarcodes and Normal sell price
    
    'tabstop array uses same as data array except for this extra
    Const conDisplayAverageSalesInc = 2      'this is Promotional price - TCS
    
    Const conReportType = "Product report for "

Dim intFileNum As Integer   ' 2 lines had to be split not to exceed VB line length limit when varibabl was added
Dim lngRecCount As Long
Dim datReportStart As Date
Dim strErrMsg As String
    
    cmdAllItems.Enabled = False
    
    With lvwPRProductReport
        .ListItems.Clear
        .Refresh
    End With
    
    If IsDateFmtOk() Then   ''' Review  Fix Reliance on date format when time permits

        Call subWriteSearchingMessageToStatusBar

        'build a query spec for all dates within the range
        datReportStart = GetDateFrom_ddmmmyy(lblPRProductReportStartDate)

        If lblPRProductReportStartDate = lblPRProductReportFinishDate Then
            sIncludedDates = "TransactionDate = " & MySqlDate(datReportStart)
            sReportingPeriod = lblPRProductReportStartDate
        Else
            sIncludedDates = "TransactionDate BETWEEN " & MySqlDate(datReportStart) & _
                                                " AND " & MySqlDate(GetDateFrom_ddmmmyy(lblPRProductReportFinishDate))
            sReportingPeriod = lblPRProductReportStartDate & " to " & lblPRProductReportFinishDate
        End If
        
        '!!! ManualFix Clearing: Array not deallocated: sbarcodeID
        Dim sbarcodeID() As String

        If optPRSelectedProducts(0) Then

            'build an array containing the ID for each selected product  - 'bevv
            For lArrayRowIndex = gconDisplayFirstItem To lstPRProductList.ListCount - 1

                If lstPRProductList.Selected(lArrayRowIndex) Then
                    ReDim Preserve sbarcodeID(iNumberOfProductSelected)
                    sbarcodeID(iNumberOfProductSelected) = fsStockIDFrom(lstPRProductList.List(lArrayRowIndex))
                    iNumberOfProductSelected = iNumberOfProductSelected + 1
                    iBarcodeIDCount = iBarcodeIDCount + 1
                    Dim sProductBarcodeList As String
                    sProductBarcodeList = sProductBarcodeList & (lstPRProductList.List(lArrayRowIndex)) & ", "
                End If

            Next lArrayRowIndex

        Else 'all products option was selected
            iBarcodeIDCount = iNumberOfProductSelected
            iNumberOfProductSelected = lstPRProductList.ListCount
            ReDim Preserve sbarcodeID(iNumberOfProductSelected - 1)

            For lArrayRowIndex = gconDisplayFirstItem To lstPRProductList.ListCount - 1
                sbarcodeID(lArrayRowIndex) = fsStockIDFrom(lstPRProductList.List(lArrayRowIndex))
            Next lArrayRowIndex

        End If
    
        If iNumberOfProductSelected Then
            'to check that at least one product has been selected
        Else 'no stock selected
            MsgBox "No stock selected", vbExclamation, gconReportManager
            GoTo Procedure_Exit
        End If ' any non-summarised stock selected?

        '*******************************************************************
        If optPRSendProductReportToFile Then

            If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                intFileNum = FreeFile   ' Get unused file
                Open gsProductReportPathAndFilename For Output As #intFileNum
            Else
                GoTo Procedure_Exit
            End If
        End If
        
        If optPRProductReportNotSummarised(0) Then '********************************************************
            ReDim lFranchiseID(gconZeroValue) As Long

            If optPRReportonSelectedFranchises(0) Then

                'build an array containing the ID for each selected franchise  - 'bevv
                For lArrayRowIndex = gconDisplayFirstItem To lstPRProductReportsFranchiseBusinessName.ListCount - 1

                    If lstPRProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) Then
                        iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                        ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)
                        lFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstPRProductReportsFranchiseBusinessName.List(lArrayRowIndex))
                        sIncludedFranchiseNames = sIncludedFranchiseNames & lstPRProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
                    End If

                Next lArrayRowIndex

                If iNumberOfFranchisesIncluded Then
                    'get rid of the last delimiters
                    sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
            
                    sFranchiseMessageBox = " for " & sIncludedFranchiseNames
                End If

            Else 'all franchises option was selected
                iNumberOfFranchisesIncluded = lstPRProductReportsFranchiseBusinessName.ListCount
                ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)

                For lArrayRowIndex = gconDisplayFirstItem To lstPRProductReportsFranchiseBusinessName.ListCount - 1
                    iCountOfFranchises = iCountOfFranchises + 1
                    lFranchiseID(iCountOfFranchises) = fsFranchiseIDFrom(lstPRProductReportsFranchiseBusinessName.List(lArrayRowIndex))
                Next lArrayRowIndex
            
                sFranchiseMessageBox = ""
            
                sIncludedFranchiseNames = gconAllFranchises
            End If
            
            If iNumberOfFranchisesIncluded Then 'franchises are included
                If iNumberOfFranchisesIncluded > 1 Then
                    sPlural = "s"
                End If
            
                Dim iCurrentFranchise As Integer

                '*********************************************************************
                If optPRSendProductReportToDisplay Then   'fart
                    'put header date and Franchises here
                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                    gvListItem.Text = sIncludedFranchiseNames & gsReportPeriodWording & sReportingPeriod
                ElseIf optPRSendProductReportToPrinter Then
                    'put print header here
                    Printer.Print sIncludedFranchiseNames & gsReportPeriodWording & sReportingPeriod
                    Printer.Print vbCrLf
                Else
                    'output to textfile place header and date here
                    Print #intFileNum, sIncludedFranchiseNames & gsReportPeriodWording & sReportingPeriod
                    Print #intFileNum, vbCrLf
                End If

                '*********************************************************************
                'Array placed here to print headings for file and printer and then erased
                'after each use
                ReDim iArrTabStop(conDisplayAverageSalesInc To conNormalSell) As Integer
        
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults!DocketPrinterEnabled Then
                    iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + Len(gconTruncateCharacter) + gconTruncateExtensionWidth + Len(gconSpace) + Len(gconStandardQuantityFormat)
                
                    iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conQuantity) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                                                         
                    iArrTabStop(conNormalSell) = iArrTabStop(conDisplayAverageSalesInc) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conNormalSell) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                Else
                    iArrTabStop(conQuantity) = 48
                    iArrTabStop(conDisplayAverageSalesInc) = 58
                    iArrTabStop(conNormalSell) = 68
                    iArrTabStop(conTotalSalesInc) = 80
                End If

                If optPRSendProductReportToPrinter Then    'TO PRINTER################
                    On Error GoTo NotSummarisedPrinterErrorHandler
                
                    cdlTSGDataWarehouse.ShowPrinter
                    Me.Refresh
                
                    Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                    '                    'leave a dual gap
                    Printer.Print vbCrLf
                    'headings
                    Printer.Print "Product"; Tab(iArrTabStop(conQuantity) - Len("  Qty")); "  Qty"; Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Prom unit")); "Prom unit"; Tab(iArrTabStop(conNormalSell) - Len("Norm unit")); "Norm unit"; Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); "Tot (inc)"
                    'leave a gap - Norm unit included in print - TCS
                    Erase iArrTabStop
                ElseIf optPRSendProductReportToFile Then
                    Print #intFileNum, "Tobacco Station" & sPlural & " - " & sIncludedFranchiseNames
                    'leave a dual gap
                    Print #intFileNum, vbCrLf
                
                    Print #intFileNum, conReportType & sReportingPeriod
                    Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                    'leave a dual gap
                    Print #intFileNum, vbCrLf

                    If chkPRProductReportTabDelimited Then
                        Print #intFileNum, "Product"; vbTab; "Qty"; vbTab; "Prom unit"; vbTab; "Norm unit"; vbTab; "Tot (inc)"
                    Else 'normal tabs
                        Print #intFileNum, "Product"; Tab(iArrTabStop(conQuantity) - Len("  Qty")); "  Qty"; Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Prom unit")); "Prom unit"; Tab(iArrTabStop(conNormalSell) - Len("Norm unit")); "Norm unit"; Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); "Tot (inc)"
                        Erase iArrTabStop
                    End If
                End If 'sent prod report to file #########################
                
                For iCurrentFranchise = 1 To iNumberOfFranchisesIncluded
                    For iCurrentProduct = 0 To (iNumberOfProductSelected - 1)
                        sSQLQuery = "SELECT DISTINCT Barcode FROM livedata " & vbNewLine & _
                                    "WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & _
                                    " AND (" & sIncludedDates & ")" & _
                                    " AND (Barcode = " & SqlQ(sbarcodeID(iCurrentProduct)) & ")" & _
                                    " AND (Quantity <> 0)"

                        lngRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
                    
                        If optPRSendProductReportToDisplay Then
                            If GetFranName(lFranchiseID(iCurrentFranchise)) <> stempName Then
                                Set gvListItem = lvwPRProductReport.ListItems.Add()
                                gvListItem.Text = gconSpace
                                Set gvListItem = lvwPRProductReport.ListItems.Add()
                                gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise)) 'puts name in display below header
                                Set gvListItem = lvwPRProductReport.ListItems.Add()
                                gvListItem.Text = gconSpace
                            End If
                        End If
                        
                        If lngRecCount Then
                            Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                                  pSource:=sSQLQuery, _
                                                                                  pSourceType:=adCmdText, _
                                                                                  pRstType:=eReadOnlyDynamic, _
                                                                                  pErrMsg:=strErrMsg) 'required for movelast etc...
                        
                            Call subWriteSizingArraysMessageToStatusBar
                        
                            'use the tabstop array to store the right justified position
                            ReDim iArrTabStop(conDisplayAverageSalesInc To conNormalSell) As Integer

                            'truncate if the docket printer is enabled in the defaults
                            If g.rstAppDefaults!DocketPrinterEnabled Then
                                iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + Len(gconTruncateCharacter) + gconTruncateExtensionWidth + Len(gconSpace) + Len(gconStandardQuantityFormat)
                            
                                iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conQuantity) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                                                                     
                                iArrTabStop(conNormalSell) = iArrTabStop(conDisplayAverageSalesInc) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                            
                                iArrTabStop(conTotalSalesInc) = iArrTabStop(conNormalSell) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                            Else
                                iArrTabStop(conQuantity) = 48
                                iArrTabStop(conDisplayAverageSalesInc) = 58
                                iArrTabStop(conNormalSell) = 68
                                iArrTabStop(conTotalSalesInc) = 78
                            End If
                        
                            'size the array
                            'do not remove this
                            lTotalNumberOfBarcodesForTheReportingPeriod = gconZeroValue  'to test for selecting barcodes in which case a full count unneccessary TCS

                            Do Until rstDistinctBarcodesForTheReportingPeriod.EOF

                                If fbProductIsIncludedInThisProductReport(rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                                    lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod + 1
                                End If

                                rstDistinctBarcodesForTheReportingPeriod.MoveNext
                            Loop

                            rstDistinctBarcodesForTheReportingPeriod.MoveFirst
                        
                            ReDim Varrsalesdata(conSortIndex To conCountSameBarcodes, 1 To lTotalNumberOfBarcodesForTheReportingPeriod) As Variant
                        
                            Call subWriteCollatingMessageToStatusBar
                        
                            cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                            lTotalItemsSoldForThePeriod = gconZeroValue
                            lArrayRowIndex = gconZeroValue
                            lTotalCustomersForThePeriod = gconZeroValue
                            
                            Do Until rstDistinctBarcodesForTheReportingPeriod.EOF

                                sSQLQuery = "SELECT * FROM LiveData " & vbNewLine & _
                                            "WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & _
                                            " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) & ")" & _
                                            " AND (" & sIncludedDates & ")"
                                lngRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
                    
                    'to test against selected products in product list
                                
                                'Test for Zero records
                                If lngRecCount Then
                                    Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                                         pSource:=sSQLQuery, _
                                                                                         pSourceType:=adCmdText, _
                                                                                         pErrMsg:=strErrMsg)
    
                                    If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                                    
                                        lArrayRowIndex = lArrayRowIndex + 1
                                        Call subDisplayCurrentRecordToUser(lArrayRowIndex, lTotalNumberOfBarcodesForTheReportingPeriod)
                                    
                                        If optPRDescription(0) Then
                                        
                                            'udtSalesData(lArrayRowIndex).Description = ""
                                        
                                            Varrsalesdata(conProduct, lArrayRowIndex) = fsDescriptionFrom(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))
                                        Else
                                            Varrsalesdata(conProduct, lArrayRowIndex) = rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)
                                        End If
                                
                                        Do Until rstAllSameBarcodesForTheReportingPeriod.EOF

                                            'avert an overflow divide by zero
                                            If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField) <> 0 Then
                                                Varrsalesdata(conQuantity, lArrayRowIndex) = Varrsalesdata(conQuantity, lArrayRowIndex) + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                            
                                                lTotalItemsSoldForThePeriod = lTotalItemsSoldForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
    
                                                Varrsalesdata(conTotalSalesInc, lArrayRowIndex) = Varrsalesdata(conTotalSalesInc, lArrayRowIndex) + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
            
                                                Varrsalesdata(conNormalSell, lArrayRowIndex) = Varrsalesdata(conNormalSell, lArrayRowIndex) + rstAllSameBarcodesForTheReportingPeriod("NormalSellInc")   '- for NormalSell column TCS
                                            
                                                cTotalSalesIncludingTaxForThePeriod = cTotalSalesIncludingTaxForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                            End If
    
                                            Varrsalesdata(conCountSameBarcodes, lArrayRowIndex) = lngRecCount
                                            
                                            rstAllSameBarcodesForTheReportingPeriod.MoveNext

                                            'to store a count of all barcodes but no copies -   TCS
                                        Loop
                                        rstAllSameBarcodesForTheReportingPeriod.Close

                                                                                 
                                    End If

                                End If

                                rstDistinctBarcodesForTheReportingPeriod.MoveNext

                                DoEvents
                            Loop
                            Set rstAllSameBarcodesForTheReportingPeriod = Nothing
                            
                            If lTotalCustomersForThePeriod <> gconZeroValue Then
                                'don't want to sort this as an array component
                                lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod - 1
                            End If
                        
                            Call subWriteSortingMessageToStatusBar
                            
                            Do 'at least one sort by description pass
                                bIndexSwapped = False
                            
                            '   Traverse array from first row to 2nd last row comparing the current row with the next
                                For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod - 1 '''?
                                    If (Varrsalesdata(conProduct, lArrayRowIndex) > Varrsalesdata(conProduct, lArrayRowIndex + 1)) Then 'swap

                                        For iArrayColumnIndex = conProduct To conNormalSell
                                            vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                            Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                            Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                            bIndexSwapped = True
                                        Next iArrayColumnIndex

                                    End If

                                Next lArrayRowIndex

                            Loop While bIndexSwapped
                        
                            If optPRSendProductReportToDisplay Then
                                If GetFranName(lFranchiseID(iCurrentFranchise)) <> stempName Then
                                    gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise)) 'puts name in display below header
                                    gvListItem.Text = gconSpace
                                End If

                                For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod

                                    If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                                        Set gvListItem = lvwPRProductReport.ListItems.Add()
                                        gvListItem.Text = Varrsalesdata(conProduct, lArrayRowIndex)
                                        Call gsubAddSubItemToListview(Varrsalesdata(conQuantity, lArrayRowIndex), 1)
                                        Call gsubAddSubItemToListview(Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex) / Varrsalesdata(conQuantity, lArrayRowIndex), gcon5DigitDollarFormat), 2)    ' - this is promotional price
                                        Call gsubAddSubItemToListview(Format(Varrsalesdata(conNormalSell, lArrayRowIndex) / Varrsalesdata(conCountSameBarcodes, lArrayRowIndex), gcon5DigitDollarFormat), 3)   ' - for NormalSell field
                                        Call gsubAddSubItemToListview(Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 4)
                                    End If

                                Next lArrayRowIndex

                                If chkIncludeTotalCustomerCount Then
                                    If lTotalCustomersForThePeriod <> gconZeroValue Then
                                        'leave another gap then total customers
                                        Set gvListItem = lvwPRProductReport.ListItems.Add()
                                        gvListItem.Text = gconSpace
                                        
                                        Set gvListItem = lvwPRProductReport.ListItems.Add()
                                        gvListItem.Text = "Total customers"
                                        Call gsubAddSubItemToListview(lTotalCustomersForThePeriod, 1)
                                    End If
                                End If

                                Me.Refresh
                            ElseIf optPRSendProductReportToPrinter Then '#################################
                                'Printer.Print GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                                'leave a gap

                                If GetFranName(lFranchiseID(iCurrentFranchise)) <> stempName Then
                                    Printer.Print vbCrLf
                                    Printer.Print GetFranName(lFranchiseID(iCurrentFranchise)) ' & gsReportPeriodWording & sReportingPeriod
                                    Printer.Print vbCrLf
                                End If

                                For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod

                                    If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                                        Printer.Print Varrsalesdata(conProduct, lArrayRowIndex); Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                                      Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                           Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                                    End If   'placed Normalsell value here - TCS

                                Next lArrayRowIndex
                                
                                'leave a gap
                                'Printer.Print gconSpace
                                
                                'expose the totals
                                '                                Printer.Print "Total"; '                                            Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); '                                            Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); '                                            Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); '                                            Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                                
                                If chkIncludeTotalCustomerCount Then
                                    If lTotalCustomersForThePeriod <> gconZeroValue Then
                                        'leave a dual gap
                                        Printer.Print vbCrLf
                                        'sbarcodeID (iCurrentProduct)
                                        'expose number of customers
                                        Printer.Print "Total customers"; Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                                    End If
                                End If

                            Else 'must be to file

                                If GetFranName(lFranchiseID(iCurrentFranchise)) <> stempName Then
                                    Print #intFileNum, gconSpace
                                    Print #intFileNum, GetFranName(lFranchiseID(iCurrentFranchise)) '& gsReportPeriodWording & sReportingPeriod
                                    'leave a gap
                                    Print #intFileNum, gconSpace
                                End If
                                
                                For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod

                                    If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                                        If chkPRProductReportTabDelimited Then
                                            Print #intFileNum, Varrsalesdata(conProduct, lArrayRowIndex); vbTab; Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); vbTab; Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); vbTab; Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat); vbTab; Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                                        Else 'normal report
                                            Print #intFileNum, Varrsalesdata(conProduct, lArrayRowIndex); Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                                          Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                               Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                                        End If
                                    End If

                                Next lArrayRowIndex
                                
                                If chkIncludeTotalCustomerCount Then
                                    If lTotalCustomersForThePeriod <> gconZeroValue Then
                                        'leave a dual gap
                                        Print #intFileNum, gconSpace
                                        Print #intFileNum, gconSpace

                                        'expose number of customers
                                        If chkPRProductReportTabDelimited Then
                                            Print #intFileNum, "Total customers"; vbTab; Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                                        Else 'normal report
                                            Print #intFileNum, "Total customers"; Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                                        End If
                                    End If
                                End If
                            
                            End If 'report destination
                            
                            On Error GoTo 0
                            'conserve memory
                            Erase iArrTabStop
                            Erase Varrsalesdata
                            
                            rstDistinctBarcodesForTheReportingPeriod.Close
                            
                        Else 'no transactions for the date Need to put a trap in here

                            If GetFranName(lFranchiseID(iCurrentFranchise)) <> stempName Then
                                If optPRSendProductReportToDisplay Then
                                    'do nothing
                                    '                                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                                    '                                    gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise))
                                ElseIf optPRSendProductReportToPrinter Then
                                    Printer.Print vbCrLf
                                    Printer.Print GetFranName(lFranchiseID(iCurrentFranchise)) ' & gsReportPeriodWording & sReportingPeriod
                                    Printer.Print vbCrLf
                                Else

                                    If optPRSendProductReportToFile Then
                                        Print #intFileNum, gconSpace
                                        Print #intFileNum, GetFranName(lFranchiseID(iCurrentFranchise)) ' & gsReportPeriodWording & sReportingPeriod
                                        'leave a gap
                                        Print #intFileNum, gconSpace
                                    End If
                                End If
                            End If

                            If optPRDescription(0) Then        ' Added Barcode to No transaction display TCS
                                If optPRSendProductReportToDisplay Then
                                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                                    gvListItem.Text = fsDescriptionFrom(sbarcodeID(iCurrentProduct)) & ": No sales"
                                ElseIf optPRSendProductReportToPrinter Then    '#############################
                                    Printer.Print fsDescriptionFrom(sbarcodeID(iCurrentProduct)) & ": No sales"
                                Else 'is destined for the file
                                    Print #intFileNum, fsDescriptionFrom(sbarcodeID(iCurrentProduct)) & ": No sales"
                                End If

                            Else

                                If optPRSendProductReportToDisplay Then
                                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                                    gvListItem.Text = sbarcodeID(iCurrentProduct) & ": No sales"
                                ElseIf optPRSendProductReportToPrinter Then    '#############################
                                    Printer.Print sbarcodeID(iCurrentProduct) & ": No sales"
                                Else 'is destined for the file
                                    Print #intFileNum, sbarcodeID(iCurrentProduct) & ": No sales"
                                End If
                            End If
                        End If 'any transactions for the report date ?
                        
                        With stb
                            .SimpleText = ""
                            .Refresh
                        End With

                        stempName = GetFranName(lFranchiseID(iCurrentFranchise))
                    Next iCurrentProduct
                    Set rstDistinctBarcodesForTheReportingPeriod = Nothing
                    DoEvents
                Next iCurrentFranchise
                
TidyUpnotSummarised:
                On Error GoTo 0
                
                If optPRSendProductReportToDisplay Then
                    'do nothing
                ElseIf optPRSendProductReportToPrinter Then     '##############################
                    Printer.EndDoc
                    MsgBox "Report was successfully submitted to the selected printer", vbInformation, gconReportManager
                Else 'was to file
                    Call subSetProductReportViewButton
                    MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & ". Use the 'View' button to display", vbInformation, gconReportManager
                End If

            Else 'no franchise selected
                MsgBox "No franchise selected", vbExclamation, gconReportManager
            End If 'any non-summarised franchises selected ?

        Else 'summarised***************************************************
            '--------------------------------------------------------------------------------------------------------------
            '  AUrban SUMMARISED REPORT: Procedure is a candidate for splitting above and below here into two procedures
            '--------------------------------------------------------------------------------------------------------------

            If optPRReportonSelectedFranchises(0) Then

                'build a query spec for all selected franchises
                For lArrayRowIndex = gconDisplayFirstItem To lstPRProductReportsFranchiseBusinessName.ListCount - 1

                    If lstPRProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) Then
                        iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1

                        sIncludedFranchiseNames = sIncludedFranchiseNames & lstPRProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "

                        sIncludedFranchiseIDs = sIncludedFranchiseIDs & gconLiveDataTableTSGFranchiseIDField & " = " & fsFranchiseIDFrom(lstPRProductReportsFranchiseBusinessName.List(lArrayRowIndex)) & " OR "
                    End If

                Next lArrayRowIndex
                
                If iNumberOfFranchisesIncluded Then
                    'get rid of the last delimiters
                    sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
                    
                    sIncludedFranchiseIDs = Left(sIncludedFranchiseIDs, Len(sIncludedFranchiseIDs) - Len(" OR "))
                
                    sFranchiseMessageBox = " for " & sIncludedFranchiseNames
                End If

                sSQLQuery = "SELECT DISTINCT " & gconLiveDataTableBarcodeField & gconSpace & "FROM " & gconLiveDataTable & gconSpace & "WHERE (" & sIncludedFranchiseIDs & ")" & gconSpace & "AND (" & sIncludedDates & ")" & gconSpace & "AND (" & gconLiveDataTableQuantityField & " <> 0)"
            Else ' all franchises option selected
                iNumberOfFranchisesIncluded = lstPRProductReportsFranchiseBusinessName.ListCount
                
                sFranchiseMessageBox = ""
                
                sIncludedFranchiseNames = gconAllFranchises
                
                sSQLQuery = "SELECT DISTINCT " & gconLiveDataTableBarcodeField & gconSpace & "FROM " & gconLiveDataTable & gconSpace & "WHERE (" & sIncludedDates & ")" & gconSpace & "AND (" & gconLiveDataTableQuantityField & " <> 0)"
            End If
            
            If iNumberOfFranchisesIncluded = 0 Then 'no transactions for the date
                MsgBox "No sales transactions" & sFranchiseMessageBox & gsReportPeriodWording & sReportingPeriod, vbInformation, gconReportManager
            Else    ' franchises are included
                If iNumberOfFranchisesIncluded > 1 Then
                    sPlural = "s"
                End If

                '*********************************************************************
                If optPRSendProductReportToDisplay Then   'fart
                    'put header date and Franchises here
                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                    gvListItem.Text = sIncludedFranchiseNames & gsReportPeriodWording & sReportingPeriod
                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                ElseIf optPRSendProductReportToPrinter Then
                    'put print header here
                    Printer.Print sIncludedFranchiseNames & gsReportPeriodWording & sReportingPeriod
                    Printer.Print vbCrLf
                Else
                    'output to textfile place header and date here
                    Print #intFileNum, sIncludedFranchiseNames & gsReportPeriodWording & sReportingPeriod
                    Print #intFileNum, vbCrLf
                End If

                Call subWriteSizingArraysMessageToStatusBar
                    
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conDisplayAverageSalesInc To conNormalSell) As Integer

                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + Len(gconTruncateCharacter) + gconTruncateExtensionWidth + Len(gconSpace) + Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conQuantity) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                    
                    iArrTabStop(conNormalSell) = iArrTabStop(conDisplayAverageSalesInc) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                    
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conDisplayAverageSalesInc) + Len(gcon5DigitDollarFormat) + Len(gconSpace)
                    
                Else
                    iArrTabStop(conQuantity) = 48
                    iArrTabStop(conDisplayAverageSalesInc) = 58
                    iArrTabStop(conNormalSell) = 68
                    iArrTabStop(conTotalSalesInc) = 78
                End If
                    
                'size the array
                Dim rst As ADODB.Recordset
                Dim intProductsIncluded As Integer    'NOT ALL SELECTED PRODUCTS ARE INCLUDED SOME HAVE ZERO QUANTITIES
                Dim intLoop As Integer

                For intLoop = 0 To iNumberOfProductSelected - 1
                    Set rst = GetRst(pCnn:=g.cnnDW, _
                                     pSource:=sSQLQuery & " AND (Barcode = " & SqlQ(sbarcodeID(intLoop)) & ")", _
                                     pSourceType:=adCmdText, _
                                     pErrMsg:=strErrMsg)
                    If Not (rst.BOF And rst.EOF) Then
                        intProductsIncluded = intProductsIncluded + 1
                    End If
                Next
                    
                If intProductsIncluded Then
                    ReDim Varrsalesdata(conSortIndex To conCountSameBarcodes, 1 To intProductsIncluded) As Variant
            '   Else
            '       No products included -> we don't need an array to populate
            '       If we try to Redim array with intProductsIncluded = 0 we would get a runtime error
                End If

                Call subWriteCollatingMessageToStatusBar
                    
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                lTotalItemsSoldForThePeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                lTotalCustomersForThePeriod = gconZeroValue
                
                For iCurrentProduct = 0 To iNumberOfProductSelected - 1

                    If sbarcodeID(iCurrentProduct) <> "" Then
                        If optPRReportonSelectedFranchises(0) Then
                            sSQLQuery = "SELECT * FROM " & gconLiveDataTable & " WHERE (" & sIncludedFranchiseIDs & ") AND (" & sIncludedDates & ") AND (" & gconLiveDataTableBarcodeField & " = " & SqlQ(sbarcodeID(iCurrentProduct)) & ")"
                            'put the sbarcodeID in to select barcodes selected in listbox
                        Else 'all franchises, no requirement to discriminate (performance reasons)
                            sSQLQuery = "SELECT * FROM " & gconLiveDataTable & " WHERE (" & sIncludedDates & ") AND (" & gconLiveDataTableBarcodeField & " = " & SqlQ(sbarcodeID(iCurrentProduct)) & ")"
                        End If
                        
                        lngRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
                        'test for zero records for that franchise - TCS
                        If lngRecCount Then
                            Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                                 pSource:=sSQLQuery, _
                                                                                 pSourceType:=adCmdText, _
                                                                                 pErrMsg:=strErrMsg)

                            If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) <> "TOTALCUSTOMERS" Then
                                If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                                    lArrayRowIndex = lArrayRowIndex + 1
                                    
                                    Call subDisplayCurrentRecordToUser(iCurrentProduct + 1, iNumberOfProductSelected)
                                    
                                    If optPRDescription(0) Then
                                        'udtSalesData(lArrayRowIndex).Description = ""
                                        Varrsalesdata(conProduct, lArrayRowIndex) = fsDescriptionFrom(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))
                                    Else
                                        Varrsalesdata(conProduct, lArrayRowIndex) = rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)
                                    End If

                                    Do Until rstAllSameBarcodesForTheReportingPeriod.EOF

                                        'avert an overflow divide by zero
                                        If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField) <> 0 Then
                                            Varrsalesdata(conQuantity, lArrayRowIndex) = Varrsalesdata(conQuantity, lArrayRowIndex) + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                                
                                            lTotalItemsSoldForThePeriod = lTotalItemsSoldForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                                
                                            Varrsalesdata(conTotalSalesInc, lArrayRowIndex) = Varrsalesdata(conTotalSalesInc, lArrayRowIndex) + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                                
                                            Varrsalesdata(conNormalSell, lArrayRowIndex) = Varrsalesdata(conNormalSell, lArrayRowIndex) + rstAllSameBarcodesForTheReportingPeriod("NormalSellInc")   '- for NormalSell column TCS
                                                
                                            cTotalSalesIncludingTaxForThePeriod = cTotalSalesIncludingTaxForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                        End If

                                        Varrsalesdata(conCountSameBarcodes, lArrayRowIndex) = lngRecCount
                                        
                                        rstAllSameBarcodesForTheReportingPeriod.MoveNext

                                        DoEvents
                                    Loop

                                End If
                            Else
                                Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                                    lTotalCustomersForThePeriod = lTotalCustomersForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                    rstAllSameBarcodesForTheReportingPeriod.MoveNext
                                Loop
                            End If
                            
                            rstAllSameBarcodesForTheReportingPeriod.Close
                            
                        Else 'no transactions for the date

                            If optPRDescription(0) Then
                                If optPRSendProductReportToDisplay Then
                                    Set gvListItem = lvwPRProductReport.ListItems.Add() 'burp
                                    gvListItem.Text = fsDescriptionFrom(sbarcodeID(iCurrentProduct)) & ": No sales"
                                ElseIf optPRSendProductReportToPrinter Then    '#############################
                                    Printer.Print fsDescriptionFrom(sbarcodeID(iCurrentProduct)) & ": No sales"
                                Else 'is destined for the file
                                    Print #intFileNum, fsDescriptionFrom(sbarcodeID(iCurrentProduct)) & ": No sales"
                                End If

                            Else

                                If optPRSendProductReportToDisplay Then
                                    Set gvListItem = lvwPRProductReport.ListItems.Add() 'burp
                                    gvListItem.Text = sbarcodeID(iCurrentProduct) & ": No sales"
                                ElseIf optPRSendProductReportToPrinter Then    '#############################
                                    Printer.Print sbarcodeID(iCurrentProduct) & ": No sales"
                                Else 'is destined for the file
                                    Print #intFileNum, sbarcodeID(iCurrentProduct) & ": No sales"
                                End If
                            End If

                        End If 'any transactions for the report date ?
                    End If

                    DoEvents

                Next iCurrentProduct
                
                Set rstAllSameBarcodesForTheReportingPeriod = Nothing
                
                If lTotalCustomersForThePeriod <> gconZeroValue Then
                    'don't want to sort this as an array component
                    lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod - 1
                End If
                
                Call subWriteSortingMessageToStatusBar
                
                'sort by description
                'if only one product selected then no need to sort
                If intProductsIncluded > 1 And optPRReportonSelectedFranchises(1) Then

                    Do
                        bIndexSwapped = False

                    '   Traverse array from first row to 2nd last row comparing the current row with the next
                        For lArrayRowIndex = 1 To (intProductsIncluded - 1)

                            If (Varrsalesdata(conProduct, lArrayRowIndex) > Varrsalesdata(conProduct, lArrayRowIndex + 1)) Then 'swap

                                For iArrayColumnIndex = conProduct To conTotalSalesInc
                                    vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                    Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                    Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                    bIndexSwapped = True
                                Next iArrayColumnIndex

                            End If

                        Next lArrayRowIndex

                    Loop While bIndexSwapped

                End If
                    
                If optPRSendProductReportToDisplay Then

                    For lArrayRowIndex = 1 To intProductsIncluded

                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            Set gvListItem = lvwPRProductReport.ListItems.Add()
                            gvListItem.Text = Varrsalesdata(conProduct, lArrayRowIndex)
                            Call gsubAddSubItemToListview(Varrsalesdata(conQuantity, lArrayRowIndex), 1)
                            Call gsubAddSubItemToListview(Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex) / Varrsalesdata(conQuantity, lArrayRowIndex), gcon5DigitDollarFormat), 2)
                            Call gsubAddSubItemToListview(Format(Varrsalesdata(conNormalSell, lArrayRowIndex) / Varrsalesdata(conCountSameBarcodes, lArrayRowIndex), gcon5DigitDollarFormat), 3)
                            Call gsubAddSubItemToListview(Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 4)
                        End If

                    Next lArrayRowIndex
                    
                    'leave a gap then totals
                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                    
                    Set gvListItem = lvwPRProductReport.ListItems.Add()
                    gvListItem.Text = "Total items"
                    Call gsubAddSubItemToListview(lTotalItemsSoldForThePeriod, 1)
                    Call gsubAddSubItemToListview(Format(cTotalSalesIncludingTaxForThePeriod, gcon6DigitDollarFormat), 4)
                    
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave another gap then total customers
                            Set gvListItem = lvwPRProductReport.ListItems.Add()
                            gvListItem.Text = gconSpace
                            
                            Set gvListItem = lvwPRProductReport.ListItems.Add()
                            gvListItem.Text = "Total customers"
                            Call gsubAddSubItemToListview(lTotalCustomersForThePeriod, 1)
                        End If
                    End If

                ElseIf optPRSendProductReportToPrinter Then     'PRINTER
                    On Error GoTo SummarisedPrinterErrorHandler
                    
                    cdlTSGDataWarehouse.ShowPrinter
                    iNumberOfCopies = cdlTSGDataWarehouse.Copies
                    Me.Refresh
                        
                    For iPageNumber = 1 To iNumberOfCopies
                        '                       Printer.Print "Tobacco Station" & sPlural & '                                          " - " & sIncludedFranchiseNames
                        '                       'leave a dual gap
                        '                       Printer.Print vbCrLf
            
                        Printer.Print conReportType & sReportingPeriod
                        Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                        'leave a dual gap
                        Printer.Print vbCrLf
                        
                        'headings
                        Printer.Print "Product"; Tab(iArrTabStop(conQuantity) - Len("  Qty")); "  Qty"; Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Avg unit")); "Avg unit"; Tab(iArrTabStop(conNormalSell) - Len("Norm unit")); "Norm unit"; Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); "Tot (inc)"
                        'leave a gap
                        Printer.Print gconSpace
                        
                        For lArrayRowIndex = LBound(Varrsalesdata, 2) To UBound(Varrsalesdata, 2)
                            If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                                Printer.Print Varrsalesdata(conProduct, lArrayRowIndex); Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conTotalSalesInc) - _
                                   Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            End If

                        Next lArrayRowIndex
                        
                        'leave a gap
                        Printer.Print gconSpace
                        
                        'expose the totals
                        Printer.Print "Total"; Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                        
                        If chkIncludeTotalCustomerCount Then
                            If lTotalCustomersForThePeriod <> gconZeroValue Then
                                'leave a dual gap
                                Printer.Print vbCrLf
                        
                                'expose number of customers
                                Printer.Print "Total customers"; Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            End If
                        End If

                    Next iPageNumber

                    Printer.EndDoc
                    
                    MsgBox "Report was successfully submitted to the selected printer", vbInformation, gconReportManager
                Else 'must be to file
                    Print #intFileNum, "Tobacco Station" & sPlural & " - " & sIncludedFranchiseNames
                    'leave a dual gap
                    Print #intFileNum, vbCrLf
                            
                    Print #intFileNum, conReportType & sReportingPeriod
                    Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                    'leave a dual gap
                    Print #intFileNum, vbCrLf
                    
                    'headings
                    If chkPRProductReportTabDelimited Then
                        Print #intFileNum, "Product"; vbTab; "Qty"; vbTab; "Prom unit"; vbTab; "Norm unit"; vbTab; "Tot (inc)"
                    Else 'normal tabs
                        Print #intFileNum, "Product"; Tab(iArrTabStop(conQuantity) - Len("  Qty")); "  Qty"; Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Avg unit")); "Avg unit"; Tab(iArrTabStop(conNormalSell) - Len("Norm unit")); "Norm unit"; Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); "Tot (inc)"
                    End If
                    
                    'leave a gap
                    Print #intFileNum, gconSpace

                    For lArrayRowIndex = 1 To intProductsIncluded

                        If Varrsalesdata(conQuantity, lArrayRowIndex) <> gconZeroValue Then
                            If chkPRProductReportTabDelimited Then
                                Print #intFileNum, Varrsalesdata(conProduct, lArrayRowIndex); vbTab; Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); vbTab; Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); vbTab; Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat); vbTab; Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            Else 'normal report
                                Print #intFileNum, Varrsalesdata(conProduct, lArrayRowIndex); Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / Val(Varrsalesdata(conCountSameBarcodes, lArrayRowIndex)), gcon5DigitDollarFormat); Tab(iArrTabStop(conTotalSalesInc) - _
                                   Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            End If
                        End If

                        DoEvents
                    Next lArrayRowIndex
                    
                    'leave a gap
                    Print #intFileNum, gconSpace
                    
                    'expose the totals
                    If chkPRProductReportTabDelimited Then
                        Print #intFileNum, "Total"; vbTab; Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); vbTab; Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    Else 'normal report
                        Print #intFileNum, "Total"; Tab(iArrTabStop(conQuantity) - Len(Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat))); Format(lTotalItemsSoldForThePeriod, gconStandardQuantityFormat); Tab(iArrTabStop(conTotalSalesInc) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                    End If
                        
                    If chkIncludeTotalCustomerCount Then
                        If lTotalCustomersForThePeriod <> gconZeroValue Then
                            'leave a dual gap
                            Print #intFileNum, gconSpace
                            Print #intFileNum, gconSpace

                            'expose number of customers
                            If chkPRProductReportTabDelimited Then
                                Print #intFileNum, "Total customers"; vbTab; Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            Else 'normal report
                                Print #intFileNum, "Total customers"; Tab(iArrTabStop(conTotalSalesInc) - Len(Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat))); Format(lTotalCustomersForThePeriod, gconStandardQuantityFormat)
                            End If
                        End If
                    End If

                    Call subSetProductReportViewButton
                    
                    MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & ". Use the 'View' button to display", vbInformation, gconReportManager
                End If 'report destination
    
TidyUpSummarised:
                On Error GoTo 0
                
                'conserve memory
                Erase iArrTabStop
                Erase Varrsalesdata
                
            End If 'any transactions for the report date ?
        
            With stb
                .SimpleText = ""
                .Refresh
            End With
            
        End If 'summarised or notsummarised?
    End If 'date format correct?
    
    Close #intFileNum
    cmdAllItems.Enabled = True
    
Procedure_Exit:
    Exit Sub

NotSummarisedPrinterErrorHandler:
    
    Printer.KillDoc

    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If

    Resume TidyUpnotSummarised

SummarisedPrinterErrorHandler:
    
    Printer.KillDoc

    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If

    Resume TidyUpSummarised
    Resume  ' Not executed but assists when debugging in IDE

End Sub

Private Sub cmdPRView_Click()
    subOpenFile gsProductReportPathAndFilename
End Sub

Private Sub cmdPurgeNielsenReportList_Click()
    MsgBox "Not functioning for new reporting. Code needs replacecment"   ''' Review
End Sub

Private Sub cmdSaveFranchiseDetails_Click()
    LockDCTabFranchiseCtls pLocked:=True
    subSaveFranchiseDetails bAddNewFranchise:=False
    subDisplayFranchiseDetails
End Sub

Private Sub cmdSaveStockDetails_Click()
    subSaveStockDetails pAddNewItem:=False
End Sub

Private Sub cmdSettingsPassword_Click(Index As Integer)
    If txtConfirmPassword.Text = _
        "asd" & Right(fsVersion, 1) Then
        Call AcceptSettings
    Else
        MsgBox "Incorrect password", vbCritical
    End If
    Call disableSettingsFields

End Sub

Private Sub cmdStickReport_Click()
'--------------------------------------------------------------------------------------------------------------
'  AUrban Procedure is a candidate for splitting into two procedures (Summarised Rpt and Not Summarised Rpt)
'  AUrban (cmdAllItems_Click, cmdMarketShare_Click & cmdStickReport_Click are similar. Prob cut & pasted and modified)
'--------------------------------------------------------------------------------------------------------------
Dim strErrMsg As String

    Dim iCurrentProductSupplierID As Integer, _
        iCurrentFranchise As Integer, iPlaceHolder As Integer, _
        iSupplierIndex As Integer, _
        iTotalSuppliersForTheTransactionPeriod As Integer

    Dim dThousandsOfSticksThisSupplierForTheTransactionPeriod As Double

    Dim bIndexSwapped As Boolean
    
    Dim cTotalSalesIncludingTaxForThePeriod As Currency
        
    Dim iArrayColumnIndex As Integer, _
        iNumberOfCopies As Integer, _
        iNumberOfFranchisesIncluded As Integer, _
        iPageNumber As Integer
    
    Dim lArrayRowIndex As Long
    Dim BarcodeArrayIndex As Long

    Dim lTotalNumberOfProductsForTheTransactionPeriod As Long
        
    Dim dThousandsOfSticksThisTransaction As Double, _
        dTotalThousandsOfSticksForTheTransactionPeriod As Double
    
Dim rstDistinctProductsForTheTransactionPeriod As ADODB.Recordset
Dim rstAllSameProductsForTheTransactionPeriod As ADODB.Recordset
        
    Dim sFranchiseMessageBox As String, _
        sIncludedDates As String, _
        sIncludedFranchiseIDs As String, _
        sIncludedFranchiseNames As String, _
        sPlural As String
    Dim sTransactionPeriod As String
    Dim sSQLQuery As String
    Dim fHeadingDone As Boolean
    Dim vPlaceHolder As Variant
    ReDim sArrBarcode(0) As String      ' PALb50
    
    Dim intFileNum As Integer
    
    'data array
    Const conSortIndex = 1
    Const conProduct = 2
    Const conSupplier = 3
    Const conThousandsOfSticks = 4
    Const conTotalRebate = 5
    
    'listview display array...
    Const conSubSupplier = 1, _
          conSubMarketShare = 2, _
          conSubThousandsOfSticks = 3, _
          conSubAverageRebate = 4, _
          conSubTotalRebate = 5
    
    Const conReportType = "Stick report for "
    
Dim datTransactionStart As Date
    
    cmdStickReport.Enabled = False
    
    With lvwStickReport
        .ListItems.Clear
        .Refresh
    End With

    If Not IsDateFmtOk() Then   ''' Review Fix Reliance on date format when time permits
        MsgBox "incorrect system date format"
        Exit Sub
    End If
    
    Call subWriteSearchingMessageToStatusBar
    
    'build a query spec for all dates within the range
    datTransactionStart = GetDateFrom_ddmmmyy(lblStickReportStartDate)
    If lblStickReportStartDate = lblStickReportFinishDate Then
        sIncludedDates = "TransactionDate = " & MySqlDate(datTransactionStart)
        sTransactionPeriod = lblStickReportStartDate
    Else
        sIncludedDates = "TransactionDate BETWEEN " & MySqlDate(datTransactionStart) & _
                                            " AND " & MySqlDate(GetDateFrom_ddmmmyy(lblStickReportFinishDate))
        sTransactionPeriod = lblStickReportStartDate & " to " & lblStickReportFinishDate
    End If
    
    If Not optStickReportSummaryType(2) Then
       ReDim lFranchiseID(gconZeroValue) As Long
        
            'build an array containing the ID for each selected franchise
        For lArrayRowIndex = gconDisplayFirstItem To lstStickReportsFranchiseBusinessName.ListCount - 1
            If (lstStickReportsFranchiseBusinessName.Selected(lArrayRowIndex) And optStickReportOnSelectedFranchisesOnly(0)) Or _
                (Not optStickReportOnSelectedFranchisesOnly(0)) Then
                iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)
                lFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstStickReportsFranchiseBusinessName.List(lArrayRowIndex))
                sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                          lstStickReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
            End If
        Next lArrayRowIndex
        
        If iNumberOfFranchisesIncluded = 0 Then
            MsgBox "No franchise selected", vbExclamation, gconReportManager
            cmdStickReport.Enabled = True
            Exit Sub
        End If
        
        If optStickReportOnSelectedFranchisesOnly(0) Then
            'get rid of the last delimiters
            sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
            sFranchiseMessageBox = " for " & sIncludedFranchiseNames
        Else
            sIncludedFranchiseNames = gconAllFranchises
            sFranchiseMessageBox = ""
        End If
        
        If iNumberOfFranchisesIncluded > 1 Then
            sPlural = "s"
        End If
         
        If optSendStickReportToPrinter Then
            On Error GoTo NotSummarisedPrinterErrorHandler
            
            cdlTSGDataWarehouse.ShowPrinter
            Me.Refresh
            
            Printer.Print "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
            'leave a dual gap
            Printer.Print vbCrLf
            
            Printer.Print conReportType & sTransactionPeriod
            Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
            'leave a dual gap
            Printer.Print vbCrLf
        ElseIf optSendStickReportToFile Then
            If fbNewStickReportDocumentEnvironmentWasSuccessfullyPrepared Then
                intFileNum = FreeFile   ' Get unused file
                Open gsStickReportPathAndFilename For Output As #intFileNum
                Print #intFileNum, "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
                'leave a dual gap
                Print #intFileNum, vbCrLf
            
                Print #intFileNum, conReportType & sTransactionPeriod
                Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                'leave a dual gap
                Print #intFileNum, vbCrLf
            Else 'environment was not created
                MsgBox "Report was aborted", vbExclamation
                GoTo NotSummarisedTidyUp
            End If 'environement created ?
        End If 'sent prod report to file
                                    
        For iCurrentFranchise = 1 To iNumberOfFranchisesIncluded
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & vbNewLine & _
                        "WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & gconSpace & _
                         " AND (" & sIncludedDates & ") " & _
                         " AND (Quantity <> 0)"
            
            Set rstDistinctProductsForTheTransactionPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                    pSource:=sSQLQuery, _
                                                                    pSourceType:=adCmdText, _
                                                                    pErrMsg:=strErrMsg)
            If Not (rstDistinctProductsForTheTransactionPeriod.BOF And _
                    rstDistinctProductsForTheTransactionPeriod.EOF) Then
                Call subWriteSizingArraysMessageToStatusBar
                ' Before sizing the array, check if there are any actual 'cigarettes'
                lTotalNumberOfProductsForTheTransactionPeriod = gconZeroValue
                Do Until rstDistinctProductsForTheTransactionPeriod.EOF
                    If fbProductIsIncludedInThisStickReport(rstDistinctProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)) Then
                        lTotalNumberOfProductsForTheTransactionPeriod = lTotalNumberOfProductsForTheTransactionPeriod + 1
                        ' Resize the array that holds the list of unique barcodes PALb50
                        ReDim Preserve sArrBarcode(lTotalNumberOfProductsForTheTransactionPeriod) 'PAL
                        sArrBarcode(lTotalNumberOfProductsForTheTransactionPeriod) = rstDistinctProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField) 'PAL
                    
                    End If
                    rstDistinctProductsForTheTransactionPeriod.MoveNext
                Loop
                rstDistinctProductsForTheTransactionPeriod.Close ' PALb50
                Set rstDistinctProductsForTheTransactionPeriod = Nothing
                If lTotalNumberOfProductsForTheTransactionPeriod <> 0 Then
                    ReDim Varrsalesdata(conSortIndex To conTotalRebate, _
                                        1 To lTotalNumberOfProductsForTheTransactionPeriod) As Variant
                    'use the tabstop array to store the right justified position
                    ReDim iArrTabStop(conSubSupplier To conSubTotalRebate) As Integer
                    'truncate if the docket printer is enabled in the defaults
                    If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                        iArrTabStop(conSubSupplier) = gconTruncateDescriptionBriefAt + _
                                                      Len(gconTruncateCharacter) + _
                                                      gconTruncateExtensionWidth + _
                                                      Len(gconSpace) + _
                                                      Len(gconStandardQuantityFormat)
                        
                        iArrTabStop(conSubMarketShare) = iArrTabStop(conSubSupplier) + _
                                                         Len(gconStandard4x3Format) + _
                                                         Len(gconSpace)
                        
                        iArrTabStop(conThousandsOfSticks) = iArrTabStop(conSubMarketShare) + _
                                                         Len(gconStandard4x3Format) + _
                                                         Len(gconSpace)
                        
                        iArrTabStop(conSubAverageRebate) = iArrTabStop(conThousandsOfSticks) + _
                                                           Len(gcon5DigitDollarFormat) + _
                                                           Len(gconSpace)
                        
                        iArrTabStop(conSubTotalRebate) = iArrTabStop(conSubAverageRebate) + _
                                                         Len(gcon6DigitDollarFormat) + _
                                                         Len(gconSpace)
                    Else
                        'note, market share (inclusive) onwards are RIGHT justified tabs
                        iArrTabStop(conSubSupplier) = 42
                        iArrTabStop(conSubMarketShare) = 67
                        iArrTabStop(conSubThousandsOfSticks) = 80
                        iArrTabStop(conSubAverageRebate) = 94
                        iArrTabStop(conSubTotalRebate) = 108
                    End If
                    
                    
                    Call subWriteCollatingMessageToStatusBar
                    
                    cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                    dTotalThousandsOfSticksForTheTransactionPeriod = gconZeroValue
                    lArrayRowIndex = gconZeroValue
                    
                    ' Do Until rstDistinctProductsForTheTransactionPeriod.EOF PALb50
                    For BarcodeArrayIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod ' PALb50
                        sSQLQuery = "SELECT * FROM LiveData " & vbNewLine & _
                                    "WHERE FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & gconSpace & _
                                     " AND (" & sIncludedDates & ") " & _
                                     " AND (Barcode = " & SqlQ(sArrBarcode(BarcodeArrayIndex)) & ")" ' PALb50
                        Set rstAllSameProductsForTheTransactionPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                               pSource:=sSQLQuery, _
                                                                               pSourceType:=adCmdText, _
                                                                               pErrMsg:=strErrMsg)
                        'has to be more than zero records, so don't waste time testing for it
                        ' If fbProductIsIncludedInThisStickReport(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)) Then PALb50
                        lArrayRowIndex = lArrayRowIndex + 1
                        Call subDisplayCurrentRecordToUser( _
                             lArrayRowIndex, _
                             lTotalNumberOfProductsForTheTransactionPeriod)
                        If optStickReportDescription(0) Then
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                fsDescriptionFrom(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField))
                        Else
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)
                        End If
                        
                        Varrsalesdata(conSupplier, lArrayRowIndex) = _
                           fsSupplierIDFrom(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField))
                        
                        Do Until rstAllSameProductsForTheTransactionPeriod.EOF
                            'avert an overflow divide by zero
                            If rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableQuantityField) <> gconZeroValue Then
    
                                'performance consideration as only 1 query is required
                                dThousandsOfSticksThisTransaction = flSticksFrom(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)) * _
                                                                    rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableQuantityField) _
                                                                    / 1000
                                
                                Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) = _
                                    Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) + _
                                    dThousandsOfSticksThisTransaction
                                
                                dTotalThousandsOfSticksForTheTransactionPeriod = _
                                    dTotalThousandsOfSticksForTheTransactionPeriod + _
                                    dThousandsOfSticksThisTransaction
    
                            End If
                            rstAllSameProductsForTheTransactionPeriod.MoveNext
                        Loop
                        ' End If PALb50
                        'Loop
                        rstAllSameProductsForTheTransactionPeriod.Close
                        Set rstAllSameProductsForTheTransactionPeriod = Nothing
                        ' rstDistinctProductsForTheTransactionPeriod.MoveNext PALb50
                    'Loop ' PALb50
                    Next BarcodeArrayIndex  ' PALb50
                    
                    Call subWriteSortingMessageToStatusBar
                    
                    'sort by description
                    Do
                        bIndexSwapped = False
                        For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod - 1
                            If (Varrsalesdata(conProduct, lArrayRowIndex) > _
                                Varrsalesdata(conProduct, lArrayRowIndex + 1)) Then 'swap
                                For iArrayColumnIndex = conProduct To conTotalRebate
                                    vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                    Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                    Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                    bIndexSwapped = True
                                Next iArrayColumnIndex
                            End If
                        Next lArrayRowIndex
                    Loop While bIndexSwapped
                    
                    'determine market share (requires number of included suppliers)
                    
                    'fill the first supplyar array with the first product supplyar
                    iTotalSuppliersForTheTransactionPeriod = 1
                    ReDim iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) As Integer
                    iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) = Varrsalesdata(conSupplier, 1)
                    
                    For lArrayRowIndex = 2 To lTotalNumberOfProductsForTheTransactionPeriod
                        iCurrentProductSupplierID = Varrsalesdata(conSupplier, lArrayRowIndex)
                        
                        'compare the current producsupplier to every one already recorded,
                        'if not there then increment array and add it
                        For iSupplierIndex = 1 To iTotalSuppliersForTheTransactionPeriod
                            If iCurrentProductSupplierID = iArrIncludedSuppliers(iSupplierIndex) Then
                                GoTo NotSummarisedSupplierAlreadyRecorded
                            End If
                        Next iSupplierIndex
                        'if here, then the current supplyar has not already been recorded
                        iTotalSuppliersForTheTransactionPeriod = iTotalSuppliersForTheTransactionPeriod + 1
                        ReDim Preserve iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) As Integer
                        iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) = iCurrentProductSupplierID
NotSummarisedSupplierAlreadyRecorded:
                    Next lArrayRowIndex
                    
                    'sort supplyar array
                    Do
                        bIndexSwapped = False
                        For lArrayRowIndex = 1 To iTotalSuppliersForTheTransactionPeriod - 1
                            If (iArrIncludedSuppliers(lArrayRowIndex) > _
                                iArrIncludedSuppliers(lArrayRowIndex + 1)) Then 'swap
                                iPlaceHolder = iArrIncludedSuppliers(lArrayRowIndex)
                                iArrIncludedSuppliers(lArrayRowIndex) = iArrIncludedSuppliers(lArrayRowIndex + 1)
                                iArrIncludedSuppliers(lArrayRowIndex + 1) = iPlaceHolder
                                bIndexSwapped = True
                            End If
                        Next lArrayRowIndex
                    Loop While bIndexSwapped
    
                    If optSendStickReportToDisplay Then
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise))
                        For iSupplierIndex = 1 To iTotalSuppliersForTheTransactionPeriod
                            dThousandsOfSticksThisSupplierForTheTransactionPeriod = gconZeroValue
                            For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod
                                If Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) <> gconZeroValue Then
                                    If Varrsalesdata(conSupplier, lArrayRowIndex) = iArrIncludedSuppliers(iSupplierIndex) Then
                                        
                                        If optStickReportSummaryType(0) Then
                                        Set gvListItem = lvwStickReport.ListItems.Add()
                                        gvListItem.Text = Varrsalesdata(conProduct, lArrayRowIndex)
                                        
                                        Call gsubAddSubItemToListview( _
                                                 fsSupplierNameFrom(Varrsalesdata(conSupplier, lArrayRowIndex)), conSubSupplier)
                                                                
                                        Call gsubAddSubItemToListview( _
                                                Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / _
                                                       dTotalThousandsOfSticksForTheTransactionPeriod * 100 _
                                                       , gconStandard4x3Format), conSubMarketShare)
                                                                                                    
                                        Call gsubAddSubItemToListview( _
                                                Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) _
                                                       , gconStandard4x3Format), conSubThousandsOfSticks)
                                        End If
                                        dThousandsOfSticksThisSupplierForTheTransactionPeriod = dThousandsOfSticksThisSupplierForTheTransactionPeriod + _
                                                                                                Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)
                                        
                                    End If 'is this product for this supplyar ?
                                End If 'qty = zero
                            Next lArrayRowIndex
                            
                            Set gvListItem = lvwStickReport.ListItems.Add()
                            gvListItem.Text = "Total"
                            Call gsubAddSubItemToListview( _
                                     fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)), conSubSupplier)
    
                            Call gsubAddSubItemToListview( _
                                     Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                            dTotalThousandsOfSticksForTheTransactionPeriod * 100 _
                                            , gconStandard4x3Format), conSubMarketShare)
                            
                            Call gsubAddSubItemToListview( _
                                     Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, _
                                            gconStandard4x3Format), conSubThousandsOfSticks)
                            'leave a gap
                            If optStickReportSummaryType(0) Then
                                Set gvListItem = lvwStickReport.ListItems.Add()
                                gvListItem.Text = gconSpace
                            End If
                        Next iSupplierIndex
                        
                        'expose the totals
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = "Total"
                        Call gsubAddSubItemToListview( _
                                 Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format), conSubThousandsOfSticks)
                        
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = gconSpace
                        Me.Refresh
                    ElseIf optSendStickReportToPrinter Then
                        On Error GoTo NotSummarisedPrinterErrorHandler
                        
                        cdlTSGDataWarehouse.ShowPrinter
                        iNumberOfCopies = cdlTSGDataWarehouse.Copies
                        Me.Refresh
                        
                        For iPageNumber = 1 To iNumberOfCopies
                            'headings
                            Printer.Print "Product"; _
                                           Tab(iArrTabStop(conThousandsOfSticks) - Len("  Qty")); _
                                          "  Qty"; _
                                           Tab(iArrTabStop(2) - Len("Avg unit")); _
                                          "Avg unit"; _
                                           Tab(iArrTabStop(conTotalRebate) - Len("Tot (inc)")); _
                                          "Tot (inc)"
                            'leave a gap
                            Printer.Print gconSpace
                            
                            For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod
                                If Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) <> gconZeroValue Then
                                    Printer.Print _
                                        Varrsalesdata(conProduct, lArrayRowIndex); _
                                        Tab(iArrTabStop(conThousandsOfSticks) - Len(Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                        Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandardQuantityFormat); _
                                        Tab(iArrTabStop(2) - Len(Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)) / _
                                               Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                        Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)) / _
                                               Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                        Tab(iArrTabStop(conTotalRebate) - Len(Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                        Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)), gcon5DigitDollarFormat)
                                End If
                            Next lArrayRowIndex
                            
                            'leave a gap
                            Printer.Print gconSpace
                            
                            'expose the totals
                            Printer.Print "Total"; _
                                        Tab(iArrTabStop(conThousandsOfSticks) - Len(Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandardQuantityFormat))); _
                                        Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandardQuantityFormat); _
                                        Tab(iArrTabStop(conTotalRebate) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); _
                                        Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                            
                        Next iPageNumber
                        Printer.EndDoc
                        MsgBox "Report was successfully submitted to the selected printer", _
                                vbInformation, gconReportManager
                    Else 'must be to file
                        'headings
                        If Not fHeadingDone Then
                            fHeadingDone = True
                            If chkStickReportTabDelimited Then
                                Print #intFileNum, "Product"; _
                                           vbTab; _
                                          "Supplier"; _
                                           vbTab; _
                                          "Market %"; _
                                           vbTab; _
                                          "Sticks/1000"; _
                                           vbTab
                            Else
                                Print #intFileNum, "Product"; _
                                           Tab(iArrTabStop(conSubSupplier)); _
                                          "Supplier"; _
                                           Tab(iArrTabStop(conSubMarketShare) - Len("Market %")); _
                                          "Market %"; _
                                           Tab(iArrTabStop(conSubThousandsOfSticks) - Len("Sticks/1000")); _
                                          "Sticks/1000"
                            End If
                        End If
                        'leave a gap
                        'Print #intFileNum, gconSpace
                        Print #intFileNum, GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                        'leave a gap
                        Print #intFileNum, gconSpace
                    
                   
                        For iSupplierIndex = 1 To iTotalSuppliersForTheTransactionPeriod
                            dThousandsOfSticksThisSupplierForTheTransactionPeriod = gconZeroValue
                            For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod
                                If Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) <> gconZeroValue Then
                                    If Varrsalesdata(conSupplier, lArrayRowIndex) = iArrIncludedSuppliers(iSupplierIndex) Then
                                        If chkStickReportTabDelimited Then
                                            If optStickReportSummaryType(0) Then
                                            Print #intFileNum, _
                                                Varrsalesdata(conProduct, lArrayRowIndex); _
                                                vbTab; _
                                                fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                                vbTab; ; _
                                                Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), _
                                                vbTab; _
                                                Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandard4x3Format)
                                            End If
                                                
                                                dThousandsOfSticksThisSupplierForTheTransactionPeriod = dThousandsOfSticksThisSupplierForTheTransactionPeriod + _
                                                                                                            Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)
                                                
                                        Else
                                            If optStickReportSummaryType(0) Then
                                                Print #intFileNum, _
                                                Varrsalesdata(conProduct, lArrayRowIndex); _
                                                Tab(iArrTabStop(conSubSupplier)); _
                                                fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                                Tab(iArrTabStop(conSubMarketShare) - Len(Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format))); _
                                                Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), _
                                                Tab(iArrTabStop(conSubThousandsOfSticks) - Len(Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandard4x3Format))); _
                                                Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandard4x3Format)
                                            End If
                                            dThousandsOfSticksThisSupplierForTheTransactionPeriod = dThousandsOfSticksThisSupplierForTheTransactionPeriod + _
                                                                                                            Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)
                                        End If 'tab delimited ?
                                    End If 'for this supplyar ?
                                End If 'above zerovalue ?
                            Next lArrayRowIndex
                            
                            'expose totals for this supplyar
                            If chkStickReportTabDelimited Then
                                Print #intFileNum, "Total"; _
                                           vbTab; _
                                           fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                           vbTab; _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                                              dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), ; _
                                           vbTab; _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, gconStandard4x3Format);
                            Else
                                Print #intFileNum, "Total"; _
                                           Tab(iArrTabStop(conSubSupplier)); _
                                           fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                           Tab(iArrTabStop(conSubMarketShare) - Len(Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                                              dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format))); _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                                              dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), ; _
                                           Tab(iArrTabStop(conSubThousandsOfSticks) - Len(Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, gconStandard4x3Format))); _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, gconStandard4x3Format);
                            End If
                            
                            Print #intFileNum,
                 
                            'leave gap
                            If optStickReportSummaryType(0) Then
                            Print #intFileNum, gconSpace
                            End If
                        Next iSupplierIndex
                        
                        'expose the totals
                        If chkStickReportTabDelimited Then
                            Print #intFileNum, "Total"; _
                                       vbTab; vbTab; vbTab; _
                                       Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format) '; _
                                       'vbtab; _
                                       'Format(cTotalRebateForTheTransactionPeriod / dTotalThousandsOfSticksForTheTransactionPeriod, gcon5DigitDollarFormat); _
                                       'vbtab; _
                                       'Format(cTotalRebateForTheTransactionPeriod, gcon6DigitDollarFormat)
                        Else 'normal
                            Print #intFileNum, "Total"; _
                                       Tab(iArrTabStop(conSubThousandsOfSticks) - Len(Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format))); _
                                       Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format) '; _
                                       'Tab(iArrTabStop(conSubAverageRebate) - Len(Format(cTotalRebateForTheTransactionPeriod / dTotalThousandsOfSticksForTheTransactionPeriod, gcon5DigitDollarFormat))); _
                                       'Format(cTotalRebateForTheTransactionPeriod / dTotalThousandsOfSticksForTheTransactionPeriod, gcon5DigitDollarFormat); _
                                       'Tab(iArrTabStop(conSubTotalRebate) - Len(Format(cTotalRebateForTheTransactionPeriod, gcon6DigitDollarFormat))); _
                                       'Format(cTotalRebateForTheTransactionPeriod, gcon6DigitDollarFormat)
                        End If 'tab delimited ?
                        Print #intFileNum, vbCrLf
                    End If 'report destination
                    On Error GoTo 0
                    'conserve memory
                    Erase iArrTabStop
                    Erase Varrsalesdata
                    Erase sArrBarcode ' PALb50
                Else
                    If optSendStickReportToDisplay Then
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = "No sticks for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = gconSpace
                    ElseIf optSendStickReportToPrinter Then
                        Printer.Print "No sticks for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                        Printer.Print vbCrLf
                    Else 'is destined for the file
                        Print #intFileNum, "No sticks for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                        Print #intFileNum, vbCrLf
                    End If
                End If
            Else 'no transactions for the date
                If optSendStickReportToDisplay Then
                    Set gvListItem = lvwStickReport.ListItems.Add()
                    gvListItem.Text = "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                    Set gvListItem = lvwStickReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                ElseIf optSendStickReportToPrinter Then
                    Printer.Print "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                    Printer.Print vbCrLf
                Else 'is destined for the file
                    Print #intFileNum, "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sTransactionPeriod
                    Print #intFileNum, vbCrLf
                End If
            End If 'any transactions for the report date ?
            With stb
                .SimpleText = ""
                .Refresh
            End With
            'rstDistinctProductsForTheTransactionPeriod.Close PALb50
        Next iCurrentFranchise

NotSummarisedTidyUp:
        On Error GoTo 0
        
        If optSendStickReportToDisplay Then
            'do nothing
        ElseIf optSendStickReportToPrinter Then
            Printer.EndDoc
            MsgBox "Report was successfully submitted to the selected printer", _
                    vbInformation, gconReportManager
        Else 'was to file
            Close #intFileNum
            Call subSetStickReportViewButton
            MsgBox "Report was successfully sent to - " & gsStickReportPathAndFilename & _
                   ". Use the 'View' button to display", _
                    vbInformation, gconReportManager
        End If
        
'==============================================================================================================
    Else 'summarised
'--------------------------------------------------------------------------------------------------------------
'  AUrban SUMMARISED REPORT: Procedure is a candidate for splitting above and below here into two procedures
'--------------------------------------------------------------------------------------------------------------
        If optStickReportOnSelectedFranchisesOnly(0) Then
            'build a query spec for all selected franchises
            For lArrayRowIndex = gconDisplayFirstItem To lstStickReportsFranchiseBusinessName.ListCount - 1
                If lstStickReportsFranchiseBusinessName.Selected(lArrayRowIndex) Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    
                    sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                              lstStickReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
                    
                    sIncludedFranchiseIDs = sIncludedFranchiseIDs & _
                                            gconLiveDataTableTSGFranchiseIDField & " = " & _
                                            fsFranchiseIDFrom(lstStickReportsFranchiseBusinessName.List(lArrayRowIndex)) & " OR "
                End If
            Next lArrayRowIndex
            
            If iNumberOfFranchisesIncluded Then
                'get rid of the last delimiters
                sIncludedFranchiseNames = Left(sIncludedFranchiseNames, _
                                          Len(sIncludedFranchiseNames) - Len(", "))
                
                sIncludedFranchiseIDs = Left(sIncludedFranchiseIDs, _
                                        Len(sIncludedFranchiseIDs) - Len(" OR "))
            
                sFranchiseMessageBox = " for " & sIncludedFranchiseNames
            End If
            
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & vbNewLine & _
                        "WHERE (" & sIncludedFranchiseIDs & ") " & _
                          "AND (" & sIncludedDates & ") " & _
                          "AND (Quantity  <> 0)"
        Else 'all franchises option was selected
            iNumberOfFranchisesIncluded = lstStickReportsFranchiseBusinessName.ListCount
            
            sFranchiseMessageBox = ""
            
            sIncludedFranchiseNames = gconAllFranchises
            
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & vbNewLine & _
                        "WHERE (" & sIncludedDates & ") AND (Quantity <> 0)"
        End If
        
        If iNumberOfFranchisesIncluded Then 'franchises are included
            If iNumberOfFranchisesIncluded > 1 Then
                sPlural = "s"
            End If
            
            Set rstDistinctProductsForTheTransactionPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                    pSource:=sSQLQuery, _
                                                                    pSourceType:=adCmdText, _
                                                                    pErrMsg:=strErrMsg)
                                                                       
            If Not (rstDistinctProductsForTheTransactionPeriod.BOF And _
                    rstDistinctProductsForTheTransactionPeriod.EOF) Then
                Call subWriteSizingArraysMessageToStatusBar
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conSubSupplier To conSubTotalRebate) As Integer
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conSubSupplier) = gconTruncateDescriptionBriefAt + _
                                                  Len(gconTruncateCharacter) + _
                                                  gconTruncateExtensionWidth + _
                                                  Len(gconSpace) + _
                                                  Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conSubMarketShare) = iArrTabStop(conSubSupplier) + _
                                                     Len(gconStandard4x3Format) + _
                                                     Len(gconSpace)
                    
                    iArrTabStop(conThousandsOfSticks) = iArrTabStop(conSubMarketShare) + _
                                                     Len(gconStandard4x3Format) + _
                                                     Len(gconSpace)
                    
                    iArrTabStop(conSubAverageRebate) = iArrTabStop(conThousandsOfSticks) + _
                                                       Len(gcon5DigitDollarFormat) + _
                                                       Len(gconSpace)
                    
                    iArrTabStop(conSubTotalRebate) = iArrTabStop(conSubAverageRebate) + _
                                                     Len(gcon6DigitDollarFormat) + _
                                                     Len(gconSpace)
                Else
                    'note, market share (inclusive) onwards are RIGHT justified tabs
                    iArrTabStop(conSubSupplier) = 42
                    iArrTabStop(conSubMarketShare) = 67
                    iArrTabStop(conSubThousandsOfSticks) = 80
                    iArrTabStop(conSubAverageRebate) = 94
                    iArrTabStop(conSubTotalRebate) = 108
                End If
                
                'size the array
                Do Until rstDistinctProductsForTheTransactionPeriod.EOF
                    If fbProductIsIncludedInThisStickReport(rstDistinctProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)) Then
                        lTotalNumberOfProductsForTheTransactionPeriod = lTotalNumberOfProductsForTheTransactionPeriod + 1
                        ' Resize the array that holds the list of unique barcodes PALb50
                        ReDim Preserve sArrBarcode(lTotalNumberOfProductsForTheTransactionPeriod) 'PAL
                        sArrBarcode(lTotalNumberOfProductsForTheTransactionPeriod) = rstDistinctProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField) 'PAL
                    End If
                    rstDistinctProductsForTheTransactionPeriod.MoveNext
                Loop
                rstDistinctProductsForTheTransactionPeriod.Close ' PALb50
                Set rstDistinctProductsForTheTransactionPeriod = Nothing
                
                ReDim Varrsalesdata(conSortIndex To conTotalRebate, _
                                    1 To lTotalNumberOfProductsForTheTransactionPeriod) As Variant
                
                Call subWriteCollatingMessageToStatusBar
                
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                dTotalThousandsOfSticksForTheTransactionPeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                
                For BarcodeArrayIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod ' PALb50
                'Do Until rstDistinctProductsForTheTransactionPeriod.EOF
                    If optStickReportOnSelectedFranchisesOnly(0) Then
                       sSQLQuery = "SELECT * FROM LiveData " & _
                                    "WHERE (" & sIncludedFranchiseIDs & ") " & _
                                     " AND (" & sIncludedDates & ") " & _
                                     " AND (Barcode  = " & SqlQ(sArrBarcode(BarcodeArrayIndex)) & ")" ' PALb50
                    Else 'all franchises, no requirement to discriminate (performance reasons)
                        sSQLQuery = "SELECT * FROM LiveData " & _
                                    "WHERE (" & sIncludedDates & ") " & _
                                     " AND (Barcode = " & SqlQ(sArrBarcode(BarcodeArrayIndex)) & ")" ' PALb50
                    End If
                    
                    Set rstAllSameProductsForTheTransactionPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                           pSource:=sSQLQuery, _
                                                                           pSourceType:=adCmdText, _
                                                                           pErrMsg:=strErrMsg)
                    'has to be more than zero records, so don't waste time testing for it
                    'If fbProductIsIncludedInThisStickReport(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)) Then
                        lArrayRowIndex = lArrayRowIndex + 1
                        If optStickReportDescription(0) Then
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                fsDescriptionFrom(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField))
                        Else
                            Varrsalesdata(conProduct, lArrayRowIndex) = _
                                rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)
                        End If
                        
                        Varrsalesdata(conSupplier, lArrayRowIndex) = _
                           fsSupplierIDFrom(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField))
                        
                        Do Until rstAllSameProductsForTheTransactionPeriod.EOF
                            'avert an overflow divide by zero
                            If rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableQuantityField) <> gconZeroValue Then

                                'performance consideration as only 1 query is required
                                dThousandsOfSticksThisTransaction = flSticksFrom(rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableBarcodeField)) * _
                                                                    rstAllSameProductsForTheTransactionPeriod(gconLiveDataTableQuantityField) _
                                                                    / 1000
                                
                                Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) = _
                                    Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) + _
                                    dThousandsOfSticksThisTransaction
                                
                                dTotalThousandsOfSticksForTheTransactionPeriod = _
                                    dTotalThousandsOfSticksForTheTransactionPeriod + _
                                    dThousandsOfSticksThisTransaction

                            End If
                            rstAllSameProductsForTheTransactionPeriod.MoveNext
                        Loop
                    'End If
                    'Loop
                    rstAllSameProductsForTheTransactionPeriod.Close
                    Set rstAllSameProductsForTheTransactionPeriod = Nothing
                    'rstDistinctProductsForTheTransactionPeriod.MoveNext
                'Loop
                Next BarcodeArrayIndex
                
                Call subWriteSortingMessageToStatusBar
                
                'sort by description
                Do
                    bIndexSwapped = False
                    For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod - 1
                        If (Varrsalesdata(conProduct, lArrayRowIndex) > _
                            Varrsalesdata(conProduct, lArrayRowIndex + 1)) Then 'swap
                            For iArrayColumnIndex = conProduct To conTotalRebate
                                vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                bIndexSwapped = True
                            Next iArrayColumnIndex
                        End If
                    Next lArrayRowIndex
                Loop While bIndexSwapped
                
                'determine market share (requires number of included suppliers)
                
                'fill the first supplyar array with the first product supplyar
                iTotalSuppliersForTheTransactionPeriod = 1
                ReDim iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) As Integer
                iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) = Varrsalesdata(conSupplier, 1)
                
                For lArrayRowIndex = 2 To lTotalNumberOfProductsForTheTransactionPeriod
                    iCurrentProductSupplierID = Varrsalesdata(conSupplier, lArrayRowIndex)
                    
                    'compare the current producsupplier to every one already recorded,
                    'if not there then increment array and add it
                    For iSupplierIndex = 1 To iTotalSuppliersForTheTransactionPeriod
                        If iCurrentProductSupplierID = iArrIncludedSuppliers(iSupplierIndex) Then
                            GoTo SummarisedSupplierAlreadyRecorded
                        End If
                    Next iSupplierIndex
                    'if here, then the current supplyar has not already been recorded
                    iTotalSuppliersForTheTransactionPeriod = iTotalSuppliersForTheTransactionPeriod + 1
                    ReDim Preserve iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) As Integer
                    iArrIncludedSuppliers(iTotalSuppliersForTheTransactionPeriod) = iCurrentProductSupplierID
SummarisedSupplierAlreadyRecorded:
                Next lArrayRowIndex
                
                'sort supplyar array
                Do
                    bIndexSwapped = False
                    For lArrayRowIndex = 1 To iTotalSuppliersForTheTransactionPeriod - 1
                        If (iArrIncludedSuppliers(lArrayRowIndex) > _
                            iArrIncludedSuppliers(lArrayRowIndex + 1)) Then 'swap
                            iPlaceHolder = iArrIncludedSuppliers(lArrayRowIndex)
                            iArrIncludedSuppliers(lArrayRowIndex) = iArrIncludedSuppliers(lArrayRowIndex + 1)
                            iArrIncludedSuppliers(lArrayRowIndex + 1) = iPlaceHolder
                            bIndexSwapped = True
                        End If
                    Next lArrayRowIndex
                Loop While bIndexSwapped

                If optSendStickReportToDisplay Then
                    For iSupplierIndex = 1 To iTotalSuppliersForTheTransactionPeriod
                        dThousandsOfSticksThisSupplierForTheTransactionPeriod = gconZeroValue
                        For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod
                            If Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) <> gconZeroValue Then
                                If Varrsalesdata(conSupplier, lArrayRowIndex) = iArrIncludedSuppliers(iSupplierIndex) Then
                                    
                                    Set gvListItem = lvwStickReport.ListItems.Add()
                                    gvListItem.Text = Varrsalesdata(conProduct, lArrayRowIndex)
                                    
                                    Call gsubAddSubItemToListview( _
                                             fsSupplierNameFrom(Varrsalesdata(conSupplier, lArrayRowIndex)), conSubSupplier)
                                                            
                                    Call gsubAddSubItemToListview( _
                                            Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / _
                                                   dTotalThousandsOfSticksForTheTransactionPeriod * 100 _
                                                   , gconStandard4x3Format), conSubMarketShare)
                                                                                                
                                    Call gsubAddSubItemToListview( _
                                            Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) _
                                                   , gconStandard4x3Format), conSubThousandsOfSticks)
                                                            
                                    dThousandsOfSticksThisSupplierForTheTransactionPeriod = dThousandsOfSticksThisSupplierForTheTransactionPeriod + _
                                                                                            Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)
                                    
                                End If 'is this product for this supplyar ?
                            End If 'qty = zero
                        Next lArrayRowIndex
                        
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = "Total"
                        Call gsubAddSubItemToListview( _
                                 fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)), conSubSupplier)

                        Call gsubAddSubItemToListview( _
                                 Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                        dTotalThousandsOfSticksForTheTransactionPeriod * 100 _
                                        , gconStandard4x3Format), conSubMarketShare)
                        
                        Call gsubAddSubItemToListview( _
                                 Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, _
                                        gconStandard4x3Format), conSubThousandsOfSticks)
                        'leave a gap
                        Set gvListItem = lvwStickReport.ListItems.Add()
                        gvListItem.Text = gconSpace
                    
                    Next iSupplierIndex
                    
                    'expose the totals
                    Set gvListItem = lvwStickReport.ListItems.Add()
                    gvListItem.Text = "Total"
                    Call gsubAddSubItemToListview( _
                             Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format), conSubThousandsOfSticks)
                    
                ElseIf optSendStickReportToPrinter Then
                    
                    On Error GoTo SummarisedPrinterErrorHandler
                    
                    cdlTSGDataWarehouse.ShowPrinter
                    iNumberOfCopies = cdlTSGDataWarehouse.Copies
                    Me.Refresh
                    
                    For iPageNumber = 1 To iNumberOfCopies
                        Printer.Print "Product"; _
                                       Tab(iArrTabStop(conThousandsOfSticks) - Len("  Qty")); _
                                      "  Qty"; _
                                       Tab(iArrTabStop(2) - Len("Avg unit")); _
                                      "Avg unit"; _
                                       Tab(iArrTabStop(conTotalRebate) - Len("Tot (inc)")); _
                                      "Tot (inc)"
                        'leave a gap
                        Printer.Print gconSpace
                        
                        For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod
                            If Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) <> gconZeroValue Then
                                Printer.Print _
                                    Varrsalesdata(conProduct, lArrayRowIndex); _
                                    Tab(iArrTabStop(conThousandsOfSticks) - Len(Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                    Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    Tab(iArrTabStop(2) - Len(Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conTotalRebate) - Len(Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalRebate, lArrayRowIndex)), gcon5DigitDollarFormat)
                            End If
                        Next lArrayRowIndex
                        
                        'leave a gap
                        Printer.Print gconSpace
                        
                        'expose the total
                        Printer.Print "Total"; _
                                    Tab(iArrTabStop(conThousandsOfSticks) - Len(Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandardQuantityFormat))); _
                                    Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandardQuantityFormat); _
                                    Tab(iArrTabStop(conTotalRebate) - Len(Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat))); _
                                    Format(cTotalSalesIncludingTaxForThePeriod, gcon5DigitDollarFormat)
                        
                    Next iPageNumber
                    Printer.EndDoc
                    MsgBox "Report was successfully submitted to the selected printer", _
                            vbInformation, gconReportManager
                Else 'must be to file
                    If fbNewStickReportDocumentEnvironmentWasSuccessfullyPrepared Then
                        intFileNum = FreeFile   ' Get unused file
                        Open gsStickReportPathAndFilename For Output As #intFileNum
                        Print #intFileNum, "Tobacco Station" & sPlural & _
                                  " - " & sIncludedFranchiseNames
                        'leave a dual gap
                        Print #intFileNum, vbCrLf
                        
                        Print #intFileNum, conReportType & sTransactionPeriod
                        Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                        'leave a dual gap
                        Print #intFileNum, vbCrLf
                        
                        'headings
                        If chkStickReportTabDelimited Then
                            Print #intFileNum, "Product"; _
                                       vbTab; _
                                      "Supplier"; _
                                       vbTab; _
                                      "Market %"; _
                                       vbTab; _
                                      "Sticks/1000"; _
                                       vbTab
                        Else
                            Print #intFileNum, "Product"; _
                                       Tab(iArrTabStop(conSubSupplier)); _
                                      "Supplier"; _
                                       Tab(iArrTabStop(conSubMarketShare) - Len("Market %")); _
                                      "Market %"; _
                                       Tab(iArrTabStop(conSubThousandsOfSticks) - Len("Sticks/1000")); _
                                      "Sticks/1000"
                        End If
                        'leave a gap
                        Print #intFileNum, gconSpace
                    
                        For iSupplierIndex = 1 To iTotalSuppliersForTheTransactionPeriod
                            dThousandsOfSticksThisSupplierForTheTransactionPeriod = gconZeroValue
                            For lArrayRowIndex = 1 To lTotalNumberOfProductsForTheTransactionPeriod
                                If Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) <> gconZeroValue Then
                                    If Varrsalesdata(conSupplier, lArrayRowIndex) = iArrIncludedSuppliers(iSupplierIndex) Then
                                        If chkStickReportTabDelimited Then
                                            Print #intFileNum, _
                                                Varrsalesdata(conProduct, lArrayRowIndex); _
                                                vbTab; _
                                                fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                                vbTab; ; _
                                                Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), _
                                                vbTab; _
                                                Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandard4x3Format)
                                        Else
                                            Print #intFileNum, _
                                                Varrsalesdata(conProduct, lArrayRowIndex); _
                                                Tab(iArrTabStop(conSubSupplier)); _
                                                fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                                Tab(iArrTabStop(conSubMarketShare) - Len(Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format))); _
                                                Format(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex) / dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), _
                                                Tab(iArrTabStop(conSubThousandsOfSticks) - Len(Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandard4x3Format))); _
                                                Format(Val(Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)), gconStandard4x3Format)
                                        End If 'tab delimited ?
                                        dThousandsOfSticksThisSupplierForTheTransactionPeriod = dThousandsOfSticksThisSupplierForTheTransactionPeriod + _
                                                                                                Varrsalesdata(conThousandsOfSticks, lArrayRowIndex)
                                    End If 'for this supplyar ?
                                End If 'above zerovalue ?
                            Next lArrayRowIndex
                            
                            'expose total for this supplyar
                            If chkStickReportTabDelimited Then
                                Print #intFileNum, "Total"; _
                                           vbTab; _
                                           fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                           vbTab; _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                                              dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), ; _
                                           vbTab; _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, gconStandard4x3Format);
                            Else
                                Print #intFileNum, "Total"; _
                                           Tab(iArrTabStop(conSubSupplier)); _
                                           fsSupplierNameFrom(iArrIncludedSuppliers(iSupplierIndex)); _
                                           Tab(iArrTabStop(conSubMarketShare) - Len(Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                                              dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format))); _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod / _
                                                              dTotalThousandsOfSticksForTheTransactionPeriod * 100, gconStandard4x3Format), ; _
                                           Tab(iArrTabStop(conSubThousandsOfSticks) - Len(Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, gconStandard4x3Format))); _
                                           Format(dThousandsOfSticksThisSupplierForTheTransactionPeriod, gconStandard4x3Format);
                            End If
                            Print #intFileNum,
                            'leave a gap
                            Print #intFileNum, gconSpace
                        Next iSupplierIndex
                        
                        'expose the totals
                        If chkStickReportTabDelimited Then
                            Print #intFileNum, "Total"; _
                                       vbTab; vbTab; vbTab; _
                                       Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format) '; _
                                       'vbtab;
                                       'Format(cTotalRebateForTheTransactionPeriod / dTotalThousandsOfSticksForTheTransactionPeriod, gcon5DigitDollarFormat); _
                                       'vbtab; _
                                       'Format(cTotalRebateForTheTransactionPeriod, gcon6DigitDollarFormat)
                        Else
                            Print #intFileNum, "Total"; _
                                       Tab(iArrTabStop(conSubThousandsOfSticks) - Len(Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format))); _
                                       Format(dTotalThousandsOfSticksForTheTransactionPeriod, gconStandard4x3Format) '; _
                                       'Tab(iArrTabStop(conSubAverageRebate) - Len(Format(cTotalRebateForTheTransactionPeriod / dTotalThousandsOfSticksForTheTransactionPeriod, gcon5DigitDollarFormat))); _
                                       'Format(cTotalRebateForTheTransactionPeriod / dTotalThousandsOfSticksForTheTransactionPeriod, gcon5DigitDollarFormat); _
                                       'Tab(iArrTabStop(conSubTotalRebate) - Len(Format(cTotalRebateForTheTransactionPeriod, gcon6DigitDollarFormat))); _
                                       'Format(cTotalRebateForTheTransactionPeriod, gcon6DigitDollarFormat)
                        End If
                        Close #intFileNum
                        Call subSetStickReportViewButton
                        MsgBox "Report was successfully sent to - " & gsStickReportPathAndFilename & _
                               ". Use the 'View' button to display", _
                                vbInformation, gconReportManager
                    Else 'environment was not created
                        MsgBox "Report was aborted", vbExclamation
                    End If 'environement created ?
                End If 'report destination

SummarisedTidyUp:
                On Error GoTo 0
                'conserve memory
                Erase iArrTabStop
                Erase Varrsalesdata
                Erase sArrBarcode
            Else 'no transactions for the date
                MsgBox "No sales transactions" & sFranchiseMessageBox & gsReportPeriodWording & sTransactionPeriod, _
                        vbInformation, gconReportManager
            End If 'any transactions for the report date ?
        
            With stb
                .SimpleText = ""
                .Refresh
            End With
            
            'rstDistinctProductsForTheTransactionPeriod.Close PALb50
        Else
            MsgBox "No franchise selected", _
                    vbExclamation, gconReportManager
        End If 'summarised franchise selected ?
    End If 'not summarised or summarised ?
    Call subPopulateStickReportRecipientListbox
    
    Exit Sub

NotSummarisedPrinterErrorHandler:
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    
    Resume NotSummarisedTidyUp

SummarisedPrinterErrorHandler:
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    
    Resume SummarisedTidyUp
    
End Sub

Private Sub cmdStockTabDelete_Click()
Dim intPrevMousePointer As Integer
Dim lngSelCount As Long
Dim strMsg As String
Dim strSeln As String
Dim strErrMsg As String
Dim vntID As Variant
Dim colIDs As VBA.Collection
Dim rstStock As ADODB.Recordset
    
    lngSelCount = lstStcokTabSelectedSoctkExport.SelCount
    If lngSelCount Then
        If lngSelCount = lstStcokTabSelectedSoctkExport.ListCount Then
            strSeln = "All Stock Items"
        Else
            strSeln = Plural(pQty:=lngSelCount, pNounSingular:="stock item") & " selected"
        End If
        strMsg = "Delete Stock Items?" & vbNewLine & strSeln
                  
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            Me.Enabled = False
            intPrevMousePointer = SetMousePointer(vbHourglass)
            
        '   qryStock not used b/c stock selected from ListBox which may or may not later be filtered to included deleted stock
            Set rstStock = GetRst(pCnn:=g.cnnDW, _
                                  pSource:="Stock", _
                                  pSourceType:=adCmdTable, _
                                  pRstType:=eEditableDynamic, _
                                  pErrMsg:=strErrMsg)
            If Not rstStock Is Nothing Then
                If Not (rstStock.EOF And rstStock.EOF) Then
                    Set colIDs = ListBoxGetCollection(pListBox:=lstStcokTabSelectedSoctkExport, pItemData:=True, pSelected:=True)
                    For Each vntID In colIDs
                        With rstStock
                            .MoveFirst
                            .Find Criteria:="Stock_ID = " & vntID
                            If Not .EOF Then
                                    .Fields!Deleted = CBoolMySql(True)
                                    StatusBar pMsg:="Deleting " & rstStock!Description.Value, pLog:=False
                                .Update
                            End If
                        End With
                    Next vntID
                    Set colIDs = Nothing
                    
                    StatusBar pMsg:=vbNullString, pLog:=False
                    subPopulateStockListboxes pExclDeletedStk:=ChkBoxToBool(chkStockTab_IncludeDeletedStock)
                    
                    rstStock.Close
                    Set rstStock = Nothing
                End If
            End If
            
            SetMousePointer intPrevMousePointer
            Me.Enabled = True
        End If
    End If

End Sub

Private Sub cmdStockTabExport_Click()
Dim intPrevMousePointer As Integer
Dim lngSelCount As Long
Dim strMsg As String
Dim strSeln As String
Dim strErrMsg As String
Dim strFileName As String
Dim vntID As Variant
Dim colIDs As VBA.Collection
Dim rstStock As ADODB.Recordset
    
    lngSelCount = lstStcokTabSelectedSoctkExport.SelCount
    If lngSelCount Then
        If lngSelCount = lstStcokTabSelectedSoctkExport.ListCount Then
            strSeln = "All Stock Items"
        Else
            strSeln = Plural(pQty:=lngSelCount, pNounSingular:="stock item") & " selected"
        End If
        strMsg = "Export Stock Items?" & vbNewLine & strSeln
                  
        If MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
            strFileName = fdlgCommon.GetFullFileName(pMethod:=eShowSave, _
                                                     pFilename:="*.txt", _
                                                     pFilter:="txt", _
                                                     pFilterDescription:="Text File", _
                                                     pDefaultExtension:="txt")
            If Len(strFileName) Then
                Me.Enabled = False
                intPrevMousePointer = SetMousePointer(vbHourglass)
                
            '   qryStock not used b/c stock selected from ListBox which may
            '   or may not later be filtered to include deleted stock
                Set rstStock = GetRst(pCnn:=g.cnnDW, _
                                      pSource:="Stock", _
                                      pSourceType:=adCmdTable, _
                                      pRstType:=eReadOnlyDynamic, _
                                      pErrMsg:=strErrMsg)
                If Not rstStock Is Nothing Then
                    If Not (rstStock.BOF And rstStock.EOF) Then
                        Set colIDs = ListBoxGetCollection(pListBox:=lstStcokTabSelectedSoctkExport, _
                                                          pItemData:=True, _
                                                          pSelected:=True)
                        For Each vntID In colIDs
                            rstStock.MoveFirst
                            rstStock.Find Criteria:="Stock_ID = " & vntID
                            If Not rstStock.EOF Then
                                WriteStockToTextFile pRstStock:=rstStock, pFullFilename:=strFileName
                                StatusBar pMsg:="Exporting " & rstStock!Description.Value, pLog:=False
                            End If
                        Next vntID
                        Set colIDs = Nothing
                        
                        StatusBar pMsg:=vbNullString, pLog:=False
                        
                        rstStock.Close
                        Set rstStock = Nothing
                    End If
                End If
                
                SetMousePointer intPrevMousePointer
                Me.Enabled = True
            End If
        End If
    End If

End Sub

Private Sub cmdTest_Click()
Dim dblResult As Double
Dim strMsg As String
Dim strErrMsg As String
Dim oSFTP As clsSFTP

    If MsgBox("Generate run-time error by submitting syntactically incorrect SQL?", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        CnnDwExecute pCommandText:="Z_SELECT * FROM EventLog"
    End If
Exit Sub

    tmrAutoDataCapture.Enabled = False
    '   PurgeLiveData
        TfrAllPreLiveDataToLiveData
    tmrAutoDataCapture.Enabled = True
Exit Sub

''   Code for setting date not to run auto capture
'    g.rstDWDefaults!LastAllFranCaptureCycleDate = GetCaptureCycleDate()
'    g.rstDWDefaults.Update
'Exit Sub
    
    strMsg = "Connect to AZTEC with FPTS?"
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Set oSFTP = New clsSFTP
        With oSFTP
            .Protocol = eFTPS
            .HostAddress = g.rstDWDefaults!AztecFtpHostAddress
            .Login = g.rstDWDefaults!AztecFtpUser
            .Password = g.rstDWDefaults!AztecFtpPwd
            .TransferType = eTfr_BINARY
            .RemoteExists "A.TXT", pErrMsg:=strErrMsg
            If Len(strErrMsg) = 0 Then
                MsgBox "Connection seemed to work", vbInformation
            Else
                MsgBox "Error Encountered: " & vbNewLine & strErrMsg, vbExclamation
            End If
        End With
        Set oSFTP = Nothing
        DoEvents
    End If
    
    strMsg = "Connect to AZTEC with SFTP?"
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        Set oSFTP = New clsSFTP
        With oSFTP
            .Protocol = eSFTP
            .HostAddress = g.rstDWDefaults!AztecFtpHostAddress
            .Login = g.rstDWDefaults!AztecFtpUser
            .Password = g.rstDWDefaults!AztecFtpPwd
            .TransferType = eTfr_BINARY
            .RemoteExists "A.TXT", pErrMsg:=strErrMsg
            If Len(strErrMsg) = 0 Then
                MsgBox "Connection seemed to work", vbInformation
            Else
                MsgBox "Error Encountered: " & vbNewLine & strErrMsg, vbExclamation
            End If
        End With
        Set oSFTP = Nothing
        DoEvents
    End If
            
    strMsg = "Create a run-time error?"
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        dblResult = 1 / 0   ' Create Divide by zero run-time error
    End If
    
    strMsg = "Optimise Db?"
    If MsgBox(strMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
        OptimiseDb
    End If

End Sub

Private Sub cmdTfrPosLiveToPreLive_Click()
Dim dtmFrom As Date
Dim dtmTo As Date
Dim strPrompt As String
Dim vdatRescheduled As Variant
Dim f As fdlgGetDateRange

    dtmFrom = DateAdd(Interval:="d", Number:=-g.rstDWDefaults!DefaultDaysOfPosLiveToTfr, Date:=Date)
    dtmTo = DateAdd(Interval:="d", Number:=-1, Date:=Date)
    strPrompt = "Select transaction date range for transfer of data " & _
                "from PosLiveData table to PreLiveData table."
    
    Set f = New fdlgGetDateRange
    vdatRescheduled = f.GetDateRange(pPrompt:=strPrompt, _
                                     pFromDate:=dtmFrom, _
                                     pToDate:=dtmTo, _
                                     pMaxToDate:=dtmTo, _
                                     pReturnType:=eMySql_BetweenClause)
    Set f = Nothing

    
    If Not IsEmpty(vdatRescheduled) Then
        TfrPosLiveToPreLive pBwDatesSqlClause:=vdatRescheduled
    End If
    
End Sub

Private Sub cmdTopSellers_Click()
'--------------------------------------------------------------------------------------------------------------
'  AUrban Procedure is a candidate for splitting into two procedures (Summarised Rpt and Not Summarised Rpt
'--------------------------------------------------------------------------------------------------------------
    
    Dim bIndexSwapped As Boolean
    
    Dim cTotalSalesIncludingTaxForThePeriod As Currency
    
    Dim iArrayColumnIndex As Integer, _
        iCurrentFranchise As Integer, _
        iNumberOfCopies As Integer, _
        iNumberOfFranchisesIncluded As Integer, _
        iPageNumber As Integer

    Dim lArrayRowIndex As Long
    Dim lTotalNumberOfBarcodesForTheReportingPeriod As Long
    Dim lTotalCustomersForThePeriod As Long, _
        lTotalItemsSoldForThePeriod As Long
    
    
Dim rstDistinctBarcodesForTheReportingPeriod As ADODB.Recordset '!!! ManualFix Clearing: Object variable not cleared: rstDistinctBarcodesForTheReportingPeriod
    Dim sFranchiseMessageBox As String, _
        sIncludedDates As String, _
        sIncludedFranchiseIDs As String, _
        sIncludedFranchiseNames As String, _
        sPlural As String
    Dim sReportingPeriod As String
    Dim sSQLQuery As String

    Dim vPlaceHolder As Variant
    Dim sReportType As String
    
    Dim intFileNum As Long
    
    'data array
    Const conSortIndex = 1, _
          conDescription = 2, _
          conQuantity = 3, _
          conNormalSell = 4, _
          conNormalSellcount = 5, _
          conTotalSalesInc = 6
    
    'tabstop array uses same as data array except for this extra
    Const conDisplayAverageSalesInc = 2
    
Dim datReportStart As Date
Dim lngRecCount As Long
Dim strErrMsg As String
    
    sReportType = "Top " & giTopSellers & " sellers for "

    If Not IsDateFmtOk() Then   ''' Review Fix Reliance on date format when time permits
        MsgBox "incorrect system date format"
        Exit Sub
    End If
    cmdTopSellers.Enabled = False
    lvwProductReport.ListItems.Clear
    lvwProductReport.Refresh
    
    Call subWriteSearchingMessageToStatusBar
    'build a query spec for all dates within the range
    datReportStart = GetDateFrom_ddmmmyy(lblProductReportStartDate)
    If lblProductReportStartDate = lblProductReportFinishDate Then
        sIncludedDates = "TransactionDate = " & MySqlDate(datReportStart)
        sReportingPeriod = lblProductReportStartDate
    Else
        sIncludedDates = "TransactionDate BETWEEN " & MySqlDate(datReportStart) & _
                                            " AND " & MySqlDate(GetDateFrom_ddmmmyy(lblProductReportFinishDate))
    
        sReportingPeriod = lblProductReportStartDate & " to " & lblProductReportFinishDate
    End If
    If optProductReportNotSummarised(0) Then
        ReDim lFranchiseID(gconZeroValue) As Long
        
            'build an array containing the ID for each selected franchise
        For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
            If (lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) And optProductReportOnSelectedFranchisesOnly(0)) Or _
                (Not optProductReportOnSelectedFranchisesOnly(0)) Then
                iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                ReDim Preserve lFranchiseID(iNumberOfFranchisesIncluded)
                lFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex))
                sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                          lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
            End If
        Next lArrayRowIndex
        
        If iNumberOfFranchisesIncluded = 0 Then
            MsgBox "No franchise selected", vbExclamation, gconReportManager
            cmdTopSellers.Enabled = True
            Exit Sub
        End If
        
        If optProductReportOnSelectedFranchisesOnly(0) Then
            'get rid of the last delimiters
            sIncludedFranchiseNames = Left(sIncludedFranchiseNames, Len(sIncludedFranchiseNames) - Len(", "))
            sFranchiseMessageBox = " for " & sIncludedFranchiseNames
        Else
            sIncludedFranchiseNames = gconAllFranchises
            sFranchiseMessageBox = ""
        End If
        
        If iNumberOfFranchisesIncluded > 1 Then
            sPlural = "s"
        End If
        If optSendProductReportToPrinter Then
            On Error GoTo NotSummarisedPrinterErrorHandler
            
            cdlTSGDataWarehouse.ShowPrinter
            Me.Refresh
            
            Printer.Print "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
            'leave a dual gap
            Printer.Print vbCrLf

            Printer.Print sReportType & sReportingPeriod
            Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
            'leave a dual gap
            Printer.Print vbCrLf
        ElseIf optSendProductReportToFile Then
            If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                intFileNum = FreeFile   ' Get unused file
                Open gsProductReportPathAndFilename For Output As #intFileNum
                Print #intFileNum, "Tobacco Station" & sPlural & _
                          " - " & sIncludedFranchiseNames
                'leave a dual gap
                Print #intFileNum, vbCrLf
            
                Print #intFileNum, sReportType & sReportingPeriod
                Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                'leave a dual gap
                Print #intFileNum, vbCrLf
            Else 'environment was not created
                MsgBox "Report was aborted", vbExclamation
                GoTo NotSummarisedTidyUp
            End If 'environement created ?
        End If 'sent prod report to file
                                    
        For iCurrentFranchise = 1 To iNumberOfFranchisesIncluded
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & vbNewLine & _
                        "WHERE (FranchiseIDTSG = " & lFranchiseID(iCurrentFranchise) & ") " & _
                         " AND (" & sIncludedDates & ")"

            Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                  pSource:=sSQLQuery, _
                                                                  pSourceType:=adCmdText, _
                                                                  pRstType:=eReadOnlyDynamic, _
                                                                  pErrMsg:=strErrMsg) 'required for movelast etc...
            
            If Not (rstDistinctBarcodesForTheReportingPeriod.BOF _
                And rstDistinctBarcodesForTheReportingPeriod.EOF) Then
                
                Call subWriteSizingArraysMessageToStatusBar
                
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conDisplayAverageSalesInc To conTotalSalesInc) As Integer
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + _
                                               Len(gconTruncateCharacter) + _
                                               gconTruncateExtensionWidth + _
                                               Len(gconSpace) + _
                                               Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conNormalSell) = iArrTabStop(conQuantity) + _
                                                             Len(gcon5DigitDollarFormat) + _
                                                             Len(gconSpace) '
                    
                    iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conNormalSell) + _
                                                             Len(gcon5DigitDollarFormat) + _
                                                             Len(gconSpace)
                    
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conDisplayAverageSalesInc) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                Else
                    iArrTabStop(conQuantity) = 42
                    iArrTabStop(conNormalSell) = 52 '
                    iArrTabStop(conDisplayAverageSalesInc) = 64 ' 58
                    iArrTabStop(conTotalSalesInc) = 78 ' 69
                End If
                
                'size the array
                lTotalNumberOfBarcodesForTheReportingPeriod = gconZeroValue
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    If fbProductIsIncludedInThisProductReport(rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                        lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod + 1
                    End If
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
                rstDistinctBarcodesForTheReportingPeriod.MoveFirst
            
                ReDim Varrsalesdata(conSortIndex To conTotalSalesInc, _
                                    1 To lTotalNumberOfBarcodesForTheReportingPeriod) As Variant
            
                Call subWriteCollatingMessageToStatusBar
            
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                lTotalItemsSoldForThePeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                lTotalCustomersForThePeriod = gconZeroValue

                Dim rstAllSameBarcodesForTheReportingPeriod As ADODB.Recordset '!!! ManualFix Clearing: Object variable not cleared: rstAllSameBarcodesForTheReportingPeriod
            
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    'If optProductReportOnSelectedFranchisesOnly(0) Then
                    sSQLQuery = "SELECT * FROM LiveData " & _
                                "WHERE (FranchiseIDTSG  = " & lFranchiseID(iCurrentFranchise) & ")" & _
                                 " AND (" & sIncludedDates & ")" & _
                                 " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod!Barcode) & ")"
                
                    lngRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:=sSQLQuery)
                    Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                         pSource:=sSQLQuery, _
                                                                         pSourceType:=adCmdText, _
                                                                         pErrMsg:=strErrMsg)
                    
                    'has to be more than zero records, so don't waste time testing for it
                    If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) <> "TOTALCUSTOMERS" Then
                        If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then

                            'lCurrentProductDisplayIndex = lCurrentProductDisplayIndex + 1
                            lArrayRowIndex = lArrayRowIndex + 1
                            Call subDisplayCurrentRecordToUser( _
                                 lArrayRowIndex, _
                                 lTotalNumberOfBarcodesForTheReportingPeriod)

                            Call subDisplayCurrentRecordToUser(lArrayRowIndex, lTotalNumberOfBarcodesForTheReportingPeriod)
                            Varrsalesdata(conDescription, lArrayRowIndex) = _
                                fsDescriptionFrom(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))
                        
                            Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                                'avert an overflow divide by zero
                                If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField) > 0 Then
                                    Varrsalesdata(conQuantity, lArrayRowIndex) = _
                                        Varrsalesdata(conQuantity, lArrayRowIndex) + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                    
                                    lTotalItemsSoldForThePeriod = _
                                        lTotalItemsSoldForThePeriod + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                    
                                    Varrsalesdata(conNormalSell, lArrayRowIndex) = _
                                        Varrsalesdata(conNormalSell, lArrayRowIndex) + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableNormalSellIncTaxField)
                                   
                                    Varrsalesdata(conTotalSalesInc, lArrayRowIndex) = _
                                        Varrsalesdata(conTotalSalesInc, lArrayRowIndex) + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                    
                                    cTotalSalesIncludingTaxForThePeriod = _
                                        cTotalSalesIncludingTaxForThePeriod + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                End If
                                Varrsalesdata(conNormalSellcount, lArrayRowIndex) = lngRecCount
                                rstAllSameBarcodesForTheReportingPeriod.MoveNext
                            Loop
                        End If
                    Else
                        Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                            lTotalCustomersForThePeriod = lTotalCustomersForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                            rstAllSameBarcodesForTheReportingPeriod.MoveNext
                        Loop
                    End If
                    rstAllSameBarcodesForTheReportingPeriod.Close
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
            
                If lTotalCustomersForThePeriod > gconZeroValue Then
                    'don't want to sort this as an array component
                    lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod - 1
                End If
            
                Call subWriteSortingMessageToStatusBar
            
                Do
                    bIndexSwapped = False
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod - 1
                        If Varrsalesdata(conTotalSalesInc, lArrayRowIndex) < _
                            Varrsalesdata(conTotalSalesInc, lArrayRowIndex + 1) Then
                            For iArrayColumnIndex = conDescription To conTotalSalesInc
                                vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                bIndexSwapped = True
                            Next iArrayColumnIndex
                        End If
                    Next lArrayRowIndex
                Loop While bIndexSwapped
            
                If lTotalNumberOfBarcodesForTheReportingPeriod > giTopSellers Then
                    lTotalNumberOfBarcodesForTheReportingPeriod = giTopSellers
                End If
            
                If optSendProductReportToDisplay Then
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = GetFranName(lFranchiseID(iCurrentFranchise))
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        Set gvListItem = lvwProductReport.ListItems.Add()
                        gvListItem.Text = Varrsalesdata(conDescription, lArrayRowIndex)
                        Call gsubAddSubItemToListview(Varrsalesdata(conQuantity, lArrayRowIndex), 1)
                        '
                        Call gsubAddSubItemToListview( _
                                 Format(Varrsalesdata(conNormalSell, lArrayRowIndex) / _
                                        Varrsalesdata(conNormalSellcount, lArrayRowIndex), gcon5DigitDollarFormat), 2)

                        Call gsubAddSubItemToListview(Format( _
                                                            Varrsalesdata(conTotalSalesInc, lArrayRowIndex) / _
                                                            Varrsalesdata(conQuantity, lArrayRowIndex) _
                                                     , gcon5DigitDollarFormat), 3)
                        Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 4)
                    Next lArrayRowIndex
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                    Me.Refresh
                ElseIf optSendProductReportToPrinter Then
                    Printer.Print GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    'leave a gap
                    Printer.Print gconSpace
                    
                    'headings
                    Printer.Print "Description"; _
                                   Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "  Qty"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Avg unit")); _
                                  "Avg unit"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "Tot (inc)"
                    'leave a gap
                    Printer.Print gconSpace
                    
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        Printer.Print _
                            Varrsalesdata(conDescription, lArrayRowIndex); _
                            Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                            Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                            Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                   Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                            Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                   Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                            Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                            Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                    Next lArrayRowIndex
                    'leave a dual gap
                    Printer.Print vbCrLf
                Else 'must be to file
                    Print #intFileNum, GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    'leave a gap
                    Print #intFileNum, gconSpace
                
                    'headings
                    If chkProductReportTabDelimited Then
                        Print #intFileNum, "Description"; _
                                   Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "Qty"; _
                                   vbTab; _
                                   "Normal"; _
                                   vbTab; _
                                  "Avg unit"; _
                                   vbTab; _
                                  "Tot (inc)"
                    Else
                        Print #intFileNum, "Description"; _
                                   Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                  "  Qty"; _
                                   Tab(iArrTabStop(conNormalSell) - Len("Normal")); _
                                  "Normal"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Avg unit")); _
                                  "Avg unit"; _
                                   Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                  "Tot (inc)"
                        Print #intFileNum, Tab(iArrTabStop(conNormalSell) - Len("Sell")); _
                                  "Sell"; _
                                   Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Sell")); _
                                  "Sell"
                    End If
                    'leave a gap
                    Print #intFileNum, gconSpace
                    
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        If chkProductReportTabDelimited Then
                            Print #intFileNum, _
                                Varrsalesdata(conDescription, lArrayRowIndex); _
                                vbTab; _
                                Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                vbTab; _
                                Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                    Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                vbTab; _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                vbTab; _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                        Else 'normal report
                            Print #intFileNum, _
                                Varrsalesdata(conDescription, lArrayRowIndex); _
                                Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conNormalSell) - Len(Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                        Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conNormalSell, lArrayRowIndex)) / _
                                        Val(Varrsalesdata(conNormalSellcount, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                        End If
                    Next lArrayRowIndex
                    Print #intFileNum, vbCrLf
                End If 'report destination
                
                On Error GoTo 0
                
                'conserve memory
                Erase iArrTabStop
                Erase Varrsalesdata
                
            Else 'no transactions for the date
                If optSendProductReportToDisplay Then
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Set gvListItem = lvwProductReport.ListItems.Add()
                    gvListItem.Text = gconSpace
                ElseIf optSendProductReportToPrinter Then
                    Printer.Print "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Printer.Print vbCrLf
                Else 'is destined for the file
                    Print #intFileNum, "No transactions for " & GetFranName(lFranchiseID(iCurrentFranchise)) & gsReportPeriodWording & sReportingPeriod
                    Print #intFileNum, vbCrLf
                End If
            End If 'any transactions for the report date ?
            
            With stb
                .SimpleText = ""
                .Refresh
            End With

            rstDistinctBarcodesForTheReportingPeriod.Close
        Next iCurrentFranchise
        
NotSummarisedTidyUp:
        
        On Error GoTo 0
        
        If optSendProductReportToDisplay Then
            'do nothing
        ElseIf optSendProductReportToPrinter Then
            Printer.EndDoc
            MsgBox "Report was successfully submitted to the selected printer", _
                    vbInformation, gconReportManager
        Else 'was to file
            Close #intFileNum
            Call subSetProductReportViewButton
            MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & _
                   ". Use the 'View' button to display", _
                    vbInformation, gconReportManager
        End If
'--------------------------------------------------------------------------------------------------------------
'  AUrban SUMMARISED REPORT: Procedure is a candidate for splitting above and below here into two procedures
'--------------------------------------------------------------------------------------------------------------
    Else 'summarised
        Call subWriteSearchingMessageToStatusBar
        
        'build a query spec for all dates within the range
        datReportStart = GetDateFrom_ddmmmyy(lblProductReportStartDate)

        If lblProductReportStartDate = lblProductReportFinishDate Then
            sIncludedDates = "TransactionDate = " & MySqlDate(datReportStart)
            sReportingPeriod = lblProductReportStartDate
        Else
            sIncludedDates = "TransactionDate BETWEEN " & MySqlDate(datReportStart) & _
                                                " AND " & MySqlDate(GetDateFrom_ddmmmyy(lblProductReportFinishDate))
            sReportingPeriod = lblProductReportStartDate & " to " & lblProductReportFinishDate
        End If
        
        If optProductReportOnSelectedFranchisesOnly(0) Then
            'build a query spec for all selected franchises
            For lArrayRowIndex = gconDisplayFirstItem To lstProductReportsFranchiseBusinessName.ListCount - 1
                If lstProductReportsFranchiseBusinessName.Selected(lArrayRowIndex) Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    
                    sIncludedFranchiseNames = sIncludedFranchiseNames & _
                                              lstProductReportsFranchiseBusinessName.List(lArrayRowIndex) & ", "
                    
                    sIncludedFranchiseIDs = sIncludedFranchiseIDs & _
                                            gconLiveDataTableTSGFranchiseIDField & " = " & _
                                            fsFranchiseIDFrom(lstProductReportsFranchiseBusinessName.List(lArrayRowIndex)) & " OR "
                End If
            Next lArrayRowIndex
            
            If iNumberOfFranchisesIncluded Then
                'get rid of the last delimiters
                sIncludedFranchiseNames = Left(sIncludedFranchiseNames, _
                                          Len(sIncludedFranchiseNames) - Len(", "))
                
                sIncludedFranchiseIDs = Left(sIncludedFranchiseIDs, _
                                        Len(sIncludedFranchiseIDs) - Len(" OR "))
            
                sFranchiseMessageBox = " for " & sIncludedFranchiseNames
            End If
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData " & vbNewLine & _
                        "WHERE (" & sIncludedFranchiseIDs & ") " & _
                         " AND (" & sIncludedDates & ")"

        Else 'all franchises option was selected
            iNumberOfFranchisesIncluded = lstProductReportsFranchiseBusinessName.ListCount
            
            sFranchiseMessageBox = ""
            
            sIncludedFranchiseNames = gconAllFranchises
            
            sSQLQuery = "SELECT DISTINCT Barcode FROM LiveData WHERE " & sIncludedDates
        End If
        
        If iNumberOfFranchisesIncluded Then 'franchises are included
            If iNumberOfFranchisesIncluded > 1 Then
                sPlural = "s"
            End If
            
            Set rstDistinctBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                  pSource:=sSQLQuery, _
                                                                  pSourceType:=adCmdText, _
                                                                  pRstType:=eReadOnlyDynamic, _
                                                                  pErrMsg:=strErrMsg) 'required for movelast etc...
            If Not (rstDistinctBarcodesForTheReportingPeriod.BOF _
                And rstDistinctBarcodesForTheReportingPeriod.EOF) Then
                
                Call subWriteSizingArraysMessageToStatusBar
                
                'use the tabstop array to store the right justified position
                ReDim iArrTabStop(conDisplayAverageSalesInc To conTotalSalesInc) As Integer
                'truncate if the docket printer is enabled in the defaults
                If g.rstAppDefaults(gconDocketPrinterEnabled) Then
                    iArrTabStop(conQuantity) = gconTruncateDescriptionBriefAt + _
                                               Len(gconTruncateCharacter) + _
                                               gconTruncateExtensionWidth + _
                                               Len(gconSpace) + _
                                               Len(gconStandardQuantityFormat)
                    
                    iArrTabStop(conDisplayAverageSalesInc) = iArrTabStop(conQuantity) + _
                                                             Len(gcon5DigitDollarFormat) + _
                                                             Len(gconSpace)
                    
                    iArrTabStop(conTotalSalesInc) = iArrTabStop(conDisplayAverageSalesInc) + _
                                                    Len(gcon5DigitDollarFormat) + _
                                                    Len(gconSpace)
                Else
                    iArrTabStop(conQuantity) = 48
                    iArrTabStop(conDisplayAverageSalesInc) = 58
                    iArrTabStop(conTotalSalesInc) = 69
                End If
                
                'size the array
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    If fbProductIsIncludedInThisProductReport(rstDistinctBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then
                        lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod + 1
                    End If
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
                rstDistinctBarcodesForTheReportingPeriod.MoveFirst
                
                ReDim Varrsalesdata(conSortIndex To conTotalSalesInc, _
                                    1 To lTotalNumberOfBarcodesForTheReportingPeriod) As Variant
                
                Call subWriteCollatingMessageToStatusBar
                
                cTotalSalesIncludingTaxForThePeriod = gconZeroValue
                lTotalItemsSoldForThePeriod = gconZeroValue
                lArrayRowIndex = gconZeroValue
                lTotalCustomersForThePeriod = gconZeroValue
                
                Do Until rstDistinctBarcodesForTheReportingPeriod.EOF
                    If optProductReportOnSelectedFranchisesOnly(0) Then
                        sSQLQuery = "SELECT * FROM LiveData " & vbNewLine & _
                                    "WHERE (" & sIncludedFranchiseIDs & ") " & _
                                    " AND (" & sIncludedDates & ") " & _
                                    " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod!Barcode) & ")"
                    Else 'all franchises, no requirement to discriminate (performance reasons)
                        sSQLQuery = "SELECT * FROM LiveData " & vbNewLine & _
                                    "WHERE (" & sIncludedDates & ") " & _
                                     " AND (Barcode = " & SqlQ(rstDistinctBarcodesForTheReportingPeriod!Barcode) & ")"
                    End If
                    
                    Set rstAllSameBarcodesForTheReportingPeriod = GetRst(pCnn:=g.cnnDW, _
                                                                         pSource:=sSQLQuery, _
                                                                         pSourceType:=adCmdText, _
                                                                         pErrMsg:=strErrMsg)
                    
                    'has to be more than zero records, so don't waste time testing for it
                    If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField) <> "TOTALCUSTOMERS" Then

                        If fbProductIsIncludedInThisProductReport(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField)) Then

                            lArrayRowIndex = lArrayRowIndex + 1
                            Call subDisplayCurrentRecordToUser(lArrayRowIndex, lTotalNumberOfBarcodesForTheReportingPeriod)
                            Varrsalesdata(conDescription, lArrayRowIndex) = _
                                fsDescriptionFrom(rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableBarcodeField))
                        
                            Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                                'avert an overflow divide by zero
                                If rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField) > 0 Then
                                    Varrsalesdata(conQuantity, lArrayRowIndex) = _
                                        Varrsalesdata(conQuantity, lArrayRowIndex) + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                    
                                    lTotalItemsSoldForThePeriod = _
                                        lTotalItemsSoldForThePeriod + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                                    
                                    Varrsalesdata(conTotalSalesInc, lArrayRowIndex) = _
                                        Varrsalesdata(conTotalSalesInc, lArrayRowIndex) + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                    
                                    cTotalSalesIncludingTaxForThePeriod = _
                                        cTotalSalesIncludingTaxForThePeriod + _
                                        rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableTotalIncTaxField)
                                End If
                                rstAllSameBarcodesForTheReportingPeriod.MoveNext
                            Loop
                        End If
                    Else
                        Do Until rstAllSameBarcodesForTheReportingPeriod.EOF
                            lTotalCustomersForThePeriod = lTotalCustomersForThePeriod + rstAllSameBarcodesForTheReportingPeriod(gconLiveDataTableQuantityField)
                            rstAllSameBarcodesForTheReportingPeriod.MoveNext
                        Loop
                    End If
                    rstAllSameBarcodesForTheReportingPeriod.Close
                    rstDistinctBarcodesForTheReportingPeriod.MoveNext
                Loop
                
                If lTotalCustomersForThePeriod > 0 Then
                    'don't want to sort this as an array component
                    lTotalNumberOfBarcodesForTheReportingPeriod = lTotalNumberOfBarcodesForTheReportingPeriod - 1
                End If
                
                Call subWriteSortingMessageToStatusBar
                
                Do
                    bIndexSwapped = False
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod - 1
                        If Varrsalesdata(conTotalSalesInc, lArrayRowIndex) < _
                            Varrsalesdata(conTotalSalesInc, lArrayRowIndex + 1) Then
                            For iArrayColumnIndex = conDescription To conTotalSalesInc
                                vPlaceHolder = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex) = Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1)
                                Varrsalesdata(iArrayColumnIndex, lArrayRowIndex + 1) = vPlaceHolder
                                bIndexSwapped = True
                            Next iArrayColumnIndex
                        End If
                    Next lArrayRowIndex
                Loop While bIndexSwapped
                
                If lTotalNumberOfBarcodesForTheReportingPeriod > giTopSellers Then
                    lTotalNumberOfBarcodesForTheReportingPeriod = giTopSellers
                End If
                
                If optSendProductReportToDisplay Then
                    For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                        Set gvListItem = lvwProductReport.ListItems.Add()
                        gvListItem.Text = Varrsalesdata(conDescription, lArrayRowIndex)
                        Call gsubAddSubItemToListview(Varrsalesdata(conQuantity, lArrayRowIndex), 1)
                        Call gsubAddSubItemToListview(Format( _
                                                            Varrsalesdata(conTotalSalesInc, lArrayRowIndex) / _
                                                            Varrsalesdata(conQuantity, lArrayRowIndex) _
                                                     , gcon5DigitDollarFormat), 2)
                        Call gsubAddSubItemToListview( _
                             Format(Varrsalesdata(conTotalSalesInc, lArrayRowIndex), gcon6DigitDollarFormat), 3)
                    Next lArrayRowIndex
                ElseIf optSendProductReportToPrinter Then
                    On Error GoTo SummarisedPrinterErrorHandler
                    
                    cdlTSGDataWarehouse.ShowPrinter
                    iNumberOfCopies = cdlTSGDataWarehouse.Copies
                    Me.Refresh
                    
                    For iPageNumber = 1 To iNumberOfCopies
                        Printer.Print "Tobacco Station" & sPlural & _
                                      " - " & sIncludedFranchiseNames
                        'leave a dual gap
                        Printer.Print vbCrLf
            
                        Printer.Print sReportType & sReportingPeriod
                        Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                        'leave a dual gap
                        Printer.Print vbCrLf
                        
                        'headings
                        Printer.Print "Description"; _
                                       Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                      "  Qty"; _
                                       Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Avg unit")); _
                                      "Avg unit"; _
                                       Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                      "Tot (inc)"
                        'leave a gap
                        Printer.Print gconSpace
                        
                        For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                            Printer.Print _
                                Varrsalesdata(conDescription, lArrayRowIndex); _
                                Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                       Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                        Next lArrayRowIndex
                    Next iPageNumber
                    Printer.EndDoc
                    
                    MsgBox "Report was successfully submitted to the selected printer", _
                            vbInformation, gconReportManager
                Else 'must be to file
                    If fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared Then
                        intFileNum = FreeFile   ' Get unused file
                        Open gsProductReportPathAndFilename For Output As #intFileNum
                        Print #intFileNum, "Tobacco Station" & sPlural & _
                                  " - " & sIncludedFranchiseNames
                        'leave a dual gap
                        Print #intFileNum, vbCrLf
                        
                        Print #intFileNum, sReportType & sReportingPeriod
                        Print #intFileNum, "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
                        'leave a dual gap
                        Print #intFileNum, vbCrLf
                        
                        'headings
                        If chkProductReportTabDelimited Then
                            Print #intFileNum, "Description"; _
                                       Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                      "Qty"; _
                                       vbTab; _
                                      "Avg unit"; _
                                       vbTab; _
                                      "Tot (inc)"
                        Else
                            Print #intFileNum, "Description"; _
                                       Tab(iArrTabStop(conQuantity) - Len("  Qty")); _
                                      "  Qty"; _
                                       Tab(iArrTabStop(conDisplayAverageSalesInc) - Len("Avg unit")); _
                                      "Avg unit"; _
                                       Tab(iArrTabStop(conTotalSalesInc) - Len("Tot (inc)")); _
                                      "Tot (inc)"
                        End If
                        'leave a gap
                        Print #intFileNum, gconSpace
                        
                        For lArrayRowIndex = 1 To lTotalNumberOfBarcodesForTheReportingPeriod
                            If chkProductReportTabDelimited Then
                                Print #intFileNum, _
                                    Varrsalesdata(conDescription, lArrayRowIndex); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    vbTab; _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            Else 'normal report
                                Print #intFileNum, _
                                    Varrsalesdata(conDescription, lArrayRowIndex); _
                                    Tab(iArrTabStop(conQuantity) - Len(Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat))); _
                                    Format(Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gconStandardQuantityFormat); _
                                    Tab(iArrTabStop(conDisplayAverageSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)) / _
                                           Val(Varrsalesdata(conQuantity, lArrayRowIndex)), gcon5DigitDollarFormat); _
                                    Tab(iArrTabStop(conTotalSalesInc) - Len(Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat))); _
                                    Format(Val(Varrsalesdata(conTotalSalesInc, lArrayRowIndex)), gcon5DigitDollarFormat)
                            End If
                        Next lArrayRowIndex
                                           
                        Close #intFileNum
                        
                        Call subSetProductReportViewButton
                        
                        MsgBox "Report was successfully sent to - " & gsProductReportPathAndFilename & _
                               ". Use the 'View' button to display", _
                                vbInformation, gconReportManager
                    Else 'environment was not created
                        MsgBox "Report was aborted", vbExclamation
                    End If 'environement created ?
                End If 'report destination
    
SummarisedTidyUp:
                On Error GoTo 0
                
                'conserve memory
                Erase iArrTabStop
                Erase Varrsalesdata
            Else 'no summarised for the date
                MsgBox "No sales transactions" & sFranchiseMessageBox & gsReportPeriodWording & sReportingPeriod, _
                        vbInformation, gconReportManager
            End If 'any summarised transactions for the report date ?
        
            With stb
                .SimpleText = ""
                .Refresh
            End With
            
            rstDistinctBarcodesForTheReportingPeriod.Close
        Else 'no franchises
            MsgBox "No franchise selected", vbExclamation, gconReportManager
        End If 'any summarised franchises selected ?
    End If 'not-summarised or summarised ?
    
    cmdTopSellers.Enabled = True
    Exit Sub

NotSummarisedPrinterErrorHandler:
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    Resume NotSummarisedTidyUp

SummarisedPrinterErrorHandler:
    
    Printer.KillDoc
    If Err.Number <> gconZeroValue Then 'the error was not caused by user cancelling the dialogue
        MsgBox "Printer error - " & Err.Number & ". Report failed, try selecting " & SQ("Display"), vbCritical
    End If
    Resume SummarisedTidyUp

End Sub

Private Sub cmdViewProductReport_Click()
    subOpenFile gsProductReportPathAndFilename
End Sub

Private Sub cmdViewStickReport_Click()
    subOpenFile gsStickReportPathAndFilename
End Sub

Private Sub ConfigureBataTabButtons()
Dim bRptSelected As Boolean
Dim vnt As Variant
Dim col As VBA.Collection

'   Determining whether a report is selected
'   Get collection of string values of Report column in the selected rows
'   Remembering that some selected rows may not have a report we set the flag
'   bRptSelected = True when at least one string value in our collection <> ""
    Set col = GridGetCollection(pGrid:=grdBataRpts, pColKey:="RptType", pSelected:=True)
    For Each vnt In col
        If Len(vnt) Then
            bRptSelected = True
            Exit For
        End If
    Next vnt
    Set col = Nothing
    
    cmdBataTabSaveSelected.Enabled = bRptSelected
    cmdBataTabViewSelected.Enabled = bRptSelected
    cmdBataTabExportGrid.Enabled = (grdBataRpts.Rows > grdBataRpts.FixedRows)   ' Grid has data
    cmdBataTabPrintGrid.Enabled = (grdBataRpts.Rows > grdBataRpts.FixedRows)    ' Grid has data
    
End Sub

Private Sub ConfigureStkTabCtls()
    
'   Only alter when adding a new item (no item selected)
    If lstDescription.ListIndex < 0 Then
        Select Case cboCategory.Text
            Case gkCAT_CigCtn
                chkPackage.Value = 0
                txtStkItemDescription.Text = Trim$(stripType(txtStkItemDescription.Text)) & " Ctn"
            Case gkCAT_CigPkt
                chkPackage.Value = 1
                txtStkItemDescription.Text = Trim$(stripType(txtStkItemDescription.Text)) & " Pkt"
        End Select
    End If
    
End Sub

Private Sub CreateNielsenReports(ByVal pLastReportEndDate As Date, pCalledAutomatically As Boolean)
Dim lngDays As Long
Dim lngWeeks As Long
Dim dtmNielsenStart As Date
Dim dtmNielsenEnd As Date
Dim dtmLastNielsenRptStart As Date
Dim dtmLastNielsenRptEnd As Date
Dim dtmFirstNielsenRptStart As Date
Dim strAutoManual As String
Dim strNielsenRptsSubFolder As String
Dim strZipOfWeeklyFULLname As String
Dim strZipOfDailyFULLname As String
Dim colFullFilenamesToZip As VBA.Collection
Dim fso As Scripting.FileSystemObject

    dtmLastNielsenRptEnd = pLastReportEndDate
    dtmLastNielsenRptStart = DateAdd("d", -6, dtmLastNielsenRptEnd)

'   Only case report is not overwritten is when called automatically and the report already exists,
'   therefore ovewrite when NOT (called automatically and report already exists)
    If Not (pCalledAutomatically And FileExists(fsNielsenRptFullname(pStartDate:=dtmLastNielsenRptStart, pEndDate:=dtmLastNielsenRptEnd))) Then
        If pCalledAutomatically Then
            strAutoManual = "Automatically"
        Else
            strAutoManual = "Manually"
        End If
        StatusBar strAutoManual & " invoke Nielsen report cycle for 3 weeks up to " & Format$(pLastReportEndDate, gkFmtDateUnambiguous)
        
    '   Make sub folder for Nielsen reports
        strNielsenRptsSubFolder = GetNeilsenRptsSubFolderName(pLastReportEndDate:=dtmLastNielsenRptEnd)
        Set fso = New Scripting.FileSystemObject
        If Not fso.FolderExists(strNielsenRptsSubFolder) Then
            fso.CreateFolder Path:=strNielsenRptsSubFolder
        End If
        Set fso = Nothing
        
    '   Create Nielsen DAILY reports for last three weeks
        dtmFirstNielsenRptStart = DateAdd("ww", -2, dtmLastNielsenRptStart)
        Set colFullFilenamesToZip = New VBA.Collection
        
        StatusBar pMsg:="Collecting Nielsen daily reports"
        For lngDays = 0 To 20
            dtmNielsenStart = DateAdd("d", lngDays, dtmFirstNielsenRptStart)
            CreateNielsenRpt pStartDate:=dtmNielsenStart, pEndDate:=dtmNielsenStart
            colFullFilenamesToZip.Add fsNielsenRptFullname(pStartDate:=dtmNielsenStart, pEndDate:=dtmNielsenStart)
        Next lngDays
        
        StatusBar pMsg:=colFullFilenamesToZip.Count & " of 21 Nielsen daily reports collected."
        If colFullFilenamesToZip.Count Then
            strZipOfDailyFULLname = GetNeilsenDailyRptFullname(pLastReportEndDate:=dtmLastNielsenRptEnd)
            DeleteFile strZipOfDailyFULLname
            ZipFiles colFullFilenamesToZip, strZipOfDailyFULLname
        '   Clear collection: (Prepare for collecting Nielsen WEEKLY reports)
            Do While colFullFilenamesToZip.Count > 0
                colFullFilenamesToZip.Remove Index:=1
            Loop
        End If
    
    '   Create Nielsen WEEKLY reports for last three weeks
        StatusBar pMsg:="Collecting Nielsen weekly reports"
        For lngWeeks = 0 To 2
            dtmNielsenStart = DateAdd("ww", lngWeeks, dtmFirstNielsenRptStart)
            dtmNielsenEnd = DateAdd("d", 6, dtmNielsenStart)
            CreateNielsenRpt pStartDate:=dtmNielsenStart, pEndDate:=dtmNielsenEnd
            colFullFilenamesToZip.Add fsNielsenRptFullname(pStartDate:=dtmNielsenStart, pEndDate:=dtmNielsenEnd)
        Next lngWeeks
        
        StatusBar pMsg:=colFullFilenamesToZip.Count & " of 3 Nielsen weekly reports collected."
        If colFullFilenamesToZip.Count Then
            strZipOfWeeklyFULLname = GetNeilsenWeeklyRptFullname(pLastReportEndDate:=dtmLastNielsenRptEnd)
            DeleteFile strZipOfWeeklyFULLname
            ZipFiles colFullFilenamesToZip, strZipOfWeeklyFULLname
        '   Clear collection
            Do While colFullFilenamesToZip.Count > 0
                colFullFilenamesToZip.Remove Index:=1
            Loop
        End If
        
        Set colFullFilenamesToZip = Nothing
        StatusBar "Nielsen report cycle completed"
    End If
    
End Sub

Private Sub CreateNielsenRpt(ByVal pStartDate As Date, ByVal pEndDate As Date)
Const conDelimiter = vbTab
Dim cTotalSalesValueInc As Currency
Dim lLastFranchiseID As Long
Dim lTotalQuantity As Long
Dim lRecordcounter As Long
Dim rstSnpAllSalesForSelectedPeriod As ADODB.Recordset
Dim iFile As Integer
Dim sExceptionReportFullPathAndFilename As String
Dim strWhereClause As String
Dim sLastBarcode As String
Dim sSQLQuery As String
Dim intPrevMousePointer As Integer
Dim strDailyOrWeekly As String      ' Daily = "d" and Weekly = "" (ie is the default and original)
Dim strMsg As String
Dim strErrMsg As String

    If pStartDate = pEndDate Then
        strDailyOrWeekly = "d"
    End If
    
    cmdCreateNielsenReports.Enabled = False
    intPrevMousePointer = SetMousePointer(vbHourglass)

    If IsDateFmtOk() Then   ''' Fix Reliance on date format when time permits
        subPurgeNielsenReport pReportStartDate:=pStartDate, pReportEndDate:=pEndDate
        
        strMsg = "Running database query for Nielsen report: "
        If pStartDate = pEndDate Then
            strMsg = strMsg & Format$(pStartDate, gconStandardDateFormat)
        Else
            strMsg = Format$(pStartDate, gconStandardDateFormat) & " to " & Format$(pEndDate, gconStandardDateFormat)
        End If
        stb.SimpleText = strMsg & ", please wait...."
        
        Me.Refresh
    
        sExceptionReportFullPathAndFilename = g.strNielsenRptsFolder & "\" & _
                                             "Nielsen Exceptions" & _
                                              "_" & strDailyOrWeekly & _
                                              Format$(pEndDate, gconFmtDateInNielsenFilename) & _
                                              gconTextFileSuffix
        
        subEnsurePathExistsCreateNewIterationOfExceptionReport sExceptionReportFullPathAndFilename, g.strNielsenRptsFolder
                        
    '   Optimise where clause if start and end date are the same
        If pStartDate = pEndDate Then
            strWhereClause = "TransactionDate = " & MySqlDate(pStartDate)
        Else
            strWhereClause = "TransactionDate BETWEEN " & MySqlDate(pStartDate) & " AND " & MySqlDate(pEndDate)
        End If
    
        'open (create) the file
        iFile = openNewFile(fsNielsenRptFullname(pStartDate:=pStartDate, pEndDate:=pEndDate))
        
        'get all franchises
        sSQLQuery = "SELECT FranchiseIDTSG, Barcode, TransactionDate, Quantity, TotalInc " & _
                    "FROM livedata " & _
                    "WHERE (" & strWhereClause & ") " & _
                     " AND (Quantity <> 0) " & _
                     " AND (Barcode <> " & SqlQ("TOTALCUSTOMERS") & ") " & _
                    "ORDER BY FranchiseIDTSG, Barcode"
        Set rstSnpAllSalesForSelectedPeriod = GetRst(pCnn:=g.cnnDW, pSource:=sSQLQuery, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
        DoEvents
        'set the first values to the first tobacco product
        If Not (rstSnpAllSalesForSelectedPeriod.BOF And rstSnpAllSalesForSelectedPeriod.EOF) Then
            Do Until rstSnpAllSalesForSelectedPeriod.EOF
                If fbBarcodeIsATobaccoProduct(rstSnpAllSalesForSelectedPeriod(gconLiveDataTableBarcodeField)) Then
                    lTotalQuantity = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableQuantityField)
                    cTotalSalesValueInc = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTotalIncTaxField)
                    lLastFranchiseID = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTSGFranchiseIDField)
                    sLastBarcode = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableBarcodeField)
                    Exit Do
                End If
                rstSnpAllSalesForSelectedPeriod.MoveNext
            Loop
            DoEvents
            'compare the next record to the first
            Do Until rstSnpAllSalesForSelectedPeriod.EOF
                If fbBarcodeIsATobaccoProduct(rstSnpAllSalesForSelectedPeriod(gconLiveDataTableBarcodeField)) Then
                    DoEvents
                    If lLastFranchiseID <> _
                       rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTSGFranchiseIDField) Then
                        lRecordcounter = lRecordcounter + 1
                        stb.SimpleText = GetFranName(lLastFranchiseID) & " record " & lRecordcounter & " - processing"
                        Me.Refresh
                        'the franchise id for the last record checked was different to the one, so write the (aggregated) data, and set the variables to the current record's
                        Print #iFile, lLastFranchiseID & _
                                  conDelimiter & _
                                  GetFranName(lLastFranchiseID) & _
                                  conDelimiter & _
                                  sLastBarcode & _
                                  conDelimiter & _
                                  fsDescriptionFrom(sLastBarcode) & _
                                  conDelimiter & _
                                  Format$(pEndDate, gconFmtDateInNielsenFilename) & _
                                  conDelimiter & _
                                  lTotalQuantity & _
                                  conDelimiter & _
                                  Format(cTotalSalesValueInc, gcon22DigitDecimalFormat)
                        
                        lTotalQuantity = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableQuantityField)
                        cTotalSalesValueInc = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTotalIncTaxField)
                        lLastFranchiseID = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTSGFranchiseIDField)
                        sLastBarcode = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableBarcodeField)
                        lRecordcounter = 0
                    ElseIf sLastBarcode <> _
                        rstSnpAllSalesForSelectedPeriod(gconLiveDataTableBarcodeField) Then
                        'the barcode for the last record checked was different to the one,
                        'so write the (aggregated) data, and set the variables to the
                        'current record's
                        lRecordcounter = lRecordcounter + 1
                        stb.SimpleText = GetFranName(lLastFranchiseID) & " record " & lRecordcounter & " - processing"
                        Me.Refresh
                        Print #iFile, lLastFranchiseID & _
                                  conDelimiter & _
                                  GetFranName(lLastFranchiseID) & _
                                  conDelimiter & _
                                  sLastBarcode & _
                                  conDelimiter & _
                                  fsDescriptionFrom(sLastBarcode) & _
                                  conDelimiter & _
                                  Format(pEndDate, gconFmtDateInNielsenFilename) & _
                                  conDelimiter & _
                                  lTotalQuantity & _
                                  conDelimiter & _
                                  Format(cTotalSalesValueInc, gcon22DigitDecimalFormat)
                        
                        lTotalQuantity = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableQuantityField)
                        cTotalSalesValueInc = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTotalIncTaxField)
                        lLastFranchiseID = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTSGFranchiseIDField)
                        sLastBarcode = rstSnpAllSalesForSelectedPeriod(gconLiveDataTableBarcodeField)
                    Else 'same franchiseID and barcode as the last one tested
                         'so aggregate the other records tothe first one
                        lRecordcounter = lRecordcounter + 1
                        stb.SimpleText = GetFranName(lLastFranchiseID) & " record " & lRecordcounter & " - aggregating"
                        Me.Refresh
                        lTotalQuantity = lTotalQuantity + rstSnpAllSalesForSelectedPeriod(gconLiveDataTableQuantityField)
                        cTotalSalesValueInc = cTotalSalesValueInc + rstSnpAllSalesForSelectedPeriod(gconLiveDataTableTotalIncTaxField)
                    End If
                End If 'tobacco product ?
                rstSnpAllSalesForSelectedPeriod.MoveNext
                DoEvents
            Loop
            Print #iFile, lLastFranchiseID & _
                      conDelimiter & _
                      GetFranName(lLastFranchiseID) & _
                      conDelimiter & _
                      sLastBarcode & _
                      conDelimiter & _
                      fsDescriptionFrom(sLastBarcode) & _
                      conDelimiter & _
                      Format(pEndDate, gconFmtDateInNielsenFilename) & _
                      conDelimiter & _
                      lTotalQuantity & _
                      conDelimiter & _
                      Format(cTotalSalesValueInc, gcon22DigitDecimalFormat)
                      
        End If 'any sales for the period?
        rstSnpAllSalesForSelectedPeriod.Close
        Set rstSnpAllSalesForSelectedPeriod = Nothing
        
        Close #iFile
        With stb
            .SimpleText = "Report complete"
            .Refresh
        End With
    End If 'date format correct?

    SetMousePointer intPrevMousePointer

End Sub

Sub CreateUploadsPending(ByVal when, ByVal which_ones)
' Creates the 'uploads pending' records by looking at the selections on
' the Uploads tab. ie. creates a record for each store, for each item to be uploaded
' If the fNow flag is true, it will then actually do the upload.
'   AUrban Optimisation: Could be rewritten to remove multiple ReDim calls within loops (very expensive operation)

'   lstUploadFranchiseList is populated by subPopulateFranchiseBusinessNameListBoxes()
'   As at V400 it only populates list with Live Franchises from qryFranchiseLive

    Const conImmediately = " Immediately. "
    Dim bContinue As Boolean
    Dim Reply As Variant
    Dim sWhen As String
    Dim sUploadList As String
    Dim sFranchiseList As String
    Dim iNumberOfItemsIncluded As Integer
    Dim sFragment, sOther As String
    Dim iExistingPending As Integer
    Dim fTimesToBeSet As Boolean
    Dim sBaseName As String
    Dim iIndex
    Dim iInner As Integer
    Dim iNumberOfFranchisesIncluded As Integer
    Dim lArrayRowIndex As Long
Dim lngFranID As Long
Dim lngFranType As Long
Dim bFranMsgFlag As Boolean
Dim rstUploadTasks As ADODB.Recordset
Dim rstFran As ADODB.Recordset
    
    Dim strSQL As String
    Dim strErrMsg As String

    sFranchiseList = vbCrLf
    Dim alngFranchiseID() As Long
    
    Select Case True
        
        Case optUploadSelection(eUpldSelFrans).Value
            'build an array containing the ID for each selected franchise
            For lArrayRowIndex = 0 To lstUploadFranchiseList.ListCount - 1
                If lstUploadFranchiseList.Selected(lArrayRowIndex) Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    ReDim Preserve alngFranchiseID(1 To iNumberOfFranchisesIncluded)
                    alngFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstUploadFranchiseList.List(lArrayRowIndex))
                    sFranchiseList = sFranchiseList & lstUploadFranchiseList.List(lArrayRowIndex) & vbCrLf
                End If
            Next lArrayRowIndex
            
        Case optUploadSelection(eUpldAllFrans).Value
''' Review: Must be a simpler and quicker way of coding this case (and probably the other cases)
            For lArrayRowIndex = 0 To lstUploadFranchiseList.ListCount - 1
                iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                ReDim Preserve alngFranchiseID(1 To iNumberOfFranchisesIncluded)
                alngFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstUploadFranchiseList.List(lArrayRowIndex))
            Next lArrayRowIndex
            sFranchiseList = vbCrLf & "All Franchises"
            
        Case optUploadSelection(eUpldPmAndRmFrans)
            For lArrayRowIndex = 0 To lstUploadFranchiseList.ListCount - 1
                ' check the type of this store
                strSQL = "SELECT * FROM Franchises " & vbNewLine & _
                         "WHERE FranchiseIDTSG  = " & fsFranchiseIDFrom(lstUploadFranchiseList.List(lArrayRowIndex)) & ";"
                Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
                If rstFran!FranchiseType < 30 Then
                    iNumberOfFranchisesIncluded = iNumberOfFranchisesIncluded + 1
                    ReDim Preserve alngFranchiseID(1 To iNumberOfFranchisesIncluded)
                    alngFranchiseID(iNumberOfFranchisesIncluded) = fsFranchiseIDFrom(lstUploadFranchiseList.List(lArrayRowIndex))
                End If
                rstFran.Close
                Set rstFran = Nothing
            Next lArrayRowIndex
        ''' sFranchiseList = vbCrLf & "All Franchises"  ' V400
            sFranchiseList = vbCrLf & "Price Module & Retail Mgr franchises"
            
        Case optUploadSelection(eUpldSelStates)
            ' Look at the 'state' check boxes
            iNumberOfFranchisesIncluded = fGetStateFranchiseList(alngFranchiseID)
            sFranchiseList = sFranchiseList & " the states selected." & vbCrLf

    End Select

''' V388 The upside of calling DisplayUploadsPending() is it refreshes display, of uploads pending
''' V388 BUT creating uploads pending can occur on master so diplay will already be up to date
''' V388 NOTE there is a call to DisplayUploadsPending() at the end of the procedure
''' V388 IN WRITING THIS I DISCOVERED THE GRIM REALITY THAT SOME DB WRITING MAY BE PEROFRMED BY SLAVES
''' V388 BECAUSE CnnMode property in Connection object is not supported by MySQL ODBC Connector/Driver
''' V388 Call to DisplayUploadsPending() clears listbox selection but this is inconsequential
''' V388 b/c lvwUploadsPending are note used for anything. COULD SPEED UP WITH CALL TO GetRstVal()
''' V388 using SQL below but not really worth the bother. Will let called procedure do the work
''' V388    strSQL = "SELECT * FROM FranchiseUploads WHERE UploadedDate IS NULL"
''' V388 Culd even use iExistingPending = lvwUploadsPending.ListItems.Count but could be less accurate if
''' V388 lvwUploadsPending has not been updated, so leave it all the same until this function gets a rewrite
    iExistingPending = DisplayUploadsPending()

    If (iNumberOfFranchisesIncluded = 0) And (iExistingPending = 0) Then
        MsgBox "No franchise selected, and no other uploads pending.", vbExclamation
        Exit Sub
    End If
    
    ' Build an array of possible files to be uploaded. These things are highlighted in
    ' the 'uploads' list box
    sUploadList = vbCrLf
''' Review: Must be a simpler and quicker way than using For Loop and code below
    ReDim aszUploadItem(gconZeroValue) As String
    For lArrayRowIndex = gconDisplayFirstItem To lstUploadItemList.ListCount - 1
        If lstUploadItemList.Selected(lArrayRowIndex) Then
            iNumberOfItemsIncluded = iNumberOfItemsIncluded + 1
            ReDim Preserve aszUploadItem(iNumberOfItemsIncluded)
            aszUploadItem(iNumberOfItemsIncluded) = lstUploadItemList.List(lArrayRowIndex)
            sUploadList = sUploadList & lstUploadItemList.List(lArrayRowIndex) & vbCrLf
        End If
    Next lArrayRowIndex
   
    If when = IMMEDIATELY Then
        sWhen = conImmediately
    ElseIf when = LATER Then
        sWhen = " during the overnight capture cycle."
    End If
    
    fTimesToBeSet = False

    sFragment = ""
    
    If chkResetRemoteOpenedBy Then
        sFragment = sFragment & " as well as reset the `Opended By` flag" & vbCrLf
        fTimesToBeSet = True
    End If
    
    If iNumberOfItemsIncluded = 0 And Not fTimesToBeSet And iExistingPending = 0 Then
        MsgBox "No items selected for upload. Select something to upload" & vbCrLf & _
        "or tick at least one of the boxes.", vbExclamation
        Exit Sub
    End If
    
    If iNumberOfItemsIncluded = 0 And Not fTimesToBeSet And iExistingPending <> 0 Then
        Reply = MsgBox("You have not chosen to upload anything, but there are still uploads" & vbCrLf & _
            "pending (see list below). Click OK to continue.", vbOKCancel)
        If Reply = vbCancel Then
            Exit Sub
        End If
    End If
    
    If iNumberOfItemsIncluded <> 0 Or fTimesToBeSet Then
        If iExistingPending <> 0 Then
            sOther = " as well as the existing pending uploads "
        End If
        If which_ones = CURRENT_UPLOADS Then
            sOther = "(NOTE: the existing pending uploads will be ignored in this session)"
        End If
        Reply = MsgBox("You have chosen to upload the following files:" & vbCrLf & vbCrLf & _
            sUploadList & vbCrLf & sFragment & vbCrLf & sOther & vbCrLf & _
            "TO " & vbCrLf & vbCrLf & _
            sFranchiseList & vbCrLf & sWhen, vbOKCancel)
    End If
    
    If Reply = vbOK Then
        Set rstUploadTasks = GetRst(pCnn:=g.cnnDW, _
                                    pSource:="FranchiseUploads", _
                                    pSourceType:=adCmdTable, _
                                    pRstType:=eEditableFwdOnly, _
                                    pErrMsg:=strErrMsg)
        For iIndex = 1 To iNumberOfFranchisesIncluded
        ' First check if this is a 'live' store. If not live, we don't want to send stuff to it.
        '   Uploads are created for all live stores INCLUDING those excluded from capture cycle (i.e. FranchiseIncludedInStatistics = False)
        '   Stores may be temporarily excluded and later included in which case they will receive the uploads created for them
        '   Pending uploads are automatically purged as part of the capture cycle according to TsgDwMdb.Defaults!MonthsOfFranchiseUploads
            lngFranID = alngFranchiseID(iIndex)
            strSQL = "SELECT * FROM qryFranchiseLive WHERE FranchiseIDTSG = " & lngFranID
            Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            bContinue = Not (rstFran.BOF And rstFran.EOF)
            If bContinue Then
                lngFranType = Cn(rstFran!FranchiseType, 100)
                bFranMsgFlag = CBool(rstFran!FranchiseMessageFlag)
            End If
            
            rstFran.Close
            Set rstFran = Nothing
            If bContinue Then
                ' Create an entry for the franchise, for each file that needs to be uploaded
                For iInner = 1 To iNumberOfItemsIncluded
                    ' Only add a new pending-upload entry if one doesn't exist for this file
                    ' going up to this franchise
                    If fOKToUploadItem(pFranID:=lngFranID, _
                                       pFranType:=lngFranType, _
                                       pFranMsgFlag:=bFranMsgFlag, _
                                       pUploadFile:=aszUploadItem(iInner)) Then
''' Review Note that if a pending upload does exist for this item and fran it should probably have its UploadCurrentSession
''' Revew value set appropriately if it is not already set.
                        rstUploadTasks.AddNew
                            rstUploadTasks(gconUploadFranchiseIDField) = alngFranchiseID(iIndex)
                            rstUploadTasks(gconUploadFileField) = aszUploadItem(iInner)
                            sBaseName = fGetLastWord(aszUploadItem(iInner), "\")
                            If (Left(sBaseName, 5) = gconNewStockFilePrefix) Or _
                               (Left(sBaseName, 5) = gconWLPUpgradePrefix) Or _
                               (Left(sBaseName, 5) = gconUpdateStkFldsUpdatePrefix) Then
                                SetAttr aszUploadItem(iInner), vbReadOnly
                            End If
                            rstUploadTasks!UploadCurrentSession = CBoolMySql(which_ones = CURRENT_UPLOADS)
                        rstUploadTasks.Update
                    End If
                Next iInner
                
                If chkResetRemoteOpenedBy Then
                    If fOKToUploadItem(pFranID:=lngFranID, _
                                       pFranType:=lngFranType, _
                                       pFranMsgFlag:=bFranMsgFlag, _
                                       pUploadFile:=gconRemoteDefaultsTableDatabaseOpenedByField) Then
                        rstUploadTasks.AddNew
                            rstUploadTasks(gconUploadFranchiseIDField) = alngFranchiseID(iIndex)
                            rstUploadTasks(gconUploadFileField) = gconRemoteDefaultsTableDatabaseOpenedByField
                            rstUploadTasks!UploadCurrentSession = CBoolMySql(which_ones = CURRENT_UPLOADS)
                        rstUploadTasks.Update
                    End If
                End If
            End If
        Next iIndex
        
        rstUploadTasks.Close
        Set rstUploadTasks = Nothing
        
        DisplayUploadsPending
        
        If sWhen = conImmediately Then
            Call UploadFilesToFranchises(which_ones)
        End If
    End If
    ' clear arrays, release recordsets, close etc

End Sub

Sub DialupResults(ByVal fPrint As Boolean)
    Dim strSQL As String
    Dim strWC As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset
    
    If Not fPrint Then
        If gbEventLogRefreshIsNotAlreadyInProgress Then
            gbEventLogRefreshIsNotAlreadyInProgress = False
            With frmTSGDataWarehouse.stb
                .SimpleText = "displaying dialup results, please wait..."
                .Refresh
            End With
            With frmTSGDataWarehouse.lvwEventLog
                .ListItems.Clear
                .Refresh
            End With
        End If
    End If

    strWC = "FranchiseIncludedInStatistics"
    If optDialupResults(1) Then
        strWC = strWC & " AND (FranchiseLastDialupResult LIKE " & SqlQ("%failed%") & ")"
    End If
    strSQL = "SELECT * FROM franchises WHERE " & strWC & " ORDER BY FranchiseBusinessName"
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    Do Until rst.EOF
        If fPrint Then
            Printer.Print rst(gconFranchiseTableBusinessNameField); _
                          Tab(30); _
                          rst!FranchiseLastDialupResult
        Else
            Set gvListItem = frmTSGDataWarehouse.lvwEventLog.ListItems.Add()
            gvListItem.Text = ""
            Call gsubAddSubItemToListview(rst(gconFranchiseTableBusinessNameField), 1)
            Call gsubAddSubItemToListview(rst!FranchiseLastDialupResult, 2)
        End If
        rst.MoveNext
    Loop
    
    rst.Close
    Set rst = Nothing

    If fPrint Then
        ' close file & setup headings
        Printer.EndDoc
        Call gsubRefreshEventLogDisplay
    Else
        gbEventLogRefreshIsNotAlreadyInProgress = True
        If g.rstAppDefaults!NetworkPrinterEnabled Then
            If MsgBox("Do you want to print the dial-up results?", vbYesNo) = vbYes Then
                Call PrintDialupResults  '** PROCEDURE CALLS ITSELF VIA THIS PROCEDURE **'
            End If
        End If
    End If

    With frmTSGDataWarehouse.stb
        .SimpleText = "Ready..."
        .Refresh
    End With

End Sub

Sub disableSettingsFields()
    txtConfirmPassword.Text = vbNullString
End Sub

'--------------------------------------------------------------------------------------------------'
' displayUploadsPending
'    Returns the number of items to be uploaded as the return code.
'--------------------------------------------------------------------------------------------------'
Function DisplayUploadsPending() As Integer
Dim lngFranCount As Long
Dim lngUploadCount As Long
Dim strMsg As String
Dim strSQL As String
Dim strErrMsg As String
Dim rsUploadsPending As ADODB.Recordset

    strSQL = "SELECT COUNT(DISTINCT FranchiseID) " & vbNewLine & _
             "FROM FranchiseUploads " & vbNewLine & _
             "WHERE UploadedDate IS NULL"
    lngFranCount = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL)
    
    With lvwUploadsPending
        .ListItems.Clear
        .Refresh
    End With
    
    strSQL = "SELECT * FROM FranchiseUploads WHERE UploadedDate IS NULL"
    
    Set rsUploadsPending = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    Do Until rsUploadsPending.EOF
        Set gvListItem = lvwUploadsPending.ListItems.Add()
        gvListItem.Text = GetFranName(rsUploadsPending(gconUploadFranchiseIDField))
        Call gsubAddSubItemToListview(rsUploadsPending(gconUploadFileField), 1)
        rsUploadsPending.MoveNext
    Loop
    
    rsUploadsPending.Close
    Set rsUploadsPending = Nothing
    
    lngUploadCount = lvwUploadsPending.ListItems.Count
    strMsg = Plural(pQty:=lngUploadCount, pNounSingular:="upload") & " pending for " & lngFranCount & " franchises"
    lblUploadsPending.Caption = strMsg
    
    DisplayUploadsPending = lngUploadCount

End Function

Private Sub dtpBataTabTxDate_Change()
    RefreshBataTabGrid
End Sub

Private Sub dtpPromoEnd_Change()
    dtpPromoStart.MaxDate = dtpPromoEnd.Value
End Sub

Private Sub dtpPromoStart_Change()
    dtpPromoEnd.MinDate = dtpPromoStart.Value
End Sub

Private Sub EnablePromoEditCtls(ByVal pEnabled As Boolean)
Dim lbl As VB.Label

    fraAddPromotion.Enabled = pEnabled
    
    For Each lbl In Me.lblAddPromotion
        lbl.Enabled = pEnabled
    Next lbl
    
    dtpPromoStart.Enabled = pEnabled
    dtpPromoEnd.Enabled = pEnabled

    lstPromoSubCat.Enabled = pEnabled
    lstPromoProducts.Enabled = pEnabled
    lstPromoTabState.Enabled = pEnabled
    lstPromoTabRegion.Enabled = pEnabled And (chkPromoTabAllRegions.Value <> vbChecked)
    lstPromoTabState.Enabled = pEnabled And (chkPromoTabAllStates.Value <> vbChecked)
    
    cmdPromotion(0).Enabled = pEnabled  ' "Save Promotion"
    cmdPromotion(1).Enabled = pEnabled  ' "Clear"
    
End Sub

Function fbBarcodeIsATobaccoProduct(ByVal sBarcode As String) As Boolean
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpBarcode As ADODB.Recordset

    strSQL = "SELECT barcode FROM stock WHERE barcode  = " & SqlQ(sBarcode)
    Set rstSnpBarcode = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    
    fbBarcodeIsATobaccoProduct = Not (rstSnpBarcode.BOF And rstSnpBarcode.EOF)
    
    rstSnpBarcode.Close
    Set rstSnpBarcode = Nothing
    
End Function

Function fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared() As Boolean

    fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared = True
    If Dir(gsProductReportPathAndFilename) <> "" Then 'a previous report exists
        If MsgBox("Previous product report already exists. Do you want to overwrite [erase]", _
            vbYesNo + vbDefaultButton2) = vbYes Then
            SetAttr gsProductReportPathAndFilename, vbNormal
            Kill gsProductReportPathAndFilename
        Else
            fbNewProductReportDocumentEnvironmentWasSuccessfullyPrepared = False
        End If
    End If

End Function

Function fbNewStickReportDocumentEnvironmentWasSuccessfullyPrepared() As Boolean

    fbNewStickReportDocumentEnvironmentWasSuccessfullyPrepared = True
    If Dir(gsStickReportPathAndFilename) <> "" Then 'a previous report exists
        If MsgBox("Previous stick report already exists. Do you want to overwrite [erase]", _
            vbYesNo + vbDefaultButton2) = vbYes Then
            SetAttr gsStickReportPathAndFilename, vbNormal
            Kill gsStickReportPathAndFilename
        Else
            fbNewStickReportDocumentEnvironmentWasSuccessfullyPrepared = False
        End If
    End If

End Function

Function fbProductIsIncludedInThisProductReport(ByVal sBarcode As String) As Boolean
Dim strErrMsg As String
Dim rstSnpBarcode As ADODB.Recordset
Dim strSQL As String
    
    strSQL = "SELECT Barcode FROM Stock WHERE Barcode = " & SqlQ(sBarcode)
    Set rstSnpBarcode = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
    If Not (rstSnpBarcode.BOF And rstSnpBarcode.EOF) Then  'the product is in the database, so include it
        fbProductIsIncludedInThisProductReport = True
    Else 'determine whether non-tobacco related products are included
        If chkNonTobaccoBarcodesAreIncluded Then
            fbProductIsIncludedInThisProductReport = True
        Else
            fbProductIsIncludedInThisProductReport = False
        End If
    End If
    rstSnpBarcode.Close
    Set rstSnpBarcode = Nothing
    
End Function

Function fbProductIsIncludedInThisStickReport(ByVal sBarcode As String) As Boolean
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpBarcode As ADODB.Recordset
    
    strSQL = "SELECT Barcode, cat1 FROM Stock " & vbNewLine & _
             "WHERE (Barcode = " & SqlQ(sBarcode) & ") " & _
             " AND ((cat1 = " & SqlQ(gkCAT_CigCtn) & ") OR (cat1 = " & SqlQ(gkCAT_CigPkt) & "))"
    Set rstSnpBarcode = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstSnpBarcode.BOF And rstSnpBarcode.EOF) Then  'the product is in the database, so include it
        fbProductIsIncludedInThisStickReport = True
    End If
    rstSnpBarcode.Close
    Set rstSnpBarcode = Nothing
    
End Function

Function fConnectFranchiseMapShareDisk(ByRef prstFran As ADODB.Recordset, Optional ByVal pAttemptNumber As Integer = 1) As Boolean

''' Review ***************************************************************************************************'
''' Review NB Should have a pErrString parameter passed ByRef then the calling procedure could use the return '
''' Review Success value and pErrString parameter and make any necessary calls to subUpdateFranDialupResult() '
''' Review This proc is called from subCaptureData() and UploadFilesToOneFranchise()                          '
''' Review ***************************************************************************************************'
'   AUrban Function should be re-written to call helper functions
'   AUrban The first to establish the network connection and the 2nd to establish the network share
                    
    Const kMaxNetAddTries As Long = 3
    Dim lngNetAddRetries As Long
    Dim strFranName As String
    Dim strNetErrMsg As String

Dim bNetworkConnection As Boolean
Dim bShareDriveConnected As Boolean
Dim strFranVpnIpAddress As String
Dim strRemotePath As String

    strFranName = prstFran!FranchiseBusinessName.Value
    strFranVpnIpAddress = prstFran!VpnIpAddress.Value
    
    strRemotePath = GetRemotePath(prstFran:=prstFran)
    If IsUseLocalDriveFranFolder() Then
    '   Testing locally
        bShareDriveConnected = True
    Else
    '   Could change interface IncludeInCycle to exclude because it is also the exception
''' V397 Start
'''     If Not g.bVpnAvailable Then
'''     '   Use neither VPN or DialUp => VPN store (no dial up support) but VPN is down
'''     '   or unavailable on this machine (eg old master)
'''         StatusBar pMsg:="VPN unavailable", pFranchise:=strFranName
'''     Else
''' V397 End
        ''' Review
        '   AUrban Version 3.0.9059 Will consider multiplying the Timeout by the Attempt Number if any (or particular franchises)
        '   AUrban Version 3.0.9059 inexplicably fail IsPingSuccessful() test. Today Andy reported that although Cranbourne
        '   AUrban Version 3.0.9059 tests over the VPN during the day without fail, it has been failing VPN verification each night
        '   AUrban Version 3.0.9059 since 16Feb2007. The problem does not correlate with a new TsgDW version (from checking event log)
        '   AUrban Version 3.0.9059 (09Mar2007 Notes) 1/2 of Victoria has now dropped off the VPN due to a suspected problem with a router
        '   AUrban Version 3.0.9059 (09Mar2007 Notes) Possible that Cranbourne was an early sign of this so will hang off change for a while
        '   AUrban
        '   AUrban TURNS OUT THAT THE MACHINE WAS HIBERNATING AT NIGHT. THIS COULD BE FIXED BY GIVING IT A PING OR TWO TO WAKE IT UP
        
            StatusBar "Verify VPN connection for " & strFranVpnIpAddress & " Attempt: " & pAttemptNumber, strFranName
            bNetworkConnection = IsPingSuccessful(pIpAddress:=strFranVpnIpAddress, pTimeout:=3000)
            If Not bNetworkConnection Then
            
            '   PERHAPS WHEN WE COLLECT RECORDS WHEN TESTING FRANCHISE IT WILL BE MORE LEGITIMATE TO UPDATE FRAN DIALUP RESULT
                subUpdateFranDialupResult _
                    prstFran!FranchiseIDTSG, _
                    "(Attempt " & pAttemptNumber & ") " & Format$(Now, gkFmtDateTime) & " - Failed. Cannot verify VPN connection.)"
            End If
'''     End If  ''' V397
        
        If bNetworkConnection Then
            
            Do While (lngNetAddRetries < kMaxNetAddTries) And (Not bShareDriveConnected)
                lngNetAddRetries = lngNetAddRetries + 1
                
                StatusBar "Connect network share " & strRemotePath & " Attempt " & lngNetAddRetries, strFranName
                If NetConnectShareDisk(pRemoteName:=strRemotePath, _
                                       pErrMsg:=strNetErrMsg, _
                                       pUsr:=prstFran!FranchiseRASUsername, _
                                       pPwd:=prstFran!FranchiseRASPassword) Then
                    bShareDriveConnected = True
                Else 'network share connection was NOT successful
                    StatusBar "Unable to connect network share. " & strNetErrMsg, strFranName
                    
                '   PERHAPS WHEN WE COLLECT RECORDS WHEN TESTING FRANCHISE IT WILL BE MORE LEGITIMATE TO UPDATE FRAN DIALUP RESULT
                    subUpdateFranDialupResult _
                         prstFran!FranchiseIDTSG, _
                        "(Attempt " & pAttemptNumber & ") " & Format(Now, gkFmtDateTime) & _
                        " - Failed. Could not connect network share. " & strNetErrMsg
                End If 'is network share successfully connected ? if-endif
            'Next lngNetAddRetries
            Loop ' Do While lngNetAddRetries < kMaxNetAddTries And Not bShareDriveConnected
            
        End If
        
    End If
    
'   bShareDriveConnected is used as the Success return variable for the function
    fConnectFranchiseMapShareDisk = bShareDriveConnected
    
End Function

Function fGetStateFranchiseList(ByRef pFranchiseIDArray As Variant) As Integer
Const kListSep As String = ", "
Dim lngLoop As Long
Dim lngFranCount As Long
Dim strSQL As String
Dim strErrMsg As String
Dim strStateList As String
Dim vntFranIdArray As Variant
Dim chk As VB.CheckBox
Dim rst As ADODB.Recordset
    
    For Each chk In Me.chkUpload_State
        If chk.Value Then
            strStateList = strStateList & SQ(chk.Caption) & kListSep
        End If
    Next
    
    If Len(strStateList) Then
        strSQL = "SELECT FranchiseIDTSG FROM franchises WHERE " & vbNewLine & _
                 "FranchiseStateOfOz IN (" & Left$(strStateList, Len(strStateList) - Len(kListSep)) & ")"
        Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)

        With rst
            If Not (.BOF And .EOF) Then
            '   GetRows returns recordset as a two-dimensional array - Zero based index
            '   The first subscript identifies the field and the second identifies the record number
                vntFranIdArray = .GetRows
                lngFranCount = UBound(vntFranIdArray, 2) + 1
                ReDim pFranchiseIDArray(1 To lngFranCount)
                For lngLoop = 1 To lngFranCount
                    pFranchiseIDArray(lngLoop) = vntFranIdArray(0, lngLoop - 1)
                Next
            
            End If
        End With
        
        rst.Close
        Set rst = Nothing
    End If
    
    fGetStateFranchiseList = lngFranCount
    
End Function

Function fiSupplierIDForBarcode(ByVal sBarcode As String) As Integer
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpSupplierID As ADODB.Recordset
    
    strSQL = "SELECT supplier_id FROM Stock WHERE Barcode = " & SqlQ(sBarcode)
    Set rstSnpSupplierID = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstSnpSupplierID.BOF And rstSnpSupplierID.EOF) Then  'the product is in the database, so include it
        fiSupplierIDForBarcode = rstSnpSupplierID(gconSupplierTableSupplierIDField)
    Else 'determine whether non-tobacco related products are included
        fiSupplierIDForBarcode = gconOtherSuppliers
    End If
    rstSnpSupplierID.Close
    Set rstSnpSupplierID = Nothing

End Function

Function flSticksFrom(ByVal sBarcode As String) As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpSticks As ADODB.Recordset
    
    strSQL = "SELECT " & gconStockTableSticksField & gconSpace & _
            "FROM Stock " & _
            "WHERE " & gconStockTableBarcodeField & " = " & SqlQ(sBarcode)
    Set rstSnpSticks = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstSnpSticks.BOF And rstSnpSticks.EOF) Then
        flSticksFrom = rstSnpSticks(gconStockTableSticksField)
    End If
    rstSnpSticks.Close
    Set rstSnpSticks = Nothing

End Function

Function flSupplierIDFrom(ByVal sSupplierName As String) As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpSupplier As ADODB.Recordset
    
    strSQL = "SELECT supplier_id FROM Supplier " & vbNewLine & _
             "WHERE Supplier = " & SqlQ(sSupplierName)
    Set rstSnpSupplier = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstSnpSupplier.BOF And rstSnpSupplier.EOF) Then 'supplyar name was found in supplyar tayble
        flSupplierIDFrom = rstSnpSupplier(gconSupplierTableSupplierIDField)
    Else
        flSupplierIDFrom = gconOtherSuppliers
        MsgBox "Default supplier was used", vbInformation
    End If
    rstSnpSupplier.Close
    Set rstSnpSupplier = Nothing

End Function

'---------------------------------------------------------------------------------------------------
'  fOKToUploadItem
'
'  Check that it is OK to upload the item to the store.
'  It's OK if :
'  1. this franchise is able to receive this type of item (eg rogues don't get PM) AND
'  2. an upload pending entry doesn't already exist for the item to the store
'  If pFranID is zero, it means is there an upload for this file for ANY franchise.
'---------------------------------------------------------------------------------------------------
Function fOKToUploadItem(ByVal pFranID As Long, _
                         ByVal pFranType As Long, _
                         ByVal pFranMsgFlag As Boolean, _
                         ByVal pUploadFile As String) As Boolean
''' Review  PROCEDURE COULD DO WITH A REWRITE FOR READABILITY. POSSIBLY REPLACE ALL THE 'IF THENs' WITH A 'CASE STATEMENT'
Dim strSQL As String
Dim strErrMsg As String
Dim rstUploads As ADODB.Recordset

'   Exclude oPOS franchises from all uploads from TsgDw (includes gconUtilityExe)
    If pFranType = gkOPosFranType Then
        fOKToUploadItem = False
        Exit Function
    End If
    
    ' Check that if this is a 'msg' we're going to send, that this store actually
    ' wants to receive messages. Some stores don't want messages as indicated by the
    ' gconFranchiseMessageFlag flag
    If LCase(Left(fGetLastWord(pUploadFile, "\"), 5)) = LCase(gconNewMessageFilePrefix) Then
        If Not pFranMsgFlag Then
            fOKToUploadItem = False
            Exit Function
        End If
    End If
    
    ' Check that if the item is PM or NewStock or WLPUpdates or StkFieldUpdate or Promo reminder,
    ' don't send it to stores that do not use RM
    If (LCase(Left(fGetLastWord(pUploadFile, "\"), 5)) = LCase(gconNewStockFilePrefix)) Or _
       (LCase(Left(fGetLastWord(pUploadFile, "\"), 5)) = LCase(gconWLPUpgradePrefix)) Or _
       (LCase(Left(fGetLastWord(pUploadFile, "\"), 5)) = LCase(gconUpdateStkFldsUpdatePrefix)) Or _
       (LCase(fGetLastWord(pUploadFile, "\")) = LCase(gconPriceModule)) Then
        If pFranType > 29 Then
            fOKToUploadItem = False
            Exit Function
        End If
    End If
    ' Check that if the item is upgraderemotestatistics.exe or PM or RS or Promotion,
    ' don't send it to Cliff's stores (ie. those that don't have RS.exe, and their own special rs.mdb)
''' Review
'''*** DANGER DANGER DANGER  DANGER DANGER - DON'T UPLOAD upgraderemotestatistics.exe TO CLIFFY STORES  *
'''*** MAYBE EASIER TO DETERMINE WHAT IS ALLOWED THROUGH RATHER THAT WHAT ISN'T                         *
'''*** MIGHT NEED TO GET CLFFY TO RENAME REMOTE STATISTICS OR POSSIBLY MORE SIMPLY GET US TO STAMP      *
'''*** OUR VERSION INTERNALLY SO THAT OUR UPGRADE REMOTE STATISTICS CAN DETERMINE WHETHER OR NOT        *
'''*** IT SHOULD UPGRADE THE EXISTING VERSION OR WHETHER THAT VERSION MAY BE SOMEONE ELSES              *
'''*** ANOTHER SCHOOL OF THOUGHT WOULD BE THAT THERE IS NO DANGER BECUASE CLIFF's VERSION OF REMOTE     *
'''*** STATISTICS MAY NOT NECESSARILY HAVE THE CODE WHICH RUNS UTILITY PROGRAMS AS THEY ARRIVE          *
    If (LCase(fGetLastWord(pUploadFile, "\")) = LCase(gconUpgradeRS)) Or _
       (LCase(fGetLastWord(pUploadFile, "\")) = LCase(gsNewRSexe)) Or _
       (Left$(pUploadFile, 5) = gkPromoADD) Or _
       (Left$(pUploadFile, 8) = gkPromoDELETE) Then
        If pFranType > 59 Then
            fOKToUploadItem = False
            Exit Function
        End If
    End If
    
    strSQL = "SELECT * FROM FranchiseUploads " & vbNewLine & _
             "WHERE FranchiseID = " & pFranID & _
             " AND UploadFile = " & SqlQ(LCase(pUploadFile)) & _
             " AND UploadedDate IS NULL"
    
    Set rstUploads = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    
    fOKToUploadItem = (rstUploads.BOF And rstUploads.EOF)
    
    rstUploads.Close
    Set rstUploads = Nothing
    
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'   If user pressed Control-Shift-A on the setup form show the version
    If Shift And vbCtrlMask Then
        If Shift And vbShiftMask Then
            Select Case KeyCode
                Case vbKeyA
                    MsgBox GetVersionString()
            End Select
        End If
    End If
End Sub

Private Sub Form_Load()
Dim lngTabLoop As Long
Dim strSQL As String
Dim lngListLoop As Long
Dim strErrMsg As String
Dim rstFranTypes As ADODB.Recordset
    
'   Data Capture Tab
'   (Nb. Some ctls have visible borders in the designer to prevent overlapping )
'   (    during design, and the BorderStyle is then set to vbBSNone in code    )
    lblLocalDriveFranFolder.Visible = IsUseLocalDriveFranFolder()   ' Label to highlight testing with local franchise folder
    cmdTest.Visible = InStr(UCase$(VBA.Command$), "TESTMODE")       ' Button for testing new code, calling test code etc.
    txtStock_ID_DEV.Visible = IsIDE()
    lblStock_ID_DEV.Visible = IsIDE()
    lblVersion.BorderStyle = vbBSNone
    lblVersion.Caption = "Version " & fsVersion()
    lblDisplayedFranchise.BorderStyle = vbBSNone
    lblDisplayedFranchise.Caption = vbNullString
    
    txtActiveDatabase = "MySQL Database: " & _
                        Bracket(GetValStringValue(pVString:=g.cnnDW.ConnectionString, pVName:="DATABASE"))
    
    gCompanyIdentifier = g.rstAppDefaults!CompanyIdentifier

    With txtNewFranchiseBusinessName
        .Top = lstDataCaptureFranchiseBusinessName.Top + 30
        .Left = lstDataCaptureFranchiseBusinessName.Left + 30
        .Width = lstDataCaptureFranchiseBusinessName.Width - 60
    End With

    Me.Caption = g.strNodeType & " - " & App.Title

    Me.Show
    
    SetMousePointer vbHourglass
    
    SetGlobalVariables
    
    lstDataCaptureFranchiseBusinessName.SetFocus

'   BataRpt Tab
    grdBataRpts.MergeCol(-1) = True ' Merge  all cols (-1 for all) ie. cells along col can merge
    dtpBataTabTxDate.MinDate = GetMinBataTabTxDate()
    cboBataTabTxOrProcessedDate.ListIndex = 0 ' 0-TransactionDate,  1=ProcessedDate
    cmdBataTabUploadUnSent.Visible = False
    cmdBataTabUploadSelected.Visible = False
    
'   Promotions Tab
    chkPromoSelectFranchise_Click
    With fraPromoSelFranchise
        .Visible = ChkBoxToBool(chkPromoSelectFranchise)
        .Left = fraPromoSelectState.Left
        .Top = fraPromoSelectState.Top
    End With
    lblPromoFran.Left = lblPromoState.Left
    lblPromoFran.Top = lblPromoState.Top

'   Only Master PC is enabled for:-
'   - Data Capture (including importing Batscan files)
'   - Secure FTP of sales reports to BATA
'   - Creating Promotions
'   Only master should change data: data collection (auto or manual), adding/editing franchises
'   Only master is connected to VPN and can perform functions requiring VPN (uploading or downloading)
'   Non-master machines should be enquiries and reporting only
    tmrAutoDataCapture.Enabled = g.bMaster
    cmdCaptureData.Enabled = g.bMaster
    cmdImportBatscanFiles.Enabled = g.bMaster
    cmdBataTabUploadUnSent.Enabled = g.bMaster
    cmdBataTabUploadSelected.Enabled = g.bMaster
    EnablePromoEditCtls pEnabled:=g.bMaster
''' V397 Start
''' If g.bMaster Then
'''     g.bVpnAvailable = IsVpnAvailable()
'''     lblVpnDisabled.ForeColor = vbRed
'''     lblVpnDisabled.BorderStyle = vbBSNone
'''     lblVpnDisabled.Visible = Not g.bVpnAvailable
''' End If
''' V397 End

'   Enable/Disable tabs appropriately
'   Disable tabs as appropriate (Some tabs are disabled for the 'Business Manager PC')
    With tabMain
        .TabEnabled(TabEnum.eDataCaptureTab) = Not g.bBusinessMgrPC
        .TabEnabled(TabEnum.eSalesRptsTab) = True
        .TabEnabled(TabEnum.eStickRptsTab) = True
        .TabEnabled(TabEnum.eStockTab) = Not g.bBusinessMgrPC
        .TabEnabled(TabEnum.eBataTab) = Not g.bBusinessMgrPC
        .TabEnabled(TabEnum.eNielsenTab) = Not g.bBusinessMgrPC
        .TabEnabled(TabEnum.eVersionsTab) = Not g.bBusinessMgrPC
        .TabEnabled(TabEnum.eProductRptsTab) = True
        .TabEnabled(TabEnum.eSettingsTab) = Not g.bBusinessMgrPC
        .TabEnabled(TabEnum.eUploadsTab) = g.bMaster
        .TabEnabled(TabEnum.ePromotionsTab) = True
    End With

'   Set focus to first enabled tab
'   [If current tab is disabled in code it still retains focus AND FUNCTION. Once it has lost focus is cannot regain focus)
'   (=> move focus to first enabled tab which will move focus from current tab if it is been disabled)
    For lngTabLoop = TabEnum.eDataCaptureTab To TabEnum.ePromotionsTab
        If tabMain.TabEnabled(lngTabLoop) Then
            tabMain.Tab = lngTabLoop
            Exit For
        End If
    Next lngTabLoop
    
    strSQL = "SELECT * FROM tlkpRegion WHERE RegionID <> 0" ' Exclude ALL-STATES from available selections
    LoadCombo_Rst pCombo:=cboDCTabRegion, pCnn:=g.cnnDW, pSource:=strSQL, pDisplayFld:="RegionName", pDataFld:="RegionID"
    LoadCombo_Rst pCombo:=cboDCTabPromoGrade, pCnn:=g.cnnDW, pSource:="tlkpPromoGrade", pDisplayFld:="PromoGradeName", pDataFld:="PromoGradeID"
'   --------------------------------------------------------------------------------------------------------'
'   Note query only shows states with Franchise                                                             '
'   LoadCombo_Rst pCombo:=cboState, pCnn:=g.cnnDW, pSource:="SELECT * FROM qlkpState", pDisplayFld:="State" '
'   --------------------------------------------------------------------------------------------------------'
    
'   Populate Franchise Type combo and associated ToolTipText array
    LoadCombo_Rst pCombo:=cboFranchiseType, pCnn:=g.cnnDW, pSource:="FranTypes", pDisplayFld:="FranTypeName", pDataFld:="FranchiseType"
    ReDim m.astrFranTypeTooltip(0 To cboFranchiseType.ListCount - 1) As String
    Set rstFranTypes = GetRst(pCnn:=g.cnnDW, _
                              pSource:="FranTypes", _
                              pSourceType:=adCmdTable, _
                              pRstType:=eReadOnlyStatic, _
                              pErrMsg:=strErrMsg)   ' eReadOnlyStatic so can use find (cf FwdOnly)
    With rstFranTypes
        If Not (.BOF And .EOF) Then
            For lngListLoop = 0 To cboFranchiseType.ListCount - 1
                .MoveFirst
                .Find "FranchiseType=" & cboFranchiseType.ItemData(lngListLoop)
                If Not .EOF Then
                    m.astrFranTypeTooltip(lngListLoop) = Cn(rstFranTypes!FranTypeDescription, vbNullString)
                End If
            Next
        End If
    End With
    rstFranTypes.Close
    Set rstFranTypes = Nothing
     
    gsubRefreshEventLogDisplay

'   Helper routine subTabMainClick to be called from TabMain_Click. Routine is called from here to
'   ensure that no matter which tab the program  starts on the starting tab will be correctly initialised
    subTabMainClick pTab:=tabMain.Tab

    SetMousePointer vbDefault
    Me.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
    gsubAddToLocalEventLog App.Title & " terminated on " & g.strNodeName, pFranchise:=vbNullString
    g.rstAppDefaults.Close  ' Global rst that keeps exlusively opened cnnAppDefaults open
    CloseDatabaseCnns
End Sub

Function fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks(ByVal sOriginal As String) As String

    Dim lCharacterPosition As Long, _
        sCharacterBeingChecked As String * 1

    sOriginal = Trim(sOriginal)
    If sOriginal = "" Then 'avert a null
        fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks = ""
    ElseIf Len(sOriginal) = 1 Then
        fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks = UCase(sOriginal)
    Else
        fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks = UCase(Left(sOriginal, 1))
        For lCharacterPosition = 2 To Len(sOriginal)
            sCharacterBeingChecked = Mid(sOriginal, lCharacterPosition, 1)
            If Asc(sCharacterBeingChecked) = gconSpaceAscii Then 'it is a space
                'so add it
                fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks = fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks & gconSpace
                'reposition to uppercase the next character
                lCharacterPosition = lCharacterPosition + 1
                Do Until Asc(Mid(sOriginal, lCharacterPosition, 1)) <> gconSpaceAscii
                    lCharacterPosition = lCharacterPosition + 1
                Loop
                'upperase the character following the space
                sCharacterBeingChecked = UCase(Mid(sOriginal, lCharacterPosition, 1))
            ElseIf Asc(sCharacterBeingChecked) = gconSingleQuoteAscii Then
                sCharacterBeingChecked = Chr(gconNonDestructiveSingleQuoteAscii)
            Else
                sCharacterBeingChecked = LCase(sCharacterBeingChecked)
            End If
            fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks = _
                fsAfterEverySpaceCapFirstRemainderLowerAndReplaceQuoteMarks & _
                sCharacterBeingChecked
        Next lCharacterPosition
    End If

End Function

Function fsDescriptionFrom(ByVal sBarcode As String) As String
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset
    
    strSQL = "SELECT Barcode, Description FROM Stock " & vbNewLine & _
             "WHERE Barcode  = " & SqlQ(sBarcode)
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rst.BOF And rst.EOF) Then
        If g.rstAppDefaults(gconDocketPrinterEnabled) Then
            If Len(rst(gconStockTableDescriptionField)) > _
                                     gconTruncateDescriptionBriefAt + _
                                     Len(gconTruncateCharacter) + _
                                     gconTruncateExtensionWidth Then
                fsDescriptionFrom = _
                    Left(rst(gconStockTableDescriptionField), gconTruncateDescriptionBriefAt) & _
                    gconTruncateCharacter & _
                    Right(rst(gconStockTableDescriptionField), gconTruncateExtensionWidth)
             Else 'leave it as it is
                fsDescriptionFrom = rst(gconStockTableDescriptionField)
             End If
        Else
            fsDescriptionFrom = rst(gconStockTableDescriptionField)
        End If
    Else
        If g.rstAppDefaults(gconDocketPrinterEnabled) Then
            fsDescriptionFrom = sBarcode
        Else
            fsDescriptionFrom = "Barcode = " & sBarcode
        End If
    End If
    rst.Close
    Set rst = Nothing

End Function

Function fsFileContentsAndEOFMessage(ByVal sFullPathAndFileName As String) As String

    Dim sTextLine As String
    Dim lLineCounter As Long
    Dim intFileNum As Long
    
    intFileNum = FreeFile   ' Get unused file
    Open sFullPathAndFileName For Input As #intFileNum
    Do Until EOF(intFileNum)
        Line Input #intFileNum, sTextLine
        fsFileContentsAndEOFMessage = fsFileContentsAndEOFMessage & sTextLine & vbCrLf
        lLineCounter = lLineCounter + 1
        With stb
            .SimpleText = "Importing line " & lLineCounter
            .Refresh
        End With
    Loop
    Close #intFileNum
    If Trim(fsFileContentsAndEOFMessage) <> "" Then
        fsFileContentsAndEOFMessage = fsFileContentsAndEOFMessage & "[End of file]"
    Else
        fsFileContentsAndEOFMessage = "File contains no data"
    End If

End Function

Function fsFranchiseIDFrom(ByVal sFranchiseName As String) As String
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpFranchiseID As ADODB.Recordset
    
    strSQL = "SELECT FranchiseIDTSG FROM Franchises " & vbNewLine & _
             "WHERE FranchiseBusinessName = " & SqlQ(sFranchiseName) & ";"
    Set rstSnpFranchiseID = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)

    If Not (rstSnpFranchiseID.BOF And rstSnpFranchiseID.EOF) Then
        fsFranchiseIDFrom = rstSnpFranchiseID(gconLiveDataTableTSGFranchiseIDField)
    End If
    rstSnpFranchiseID.Close
    Set rstSnpFranchiseID = Nothing

End Function

Private Function fsNielsenFileSpecification(ByVal pDate As Date) As String
    fsNielsenFileSpecification = g.strNielsenRptsFolder & "\*" & Format$(pDate, gconFmtDateInNielsenFilename) & "*"  'to also include the zip
End Function

Function fsNielsenRptFilename(ByVal pStartDate As Date, ByVal pEndDate As Date) As String
Dim strDailyOrWeekly As String      ' Daily = "d" and Weekly = "" (ie is the default and original)

    If pStartDate = pEndDate Then
        strDailyOrWeekly = "d"
    End If
    
    fsNielsenRptFilename = gconNielsenFilePrefix & _
                                  strDailyOrWeekly & _
                                  Format$(pEndDate, gconFmtDateInNielsenFilename) & _
                                  gconTextFileSuffix
End Function

Function fsNielsenRptFullname(ByVal pStartDate As Date, ByVal pEndDate As Date) As String
    fsNielsenRptFullname = g.strNielsenRptsFolder & "\" & fsNielsenRptFilename(pStartDate:=pStartDate, pEndDate:=pEndDate)
End Function

Function fsSpacesRemovedFrom(ByVal sOriginal As String) As String

    Dim iCharacterPosition As Integer

    sOriginal = Trim(sOriginal)
    
    For iCharacterPosition = 1 To Len(sOriginal)
        If Mid(sOriginal, iCharacterPosition, 1) <> gconSpace Then
            fsSpacesRemovedFrom = fsSpacesRemovedFrom & LCase(Mid(sOriginal, iCharacterPosition, 1))
        End If
    Next iCharacterPosition

End Function

Private Function fsStockIDFrom(ByVal pStkDescription As String) As String
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpStockID As ADODB.Recordset

    strSQL = "SELECT barcode FROM Stock WHERE description = " & SqlQ(pStkDescription)
    Set rstSnpStockID = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstSnpStockID.BOF And rstSnpStockID.EOF) Then
        fsStockIDFrom = rstSnpStockID(gconLiveDataTableBarcodeField)
    Else
        fsStockIDFrom = vbNullString
    End If
    rstSnpStockID.Close
    Set rstSnpStockID = Nothing

End Function

Function fsSupplierIDFrom(ByVal sBarcode As String) As Long
Dim strErrMsg As String
Dim rstSnpSupplierID As ADODB.Recordset
Dim strSQL As String
    
    strSQL = "SELECT Barcode, supplier_id FROM Stock WHERE Barcode = " & SqlQ(sBarcode)
    Set rstSnpSupplierID = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
    If Not (rstSnpSupplierID.BOF And rstSnpSupplierID.EOF) Then
        fsSupplierIDFrom = rstSnpSupplierID(gconStockTableSupplierIDField)
    Else 'can't find the barcode in the stock table, so use a default supplyar
        fsSupplierIDFrom = gconOtherSuppliers
        MsgBox "Default supplier was used", vbInformation
    End If
    rstSnpSupplierID.Close
    Set rstSnpSupplierID = Nothing

End Function

Function fsSupplierNameFrom(ByVal lSupplierID As Long) As String
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpSupplier As ADODB.Recordset
    
    If lSupplierID = 0 Then
        fsSupplierNameFrom = "unknown"
    Else
        strSQL = "SELECT Supplier FROM Supplier " & vbNewLine & _
                 "WHERE supplier_id = " & lSupplierID
        Set rstSnpSupplier = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        If Not (rstSnpSupplier.BOF And rstSnpSupplier.EOF) Then 'supplyar name was found in supplyar tayble
            fsSupplierNameFrom = rstSnpSupplier!Supplier
        Else
            fsSupplierNameFrom = "????"
        End If
        rstSnpSupplier.Close
        Set rstSnpSupplier = Nothing
    End If

End Function

Private Function GetBlankPromoRebatesArray(ByVal pUsePromoGrades As Boolean) As Variant()
Dim lngRow As Long
Dim strErrMsg As String
Dim avntResult() As Variant
Dim rst As ADODB.Recordset
    
    If Not pUsePromoGrades Then
        ReDim avntResult(1 To 4, 1 To 1)
        avntResult(1, 1) = mkPromoGradeIdNA
        avntResult(2, 1) = vbNullString '   "N/A" empty string agrees with list of promos and printing promo list
        avntResult(3, 1) = 0
        avntResult(4, 1) = 0
    Else
        Set rst = GetRst(pCnn:=g.cnnDW, pSource:="tlkpPromoGrade", pSourceType:=adCmdTable, pErrMsg:=strErrMsg)
        If Not rst Is Nothing Then
            Do While Not rst.EOF
                lngRow = lngRow + 1
                ReDim Preserve avntResult(1 To 4, 1 To lngRow)
                avntResult(1, lngRow) = rst!PromoGradeID
                avntResult(2, lngRow) = rst!PromoGradeName
                avntResult(3, lngRow) = 0
                avntResult(4, lngRow) = 0
                rst.MoveNext
            Loop
            rst.Close
            Set rst = Nothing
        End If
    End If
    
    GetBlankPromoRebatesArray = avntResult

End Function

Private Function GetCaptureCycleDate() As Date
Const gkCaptureStartTime As Date = "12:00 AM"   ' NB Investigate effects before changing. May need to change values
                                                ' database values for TsgDw.Franchises!CaptureCycleDateOnLastDataCapture
                                                ' May be replaced with mdb fld once changes post Version 3.3.0 have settled
                                                ' however a lot of analysis will be required
Dim dat As Date

    If Time > gkCaptureStartTime Then
        dat = Date
    Else
        dat = fdtmYesterday()
    End If
        
    GetCaptureCycleDate = dat

End Function

Private Function GetDialSequenceFromState(ByVal pState As String) As Long
Dim lngResult As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rstState As ADODB.Recordset

    lngResult = 7 ' Default to lowest priority if state not found
    strSQL = "SELECT DialSequence FROM tlkpState WHERE StateOfOz = " & SqlQ(pState)
    Set rstState = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstState.BOF And rstState.EOF) Then
        lngResult = rstState.Fields!DialSequence.Value
    End If
    rstState.Close
    Set rstState = Nothing

    GetDialSequenceFromState = lngResult
    
End Function

Private Function GetFranIdColn_FP(ByVal pPromoID As Long, _
                                  ByVal pTfrStatus As FpTfrEnum) As VBA.Collection
Dim lngFranID As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset
Dim colFranIDs As VBA.Collection

    Set colFranIDs = New VBA.Collection
    
    strSQL = "SELECT FranchiseID " & vbNewLine & _
             "FROM tblFranchisePromotions " & vbNewLine & _
             "WHERE " & Bracket("PromotionID = " & pPromoID) & _
              " AND " & Bracket("TfrStatus = " & pTfrStatus)
    
    Set rst = GetRst(pCnn:=g.cnnDW, _
                     pSource:=strSQL, _
                     pSourceType:=adCmdText, _
                     pRstType:=eReadOnlyFwdOnly, _
                     pErrMsg:=strErrMsg)
                         
    If Not (rst Is Nothing) Then
        If Not (rst.BOF And rst.EOF) Then
            Do While Not rst.EOF
                lngFranID = rst.Fields!FranchiseID
                colFranIDs.Add Item:=lngFranID, Key:=CStr(lngFranID)
                rst.MoveNext
            Loop
        End If
    End If
    
    Set GetFranIdColn_FP = colFranIDs
    
End Function

Private Function GetMinBataTabTxDate() As Date
Dim strSQL As String
Dim dtmResult As Date
Dim dtmOldestSalesData As Date
Dim dtmOldestBataUpload As Date

'   MySQL Review
'   Perhaps should be automatic purging of tblBataUploads that keeps it in synch with Archive db
'   MinDate is set for control so that you can only select reports you are able to view

'   Get oldest TxDate from tblBataUploads
    strSQL = "SELECT MIN(TxDate) FROM tblBataUploads"
    dtmOldestBataUpload = GetRstVal(pCnn:=g.cnnDW, _
                                    pSource:=strSQL, _
                                    pDefaultVal:=DateSerial(Year:=2000, Month:=1, Day:=1))

    strSQL = "SELECT MIN(TransactionDate) FROM LiveData"
    dtmOldestSalesData = GetRstVal(pCnn:=g.cnnDW, _
                                   pSource:=strSQL, _
                                   pDefaultVal:=DateSerial(Year:=2000, Month:=1, Day:=1))

'   Get most recent of the two dates to ensure you can only select reports that can be viewed
    If dtmOldestSalesData < dtmOldestBataUpload Then
        dtmResult = dtmOldestBataUpload
    Else
        dtmResult = dtmOldestSalesData
    End If

    GetMinBataTabTxDate = dtmResult

End Function

Private Function GetNeilsenDailyRptFullname(ByVal pLastReportEndDate As Date) As String
    GetNeilsenDailyRptFullname = GetNeilsenRptsSubFolderName(pLastReportEndDate) & "\DailyRpts.zip"
End Function

Private Function GetNeilsenRptsSubFolderName(ByVal pLastReportEndDate As Date) As String
'   pLastReportEndDate passed so calling programs handle change of date during their processing
    GetNeilsenRptsSubFolderName = g.strNielsenRptsFolder & "\" & Format$(pLastReportEndDate, "yyyy-mm-dd")
End Function

Private Function GetNeilsenWeeklyRptFullname(ByVal pLastReportEndDate As Date) As String
    GetNeilsenWeeklyRptFullname = GetNeilsenRptsSubFolderName(pLastReportEndDate) & "\WeeklyRpts.zip"
End Function

Private Function GetPromoGradesCollection() As VBA.Collection
'   Returns collection of Promo Grade Names using String value of PromoGradeID as the Key
'   Includes all PromoGrades from tlkpPromoGrade
Dim strErrMsg As String
Dim col As VBA.Collection
Dim rst As ADODB.Recordset

    Set col = New VBA.Collection
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:="tlkpPromoGrade", pSourceType:=adCmdTable, pErrMsg:=strErrMsg)
    If Not rst Is Nothing Then
        With rst
            Do While Not .EOF
                col.Add Item:=rst!PromoGradeName.Value, Key:=CStr(rst!PromoGradeID.Value)
                .MoveNext
            Loop
        End With
    End If
    
'   Ignore error if entry already exists in lookup table
    On Error Resume Next
        col.Add Item:=vbNullString, Key:=CStr(mkPromoGradeIdNA) ' could add n/a instead of vbNullString (GetRegionsCollection & GetPromoGradesCollection)
    On Error GoTo 0
    
''' Review
''' V373 Start - Re-instate code below b/c run-time error occured b/c of 2006 promo data with 0 for PromoID
''' Code below is a strong candidate for removal but need to check TSG Data
''''   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
''''   ''' CODE SECTION CAN BE REMOVED ONCE PROMOTIONS TABLE                                   '~
''''   ''' NO LONGER CONTAINS ANY ROWS WHERE PromoGradeID = 0                                  '~
''''   Add extra item to cater for Promotions created before PromoGrade was added              '~
''''   When PromoGradeID was added to Promotions table existing rows were given a value of zero'~
''''   Ignore error if (0, vbNullString) record has been added to tlkpPromoGrade)              '~
'''    On Error Resume Next                                                                    '~
'''        col.Add Item:=vbNullString, Key:=CStr(0) ' Zero not a capitol o                     '~
'''    On Error GoTo 0                                                                         '~
'   Add extra item to cater for Promotions created before PromoGrade was added                 '~
'   When PromoGradeID was added to Promotions table existing rows were given a value of zero   '~
'   Ignore error if (0, vbNullString) record has been added to tlkpPromoGrade)                 '~
    On Error Resume Next                                                                       '~
        col.Add Item:=vbNullString, Key:=CStr(0) ' Zero not a capitol o                        '~
    On Error GoTo 0                                                                            '~
''''   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Set GetPromoGradesCollection = col

End Function

Private Function GetRegionsCollection() As VBA.Collection
'   Returns collection of Region Names using String value of RegionID as the Key
'   Includes all regions from tlkpRegion and the special case of "ALL REGIONS" with a RegionID of mkPromoRegionsAll (ie zero)
Dim strErrMsg As String
Dim col As VBA.Collection
Dim rst As ADODB.Recordset

    Set col = New VBA.Collection
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:="tlkpRegion", pSourceType:=adCmdTable, pErrMsg:=strErrMsg)
    If Not rst Is Nothing Then
        With rst
            Do While Not .EOF
                col.Add Item:=rst!RegionName.Value, Key:=CStr(rst!RegionID.Value)
                .MoveNext
            Loop
        End With
    End If
    
'   Ignore error if entries already exist in lookup table
    On Error Resume Next
        col.Add Item:=vbNullString, Key:=CStr(mkPromoRegionsNA) ' could add n/a instead of vbNullString (GetRegionsCollection & GetPromoGradesCollection)
        col.Add Item:="ALL REGIONS", Key:=CStr(mkPromoRegionsAll)
    On Error GoTo 0
    
    Set GetRegionsCollection = col

End Function

Private Function GetRemotePath(ByRef prstFran As ADODB.Recordset) As String
Dim strResult As String
Dim strFranVpnIpAddress As String

    If IsUseLocalDriveFranFolder() Then
    '   Testing with franchise RStats mdb file on local machine
        strResult = gkLocalDriveFranchiseFolder
    Else
        strFranVpnIpAddress = prstFran!VpnIpAddress
        If IsIPAddress(strFranVpnIpAddress) Then
            strResult = "\\" & strFranVpnIpAddress & "\Statistics"
        Else
            strResult = "\\" & prstFran!FranchiseNodename & "\Statistics"
        End If
    End If

    GetRemotePath = strResult
    
End Function

Private Function GetRstNonCompliantRpt(Optional ByRef pColSelFranIDs As VBA.Collection) As ADODB.Recordset
Dim lngFranCount As Long
Dim dtmSalesDate As Date
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset
    
'   Update data tables as required
    If GetTableUpdateTime("LiveData") > GetTableUpdateTime("PromoNonCompliantSales") Then
        StatusBar pMsg:=UCase$("Updating PromoNonCompliantSales in call to GetRstNonCompliantRpt()")
        LoadNonCompliantTable pFranCount:=lngFranCount
    End If
    
    dtmSalesDate = fdtmYesterday()
    strSQL = "SELECT * FROM ((PromoNonCompliantSales AS PNC LEFT JOIN Franchises AS F ON PNC.FranchiseIDTSG = F.FranchiseIDTSG" & vbNewLine & _
                           " ) LEFT JOIN Stock AS S ON PNC.Barcode = S.Barcode" & vbNewLine & _
                           ")  LEFT JOIN Promotions AS P on PNC.PromoID = P.PromoID" & vbNewLine & _
             "WHERE (PNC.TransactionDate = " & MySqlDate(dtmSalesDate) & ")" & _
              " AND (NOT StoreNotified) " & vbNewLine & _
             "ORDER BY FranchiseBusinessName, PromoSubCat, Description"
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pRstType:=eEditableDynamic, pErrMsg:=strErrMsg)
    Set rst.ActiveConnection = Nothing  ' Disconnect rst
    
    If Not (pColSelFranIDs Is Nothing) Then
        Do While Not rst.EOF
            If Not IsKeyInCollection(pKey:=CStr(rst("PNC.FranchiseIDTSG")), pCollection:=pColSelFranIDs) Then
                rst.Delete
            End If
            rst.MoveNext
        Loop
    End If
    
'   MUST use RecordCount for case without current record (ie Empty rst with EOF=True and BOF=False)
'   Perhaps this odd behaviour is to do with it being a disconnected rst. As a now disconnected rst
'   can I punt on it being able to support RecordCount even though originally connected to MySQL db
    If rst.RecordCount = 0 Then
    '   NB Have traversed all records
        rst.Close
        Set rst = Nothing
    Else
        rst.MoveFirst
    End If
    
    Set GetRstNonCompliantRpt = rst

End Function

Private Function GetSelFranCollection(ByVal pSelFranEnum As SelFranEnum, _
                             Optional ByVal pSelFld As String = vbNullString) As VBA.Collection
Dim strSQL As String
Dim strFldList As String
Dim strFldSelected As String
Dim strErrMsg As String
Dim col As VBA.Collection
Dim rst As ADODB.Recordset

' Could add a case for included but previously collected
' This could then be run prior to manual re-runs but the question is would we
' really want to see a huge list, probably not but we probably would like to
' see a number for how many have been previously collected (perhaps
' only when a number has been previously collected (eg don't log zero previously
' collected) or perhaps only when it is a manual cycle

    If Len(pSelFld) = 0 Then
        strFldList = "FranchiseIDTSG"
        strFldSelected = "FranchiseIDTSG"
    Else
        strFldList = "FranchiseIDTSG, " & pSelFld
        strFldSelected = pSelFld
    End If


    Select Case pSelFranEnum
        
       Case SelFranEnum.eSelFran_CaptureCycleExcluded
            strSQL = "SELECT " & strFldList & " FROM qryFranchiseLive" & vbNewLine & _
                     "WHERE NOT FranchiseIncludedInStatistics" & vbNewLine & _
                     "ORDER BY FranchiseBusinessName"
    
       Case SelFranEnum.eSelFran_CaptureCycleManual
            strSQL = "SELECT " & strFldList & " FROM qryFranchiseLive" & vbNewLine & _
                     "WHERE (FranchiseIncludedInStatistics " & _
                       "AND (CaptureCycleDateOnLastDataCapture < " & MySqlDate(GetCaptureCycleDate()) & "))" & vbNewLine & _
                     "ORDER BY FranchiseDialSequence, FranchiseBusinessName"
                     
        Case SelFranEnum.eSelFran_CaptureCycleAuto
            strSQL = "SELECT " & strFldList & " " & vbNewLine & _
                     "FROM qryFranchiseLive AS a LEFT JOIN qryFranchisesPendingUpload AS b ON a.FranchiseIDTSG = b.FranchiseID" & vbNewLine & _
                     "WHERE (FranchiseIncludedInStatistics " & _
                       "AND (CaptureCycleDateOnLastDataCapture < " & MySqlDate(GetCaptureCycleDate()) & "))" & vbNewLine & _
                        "OR (FranchiseIncludedInStatistics AND UploadPending)" & vbNewLine & _
                     "ORDER BY FranchiseDialSequence, FranchiseBusinessName"
   End Select
    
    Set col = New VBA.Collection
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not rst Is Nothing Then
        Do While Not rst.EOF
            col.Add Item:=rst(strFldSelected).Value, Key:=CStr(rst!FranchiseIDTSG.Value)
            rst.MoveNext
        Loop
    End If
    
    Set GetSelFranCollection = col

End Function

Private Function GetStkValue(ByVal pFldName As String, pStkId As Long) As Variant
Dim strSQL As String
Dim strErrMsg As String
Dim vntResult As Variant
Dim rst As ADODB.Recordset

    strSQL = "SELECT " & pFldName & " FROM Stock WHERE stock_id = " & pStkId
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not rst Is Nothing Then
        If Not (rst.BOF And rst.EOF) Then
            vntResult = rst(pFldName)
        End If
        rst.Close
        Set rst = Nothing
    End If

    GetStkValue = vntResult

End Function

Private Function GetTabRefreshedFlag(ByVal pTab As TabEnum) As Boolean
    GetTabRefreshedFlag = m.ablnTabRefreshed(pTab)
End Function

Private Function GetValList(ByRef pSrcFldCollection As VBA.Collection, _
                            ByVal pAddDateStamp As Boolean) As String
''' FUNCTION MAY ALSO BE LATER MODIFIED TO REPLACE TransferSalesRecord which is used to tfr recs from
''' TempData table in subCaptureData (eg TransferSalesRecord rstDWTempData, rstDWPreLiveData, False)

''' Review May be more efficient on MySQL using an INSERT statememt
  
'' INCREMENTALLY SPEED UP ALL THIS CODE. EACH REFERENCE TO COLLECT A FIELD VALUE FROM A RST IS
'' A TRIP TO THE SERVER. BY COLLECTING THEM AT ONCE WE COULD AT LEAST GET SOME SPEED BENEFITS

Const kSep As String = ", "
Dim strResult As String
Dim vntValue As Variant

'   Can't use this code as passing an object (even by value) preserves the object referenece/pointer
'   NOT the data. (i.e. the item added to the collection). Regardless of how the collection is passed
'   the calling routine object reference will point to the data containing the newly added item.
'~  If pAddDateStamp Then
'~      pSrcFldCollection.Add Item:=Date, Key:="TodaysDate"
'~  End If
    
    For Each vntValue In pSrcFldCollection
        Select Case VarType(vntValue)
            Case vbDate
                strResult = strResult & MySqlDateTime(vntValue) & kSep
            Case vbString
                strResult = strResult & SqlQ(vntValue) & kSep
            Case Else
                strResult = strResult & vntValue & kSep
        End Select
    Next vntValue
    
    If pAddDateStamp Then
        strResult = strResult & MySqlDateTime(Date) & kSep
    End If
    
    If Len(strResult) Then
        GetValList = Bracket(Left$(strResult, Len(strResult) - Len(kSep)))
    End If

End Function

Private Function GetWcValueListFromColn(ByRef pCollection As VBA.Collection) As String

 'replace with GetWcListFromColn
' Returns an empty string if pCollection is empty -> calling code can test for empty string.
' A SQL statement with an empty Value List in Where Clause {eg. SELECT ... WHERE id IN ()}
' is invalid SQL syntax.

Const kSeparator As String = ", "
Dim strFranIdList As String
Dim vntFranID As Variant

    For Each vntFranID In pCollection
        strFranIdList = strFranIdList & vntFranID & kSeparator
    Next vntFranID
    
    If Len(strFranIdList) Then
        strFranIdList = Left$(strFranIdList, Len(strFranIdList) - Len(kSeparator))
        strFranIdList = Bracket(strFranIdList)
    End If

    GetWcValueListFromColn = strFranIdList
    
End Function

Private Sub grdBataRpts_AfterSelChange(ByVal OldRowSel As Long, ByVal OldColSel As Long, ByVal NewRowSel As Long, ByVal NewColSel As Long)
    ConfigureBataTabButtons
End Sub

Private Sub grdBataRpts_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    GridSetToolTip grdBataRpts
End Sub

Private Sub grdPromoTabRebates_AfterEdit(ByVal Row As Long, ByVal col As Long)
' Values entered into Row 1 are propogated into succeeding rows
Dim lngLoop As Long

    With grdPromoTabRebates
        If Trim$(.TextMatrix(Row:=Row, col:=col)) = vbNullString Then
            .TextMatrix(Row:=Row, col:=col) = 0
        End If
        If Row = 1 Then
            For lngLoop = .FixedRows To .Rows - 1
                .TextMatrix(Row:=lngLoop, col:=col) = .TextMatrix(Row:=.FixedRows, col:=col)
            Next lngLoop
        End If
    End With
    
End Sub

Private Sub grdPromoTabRebates_BeforeEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
'   Prevent editing of Grade column
    Cancel = grdPromoTabRebates.ColKey(col) = "Grade" _
          Or grdPromoTabRebates.ColKey(col) = "PromoGradeID"
End Sub

Private Sub grdPromoTabRebates_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
'   Allow digits, decimal point & edit keys mapped by KeyAscii (BACKSPACE or ENTER) otherwsie set KeyAscii to 0
    If (col = 2) Or (col = 3) Then
        If Not (((KeyAscii >= vbKey0) And (KeyAscii <= vbKey9)) Or _
                (KeyAscii = Asc(".")) Or _
                (KeyAscii = Asc(vbBack)) Or (KeyAscii = Asc(vbCr))) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub HighlightSelectedFranchiseInEventLog()
' Highlight event Log item for franchise selected in listbox
Const lvwWholeWord As Long = 0
Dim itmFound As ComctlLib.ListItem      ' Listview item variable.

    If lstDataCaptureFranchiseBusinessName.ListIndex <> -1 Then ' Is franchise selected list box
        Set itmFound = lvwEventLog.FindItem(lstDataCaptureFranchiseBusinessName, lvwSubItem, , lvwWholeWord)
        If Not itmFound Is Nothing Then ' Item Selected
            itmFound.EnsureVisible      ' Scroll ListView to show found ListItem.
            itmFound.Selected = True    ' Select the ListItem.
        Else
        '   No log item for franchise selected in list box therefore
        '   if is an item already highlighted in event log then un-highlight it.
            Set itmFound = lvwEventLog.SelectedItem
            If Not itmFound Is Nothing Then
                itmFound.Selected = False
            End If
        End If
        Set itmFound = Nothing
    End If

End Sub

Private Function ImportBatscanFile(ByVal pFullName As String) As Boolean
' Imports EOD BatScan Sales Summary File: (GMPS###_YYYYMMDD.txt[~00000000000000])
' EOD BatScan Sales Summary Files contain both retail & wholesale sales
' Returns True for a successfully processed file and False otherwise
' Batscan files are produced by franchises using POS Solutions software

' This procedure imports Batscan files into the prelivedata table. Data is transferred from the prelivedata
' table to the livedata table during the Data Capture cycle by TfrAllPreLiveDataToLiveData()
' The same data file may be imported multiple times. In this procedure the most recent data imported for a
' Franchise/SalseDate/Barcode combination replaces any existing data for that combination in the prelivedata table.
' This replicates the processing BATA applies to Batscan files transmitted directly to them
' Dexter Tabeta of BATA confirmed in an email (Tue 13/03/07 3:15 PM) that the BATScan load program worked
' on 'insert or else update' logic such that it would overwrite the first occurence with the second occurence.
' From the 4 sample batscan files files acquired 3 had duplicate data (ie multiple records for same Franchise/SalseDate/Barcode combination)
' Existing data for a franchise/date combination is not deleted before importing data for that same combination.
' It could be argued that all data for a franchise/date combination should be deleted before importing replacement data so that if
' incorrect barcodes were corrected we don't end up with data for the originally incorrecct barcode as well as for the corrected one
' The most recently imported file replaces any other file of the same name as it moved to the relevant sub folder
'(ie Imported/NotApplicable/Failed) but ''' Review Version 3.0.9009 could possibly look at renaming with increasing number suffix

' The processing of TfrAllPreLiveDataToLiveData() has not been changed. When this procedure encounters
' duplicates information for a Franchise/Date/Barcode combination the existing data in the livedata table
' is preserved and it discards the latest record as a duplicate and places it in a Duplicates table

'--------------------------------------------------------------------------------
' EOD BatScan Sales Summary File(txt): Includes both Retail & Wholesale sales
' File Layout with constant declarations for field positions (Zero based field positions)
'--------------------
'Filename
'--------------------
' GMPS###_YYYYMMDD.txt[~00000000000000] eg GMPS420_20070115.txt~20070115165117
' BatScan Daily Summary file is named GMPS###_yyyymmdd.txt, where ### is the
' Batscan ID of the outlet, and yyyymmdd is the date of the sale.
'--------------------
'Header Record Layout                   [Eg *,703,20/07/1998,Day]
'--------------------
'onst kHdrID As Long = 0                ' Header-Id    Value will always be “*”
Const kHdrStoreID As Long = 1           ' Store ID     BATScan store ID. Three-digit left-padded with zeros.
Const kHdrDate As Long = 2              ' Date Code    Date of sale in “DD/MM/YYYY” format.  Must match the file name date string.
'onst kHdrSalesPeriod As Long = 3       ' Sales Period "Day"
'--------------------
'Detail Record Layout                   [Eg 9310029213086,2,108,202,101, ,703,98,104,]
'--------------------                   | Equivalent TSG DW fields in PreLiveData & LiveData tables
Const kDtlBarcode As Long = 0           ' Barcode
Const kDtlQty As Long = 1               ' Quantity
Const kDtlTotalCostIncGst As Long = 2   ' CostInc (cents)
Const kDtlTotalSellIncGst As Long = 3   ' TotalInc (cents)
Const kDtlNormalUnitSellIncGst As Long = 4  ' NormalSellInc (cents)
'onst kDtlOrderCode As Long = 5         ' N/A
Const kDtlStoreID As Long = 6           ' Translated to FranchiseIDTSG
'onst kDtlTotalCostExGst As Long = 7    ' N/A (cents)
'onst kDtlTotalSellExGst As Long = 8    ' N/A (cents)
'* Not available in file *              ' WholesaleQty, WholesaleActualSell
'--------------------
'Trailer Record Layout                  [Eg $,703,20/07/98,56]
'  NOT PRESENT IN SAMPLE FILE PROVIDED
'--------------------
Const kTrlID As Long = 0                ' Trailed-Id   Value will always be "$"
'onst kTrlStoreID As Long = 1           ' Store ID     BATScan store id. Three-digit left-padded with zeros.
'onst kTrlDate As Date = 2              ' Date Code    Date of sale in "DD/MM/YYYY" format.  Must match the file name date string.
'onst kTrlRecCount As Long = 3          ' Record Count Number of detail rows/records
'--------------------------------------------------------------------------------
Const kProcName As String = "ImportBatscanFile"
Dim bIsDate As Boolean
Dim bRollback As Boolean
Dim bTrailerFound As Boolean
Dim lngLoopCount As Long
Dim lngImportedCount As Long
Dim lngBatScanID As Long    'FranchiseIDBATA
Dim lngTsgFranID As Long    'FranchiseIDTSG
Dim dtmSalesDate As Date
Dim strSQL As String
Dim strErrMsg As String
Dim strFileName As String
Dim strLine As String
Dim strFranName As String
Dim strSubFolderFailed As String
Dim strSubFolderImported As String
Dim strSubFolderNotApplicable As String
Dim astrHeader() As String
Dim astrDetail() As String
Dim fso As Scripting.FileSystemObject
Dim ts As Scripting.TextStream
Dim rstDest As ADODB.Recordset
Dim rstFran As ADODB.Recordset
    
    strSubFolderFailed = g.strBatscanFolder & "\" & "Failed"
    strSubFolderImported = g.strBatscanFolder & "\" & "Imported"
    strSubFolderNotApplicable = g.strBatscanFolder & "\" & "NotApplicable"

    Set fso = New Scripting.FileSystemObject
    
    If Not g.bMaster Then
        strErrMsg = "BatScan import available only on MASTER"
    ElseIf Not fso.FileExists(pFullName) Then
        strErrMsg = "File not found: " & pFullName
    Else
    '   Ensure that appropriate folder structure exists
        If Not fso.FolderExists(g.strBatscanFolder) Then fso.CreateFolder g.strBatscanFolder
        If Not fso.FolderExists(strSubFolderFailed) Then fso.CreateFolder strSubFolderFailed
        If Not fso.FolderExists(strSubFolderImported) Then fso.CreateFolder strSubFolderImported
        If Not fso.FolderExists(strSubFolderNotApplicable) Then fso.CreateFolder strSubFolderNotApplicable
        
    '   Batscan Filname Format: GMPS###_YYYYMMDD.txt[~00000000000000] eg GMPS420_20070115.txt~20070115165117
        strFileName = fso.GetFileName(Path:=pFullName)
        
    '   IsValidBatScanFilename returns success/failure and sets ByRef params: dtmSalesDate & strErrMsg
        If Not IsValidBatScanFilename(pFilename:=strFileName, pSalesDate:=dtmSalesDate, pErrMsg:=strErrMsg) Then
            MoveFileOverWrite pSource:=pFullName, pDest:=strSubFolderNotApplicable & "\" & strFileName
        Else
        '   Translate BatScanID to a TsgFranID
            lngBatScanID = CLng(Val(Mid$(String:=strFileName, start:=5, Length:=3)))
            strSQL = "SELECT FranchiseIDTSG, FranchiseBusinessName FROM qryFranchiseBata " & _
                     "WHERE FranchiseIDBATA = " & lngBatScanID
            Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            With rstFran
                If Not (.BOF And .EOF) Then
                    lngTsgFranID = .Fields!FranchiseIDTSG.Value
                    strFranName = .Fields!FranchiseBusinessName.Value
                End If
                .Close
            End With
            Set rstFran = Nothing
            
            If lngTsgFranID = 0 Then
            '   Don't move file. Once Bata ID is setup the procedure can be re-run
                strErrMsg = "Bata ID [" & lngBatScanID & "] not setup in Franchise Table"
            Else
                Set ts = fso.OpenTextFile(pFullName, ForReading)
                If Not ts.AtEndOfStream Then
                '   Read file header
                    strLine = ts.ReadLine
                    astrHeader = Split(strLine, ",")
                    If Val(astrHeader(kHdrStoreID)) <> lngBatScanID Then
                        strErrMsg = "Store ID in filename does not match file header"
    '               ElseIf astrHeader(kHdrStoreID) <> astrTrailer(kTrlStoreID) Then
    '                   strErrMsg = "Header and Trailer Store IDs do not match"
    '               ElseIf astrTrailer(kTrlRecCount) <> (rstSrc.RecordCount - 1) Then   '??? -1 for header OR -2 for header & trailer
    '                   strErrMsg = "Records do not match record count in trailer record"
    '               ElseIf astrHeader(kHdrDate) <> astrTrailer(kTrlDate) Then
    '                   strErrMsg = "Header and Trailer dates do not match"
                    ElseIf GetDateFrom_dd_mm_yyyy(astrHeader(kHdrDate), pIsDate:=bIsDate) <> dtmSalesDate Then
                        strErrMsg = "Date in filename does not match file header"
                    Else
                        strSQL = "SELECT * FROM PreLiveData " & vbNewLine & _
                                  "WHERE (FranchiseIDTSG = " & lngTsgFranID & ")" & _
                                   " AND (TransactionDate = " & MySqlDate(dtmSalesDate) & ")"
                    '   eEditableDynamic rst type so it can perform multiple Find mehtods
                        Set rstDest = GetRst(pCnn:=g.cnnDW, _
                                             pSource:=strSQL, _
                                             pSourceType:=adCmdText, _
                                             pRstType:=eEditableDynamic, _
                                             pErrMsg:=strErrMsg)
                    '   g.cnnDW.BeginTrans
                        Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx
                        
                        Do While Not (ts.AtEndOfStream Or bRollback)
                            strLine = ts.ReadLine
                            astrDetail = Split(strLine, ",")
                            lngLoopCount = lngLoopCount + 1
                        '   SIMPLE VALIDATION OF ROW BEFORE PROCESSING
                        '   Could pass to a IsValidBatScanArray() but for the moment I have in-line validation of each row/array
                        '   It would be something like "bRollback = IsValidBatScanArray()" later followed by "if bRollback ..."
                            If astrDetail(kTrlID) = "$" Then
                            '   Is a trailer record.
                            '   May later perform processing to validate number of rows imported agrees with number
                            '   recorded in trailer record but for now will abort processing and flag we have a
                            '   different file type rather than code for a circumstance that may not eventuate
                                strErrMsg = "Trailer record encountered. Programming required"
                                bRollback = True
                                bTrailerFound = True
                            ElseIf astrDetail(kDtlStoreID) <> lngBatScanID Then
                                strErrMsg = "Store ID in filename does not match. Line " & lngLoopCount + 1 ' Takeing header line into account
                                bRollback = True
                            Else
                                With rstDest
                                    If Not (.BOF And .EOF) Then .MoveFirst
                                    .Find "Barcode = " & SqlQ(astrDetail(kDtlBarcode))
                                    If .EOF Then
                                        .AddNew
                                        lngImportedCount = lngImportedCount + 1
                                    Else
                                        Dim strMsg As String
                                        strMsg = "Ediitng an existing PreLiveData record won't work with MySQL & I " & _
                                                 "believe it's because PreLiveData table doesn't have a unique index"
                                        MsgBox strMsg, vbCritical
                                    End If
                                    '   Allow VB to perform data coercion
                                        .Fields!FranchiseIDTSG.Value = lngTsgFranID
                                        .Fields!Barcode.Value = astrDetail(kDtlBarcode)
                                        .Fields!TransactionDate.Value = dtmSalesDate
                                        .Fields!Quantity.Value = astrDetail(kDtlQty)
                                        .Fields!TotalInc.Value = astrDetail(kDtlTotalSellIncGst) / 100           ' Cnv cents to currency
                                        .Fields!NormalSellInc.Value = astrDetail(kDtlNormalUnitSellIncGst) / 100 ' Cnv cents to currency
                                        .Fields!CostInc.Value = astrDetail(kDtlTotalCostIncGst) / 100            ' Cnv cents to currency
                                    '---------------------------------------------------------------------------------------'
                                    '   EOD BatScan Sales Summary Files (GMPS###_YYYYMMDD.txt[~00000000000000]) contain       '
                                    '   total sales figures. Separate wholesale figures are not included.                     '
                                    '   Because Null values for wholesale fields cause problems for reporting functions,      '
                                    '   these fields are populated with zero if they previously held Null values.             '
                                    '   The fields are not unconditionally populated with zeroes because the function         '
                                    '   may subsequently process updated files, ALSO although the program does not currently  '
                                    '   import data from BATSCAN wholesale sumamry files (04Sep2008) it may in future, and    '
                                    '   depending on the order of processing files we would not want to obliterate valid data.'
                                    '   Apperently BATSCAN franchises don't wholesale, but some NON BATSCAN franchises        '
                                    '   supply their data to TSG in simulated BATSCAN files                                   '
                                    '-----------------------------------------------------------------------------------------'
                                        If IsNull(.Fields!WholesaleQty.Value) Then                                            '
                                            .Fields!WholesaleQty.Value = 0                                                    '
                                        End If                                                                                '
                                        If IsNull(.Fields!WholesaleActualSell.Value) Then                                     '
                                            .Fields!WholesaleActualSell.Value = 0                                             '
                                        End If                                                                                '
                                    '-----------------------------------------------------------------------------------------'
                                    .Update
                                End With
                            End If
                        Loop
                        If bRollback Then
                        '   g.cnnDW.RollbackTrans
                            Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
                        Else
                        '   g.cnnDW.CommitTrans
                            Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx
                        End If
                        rstDest.Close
                        Set rstDest = Nothing
                        Erase astrDetail()
                    End If
                    Erase astrHeader()
                End If
                ts.Close    ' Close ts and reclaim memory (must be closed before file can be moved)
                Set ts = Nothing
                If Len(strErrMsg) = 0 Then
                    MoveFileOverWrite pSource:=pFullName, pDest:=strSubFolderImported & "\" & strFileName
                Else
                '   Only move file to Failed sub folder if Trailer NOT Found. If a Trailer record is found don't
                '   move file as it may be Ok for reprocessing once some additional programming is performed
                    If Not bTrailerFound Then
                        MoveFileOverWrite pSource:=pFullName, pDest:=strSubFolderFailed & "\" & strFileName
                    End If
                End If
            End If
        End If
    End If
    Set fso = Nothing
    
    If Len(strErrMsg) Then
        StatusBar UCase$("Batscan import FAILED [" & strFileName & "] " & strErrMsg), pFranchise:=strFranName
    Else
        StatusBar "Batscan imported " & Plural(lngImportedCount, "record") & _
                             " from " & Plural(lngLoopCount, "record") & " [" & strFileName & "]", pFranchise:=strFranName
        ImportBatscanFile = True
    End If

End Function

Private Sub ImportBatScanFiles(Optional ByVal pManualImport As Boolean = False)
' Called as part of Data Capture Cycle, but can be user initiated via cmdImportBatScanFiles_Click
' Refreshing Event Log is suspended during Manual Import (no need to refresh as Event Log is displayed on another tab)
Dim bPrevEventLogRefreshEnabled As Boolean
Dim lngFileCount As Long
Dim lngProcessedFileCount As Long
Dim strFileName As String

    If pManualImport Then
        bPrevEventLogRefreshEnabled = gbEventLogRefreshIsEnabled
        gbEventLogRefreshIsEnabled = False
    End If
    
'   Formats for Different Numeric Values
'   Three sections The first section applies to positive values, the second to negative values, and the third to zeros.
    StatusBar Format$(pManualImport, ";\M\a\n\u\a\l\ ;") & _
              "BatScan import commenced: Import files from " & g.strBatscanFolder & " folder."

    strFileName = Dir$(g.strBatscanFolder & "\", vbDirectory)
    Do While strFileName <> ""  ' Start loop.
        If (GetAttr(g.strBatscanFolder & "\" & strFileName) And vbDirectory) <> vbDirectory Then
        '   It's a file.
            lngFileCount = lngFileCount + 1
            If ImportBatscanFile(pFullName:=g.strBatscanFolder & "\" & strFileName) Then
                lngProcessedFileCount = lngProcessedFileCount + 1
            End If
        End If
       strFileName = Dir        ' Get next entry.
    Loop
       
    StatusBar Format$(pManualImport, ";\M\a\n\u\a\l\ ;") & "BatScan import completed: " & _
              Plural(lngFileCount, pNounSingular:="file") & " in folder, " & _
              Plural(lngProcessedFileCount, pNounSingular:="file") & " successfully processed."
    
    If pManualImport Then
        gbEventLogRefreshIsEnabled = bPrevEventLogRefreshEnabled
        If gbEventLogRefreshIsEnabled Then
            gsubRefreshEventLogDisplay
        End If
    End If
    
End Sub

Sub IncrementStkFileNum(ByVal pStkFilename As String)
Dim lngFileNum As Long
Dim lngExtPos As Long
Dim strErrMsg As String
Dim strFileNum As String
Dim strFilebaseName As String
Dim rst As ADODB.Recordset

    lngExtPos = InStrRev(pStkFilename, ".")
    If lngExtPos > 1 Then
        strFilebaseName = Left$(pStkFilename, lngExtPos - 1)
    End If
    
    strFileNum = Right$(strFilebaseName, 3)
    
    If IsNumeric(strFileNum) Then
        lngFileNum = CLng(strFileNum)
        Set rst = GetRst(pCnn:=g.cnnDW, pSource:="StkFileNums", _
                         pSourceType:=adCmdTable, _
                         pRstType:=eEditableFwdOnly, _
                         pErrMsg:=strErrMsg)
        With rst
            If Not (.BOF And .EOF) Then
                If lngFileNum >= .Fields!FileNum.Value Then ' Increment FileNum
                    .Fields!FileNum.Value = lngFileNum + 1
                    .Update
                End If
            End If
            .Close
            Set rst = Nothing
        End With
    End If

End Sub

Private Function IsPromoApplicableToFranchise(ByVal pFranID As Long, _
                                              ByRef pPromoRst As ADODB.Recordset) As Boolean
' Test whether Promotion is applicable to Franchise by checking FranchisePromotion table
' (Did not simply add a test for PromotionGrade to pre-existing tests because promotion grade       )
' (may have changed for a franchise since the promotion was sent and the time of compliance testing )
' (- particularly if compliance testing is extended beyond yesterday                                )
' (new compliance testing is both simpler, more efficient and more accurate                         )
Dim bResult As Boolean
Dim strSQL As String
Dim rstFranPromo As ADODB.Recordset
Dim strErrMsg As String


'   Should we also ensure that it has been uploaded? Danger is for Franchise to claim rebates and get
'   them when they haven't received the promotions => haven't complied but are not picked up as non-compliant
    strSQL = "SELECT * FROM tblFranchisePromotions" & vbNewLine & _
             "WHERE (FranchiseID = " & pFranID & ") AND (PromotionID = " & pPromoRst!PromoID & ")"
    Set rstFranPromo = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    bResult = Not (rstFranPromo.BOF And rstFranPromo.EOF)
    rstFranPromo.Close
    Set rstFranPromo = Nothing
    
    IsPromoApplicableToFranchise = bResult
    
End Function

Private Function IsStkCtlsValid(ByRef pErrMsg As String) As Boolean
'   Description field can be identical for different barcodes. THIS IS NOT AN OVERSIGHT
'   Depending on how RMgr is setup at the franchise (setting something like 'accept leading zeros'?)
'   it may or may not require and accept barcodes with leading zeros. The same stock item is therefore
'   duplicated with and without leading zeros. (there are a few examples - Captain Blacks, Camel ...)
Dim bResult As Boolean
Dim bContinue As Boolean
Dim strSQL As String
Dim strErrMsg As String
Dim rstBarcode As ADODB.Recordset

    bContinue = True
    If Len(txtBarcode) < 1 Then
        strErrMsg = "Barcode is mandatory"
        txtBarcode.SetFocus
        bContinue = False
    Else 'check whether barcode already exists
        strSQL = "SELECT Barcode, Deleted FROM Stock WHERE Barcode = " & SqlQ(txtBarcode)
        Set rstBarcode = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        If Not (rstBarcode.BOF And rstBarcode.EOF) Then 'the barcode already exists
            strErrMsg = "Barcode is already in the database, re-enter"
            If CBool(rstBarcode!Deleted) Then
                strErrMsg = strErrMsg & vbNewLine & "(Barcode exists in database as deleted stock)"
            End If
            txtBarcode.SetFocus
            bContinue = False
        End If
        rstBarcode.Close
        Set rstBarcode = Nothing
    End If
    
    If bContinue Then
        Select Case True
            Case Trim$(txtStkItemDescription) = vbNullString
                strErrMsg = "Item Description is mandatory"
                txtStkItemDescription.SetFocus
            Case cboSupplier = vbNullString
                strErrMsg = "Supplier is mandatory"
                cboSupplier.SetFocus
            Case Val(Trim(txtSticks)) = 0
                strErrMsg = "Sticks quantity is mandatory"
                txtSticks.SetFocus
            Case cboCategory = vbNullString
                strErrMsg = "Category is mandatory"
                cboCategory.SetFocus
            Case cboSubCategory = vbNullString
                strErrMsg = "Sub-category is mandatory"
                cboSubCategory.SetFocus
            Case cboGoodsTax = vbNullString
                strErrMsg = "Goods tax is mandatory"
                cboGoodsTax.SetFocus
            Case cboSalesTax = vbNullString
                strErrMsg = "SalesTax is mandatory"
                cboSalesTax.SetFocus
            Case Else
                bResult = True
        End Select
    End If

    pErrMsg = strErrMsg
    IsStkCtlsValid = bResult

End Function

Private Function IsValidBatScanFilename(ByVal pFilename As String, _
                                        ByRef pSalesDate As Date, _
                                        ByRef pErrMsg As String) As Boolean
'   Returns success/failure
'   Sets ByRef params of dtmSalesDate & pErrMsg
'   Filename format: GMPS###_YYYYMMDD.txt
'   Don't bother validating date is in a valid range as it will be compared against file contents
'   Don't bother validating any characters trailing ".txt".
'   Files may come in from franchises with a long series of traling chars
Dim bResult As Boolean
Dim bValidDate As Boolean
Dim strTemp As String
Dim strErr As String

'   Validate filename format - Could possibly use a regular expression to validate the filename format
    strErr = "Incorrect filename format"
    
'   Filename format = GMPS###_YYYYMMDD.txt
    If UCase$(Left$(pFilename, 4)) = "GMPS" Then
        strTemp = Mid$(String:=pFilename, start:=5, Length:=3)
        If IsEachCharADigit(pString:=strTemp) Then
            strTemp = Mid$(String:=pFilename, start:=8, Length:=1)
            If strTemp = "_" Then
                strTemp = Mid$(String:=pFilename, start:=9, Length:=8)
                pSalesDate = GetDateFrom_yyyymmdd(pYyyyMmDd:=strTemp, pIsDate:=bValidDate)
                If Not bValidDate Then
                    pErrMsg = strErr & ": incorrect date component"
                Else
                    strTemp = Mid$(String:=pFilename, start:=17, Length:=4)
                    If strTemp = ".txt" Then
                        bResult = True
                        strErr = vbNullString
                    End If
                End If
'                Would proabably be reasonable to validate date as within one year of machine date
'                Probably not sent to BATA if it was any older
'                REMEMBER THAT THIS DATA WILL ALSO BE SENT TO AZTEC
'                Note that date in file will be validate against date in file.
            End If
        End If
    End If
    
    pErrMsg = strErr
    IsValidBatScanFilename = bResult
    
End Function

Private Function IsValidData(ByRef pColFldVals As VBA.Collection, _
                                   ByVal pMaxValue As Currency, _
                                   ByVal pMaxQty As Long) As Boolean

' Absolute values are validated against max qty and max amt because negative qtys may be used to cancel out previous sales
Dim bValid As Boolean
''' Review Return to this one day and work through each field one by one and address the root cause as to why
'''        we get Null values in fields and whether we need to transfer records with barcode = "TOTALCUSTOMERS", etc ....
    
    Select Case True
        Case (Cn(pColFldVals("Quantity"), 0) = 0) And (Cn(pColFldVals("WholesaleQty"), 0) = 0)
        '   Qty is a total quantity and WS Qty is the wholesale component
        '   Record is rejected to prevent divide by zero error when calculating WS Cost (WSQty / Quantity * Cost)
            bValid = False
        Case Abs(Cn(pColFldVals("Quantity"), 0)) > pMaxQty
            bValid = False
        Case Abs(Cn(pColFldVals("WholesaleQty"), 0)) > pMaxQty
            bValid = False
        Case Abs(Cn(pColFldVals("TotalInc"), 0)) > pMaxValue
            bValid = False
        Case Abs(Cn(pColFldVals("NormalSellInc"), 0)) > pMaxValue
            bValid = False
        Case Abs(Cn(pColFldVals("CostInc"), 0)) > pMaxValue
            bValid = False
        Case Abs(Cn(pColFldVals("WholesaleActualSell"), 0)) > pMaxValue
            bValid = False
        Case InStr(pColFldVals("Barcode"), "'") <> 0
        'Case InStr(pFldArray!Barcode, gkSqlCharStringDelimiter) <> 0
        '   Reject data where barcode has an embedded single quote
        '   Such data can cause problems when constructing SQL
            bValid = False
        Case Else
            bValid = True
    End Select

    IsValidData = bValid

End Function

Private Sub lblDCTabPromoGrade_Click()
    cboDCTabPromoGrade.Locked = False
End Sub

Private Sub lblFranchiseType_DblClick()
    cboFranchiseType.Locked = False
End Sub

Private Sub lblRegion_DblClick(Index As Integer)
    cboDCTabRegion.Locked = False
End Sub

Private Sub lblState_DblClick()
    cboState.Locked = False
End Sub

Sub ListBoxClearSelections(ByRef pListBox As VB.ListBox)
Dim lngIndex As Long

    With pListBox
        For lngIndex = 0 To .ListCount - 1
            .Selected(lngIndex) = False
        Next
    End With

End Sub

Private Sub LoadCboCtnContainingPkt(ByVal pRecordSource As String)
Dim strSQL As String

    strSQL = "SELECT Stock_ID as StkID, Barcode + ': ' + Description as StkDescription" & vbNewLine & _
             "FROM " & pRecordSource & vbNewLine & _
             "WHERE " & gconStockTableCategoryField & " = " & SqlQ(gkCAT_CigCtn) & vbNewLine & _
             "ORDER BY Description"
    
    LoadCombo_Rst pCombo:=cboCtnContainingPkt, _
                          pCnn:=g.cnnDW, _
                          pSource:=strSQL, _
                          pDisplayFld:="StkDescription", _
                          pDataFld:="StkID"

End Sub

Sub LoadGrdPromoTabRebates(ByVal pUsePromoGrades As Boolean)
    With grdPromoTabRebates
        .Rows = .FixedRows ' Clear grid
        .LoadArray GetBlankPromoRebatesArray(pUsePromoGrades)
        If (.Rows > .FixedRows) Then ' HasDataRows
            .Cell(flexcpBackColor, .FixedRows, 1, .Rows - 1) = .BackColorFixed
        End If
    End With
End Sub

'----------------------------------------------------------------------------------------
'  loadNonCompliant
'
'  Background
'  The PromoNonCompliantSales table holds non-compliant sales for one day.
'  It is populated by this function at the end of the capture cycle.
''' Review  '  This function can also be executed at the user discretion from the Promo tab. WHY???
'
'  Description
'  For each product in each active (ie today) promotion, scan through livedata and check
'  whether each sale for yesterday is compliant. If not add it to promoNonCompliantSales table
'  Note that compliance is only checked for yesterday. In the future we may decide to check for
'  compliance on past days.
'  If so, some of the processing below will need to be modified.
'
''' V369 NOTE: procedure limits promotions to currently active promotions
''' V369 NOTE: (including active today) but only looks at yesterdays sales. Is possible
''' V369 NOTE: a new promo is active today but doesn't apply to yesterday's sales
'
''' V369 NOTE: If ever going to use Non-compliant promos other than yesterdays, then
''' V369 NOTE: the procedure should probably load those periods as well since data
''' V369 NOTE: collected since the day they were populated could change results.
'
''' V369 NOTE ALL IN ALL IT IS PROBABLY ONLY WORTH REPORTING ON YESTERDAY'S NON COMPLIANTS
''' V369 AND SIMPLIFIYING THE CODE AS SUCH. IT IS ONLY THEN YOU CAN TAKE USEFUL ACTIONS FOR FRANCHISES
'
'  Returns
'   Number of non-compliant sales via function return value
'   Number of non-compliant franchises via function parameter (pFranCount)
'----------------------------------------------------------------------------------------
Function LoadNonCompliantTable(ByRef pFranCount As Long) As Long
Const kProcName As String = "LoadNonCompliantTable"
Dim rstPromo As ADODB.Recordset
Dim rstProducts As ADODB.Recordset
Dim rstSales As ADODB.Recordset
Dim rstNonCompliant As ADODB.Recordset
    
    Dim sRebateFld As String
    Dim cNonCompliants As Long
    Dim sStatus As String
    Dim iTicker As Long
    Dim retailQty As Long
    Dim whsQty As Long

Dim lngFranCount As Long
Dim strSQL As String
Dim strErrMsg As String
Dim dtmYesterday As Date
Dim strWC_TxDateYesteerday As String
Dim curMaxSell As Currency
Dim curActualSell As Currency
Dim curPromoSell As Currency
Dim curTolerance As Currency

    sStatus = "Promo "
    curTolerance = g.rstDWDefaults!PromotionPriceTolearance
    dtmYesterday = fdtmYesterday()

'   Clear non-compliant sales for yesterday (before re-populating) (Perhaps ultimately do away with this step?)
    strWC_TxDateYesteerday = "(TransactionDate = " & MySqlDate(dtmYesterday) & ")"
    strSQL = "DELETE FROM PromoNonCompliantSales WHERE " & strWC_TxDateYesteerday
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx
        CnnDwExecute pCommandText:=strSQL
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx

    strSQL = "SELECT * FROM Promotions " & vbNewLine & _
             "WHERE (PromoStart <= " & MySqlDate(dtmYesterday) & ")" & _
              " AND (PromoEnd >= " & MySqlDate(dtmYesterday) & ")"
    Set rstPromo = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    
    If Not (rstPromo.BOF And rstPromo.EOF) Then
        Set rstNonCompliant = GetRst(pCnn:=g.cnnDW, _
                                     pSource:="PromoNonCompliantSales", _
                                     pSourceType:=adCmdTable, _
                                     pRstType:=eEditableFwdOnly, _
                                     pErrMsg:=strErrMsg)
        Do Until rstPromo.EOF
            StatusBar sStatus & rstPromo!PromoName, pLog:=False
            strSQL = "SELECT * FROM qryStock " & vbNewLine & _
                     "WHERE " & gconStockTableSubCategoryField & " = " & SqlQ(rstPromo!PromoSubCat)
            Set rstProducts = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            
            Do Until rstProducts.EOF
                StatusBar sStatus & rstPromo!PromoName & " " & rstProducts!Description, pLog:=False
                strSQL = "SELECT * FROM livedata " & _
                         " WHERE " & strWC_TxDateYesteerday & _
                           " AND (Quantity > 0) " & _
                           " AND (Barcode = " & SqlQ(rstProducts!Barcode) & ")"
                Set rstSales = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
                If Not (rstSales.BOF And rstSales.EOF) Then
                    iTicker = 0
                    Do Until rstSales.EOF
                        StatusBar sStatus & rstPromo!PromoName & " " & rstProducts!Description & " " & iTicker, pFranchise:=vbNullString, pLog:=False
                        iTicker = iTicker + 1
                        ' check whether this sale had a retail component. If it didn't, ie. if the entire sale was
                        ' wholesale, then don't even bother checking any further because wholesale sales don't need
                        ' to be checked for promotional compliance.

                        whsQty = Cn(pValue:=rstSales!WholesaleQty, pReplaceWith:=0)
    
                        retailQty = rstSales!Quantity - whsQty
                        If retailQty > 0 Then
                            ' check if this store is in the same State as this promotion
                            If IsPromoApplicableToFranchise(pFranID:=rstSales!FranchiseIDTSG, pPromoRst:=rstPromo) Then
                                Select Case rstProducts(gconStockTableCategoryField)
                                    Case gkCAT_CigCtn
                                        sRebateFld = "PromoCartonDiscount"
                                    Case gkCAT_CigPkt, gkCAT_Tobac, gkCAT_Cigar
                                        sRebateFld = "PromoPacketDiscount"
                                End Select
                                
                            '   Allow some tolerance between calculated PromoSell and ActualSell
                                curPromoSell = TRound(rstSales!NormalSellInc - rstPromo(sRebateFld), pDecimalPlaces:=2)
                                curMaxSell = curPromoSell + curTolerance
                                curActualSell = TRound((rstSales!TotalInc - rstSales!WholesaleActualSell) / retailQty, pDecimalPlaces:=2)
                                If curActualSell > curMaxSell Then
                                '   This sale was non-compliant
                            ''' V369 Start
'''                         '~      ' This sale was non-compliant... check if we already know this.
'''                         '~      ' Note that this bit is only relevant if we decide in the future to implement
'''                         '~      ' non-compliant sales checking for days other than yesterday, because in the current situation, all
'''                         '~      ' non-compliant sales are deleted when we first enter this function above, so there will never be
'''                         '~      ' any records in the PromoNonCompliantSales table.
                            '~  Following commented out code would be relevant if PromoNonCompliantSales
                            '~  table was not cleared of records for Yesterday's date at top of the procedure
                            '~      strSQL = "SELECT * FROM PromoNonCompliantSales " & _
                            '~               "WHERE TransactionDate = " & MsSqlDate(rstSales!TransactionDate) & _
                            '~               " AND FranchiseIDTSG = " & rstSales!FranchiseIDTSG & _
                            '~               " AND Barcode = " & SqlQuote(rstProducts!Barcode) & _
                            '~               " AND PromoID = " & rstPromo!PromoID
                            '~  ''' Set rstNonCompliant = g.dbDW.OpenRecordset(strSQL, dbOpenDynaset)   ''' dao2AD0
                            '~      Set rstNonCompliant = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pRstType:=eEditableFwdOnly, pErrMsg:=strErrMsg)     ''' dao2AD0
                            '~      If (rstNonCompliant.BOF And rstNonCompliant.EOF) Then
                            ''' V369 End
                                        rstNonCompliant.AddNew
                                            rstNonCompliant!TransactionDate = rstSales!TransactionDate
                                            rstNonCompliant!FranchiseIDTSG = rstSales!FranchiseIDTSG
                                            rstNonCompliant!Barcode = rstSales!Barcode
                                            rstNonCompliant!PromoID = rstPromo!PromoID
                                            rstNonCompliant!Quantity = retailQty
                                            rstNonCompliant!TotalInc = rstSales!TotalInc - rstSales!WholesaleActualSell
                                            rstNonCompliant!ActualSellInc = curActualSell
                                            rstNonCompliant!NormalSellInc = TRound(rstSales!NormalSellInc, pDecimalPlaces:=2)
                                            rstNonCompliant!PromoSellInc = curPromoSell
                                            rstNonCompliant!MaxSellInc = curMaxSell
                                            rstNonCompliant!StoreNotified = CBoolMySql(False)
                                        rstNonCompliant.Update
                                    ''' cNonCompliants = cNonCompliants + 1 ''' V369 Move out of PROPOSED If (.BOF AND .EOF) Then
                            ''' V369 Start
                            '~      End If  ''' V369
                            '~    ''' End If
                            '~        rstNonCompliant.Close
                            '~        Set rstNonCompliant = Nothing
                            ''' V369 End
                                    cNonCompliants = cNonCompliants + 1     ''' V369 Move out of If (.BOF AND .EOF) Then
                                End If
                            End If
                        End If
                        rstSales.MoveNext
                    Loop
                End If
                rstSales.Close
                Set rstSales = Nothing
                rstProducts.MoveNext
                DoEvents
            Loop
            rstProducts.Close
            Set rstProducts = Nothing
            rstPromo.MoveNext
        Loop
        
SetTableUpdateTime pTableName:="PromoNonCompliantSales", pTimeStamp:=Now    ''' Review
        
        If Not rstNonCompliant Is Nothing Then
            rstNonCompliant.Close
            Set rstNonCompliant = Nothing
        End If
        
        strSQL = "SELECT COUNT(DISTINCT FranchiseIDTSG) " & vbNewLine & _
                 "FROM PromoNonCompliantSales " & vbNewLine & _
                 "WHERE " & strWC_TxDateYesteerday
        lngFranCount = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL)

    End If
    rstPromo.Close
    Set rstPromo = Nothing
    
    pFranCount = lngFranCount
    LoadNonCompliantTable = cNonCompliants
    
End Function

Private Sub LoadPromoListview(Optional ByVal pShowALL As Boolean = False, _
                              Optional ByVal pForce As Boolean = False)
Dim dtmToday As Date
Dim strSQL As String
Dim strErrMsg As String
Dim strPromoStatus As String
Dim itm As ComctlLib.ListItem
Dim colRegions As VBA.Collection
Dim colPromoGrades As VBA.Collection
Dim rstPromo As ADODB.Recordset
    
'   Time consuming -> only update display as required OR forced to
    If pForce Or (GetTableUpdateTime("Promotions") > g.udtDtmCtlUpdated.dtmPromoListView) Then
        lvwPromo.ListItems.Clear
    
        dtmToday = Date
        strSQL = "SELECT * FROM Promotions"
        If Not pShowALL Then
        '   Select current promotions
            strSQL = strSQL & " WHERE PromoEnd >= " & MySqlDate(dtmToday)
        End If
        Set rstPromo = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
        If Not (rstPromo.BOF And rstPromo.EOF) Then
            Set colRegions = GetRegionsCollection()
            Set colPromoGrades = GetPromoGradesCollection()
            Do While Not rstPromo.EOF
                Set itm = frmTSGDataWarehouse.lvwPromo.ListItems.Add(Text:=rstPromo!PromoName)
                itm.SubItems(1) = rstPromo!PromoSubCat
                itm.SubItems(2) = rstPromo!PromoStart
                itm.SubItems(3) = rstPromo!PromoEnd
                itm.SubItems(4) = Format$(rstPromo!PromoPacketDiscount, "#0.00")
                itm.SubItems(5) = Format$(rstPromo!PromoCartonDiscount, "#0.00")
                itm.SubItems(6) = rstPromo!PromoState
                itm.SubItems(7) = colRegions.Item(CStr(rstPromo!PromoRegionID.Value))
                itm.SubItems(8) = colPromoGrades.Item(CStr(rstPromo!PromoGradeID.Value))
                itm.SubItems(9) = rstPromo!PromoID
                If rstPromo!PromoEnd < dtmToday Then
                    strPromoStatus = "expired"
                Else
                    strPromoStatus = rstPromo!PromoStatus
                End If
                itm.SubItems(10) = strPromoStatus
                rstPromo.MoveNext
            Loop
            Set colRegions = Nothing
            Set colPromoGrades = Nothing
        End If
        rstPromo.Close
        Set rstPromo = Nothing
        
        lvwPromo.Refresh
    
        g.udtDtmCtlUpdated.dtmPromoListView = Now
    
    End If
    
End Sub

Sub LoadPromotionTab()
Dim bUsePromoGrades As Boolean
Dim dtmToday As Date
Dim strSQL As String

'   Master instance logs when Promotions table is updated
'   Called procedures use this log to refresh screen components on an as needed basis
    LoadPromoListview pShowALL:=False
            
'' Ideally (or not) a slave could either poll for changes when on the tab or provide a refresh
'' button on the tab for when the master has made a change and it mightn't be reflected on the slave
    
    PopulateSubCategory
    PopulateNonCompliantLView   '   Now (V376) only called from here
    strSQL = "SELECT * FROM tlkpRegion WHERE RegionID <> 0" ' Exclude ALL-STATES from available selections
    LoadListBox_Rst pListBox:=lstPromoTabRegion, pCnn:=g.cnnDW, pSource:=strSQL, pDisplayFld:="RegionName", pDataFld:="RegionID"
    LoadListBox_Rst pListBox:=lstPromoTabState, pCnn:=g.cnnDW, pSource:="qlkpState", pDisplayFld:="State"
    
    bUsePromoGrades = Not ChkBoxToBool(chkPromoSelectFranchise)
    LoadGrdPromoTabRebates pUsePromoGrades:=bUsePromoGrades
    
'   Set valid date ranges then set date values for date controls
    dtmToday = Date
    
    dtpPromoStart.MaxDate = dtmToday
    dtpPromoStart.Value = dtmToday
    
    dtpPromoEnd.MinDate = dtmToday
    dtpPromoEnd.Value = dtmToday
    
    Me.Refresh
    
End Sub

Sub LoadSettingsTab()
Dim strSQL As String
Dim strErrMsg As String
Dim rsStateSpecific As ADODB.Recordset
    Dim iIndex As Integer
    
    disableSettingsFields
    
    For iIndex = 0 To 7
        strSQL = "SELECT * FROM tlkpState " & vbNewLine & _
                 "WHERE StateOfOz = " & SqlQ(lblSateName(iIndex))
        Set rsStateSpecific = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        If Not (rsStateSpecific.BOF And rsStateSpecific.EOF) Then
            txtDialSequence(iIndex) = rsStateSpecific!DialSequence
        End If
    Next iIndex
    rsStateSpecific.Close
    Set rsStateSpecific = Nothing
    
    txtTxnStartDate = Format$(g.dtmLiveDataStart, gkFmtDateUnambiguous)
    
    txtThisNodeName.Text = g.strNodeName
    
End Sub

Sub LoadUploadTab()
    Dim sTemporaryFileList As String
    ' Franchise List box has already been done in subPopulateFranchiseBusinessNameListBoxes
    ' choise First fill in 'uploaded' listbox which basically just has everything
    ' in the 'uploads' directory
    lstUploadItemList.Clear
    sTemporaryFileList = Dir(g.strUploadsFolder & "\*.*", vbDirectory)
    If Len(sTemporaryFileList) > gconZeroValue Then
        Do Until sTemporaryFileList = ""
            If sTemporaryFileList <> "." And sTemporaryFileList <> ".." Then
                lstUploadItemList.AddItem g.strUploadsFolder & "\" & sTemporaryFileList
            End If
            sTemporaryFileList = Dir
        Loop
    End If
    
    ' Next clear all flags and check boxes.
    chkResetRemoteOpenedBy.Value = 0
     
    DisplayUploadsPending
    
    ' Next set focus on the New Message Title
    txtNewMessage.Text = vbNullString
    txtMessageTitle.Text = vbNullString
    txtMessageTitle.SetFocus

End Sub

Sub LockDCTabFranchiseCtls(ByVal pLocked As Boolean)

'   Combos
    cboFranchiseType.Locked = pLocked
    cboDCTabPromoGrade.Locked = pLocked
    cboDCTabRegion.Locked = pLocked
    cboState.Locked = pLocked

'   TextBoxes
    txtPhysicalAddress.Locked = pLocked
    txtContact.Locked = pLocked
    txtSuburb.Locked = pLocked
    txtAreaCode.Locked = pLocked
    txtModem.Locked = pLocked
    txtNodename.Locked = pLocked
    txtRASUsername.Locked = pLocked
    txtBATAFranchiseID.Locked = pLocked
    txtPhone.Locked = pLocked
    txtFaxNum.Locked = pLocked
    txtRASPassword.Locked = pLocked

End Sub

Private Sub lstDataCaptureFranchiseBusinessName_Click()
    If Not gbClickEventIsSuppressed Then
        cmdCaptureSelected.Enabled = g.bMaster And (lstDataCaptureFranchiseBusinessName.SelCount > 0)
        cmdCloseSelectedFranchises.Enabled = g.bMaster And (lstDataCaptureFranchiseBusinessName.SelCount > 0)
        subDisplayFranchiseDetails
        HighlightSelectedFranchiseInEventLog
    End If
End Sub

Private Sub lstDescription_Click()
    If Not gbClickEventIsSuppressed Then
        subDisplayStockItem
    End If
End Sub

Private Sub lstNielsenReportDisplayDate_Click()
Dim strNielsnRptFullFilename As String

    strNielsnRptFullFilename = g.strNielsenRptsFolder & "\" & lstNielsenReportDisplayDate
    
    If lstNielsenReportDisplayDate <> "" Then
        With rtxNielsenReportContents
            .Text = ""
            .Refresh
        End With
        
        On Error GoTo Procedure_Exit
        If LCase(Right(lstNielsenReportDisplayDate, 4)) = gconTextFileSuffix Then
            rtxNielsenReportContents.Text = fsFileContentsAndEOFMessage(strNielsnRptFullFilename)
        Else
            rtxNielsenReportContents.Text = "Zipfile created " & FileDateTime(strNielsnRptFullFilename) 'PAL
            ' rtxNielsenReportContents.Text = "Unable to preview files of this type" 'PAL
        End If
    End If

Procedure_Exit:
    Exit Sub

End Sub

Private Sub lstNielsenReportDisplayDate_DblClick()
    If lstNielsenReportDisplayDate <> "" Then
        subOpenFile g.strNielsenRptsFolder & "\" & lstNielsenReportDisplayDate
    End If
End Sub

Private Sub lstPromoSubCat_Click()
    Dim lArrayRowIndex As Integer
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpCategory As ADODB.Recordset

    lstPromoProducts.Clear
    For lArrayRowIndex = gconDisplayFirstItem To lstPromoSubCat.ListCount - 1
        'get product description for the sub-category in highlighted in box
        If (lstPromoSubCat.Selected(lArrayRowIndex)) Then
            strSQL = "SELECT description FROM qryStock" & vbNewLine & _
                     "WHERE " & gconStockTableSubCategoryField & " = " & SqlQ(lstPromoSubCat.List(lArrayRowIndex))
            Set rstSnpCategory = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            If Not (rstSnpCategory.BOF And rstSnpCategory.EOF) Then
                lstPromoProducts.AddItem stripType(rstSnpCategory(gconStockTableDescriptionField))
            End If
            rstSnpCategory.Close
            Set rstSnpCategory = Nothing
        End If
    Next lArrayRowIndex
End Sub

Private Sub lstStcokTabSelectedSoctkExport_Click()
    cmdStockTabExport.Enabled = (lstStcokTabSelectedSoctkExport.SelCount > 0)
    cmdStockTabDelete.Enabled = (lstStcokTabSelectedSoctkExport.SelCount > 0)
End Sub

Private Sub lstStickReportRecipient_Click()

    If Not gbClickEventIsSuppressed Then
        If lstStickReportRecipient <> "" Then
            cmdStickReport.Enabled = True
        End If
    End If

End Sub

Private Sub lstUploadItemList_Click()
'   Ensure that if TSStknnn.txt is selected that TSStknnnPkg.txt is also selected (& vice-versa)
Dim intPrevMousePointer As Integer
Dim lngInnerLoop As Long
Dim lngOuterLoop As Long
Dim strLoopFilename As String
Dim strSelectedFilename As String

    Me.Enabled = False
    intPrevMousePointer = SetMousePointer(vbHourglass)
    
    With lstUploadItemList
        If .SelCount Then
            For lngOuterLoop = 0 To .ListCount - 1
                If .Selected(lngOuterLoop) Then
                    strSelectedFilename = UCase$(fGetLastWord(.List(lngOuterLoop), "\"))
                    If UCase$(Left$(strSelectedFilename, 5)) = UCase$(gconNewStockFilePrefix) Then
                        For lngInnerLoop = 0 To .ListCount - 1
                            strLoopFilename = fGetLastWord(.List(lngInnerLoop), "\")
                            If UCase$(GetStkPkgFullFilename(strSelectedFilename)) = UCase$(strLoopFilename) Or _
                               UCase$(GetStkPkgFullFilename(strLoopFilename)) = UCase$(strSelectedFilename) Then
                                    .Selected(lngInnerLoop) = True
                            End If
                        Next lngInnerLoop
                    End If
                End If
            Next lngOuterLoop
        End If
    End With

    With stb
        .SimpleText = "Double left click to open the selected file"
        .Refresh
    End With

    SetMousePointer intPrevMousePointer
    Me.Enabled = True

End Sub

Private Sub lstUploadItemList_DblClick()
    subOpenFile lstUploadItemList
End Sub

Private Sub lvwEventLog_DblClick()

    Dim sMessage As String
    
    If gbEventLogRefreshIsEnabled Then
        gbEventLogRefreshIsEnabled = False
        sMessage = "disabled"
    Else
        gbEventLogRefreshIsEnabled = True
        sMessage = "enabled"
        Call gsubRefreshEventLogDisplay
    End If
    
    MsgBox "Event log automatic refresh is " & sMessage, vbInformation

End Sub

Private Sub lvwEventLog_ItemClick(ByVal Item As ComctlLib.ListItem)
    lvwEventLog.GetFirstVisible
    lvwEventLog.SelectedItem = Nothing  ' Unselect Item so it doesn't mismath Franchise List Box
    Set Item = Nothing
End Sub

Private Sub lvwNonCompliant_Click()
Dim bIsDataSelected As Boolean

'   Synchronise all rows for selected franchise to have the same selection status
    LvwSelChangeAssociatedRows pListView:=lvwNonCompliant, pAssociateBySubItemWithIdx:=1
    
'   Enable/Disable appropriate controls
    bIsDataSelected = ListView.LvwIsDataSelected(pListView:=lvwNonCompliant)
    cmdPromotion(9).Enabled = bIsDataSelected                           ' Print Selected button
    cmdPromoTabSaveNonCompliantSelected.Enabled = bIsDataSelected

End Sub

Private Sub lvwPromo_Click()
'Z Don't implement but test the effect
''' Review This is where some code to prevent selecting recalled promotions could go
'''        May depend on whether code is given to select allow viewing which Frans
'''         an individually seleted Fran promo has been sent to
    cmdPromotionRecall.Enabled = LvwIsDataSelected(lvwPromo)
End Sub

Private Sub lvwUploadsPending_dblclick()
Dim strSQL As String
Dim strErrMsg As String
Dim rstUpload As ADODB.Recordset

    If MsgBox("Delete the following pending uploads:" & vbCrLf & vbCrLf & _
        lvwUploadsPending.SelectedItem.Text & "  " & lvwUploadsPending.SelectedItem.SubItems(1) & _
        " ?", vbYesNo) = vbYes Then
        
        strSQL = "SELECT * FROM FranchiseUploads" & _
                 " WHERE " & gconUploadFranchiseIDField & " = " & fsFranchiseIDFrom(lvwUploadsPending.SelectedItem.Text) & _
                 " AND " & gconUploadFileField & " = " & SqlQ(lvwUploadsPending.SelectedItem.SubItems(1)) & _
                 " AND " & gconUploadDateField & " IS NULL"
                 
        Set rstUpload = GetRst(pCnn:=g.cnnDW, _
                               pSource:=strSQL, _
                               pSourceType:=adCmdText, _
                               pRstType:=eEditableFwdOnly, _
                               pErrMsg:=strErrMsg)
                               
        Do Until rstUpload.EOF
            rstUpload.Delete
            rstUpload.MoveNext
        Loop
        
        rstUpload.Close
        Set rstUpload = Nothing
        DisplayUploadsPending

    End If

End Sub

Private Sub moBataRpts_AddUnsentComplete(ByVal Msg As String)
    StatusBar pMsg:=Msg
End Sub

Private Sub moBataRpts_AfterRptLoad(oRpt As clsBataRpt, ByVal Success As Boolean)
Dim strMsg As String

    strMsg = "Loading " & oRpt.FranName & " - " & oRpt.Name
    If Not Success Then
        If oRpt.HasData Then
            strMsg = strMsg & " FAILED."
        Else
            strMsg = strMsg & " NO DATA."
        End If
    End If
    
    StatusBar pMsg:=strMsg, pLog:=False
    
End Sub

Private Sub moBataRpts_AfterRptUpload(oRpt As clsBataRpt, ByVal Success As Boolean, ByVal ErrMsg As String)
Dim strMsg As String

''' strMsg = "Uploading " & oRpt.FranName & " - " & oRpt.Name       ''' V401
    strMsg = "Processing " & oRpt.FranName & " - " & oRpt.Name      ''' V401
    If Success Then
        strMsg = strMsg & " succeeded"
    Else
    '   Use uppercase for first part of message to draw attention to problem, but maintain mixed case of
    '   returned ErrMsg to preserve diagnostic details and also place ErrMsg on new line to improve readability
        strMsg = UCase$(strMsg & " FAILED. ") & vbNewLine & ErrMsg
    End If
    
    StatusBar pMsg:=strMsg, pLog:=Not Success, pRefreshEventLogDisplay:=Not Success

End Sub

Private Sub moBataRpts_BeforeRptUpload(oRpt As clsBataRpt, ByVal UploadAttempt As Long)
''' StatusBar pMsg:="Uploading " & oRpt.FranName & " - " & oRpt.Name & " (attempt " & UploadAttempt & ")", pLog:=False  ''' V401
    StatusBar pMsg:="Processing " & oRpt.FranName & " - " & oRpt.Name & " (attempt " & UploadAttempt & ")", pLog:=False ''' V401
End Sub

Private Sub moBataRpts_OnRptLoad(oRpt As clsBataRpt)
    StatusBar pMsg:="Loading " & oRpt.FranName & " - " & oRpt.Name, pLog:=False
End Sub

Private Sub optBataProcessed_Click(Index As Integer)
    RefreshBataTabGrid
End Sub

Private Sub optDialupResults_Click(Index As Integer)
    DialupResults fPrint:=False ' false means display in eventlog window
End Sub

Private Sub optProductReportOnSelectedFranchisesOnly_Click(Index As Integer)

    If Index = 1 Then
        Call subClearProductReportDisplay
        With lstProductReportsFranchiseBusinessName
            .ListIndex = gconDoNotDisplayAnyItems
            .TopIndex = 0
            .Enabled = False
        End With
    Else
        lstProductReportsFranchiseBusinessName.Enabled = True
    End If

End Sub

Private Sub optPRReportonSelectedFranchises_Click(Index As Integer)

    If Index = 1 Then
        Call subClearStickReportDisplay
        With lstPRProductReportsFranchiseBusinessName
            .ListIndex = gconDoNotDisplayAnyItems
            .TopIndex = 0
            .Enabled = False
        End With
    ElseIf Index = 0 Then
        With lstPRProductReportsFranchiseBusinessName
            .Enabled = True
            .ListIndex = gconDisplayFirstItem
        End With
    End If

End Sub

Private Sub optPRSelectedProducts_Click(Index As Integer)

    If Index = 0 Then
        With lstPRProductList
            .Enabled = True
            .ListIndex = gconDisplayFirstItem
        End With
    Else
        Call subClearStickReportDisplay
        With lstPRProductList
            .ListIndex = gconDoNotDisplayAnyItems
            .TopIndex = 0
            .Enabled = False
        End With
    End If

End Sub

Private Sub optPRSendProductReportToDisplay_Click()
    With chkProductReportTabDelimited
        .Enabled = False
        .Value = vbUnchecked
    End With
End Sub

Private Sub optPRSendProductReportToFile_Click()
    With chkPRProductReportTabDelimited
        .Enabled = True
        .Value = vbUnchecked
    End With
End Sub

Private Sub optPRSendProductReportToPrinter_Click()
    With chkProductReportTabDelimited
        .Enabled = False
        .Value = vbUnchecked
    End With
End Sub

Private Sub optSendProductReportToDisplay_Click()
    With chkProductReportTabDelimited
        .Enabled = False
        .Value = vbUnchecked
    End With
End Sub

Private Sub optSendProductReportToFile_Click()
    With chkProductReportTabDelimited
        .Enabled = True
        .Value = vbUnchecked
    End With
End Sub

Private Sub optSendProductReportToPrinter_Click()
    With chkProductReportTabDelimited
        .Enabled = False
        .Value = vbUnchecked
    End With
End Sub

Private Sub optSendStickReportToDisplay_Click()
    With chkStickReportTabDelimited
        .Enabled = False
        .Value = vbUnchecked
    End With
End Sub

Private Sub optSendStickReportToFile_Click()
    With chkStickReportTabDelimited
        .Enabled = True
        .Value = vbUnchecked
    End With
End Sub

Private Sub optSendStickReportToPrinter_Click()
    With chkStickReportTabDelimited
        .Enabled = False
        .Value = vbUnchecked
    End With
End Sub

Private Sub optStickReportOnSelectedFranchisesOnly_Click(Index As Integer)

    If Index = 1 Then
        Call subClearStickReportDisplay
        With lstStickReportsFranchiseBusinessName
            .ListIndex = gconDoNotDisplayAnyItems
            .TopIndex = gconZeroValue
            .Enabled = False
        End With
    ElseIf Index = 0 Then
        Call subClearStickReportDisplay
        With lstStickReportsFranchiseBusinessName
            .Enabled = True
            .ListIndex = gconDisplayFirstItem
        End With
    End If

End Sub

Private Sub optUploadSelection_Click(Index As Integer)
' Enable state and individual franchise selection controls according upload option button selection
Dim lngLoop As Long
Dim chk As VB.CheckBox

    For Each chk In chkUpload_State
        With chk
            .Enabled = (Index = UploadSelModeEnum.eUpldSelStates)
        '   Clear selections as appropriate
            If Not chk.Enabled Then
                .Value = False
            End If
        End With
    Next
    
    lstUploadFranchiseList.Enabled = (Index = UploadSelModeEnum.eUpldSelFrans)
    If (Index <> UploadSelModeEnum.eUpldSelFrans) Then
        With lstUploadFranchiseList
            If .SelCount Then
            '   Clear selections
                For lngLoop = 0 To .ListCount - 1
                    .Selected(lngLoop) = False
                Next lngLoop
            End If
        End With
    End If
    
End Sub

Private Sub PopulateLstProductReportsFranchiseBusinessName(ByVal pIncludeClosedFrans As Boolean)
Dim strSQL As String

'   Franchises added to lstProductReportsFranchiseBusinessName according to chkSalesRptTab_IncludeClosedFrans
    strSQL = "SELECT FranchiseBusinessName FROM Franchises"
    If Not pIncludeClosedFrans Then
        strSQL = strSQL & vbNewLine & "WHERE Live"
    End If
    strSQL = strSQL & vbNewLine & "ORDER BY FranchiseBusinessName"

    LoadListBox_Rst pListBox:=lstProductReportsFranchiseBusinessName, pCnn:=g.cnnDW, pSource:=strSQL, pDisplayFld:="FranchiseBusinessName"

End Sub

Sub PopulateNonCompliantLView()
Dim lngFranCount As Long
Dim dtmLiveDataTableUpdate As Date
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset
    
'   Time consuming -> only update display and tables as required
    dtmLiveDataTableUpdate = GetTableUpdateTime("LiveData")
    If dtmLiveDataTableUpdate > g.udtDtmCtlUpdated.dtmNonCompliantLView Then
        If dtmLiveDataTableUpdate > GetTableUpdateTime("PromoNonCompliantSales") Then
            LoadNonCompliantTable pFranCount:=lngFranCount
'''' Catch up data to current date IS IT REQUIRED BECAUSE LOADING OF NON COMPLIANT LEAVES PREVIOUS
'''' WEEKS DATA AND LOADS FOR YESTERDAYS NON COMPLIANTS ANY TIME WE DOWNLOAD DATA - NEEDS SOME MROE THOUGHT ''' V369
        End If
        
        lvwNonCompliant.ListItems.Clear
        
        strSQL = "SELECT * FROM PromoNonCompliantSales " & vbNewLine & _
                 "WHERE TransactionDate = " & MySqlDate(fdtmYesterday)
        Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
        Do Until rst.EOF
        ''' SWITCHED LINES BELOW B/C IT TURNS OUT THAT NON-COMPLIANT SALES ARE SOMETIMES
        ''' POPULATED WITH MORE THAN ONE ROW OF IDENTICAL FranID/Barcode combination
            Set gvListItem = frmTSGDataWarehouse.lvwNonCompliant.ListItems.Add()
            gvListItem.Text = GetFranName(rst!FranchiseIDTSG)
''' Review Could perhaps be optimised through bypassing calls to gsubAddSubItemToListview
            gsubAddSubItemToListview rst!FranchiseIDTSG, 1
            gsubAddSubItemToListview fsDescriptionFrom(rst!Barcode), 2
            gsubAddSubItemToListview Format(rst!NormalSellInc, "###0.00##"), 3
            gsubAddSubItemToListview Format(rst!ActualSellInc, "###0.00##"), 4
            gsubAddSubItemToListview Format(rst!PromoSellInc, "###0.00##"), 5
            gsubAddSubItemToListview rst!TransactionDate, 6
            rst.MoveNext
            DoEvents
        Loop
        rst.Close
        Set rst = Nothing
        
        cmdPromoTabSaveNonCompliantAll.Enabled = (lvwNonCompliant.ListItems.Count > 0)
        
        lvwNonCompliant.Refresh

        g.udtDtmCtlUpdated.dtmNonCompliantLView = Now

    End If
    
End Sub

Sub PopulateSubCategory()
Dim strErrMsg As String
Dim rstCategory As ADODB.Recordset
Dim strSQL As String

    lstPromoSubCat.Clear
    lstPromoSubCat.Refresh
    'populate sub-category list box
    strSQL = "SELECT DISTINCT cat2 FROM qryStock " & vbNewLine & _
             "WHERE cat2 NOT LIKE " & SqlQ("<%")
    Set rstCategory = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstCategory.BOF And rstCategory.EOF) Then 'populate the combo box
        Do Until rstCategory.EOF
            lstPromoSubCat.AddItem rstCategory(gconStockTableSubCategoryField)
            rstCategory.MoveNext
        Loop
        lstPromoSubCat.ListIndex = gconDisplayFirstItem
    End If
    rstCategory.Close
    Set rstCategory = Nothing
    
End Sub

Private Sub PrintDialupResults()
    'MsgBox ("No printer available")
    On Error GoTo noprinter
    cdlTSGDataWarehouse.ShowPrinter
    Printer.FontName = "helvetica"
    Printer.FontSize = 8
    Printer.Orientation = vbPRORLandscape
    
    Printer.Print "Tobacco Station - Dialup Results For " & fsYesterdaysDate
    'leave a dual gap
    Printer.Print vbCrLf

    Printer.Print "Generated - " & Format(Now, gconStandardDateFormat & " HH:MM:SS")
    'leave a dual gap
    Printer.Print vbCrLf
    DialupResults fPrint:=True  ' true means print (or put to floppy)
    
noprinter:

End Sub

Sub printNonCompliantsReport(Optional pColSelFranIDs As VBA.Collection, Optional pFileNum As Integer = -1)
' Prints non compliant report based on records from PromoNonCompliantSales table
' If pColSelFranIDs is passed it only prints records for Franchises with FranIDs
' in the passed collection, otherwise it will print all the records in the table
    Dim cLines As Integer
    Dim sCurrFranch As String
    Dim sPrevFranch As String
Dim rst As ADODB.Recordset
    
    Set rst = GetRstNonCompliantRpt(pColSelFranIDs:=pColSelFranIDs)
    
    If rst Is Nothing Then
        MsgBox ("No non-compliant records. Press show")
    Else
    
        If pFileNum < 0 Then
            On Error GoTo noNCxprinter
            cdlTSGDataWarehouse.ShowPrinter
            Printer.Orientation = vbPRORLandscape
            Printer.FontName = "Courier New"
            Printer.FontSize = 11
            Printer.FontBold = True
            cLines = 6
        End If
        
        printNonCompliantsReportHeading pFileNum:=pFileNum
        
        Do Until rst.EOF
            sCurrFranch = rst!FranchiseBusinessName
            If sPrevFranch <> sCurrFranch Then
                If pFileNum < 0 Then
                    Printer.Print " "
                    Printer.Print sCurrFranch
                    cLines = cLines + 2
                    cLines = cLines + 1
                Else
                    Print #pFileNum, " "
                    Print #pFileNum, sCurrFranch
                End If
            End If
            
            If pFileNum < 0 Then
                    Printer.Print Tab(6); rst!PromoName; _
                                  Tab(20); rst!PromoSubCat; _
                                  Tab(27); rst!PromoStart; _
                                  Tab(38); rst!PromoEnd; _
                                  Tab(48); Format(rst!PromoCartonDiscount, "##0.00##"); _
                                  Tab(55); Format(rst!PromoPacketDiscount, "#0.00##"); _
                                  Tab(62); rst!Description; _
                                  Tab(103); Format(rst!NormalSellInc, "###0.00##"); _
                                  Tab(111); Format(rst!ActualSellInc, "###0.00##"); _
                                  Tab(118); Format(rst!PromoSellInc, "###0.00##")
            Else
                    Print #pFileNum, Tab(6); rst!PromoName; _
                                   Tab(20); rst!PromoSubCat; _
                                   Tab(27); rst!PromoStart; _
                                   Tab(38); rst!PromoEnd; _
                                   Tab(48); Format(rst!PromoCartonDiscount, "##0.00##"); _
                                   Tab(55); Format(rst!PromoPacketDiscount, "#0.00##"); _
                                   Tab(62); rst!Description; _
                                   Tab(103); Format(rst!NormalSellInc, "###0.00##"); _
                                   Tab(111); Format(rst!ActualSellInc, "###0.00##"); _
                                   Tab(118); Format(rst!PromoSellInc, "###0.00##")
            End If
            
            rst.MoveNext
            sPrevFranch = sCurrFranch
            If pFileNum < 0 Then
                If cLines > 40 Then
                    Printer.Print Tab(50); "Page "; Printer.Page
                    Printer.NewPage
                    printNonCompliantsReportHeading
                    cLines = 6
                End If
            End If
        Loop
        
        rst.Close
        Set rst = Nothing
        If pFileNum < 0 Then
            Printer.EndDoc
            MsgBox "Non-Compliant franchises have been printed." & vbCrLf & vbCrLf
        End If
    End If
    
noNCxprinter:
    Exit Sub

End Sub

Sub printNonCompliantsReportHeading(Optional pFileNum As Integer = -1)

If pFileNum < 0 Then
       Printer.Print "Tobacco Station Non-Compliant Sales for " & Format$(fdtmYesterday, gkFmtDateUnambiguous)
       Printer.Print "Generated - " & Format$(Now, gkFmtDateUnambiguous & " HH:MM:SS")
       Printer.Print "                                                                                                   --------Average-------"
       Printer.Print "                                              Carton Packet                                         Normal  Actual  Promo"
       Printer.Print "     Promotion     SubCat Start      End      Rebate Rebate  Description                              Sell    Sell   Sell"
       Printer.Print "     --------------------------------------------------------------------------------------------------------------------"
Else
    Print #pFileNum, "Tobacco Station Non-Compliant Sales for " & Format$(fdtmYesterday, gkFmtDateUnambiguous)
    Print #pFileNum, "Generated - " & Format$(Now, gkFmtDateUnambiguous & " HH:MM:SS")
    Print #pFileNum, "                                                                                                   --------Average-------"
    Print #pFileNum, "                                              Carton Packet                                         Normal  Actual  Promo"
    Print #pFileNum, "     Promotion     SubCat Start      End      Rebate Rebate  Description                              Sell    Sell   Sell"
    Print #pFileNum, "     --------------------------------------------------------------------------------------------------------------------"
End If

End Sub

Sub PrintPromoList(Optional ByVal pToFile As Boolean = False)
    Dim lngLineCnt As Long
    Dim colRegions As VBA.Collection
Dim rstPromos As ADODB.Recordset
Dim strErrMsg As String
Dim strSQL As String
    
    Set colRegions = GetRegionsCollection()
    
    strSQL = "SELECT * FROM Promotions " & vbNewLine & _
             "WHERE PromoEnd >= " & MySqlDate(Date)         ''' Review - PromoStatus may not be up to date if Master hasn't
                                                            '''         updated and database hasn't been transferred to SLAVEs
    Set rstPromos = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    
    If (rstPromos.BOF And rstPromos.EOF) Then
        MsgBox "No Promotions"
    Else
    
        If pToFile Then
        '   Do nothing YET!
        Else
            On Error GoTo Procedure_Exit
                cdlTSGDataWarehouse.ShowPrinter
                Printer.FontName = "courier new"
                Printer.FontSize = 11
                Printer.FontBold = True
                Printer.Orientation = vbPRORLandscape
                PrintPromoListHeading pToFile:=pToFile
                lngLineCnt = 6
                Do Until rstPromos.EOF
                    Printer.Print _
                    Tab(4); rstPromos!PromoName; _
                    Tab(18); rstPromos!PromoSubCat; _
                    Tab(25); rstPromos!PromoStart; _
                    Tab(35); rstPromos!PromoEnd; _
                    Tab(46); Format$(rstPromos!PromoCartonDiscount, "##0.00"); _
                    Tab(54); Format$(rstPromos!PromoPacketDiscount, "#0.00"); _
                    Tab(61); rstPromos!PromoState; _
                    Tab(67); colRegions.Item(CStr(rstPromos!PromoRegionID))
                    lngLineCnt = lngLineCnt + 1
                    If lngLineCnt > 36 Then
                        Printer.Print Tab(50); "Page "; Printer.Page
                        Printer.NewPage
                        PrintPromoListHeading pToFile:=pToFile
                        lngLineCnt = 6
                    End If
                    rstPromos.MoveNext
                Loop
                rstPromos.Close
                Printer.EndDoc
                MsgBox "Promotions have been printed." & vbCrLf & vbCrLf
        End If
    End If

Procedure_Exit:
    Set rstPromos = Nothing
    Set colRegions = Nothing
    
End Sub

Private Sub PrintPromoListHeading(Optional ByVal pToFile As Boolean = False)

    If pToFile Then
    '   Do nothing YET! - SEE CALLING PROCEDURE
    Else
        Printer.Print "Tobacco Station Promotions for " & Format$(Date, gkFmtDateUnambiguous)
        Printer.Print "Generated - " & Format(Now, gkFmtDateUnambiguous & " HH:MM:SS")
        Printer.Print
        Printer.Print "                                           Carton  Packet"      'MAJOR Regional
        Printer.Print "   Promotion     SubCat Start     End      Rebate  Rebate  State Region"
        Printer.Print "   ----------------------------------------------------------------------------"
    End If

End Sub

Sub PrintRejectedData()
    
    Dim intFileNum As Long
    Dim rsRejects As ADODB.Recordset
    Dim sSkeletonText As String
    Dim lPos As Long
    '!!! ManualFix Clearing: Object variable not cleared: rsFranch
Dim rsFranch As ADODB.Recordset
    Dim sCurrFranch, sPrevFranch As String
    Dim cFaxSkeleton As String
    Dim sAction, sExtra As String
Dim strSQL As String
Dim strErrMsg As String
Dim bPreviousMeEnabled As Boolean
Dim intPrevMousePointer As Integer
                                                                    
    bPreviousMeEnabled = SetFormEnabled(pForm:=Me, pEnabled:=False)
    intPrevMousePointer = SetMousePointer(vbHourglass)
    
    cFaxSkeleton = g.strAppRoot & "fax-templates\RejectedSalesFax.txt"
    strSQL = "SELECT * FROM RejectData ORDER BY FranchiseIDTSG"
    Set rsRejects = GetRst(pCnn:=g.cnnDW, _
                           pSource:=strSQL, _
                           pSourceType:=adCmdText, _
                           pRstType:=eEditableDynamic, _
                           pErrMsg:=strErrMsg)
    If (rsRejects.BOF And rsRejects.EOF) Then
       MsgBox "There are no rejecteed sales data records."
       GoTo Procedure_Exit
    End If
    If gbEventLogRefreshIsNotAlreadyInProgress Then
        gbEventLogRefreshIsNotAlreadyInProgress = False
        With frmTSGDataWarehouse.lvwEventLog
            .ListItems.Clear
            .Refresh
        End With
    End If
    Do Until rsRejects.EOF
        Set gvListItem = frmTSGDataWarehouse.lvwEventLog.ListItems.Add()
        gvListItem.Text = Cn(rsRejects!TransactionDate, vbNullString)   ' Cn() to accommodate rogue data which had shown as a bug
        Call gsubAddSubItemToListview(GetFranName(rsRejects(gconFranchiseTableTSGFranchiseIDField)), 1)
        Call gsubAddSubItemToListview("Qty = " & rsRejects(gconLiveDataTableQuantityField) & _
             "    Amnt = $" & rsRejects(gconLiveDataTableTotalIncTaxField), 2)
        rsRejects.MoveNext
    Loop
    
    If g.rstAppDefaults!NetworkPrinterEnabled Then
        sAction = " printed."
        If MsgBox("Do you want to print the faxes for these 'rejects'?", vbYesNo) = vbNo Then
            gbEventLogRefreshIsNotAlreadyInProgress = True
            GoTo Procedure_Exit
        End If
        
        rsRejects.MoveFirst
        
        If Dir(cFaxSkeleton) = "" Then
            MsgBox "Cannot find cFaxSkeleton", vbCritical
            GoTo Procedure_Exit
        End If
        
        On Error GoTo Procedure_Exit
        cdlTSGDataWarehouse.ShowPrinter
        
       
        Do Until rsRejects.EOF
            Printer.FontName = "courier new"
            Printer.FontSize = 11
            Printer.FontBold = True
            Printer.Orientation = vbPRORPortrait

            sCurrFranch = GetFranName(rsRejects(gconFranchiseTableTSGFranchiseIDField))
            intFileNum = FreeFile   ' Get unused file
            Open cFaxSkeleton For Input As #intFileNum
            Do Until EOF(intFileNum)
                Line Input #intFileNum, sSkeletonText
                lPos = InStr(1, sSkeletonText, "<ScanError>", 1)
                If lPos > 0 Then
                    sPrevFranch = sCurrFranch
                    Do Until rsRejects.EOF Or sCurrFranch <> sPrevFranch
                        Printer.Print Tab(6); rsRejects!TransactionDate; Tab(15); _
                        rsRejects(gconLiveDataTableBarcodeField); Tab(29); _
                        fsDescriptionFrom(rsRejects(gconLiveDataTableBarcodeField)); Tab(59); _
                        rsRejects(gconLiveDataTableQuantityField); Tab(70); _
                        rsRejects(gconLiveDataTableTotalIncTaxField)
                        rsRejects.MoveNext
                        sPrevFranch = sCurrFranch
                        If Not rsRejects.EOF Then
                            sCurrFranch = GetFranName(rsRejects(gconFranchiseTableTSGFranchiseIDField))
                        End If
                    Loop
                Else
                    lPos = InStr(1, sSkeletonText, "<FranchiseName>", 1)
                    If lPos > 0 Then
                        sSkeletonText = Left(sSkeletonText, lPos - 1) & _
                                        sCurrFranch
                    End If
                    lPos = InStr(1, sSkeletonText, "<FaxNumber>", 1)
                    If lPos > 0 Then
                        strSQL = "SELECT * FROM franchises " & vbNewLine & _
                                 "WHERE FranchiseIDTSG = " & rsRejects!FranchiseIDTSG
                        Set rsFranch = GetRst(pCnn:=g.cnnDW, _
                                              pSource:=strSQL, _
                                              pSourceType:=adCmdText, _
                                              pErrMsg:=strErrMsg)
                        If Not (rsFranch.BOF And rsFranch.EOF) Then
                            If rsFranch(gconFranchiseTableFaxField) <> "unknown" Then
                                sSkeletonText = Left(sSkeletonText, lPos - 1) & _
                                rsFranch(gconFranchiseTableAreaCodeField) & " " & _
                                rsFranch(gconFranchiseTableFaxField)
                            End If
                        End If
                        rsFranch.Close
                        Set rsFranch = Nothing
                    End If
                    lPos = InStr(1, sSkeletonText, "<Date>", 1)
                    If lPos > 0 Then
                        sSkeletonText = Left(sSkeletonText, lPos - 1) & Date
                    End If
                    Printer.Print sSkeletonText
                End If
            Loop
            Close #intFileNum
            Printer.EndDoc
        Loop
    Else
        sAction = " dsiplayed."
    End If
    
    ''gbEventLogRefreshIsNotAlreadyInProgress = True
    Call gsubRefreshEventLogDisplay
   
    sExtra = ""
    If Not g.bMaster Then
        sExtra = vbCrLf & "(Please also delete these from the MASTER database.)"
    End If
    If MsgBox("List of rejected sales data has been" & sAction & vbCrLf & vbCrLf & _
              "Delete list of rejected data from database?" & sExtra, vbYesNo) = vbYes Then
        rsRejects.MoveFirst
        Do Until rsRejects.EOF
            rsRejects.Delete
            rsRejects.MoveNext
        Loop
    End If
    
    rsRejects.Close
    Set rsRejects = Nothing
    
Procedure_Exit:
    SetMousePointer pMousePointer:=intPrevMousePointer
    SetFormEnabled pForm:=Me, pEnabled:=bPreviousMeEnabled
    Exit Sub

End Sub

Private Function PromotionRecall(ByVal pPromotionID As Long, pErrMsg As String) As Boolean              '*
'~~~~~~ Asterisks marks furthest column I can use with laptop tilted on side without having to scroll   '*
'~~~~~~ (with debug toolbar docked on left hand side of code window and using Courier New Font Size 8   '*
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'*
'   Could log each promo event (uploading, activating, etc) but would become complicated and is already
'   logged by RStats and can be read by TsgMsgCentre if and when TSG want to upgrade TsgMsgCentre to MySQL
Const kProcName As String = "PromotionRecall"
Dim lngFPCount As Long
Dim strSQL As String
Dim strErrMsg As String
Dim strStatusBarExitMsg As String
Dim strSqlFPCount As String
Dim strFranIdWcValueList As String
Dim strSqlDelUnsent As String
Dim colFranIDs As VBA.Collection

'   Could have a failsafe/double check of Promotions!PromoStatus
'    that we are not calling the proc with an already recalled promo

'   Set Promotions.PromoStatus flag to PROMO_RECALLED so oPOS suite of SW
'   doesn't create FP records for Promotions in process of being recalled
'   If not done first would be a window during the process where it might occur
'   Note that all db changes will be rolled back if there is an error with db.
    strSQL = "UPDATE Promotions " & vbNewLine & _
              " SET PromoStatus  = " & SqlQ(PROMO_RECALLED) & " " & vbNewLine & _
             "WHERE PromoID = " & pPromotionID

'-  Begin data transaction *************************************'*
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx '*
    On Error GoTo Procedure_Error_Rollback                      '*
'   ************************************************************'*
    
    CnnDwExecute pCommandText:=strSQL
    If g.cnnDW.Errors.Count Then
        strErrMsg = "Failed to update Promotions!PromoStatus. " & g.cnnDW.Errors(0).Description
    Else
    '   When recalling we don't want to restrict fran selection. If promo was
    '   tfrd to fran we want recell it regardless of subsequent changes to fran
    '   LEFT Join deletes all FP records regardless of matching FU records (RStats Frans) or not (oPos Frans)
    '   Because we selected (FP.TfrStatus = FpTfrRequested) we know
    '   neither oPOS or TsgDW has transfered promo so we can delete the request
    
        strSqlDelUnsent = "DELETE FP.*, FU.* " & vbNewLine & _
                          "FROM tblFranchisePromotions FP LEFT JOIN FranchiseUploads FU " & vbNewLine & _
                           " ON FP.FranchiseID = FU.FranchiseID " & _
                           " AND CONCAT('PROMO', FP.PromotionID) = FU.UploadFile " & vbNewLine & _
                          "WHERE (FP.PromotionID = " & pPromotionID & ") " & _
                           " AND (FP.TfrStatus = " & FpTfrEnum.FpTfrRequested & ") " & _
                           " AND (FU.UploadedDate IS NULL)"

        CnnDwExecute pCommandText:=strSqlDelUnsent
        If g.cnnDW.Errors.Count Then
            strErrMsg = "Failed to delete FP & related FU records. " & g.cnnDW.Errors(0).Description
        Else
        '   Testing shows despite previous SQL statement being executed (in a txn), the txn applies to this
        '   cnn and can be rolled back and the following code returns lngFPCount as if it wasn't in a txn.
            strSqlFPCount = "SELECT COUNT(*) FROM tblFranchisePromotions " & vbNewLine & _
                            "WHERE (PromotionID = " & pPromotionID & ")"
            lngFPCount = GetRstVal(pCnn:=g.cnnDW, pSource:=strSqlFPCount, pDefaultVal:=-1)
            If lngFPCount = 0 Then
            '   All FP records were unsent so we can delete related Promotion table recrod
            '   Deletion of Promo record logged to EventLog at end of proc after commiting
            '   Txn otherwise would be no trace of it ever having existed
                strSQL = "DELETE FROM Promotions WHERE PromoID = " & pPromotionID
                CnnDwExecute pCommandText:=strSQL
                If g.cnnDW.Errors.Count Then
                    strErrMsg = "Failed to delete Promotion record. " & g.cnnDW.Errors(0).Description
                Else
                    strStatusBarExitMsg = "DELETED Promotion " & pPromotionID & ". " & _
                                          "(Promo was NOT transferred to any franchises)"
                End If
            Else
            '   Some records MAY have (FP.TfrStatus = FpTfrCompleted) [i.e. NOT ALL are FpTfrRequested]
            '   Any FP records with (FP.TfrStatus = FpTfrCompleted) require a RecallRequest
            '   to be added to FranchiseUploads and appropriate changes to tblFranchisePromotions
                Set colFranIDs = GetFranIdColn_FP(pPromoID:=pPromotionID, pTfrStatus:=FpTfrEnum.FpTfrCompleted)
                strFranIdWcValueList = GetWcValueListFromColn(colFranIDs)
                Set colFranIDs = Nothing

            '   Cater for no FpTfrCompleted records requiring recall (Perhaps a previously recalled promo)
            '   therefore strErrMsg isn't populated
                If Len(strFranIdWcValueList) Then
                    '*  Add to tblFranchiseUploads and edits tblFranchisePromotions when recalling
                        AddPromoToFUandFP pAction:=gkPromoDELETE, _
                                          pPromoID:=pPromotionID, _
                                          pFranIdWcValueList:=strFranIdWcValueList, _
                                          pErrMsg:=strErrMsg
                End If
            End If
        End If
    End If

'   Set time stamp and display event log msg after commiting tx
SetTableUpdateTime pTableName:="Promotions", pTimeStamp:=Now    ''' Review better off implemented by TRIGGERS

'-  Resolve data transaction *******************************************'*
    If Len(strErrMsg) Then                                              '*
        Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx  '*
    Else                                                                '*
        Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx    '*
    End If                                                              '*
'-  ********************************************************************'*
    
    If Len(strStatusBarExitMsg) Then
        StatusBar strStatusBarExitMsg
    End If

Procedure_Exit:
    If Len(strErrMsg) Then pErrMsg = kProcName & "() -> " & strErrMsg ' prepend calling proc/stack
    PromotionRecall = (Len(strErrMsg) = 0)
    Exit Function

Procedure_Error_Rollback:
    If Err.Number Then strErrMsg = Trim$(strErrMsg & " " & Err.Source & " " & Err.Number & ": " & Err.Description)
    If g.cnnDW.Errors.Count Then strErrMsg = Trim$(strErrMsg & " " & g.cnnDW.Errors(0).Description)
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
    Resume Procedure_Exit
    
End Function

Private Sub PurgeBataUploadLogs(ByVal pMonthsToKeep As Long)
''' Review ********************************************************************************'
''' Review CALL THIS CODE AFTER AT LEAST ONE NIGHT WITH THE NEW VERSION SO WE KNOW HOW    '
''' Review MUCH SPEED IMPROVEMENT IN Bata Uploads IS FROM LIMITING SIZE OF RSTS TO UPDATE '
''' Review ********************************************************************************'
Dim lngCount As Long
Dim dtmOldestDate As Date
Dim strSQL As String
Dim strErrMsg As String
Dim vntTableName As Variant
Dim astrTableNames() As String
Dim rst As ADODB.Recordset
    
    dtmOldestDate = DateAdd("m", pMonthsToKeep * -1, Date)
    astrTableNames = Array("tblBataUploads", "tblBataREUploads")
    
    For Each vntTableName In astrTableNames
        
        StatusBar "PURGING " & vntTableName & " records older than " & Format$(dtmOldestDate, "dd mmm yyyy") & _
                  "  [Keeping " & Plural(pQty:=pMonthsToKeep, pNounSingular:="month") & "]"
     
        strSQL = "SELECT * FROM " & vntTableName & " WHERE DateTime < " & MySqlDate(dtmOldestDate)
        Set rst = GetRst(pCnn:=g.cnnDW, _
                         pSource:=strSQL, _
                         pSourceType:=adCmdText, _
                         pRstType:=eEditableFwdOnly, _
                         pErrMsg:=strErrMsg)
                         
        Do While Not rst.EOF
            rst.Delete
            lngCount = lngCount + 1

            If lngCount Mod 100 = 0 Then
                StatusBar "Purging " & vntTableName & ": " & lngCount & " records deleted", pLog:=False
                DoEvents
            End If

            rst.MoveNext
        Loop
        
        StatusBar Plural(pQty:=lngCount, pNounSingular:="record") & " deleted." ' In case of 0 records
        rst.Close
    Next vntTableName
    
    Set rst = Nothing

End Sub

Private Sub PurgeEventLog(ByVal pMonthsToKeep As Long)
Dim lngOldestDate As Long
Dim lngMaxSeqDeleted As Long
Dim lngRecordsDeleted As Long
Dim dtmOldestDate As Date
Dim strSQL As String

'   Keep pMonthsToKeep full months plus todays log (i.e. if pMonthsToKeep=0 we keep today)
    dtmOldestDate = DateAdd("m", pMonthsToKeep * -1, Date)
    StatusBar "PURGING event log records older than " & Format$(dtmOldestDate, "dd mmm yyyy") & _
              "  [Keeping " & Plural(pQty:=pMonthsToKeep, pNounSingular:="month") & "]"
              
'   Select < lngOldestDate to cater for gaps (missing dates) in event log
    lngOldestDate = CLng(dtmOldestDate)
    strSQL = "SELECT MAX(Sequence) " & vbNewLine & _
             "FROM EventLog " & vbNewLine & _
             "WHERE lngDate < " & lngOldestDate
    lngMaxSeqDeleted = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=0)
    
'   Skip rare case where above strSQL returns no records (ie lngMaxSeqDeleted = 0)
    If lngMaxSeqDeleted > 0 Then
    '   Add one to retuned lngDate so SQL selection can be "<" cf "<="
        strSQL = "DELETE FROM EventLog WHERE Sequence < " & lngMaxSeqDeleted + 1
        CnnDwExecute pCommandText:=strSQL, pRecordsAffected:=lngRecordsDeleted
    End If
    
    StatusBar "Event log records older than " & Format$(dtmOldestDate, "dd mmm yyyy") & " purged"
    
End Sub

Private Sub PurgeFranchiseUploads(ByVal pMonthsToKeep As Long)
Dim dtmOldestDate As Date
Dim strSQL As String
    
    dtmOldestDate = DateAdd("m", pMonthsToKeep * -1, Date)
    
    StatusBar "PURGING Franchise Upload records older than " & Format$(dtmOldestDate, "dd mmm yyyy") & _
              "  [Keeping " & Plural(pQty:=pMonthsToKeep, pNounSingular:="month") & "]"
    
    strSQL = "DELETE FROM FranchiseUploads " & vbNewLine & _
             "WHERE Not ((UploadedDate Is Null) Or (UploadedDate >= " & MySqlDate(dtmOldestDate) & "))"
    
    CnnDwExecute pCommandText:=strSQL
    
    StatusBar "Franchise upload records older than " & Format$(dtmOldestDate, "dd mmm yyyy") & " purged"

End Sub


Private Sub PurgeLiveData()
Dim lngRecsDeleted As Long
Dim lngDaysKept As Long
Dim strSQL As String
    
    lngDaysKept = Date - g.dtmLiveDataStart
    StatusBar "PURGING live data records older than " & Format$(g.dtmLiveDataStart, "dd mmm yyyy") & _
              "  [Keeping " & Plural(pQty:=lngDaysKept, pNounSingular:="day") & "]"
    
    strSQL = "DELETE FROM livedata WHERE TransactionDate < " & MySqlDate(g.dtmLiveDataStart)
    
    CnnDwExecute pCommandText:=strSQL, pRecordsAffected:=lngRecsDeleted
    
    dtpBataTabTxDate.MinDate = GetMinBataTabTxDate() ' Date picker control will change control value if necessary  ''' V401

    StatusBar Plural(pQty:=lngRecsDeleted, pNounSingular:="record") & " deleted." ' In case of 1 record
    
End Sub

Private Sub PurgeNonCompliantPromos()
'   Keep a week of Non Compliant data (Yesterday and six days prior to yesterday)
Dim dtmOldestDate As Date
Dim strSQL As String

'   Interval parameter: w = Weekday, ww = week, (d = day, y = year, h = hour, ...)
    dtmOldestDate = DateAdd(Interval:="ww", Number:=-1, Date:=Date)
    
    StatusBar "PURGING Non Compliant Promo records older than " & _
              Format$(dtmOldestDate, "dd mmm yyyy") & "  [Keeping one week]"

    strSQL = "DELETE FROM PromoNonCompliantSales " & vbNewLine & _
             "WHERE (TransactionDate < " & MySqlDate(dtmOldestDate) & ")"
    
    CnnDwExecute pCommandText:=strSQL

    StatusBar "Non compliant promo records older than " & Format$(dtmOldestDate, "dd mmm yyyy") & " purged"

End Sub

Private Sub PurgeUploadsPending(ByVal fAll, ByVal pFranID As Long)
Dim strSQL As String
Dim strErrMsg As String
Dim rsUploadsPending As ADODB.Recordset
    
    If fAll Then
        strSQL = "SELECT * FROM FranchiseUploads " & _
                 " WHERE " & gconUploadDateField & " IS NULL"
    Else
        strSQL = "SELECT * FROM FranchiseUploads " & _
                 "WHERE " & gconUploadDateField & " IS NULL AND " & _
                          gconUploadFranchiseIDField & " = " & pFranID
    End If

    Set rsUploadsPending = GetRst(pCnn:=g.cnnDW, _
                                  pSource:=strSQL, _
                                  pSourceType:=adCmdText, _
                                  pRstType:=eEditableFwdOnly, _
                                  pErrMsg:=strErrMsg)
    Do Until rsUploadsPending.EOF
        rsUploadsPending.Delete
        rsUploadsPending.MoveNext
    Loop
    rsUploadsPending.Close
    Set rsUploadsPending = Nothing
    
    With lvwUploadsPending
        .ListItems.Clear
        .Refresh
    End With

End Sub

Private Sub RefreshBataTabGrid()
Dim bTxDateColHidden As Boolean
Dim lngFranCount As Long
Dim strSelection As String
Dim avnt() As Variant
Dim prmTxDate As ADODB.Parameter
Dim prmSelection As ADODB.Parameter
Dim com As ADODB.Command
Dim rst As ADODB.Recordset
    
'.  Collect data
'
'   Get max num of frans we need to accommodate
    lngFranCount = GetRecordCount(pCnn:=g.cnnDW, pSource:="Franchises")

'   Determine selection (Uploaded, NOT Uploaded or ALL)
    Select Case True
        Case optBataProcessed(0): strSelection = "Processed"     ' Processed (Uploaded or Disk File Created)
        Case optBataProcessed(1): strSelection = "NOT Processed" ' NOT Processed
        Case optBataProcessed(2): strSelection = "ALL"           ' All Bata Franchises
    End Select
    
'   Create and populate command parameters
    Set prmTxDate = New ADODB.Parameter
    With prmTxDate
    ''' .Name = "pTxDate"
        .Name = "pDate"
        .Type = ADODB.DataTypeEnum.adDate
        .Direction = adParamInput
        .Value = dtpBataTabTxDate.Value
    End With
    
    Set prmSelection = New ADODB.Parameter
    With prmSelection
        .Name = "pSelection"
        .Type = ADODB.DataTypeEnum.adBSTR
        .Direction = adParamInput
        .Value = strSelection
    End With
    
'   Create command
    Set com = New ADODB.Command
    With com
        Set .ActiveConnection = g.cnnDW
    ''' .CommandText = "qryBataUploadsGrid"
        If InStr(cboBataTabTxOrProcessedDate.Text, "Transaction") Then
            .CommandText = "qryBataRptLogGridSelTxDate"
        Else
            .CommandText = "qryBataRptLogGridSelPrDate"
        End If
        .CommandType = adCmdStoredProc
        .Parameters.Append prmTxDate
        If InStr(cboBataTabTxOrProcessedDate.Text, "Transaction") Then
            .Parameters.Append prmSelection
        End If
    End With

'   Only way to open rst with a command in ADO and specify CursorType and Locking Scheme
'   even though Default Values are CursorType:=adOpenForwardOnly & LockType:=adLockReadOnly
'   Also, could perhaps alter GetRst() or write another similar proc for opening a rst
'   with a command as it would be helpful for trapping any errors in a standardised way
'   (Given default vals of CursorType & LockType alternative call could be "Set rst = com.Execute")
    Set rst = New ADODB.Recordset
    With rst
        .LockType = adLockReadOnly      ' Should by default open as ReadOnly
        .CursorType = adOpenForwardOnly ' Should by default open as ForwardOnly
        .Open Source:=com, Options:=adCmdStoredProc
        If Not (.BOF And .EOF) Then
        '   [If you] request more rows than are available GetRows returns only the number of available rows.
            avnt = .GetRows(lngFranCount * 2) ' Franchises * 2 (for each rpt type) will always be enough
        End If
    End With
    rst.Close: Set rst = Nothing
    Set com = Nothing
    
'.  Load grid with data
'
    With grdBataRpts
        .Redraw = flexRDNone    ' Suspend redraw
        .Rows = .FixedRows      ' Clear grid
    '   Set Hidden flag for TxDate column
        bTxDateColHidden = InStr(cboBataTabTxOrProcessedDate.Text, "Transaction")
        .ColHidden(.ColIndex("TxDate")) = bTxDateColHidden          ' Set visibility TxDate Col
        .ColHidden(.ColIndex("Uploaded")) = Not bTxDateColHidden    ' Set visibility ProcessedTime Col
        
        If Not IsEmptyArray(avnt) Then
            .LoadArray avnt              ' Load grid
        End If
        .AutoResize = True
        .Redraw = flexRDBuffered         ' Redraw grid
    End With
    
'   Update total Franchises label
    lblBataTabFranCount = Plural(pQty:=GridGetCollection(pGrid:=grdBataRpts, pColKey:="FranID", pSelected:=False).Count, _
                                 pNounSingular:="BATA Franchise")
    
    ConfigureBataTabButtons
    
    SetTabRefreshedFlag pTab:=TabEnum.eBataTab, pRefreshed:=True
    
End Sub

Private Function RemotePromotionRecall(ByVal pFranID As Long, _
                                       ByVal pFranName As String, _
                                       ByRef pCnnRemote As ADODB.Connection, _
                                       ByVal pDelPromoTag As String) As Boolean
' Recall a promotion uploaded to a store
' 1. If promotion has NOT been displayed to user (via RStats) delete it from the stores database.
' 2. If promotion has been displayed to user (via RStats) edit PromoEnd to Yesterday
'    so RemoteStatitcs notifies promo as expired on next polling
Dim bResult As Boolean
Dim lngPromoID As Long
Dim eFPTfrStatus As FpTfrEnum
Dim dtmYesterday As Date
Dim strMsg As String
Dim strSQL As String
Dim rstRemote As ADODB.Recordset
Dim strErrMsg As String

'   Set default FpTfrStatus value for when procedure executes without failure but doesn't
'   find a promo with New or Notified status. In this case Promo the has either expired,
'   or for an unknown reason doesn't exist and may was well be flagged as recalled
    eFPTfrStatus = FpTfrEnum.FpRecalled
    
    On Error GoTo Procedure_Error
    
'   pDelPromoTag = "DELPROMO" & PromoID  where PromoID is a whole number
    lngPromoID = Val(Right$(pDelPromoTag, Len(pDelPromoTag) - Len(gkPromoDELETE)))
    strSQL = "SELECT * FROM Promotions WHERE PromoID = " & lngPromoID
    Set rstRemote = GetRst(pCnn:=pCnnRemote, _
                           pSource:=strSQL, _
                           pSourceType:=adCmdText, _
                           pRstType:=eEditableFwdOnly, _
                           pErrMsg:=strErrMsg)
'   Review: see notes in analagous part of UploadPromotion
    With rstRemote
        If Not (.BOF And .EOF) Then
            strMsg = "RECALLED Promotion " & lngPromoID & " from " & pFranName & ". "
            Select Case Cn(.Fields!PromoStatus, vbNullString)
                Case "New"
                '   Promotion has NOT been displayed by Remote Statistics (or dispalyed but user
                '   closed reminder form with the X close box rather than with buttons provided)
                '   DELETE Promotion record from remote statistics database
                    .Delete
                    eFPTfrStatus = FpTfrEnum.FpRecalled
                    strMsg = strMsg & "(Promo was NOT been displayed/applied - remote record deleted)"
                    StatusBar pMsg:=strMsg, pFranchise:=pFranName
                Case "Notified"
                '   Promotion has been dispalyed by Remote Statistics. Instead of deleting the record
                '   without informing user that promotion has been recalled, we change the PromoEnd date
                '   so that on next polling RemoteStatistics informs user that the promotion has expired
                    dtmYesterday = fdtmYesterday()
                    .Fields!PromoEnd = dtmYesterday
                    .Update
                    eFPTfrStatus = FpTfrEnum.FpRecallRequestUploaded
                    strMsg = strMsg & "(Promo had been displayed/applied - remote record End Date set to " & _
                             Format$(dtmYesterday, gkFmtDateUnambiguous) & ")"
                    StatusBar pMsg:=strMsg, pFranchise:=pFranName
                Case Else ' "Expired"
                '   No Action as user already informed promotion has expired
            End Select
        End If
        .Close
    End With
    Set rstRemote = Nothing
    bResult = True  ' SHOULD FLAG SHOULD BE WITHIN THE IF NOT(.BOF AND EOF) AND ONLY WHEN AN ACTION HAS BEEN TAKEN
                    ' OR SHOULD WE SIMPLY LOG AN EVEN IN THE EVENT LOG WHEN THERE IS NOTHING TO RECALL
                    ' AND MARK IT AS SUCCESSFULLY RECALLED BECAUSE IT IS OR SHOULD THERE BE A NEW STATUS
    

    SetFPTfrStatus pFranID:=pFranID, pPromoID:=lngPromoID, pTfrStatus:=eFPTfrStatus

Procedure_Exit:
    RemotePromotionRecall = bResult
    Exit Function
    
Procedure_Error:
    bResult = False
    StatusBar "Error Recalling Promotion - Err: " & Err.Number & " " & Err.Description, pFranName
    Resume Procedure_Exit

End Function

' This routine resets the OpenedByField in the defaults table on the remote DB
Private Sub ReSetRemoteOpenedByField(ByVal pFranName As String, _
                             ByRef pCnnRemote As ADODB.Connection, _
                             ByVal pOpenedBy As String)
Dim strErrMsg As String
Dim rsRemoteDefaults As ADODB.Recordset

    On Error GoTo Procedure_Error
    Set rsRemoteDefaults = GetRst(pCnn:=pCnnRemote, _
                                  pSource:="Defaults", _
                                  pSourceType:=adCmdTable, _
                                  pRstType:=eEditableFwdOnly, _
                                  pErrMsg:=strErrMsg)
        
        rsRemoteDefaults(pOpenedBy) = vbNullString
    rsRemoteDefaults.Update
    rsRemoteDefaults.Close
    Set rsRemoteDefaults = Nothing

Procedure_Exit:
    Exit Sub
    
Procedure_Error:
    StatusBar "error resetting remote value." & pOpenedBy, pFranName
    Resume Procedure_Exit

End Sub

Sub SaveNewPromotion()
Const kProcName As String = "SaveNewPromotion"
Const kUseInProcedureTxProcessing As Boolean = False
Dim bCtnRebate As Boolean
Dim bPktRebate As Boolean
Dim bSelectedFrans As Boolean
Dim intPrevMousePointer  As Integer
Dim lngCtnIdx As Long
Dim lngPktIdx As Long
Dim lngLoop As Long
Dim lngSubCatIdx As Long
Dim strErrMsg As String
Dim vntState As Variant
Dim vntRegionID As Variant
Dim colSelStates As VBA.Collection
Dim colSelRegionIDs As VBA.Collection

''' Review *** Revisit & determine whether we really need to test for a Pkt rebate being included (see Tobac, Cigars, ...)  ***
    With grdPromoTabRebates
        lngCtnIdx = .ColIndex("Carton")
        lngPktIdx = .ColIndex("Packet")
        For lngLoop = .FixedRows To .Rows - 1
            If .TextMatrix(Row:=lngLoop, col:=2) > 0 Then
                bCtnRebate = True
            End If
            If .TextMatrix(Row:=lngLoop, col:=3) > 0 Then
                bPktRebate = True
            End If
        Next lngLoop
    End With

    bSelectedFrans = ChkBoxToBool(chkPromoSelectFranchise)

    If Len(Trim$(txtPromoName.Text)) = 0 Then
        MsgBox "Must have a name", vbExclamation
        txtPromoName.SetFocus
    ElseIf Not bSelectedFrans And ((lstPromoTabState.SelCount = 0) And (chkPromoTabAllStates.Value <> vbChecked)) Then
        MsgBox "No states selected", vbExclamation
        lstPromoTabState.SetFocus
    ElseIf Not bSelectedFrans And ((lstPromoTabRegion.SelCount = 0) And (chkPromoTabAllRegions.Value <> vbChecked)) Then
        MsgBox "No regions selected", vbExclamation
        If lstPromoTabRegion.Enabled Then
            lstPromoTabRegion.SetFocus
        Else
            chkPromoTabAllRegions.SetFocus
        End If
    ElseIf lstPromoProducts.ListCount = 0 Then
        MsgBox "No products selected", vbExclamation
        lstPromoSubCat.SetFocus
    ElseIf Not bCtnRebate Then                                      ''' Review Maybe just give a warning for zero rebate and ask if they want to continue
        MsgBox "Must have a carton discount", vbExclamation         ''' Review just give a warning for zero rebate and ask if they want to continue
        grdPromoTabRebates.SetFocus                                 ''' Review just give a warning for zero rebate and ask if they want to continue
    ElseIf Not bPktRebate Then                                      ''' Review Maybe just give a warning for zero rebate and ask if they want to continue
        MsgBox "Must have a valid packet discount", vbExclamation   ''' Review Maybe just give a warning for zero rebate and ask if they want to continue
        grdPromoTabRebates.SetFocus                                 ''' Review Maybe just give a warning for zero rebate and ask if they want to continue
    ElseIf dtpPromoEnd.Value < dtpPromoStart.Value Then
        MsgBox "Promotion End Date must be later than or the same date as the Start Date", vbExclamation
        dtpPromoEnd.SetFocus
    ElseIf ChkBoxToBool(chkPromoSelectFranchise) And (lstPromoFranchise.SelCount = 0) Then
        MsgBox "You checked 'Select Specific Franchises' but haven't selected any franchises.", vbExclamation
        lstPromoFranchise.SetFocus
    ElseIf MsgBox("Are you sure the Promotion is between " & _
                   Format$(dtpPromoStart.Value, gkFmtDateUnambiguous) & _
         " and " & Format$(dtpPromoEnd.Value, gkFmtDateUnambiguous) & "?", vbYesNo) <> vbYes Then
        dtpPromoEnd.SetFocus
    Else
        Me.Enabled = False
        intPrevMousePointer = SetMousePointer(vbHourglass)
        
        If ChkBoxToBool(chkPromoSelectFranchise) Then
            Set colSelStates = New Collection:      colSelStates.Add mkPromoStatesNA
            Set colSelRegionIDs = New Collection:   colSelRegionIDs.Add mkPromoRegionsNA
        '   PromoGrade will be populated from grid which is populated with a single row with
        '   PromoGradeID equal to mkPromoGradeIdNA when chkPromoSelectFranchise is selected
        Else
            If ChkBoxToBool(chkPromoTabAllStates) Then
                Set colSelStates = New Collection
                colSelStates.Add mkPromoStatesAll
            Else
                Set colSelStates = ListBoxGetCollection(pListBox:=lstPromoTabState, pItemData:=False, pSelected:=True)
            End If
        
            If ChkBoxToBool(chkPromoTabAllRegions) Then
                Set colSelRegionIDs = New Collection
                colSelRegionIDs.Add mkPromoRegionsAll
            Else
                Set colSelRegionIDs = ListBoxGetCollection(pListBox:=lstPromoTabRegion, pItemData:=True, pSelected:=True)
            End If
        End If
    
        If kUseInProcedureTxProcessing Then
        '   Wrap everthing with same PromoName into a single Txn
        '-  Begin data transaction *************************************'*
            Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eBeginTx '*
            On Error GoTo Procedure_Error_Rollback                      '*
        '   **************************************************************

        End If
        
        For lngSubCatIdx = 0 To lstPromoSubCat.ListCount - 1
            If (lstPromoSubCat.Selected(lngSubCatIdx)) Then
                With grdPromoTabRebates
                    For lngLoop = .FixedRows To .Rows - 1
                    '   If either Packet or Carton rebate is greater than 0
                        If (.ValueMatrix(Row:=lngLoop, col:=lngCtnIdx) > 0) Or _
                           (.ValueMatrix(Row:=lngLoop, col:=lngPktIdx) > 0) Then
                            For Each vntState In colSelStates
                                For Each vntRegionID In colSelRegionIDs
                                    If Not AddNewPromo(pPromoName:=txtPromoName.Text, _
                                                    pSubCat:=lstPromoSubCat.List(lngSubCatIdx), _
                                                    pPromoStart:=dtpPromoStart.Value, _
                                                    pPromoEnd:=dtpPromoEnd.Value, _
                                                    pCtnDiscount:=.ValueMatrix(Row:=lngLoop, col:=2), _
                                                    pPktDiscount:=.ValueMatrix(Row:=lngLoop, col:=3), _
                                                    pRegionID:=vntRegionID, _
                                                    pState:=vntState, _
                                                    pPromoGradeID:=.ValueMatrix(Row:=lngLoop, col:=0), _
                                                    pErrMsg:=strErrMsg) Then
                                        If kUseInProcedureTxProcessing Then
                                            GoTo Procedure_Rollback
                                        Else
                                            strErrMsg = "Failed to save promotion " & txtPromoName & " for " & _
                                                         "State " & vntState & _
                                                         ", Region " & Substitute(vntRegionID, mkPromoRegionsAll, "All") & _
                                                         ", PromoGradeID " & .ValueMatrix(Row:=lngLoop, col:=0) & vbNewLine & _
                                                         strErrMsg
                                             StatusBar pMsg:=Replace$(Expression:=strErrMsg, Find:=vbNewLine, Replace:=" ")
                                        End If
                                    End If
                                Next vntRegionID
                            Next vntState
                        End If
                    Next lngLoop
                End With
            End If
        Next lngSubCatIdx
                
        If kUseInProcedureTxProcessing Then
            Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eCommitTx
        End If
        
        Set colSelStates = Nothing
        Set colSelRegionIDs = Nothing
        
        LoadPromoListview pShowALL:=False
    
        Me.Enabled = True   ' Re-Enable form before ClearCreatePromoCtls calls txtPromoName.SetFocus)
        ClearCreatePromoCtls ' Question is do I want to ClearCreatePromoCtls when there has been an error ???? in which case put in Procedure exit
   
    End If
    
Procedure_Exit:
'   Re-Enable form and reset mouse pointer in cse we got here via error handler
    If Len(strErrMsg) Then
        If Not kUseInProcedureTxProcessing Then
            MsgBox "Failed to add some promotion combinations. Check event log for details.", vbInformation
        Else
            StatusBar pMsg:=kProcName & "() -> " & strErrMsg
            MsgBox strErrMsg, vbExclamation
        End If
    End If
    Me.Enabled = True
    SetMousePointer intPrevMousePointer
    Exit Sub

Procedure_Error_Rollback:
'   Rollback all promo combinations with samePromoName
    If Err.Number Then strErrMsg = Trim$(strErrMsg & " " & Err.Source & " " & Err.Number & ": " & Err.Description)
    If g.cnnDW.Errors.Count Then strErrMsg = Trim$(strErrMsg & " " & g.cnnDW.Errors(0).Description)
    strErrMsg = "Failed to save promotion " & txtPromoName & vbNewLine & strErrMsg
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
    Resume Procedure_Exit
    
Procedure_Rollback:
'   Rollback all promo combinations with samePromoName
    strErrMsg = "Failed to save promotion " & txtPromoName & vbNewLine & strErrMsg
    Tx pCnn:=g.cnnDW, pProcName:=kProcName, pTxAction:=eRollbackTx
    GoTo Procedure_Exit

End Sub

Sub SaveNonCompliantRptToCsvFile(Optional pColSelFranIDs As VBA.Collection)
' Exports non compliant report based on records from PromoNonCompliantSales table
' If pColSelFranIDs is passed it only exports records for Franchises with FranIDs
' in the passed collection, otherwise it will export all the records in the table
Dim bPrevFormEnabled As Boolean
Dim intPrevMousePointer As Integer
Dim strRpt As String
Dim strFldSep As String
Dim strFileName As String
Dim rst As ADODB.Recordset
Dim fdlg As fdlgCommon

    bPrevFormEnabled = SetFormEnabled(pForm:=Me, pEnabled:=False)
    intPrevMousePointer = SetMousePointer(vbHourglass)

    strFldSep = ", "
    Set rst = GetRstNonCompliantRpt(pColSelFranIDs:=pColSelFranIDs)
    If Not (rst Is Nothing) Then
        Set fdlg = New fdlgCommon
        strFileName = fdlg.GetFullFileName(pMethod:=eShowSave, _
                                           pFilename:="NonCompliantRpt_" & Format$(Date, gkFmtDateUnambiguous) & ".csv", _
                                           pFilter:="*.csv", _
                                           pFilterDescription:="Comma Separated Values (*.csv)", _
                                           pDefaultExtension:="csv")
        Set fdlg = Nothing
        If Len(strFileName) Then
            strRpt = "Franchise" & strFldSep & "Promotion" & strFldSep & "Sub Cat" & strFldSep & "Start" & strFldSep & "End" & strFldSep & _
                     "Ctn Rebate" & strFldSep & "Pkt Rebate" & strFldSep & "Description" & strFldSep & _
                     "Normal Sell" & strFldSep & "Actual Sell" & strFldSep & "Promo Sell" & vbNewLine
            
            Do While Not rst.EOF
                strRpt = strRpt & rst!FranchiseBusinessName & strFldSep & _
                                  rst!PromoName & strFldSep & _
                                  rst!PromoSubCat & strFldSep & _
                                  Format$(rst!PromoStart, gkFmtDateUnambiguous) & strFldSep & _
                                  Format$(rst!PromoEnd, gkFmtDateUnambiguous) & strFldSep & _
                                  Format$(Cn(rst!PromoCartonDiscount, 0), "#0.00##") & strFldSep & _
                                  Format$(rst!PromoPacketDiscount, "#0.00##") & strFldSep & _
                                  rst!Description & strFldSep & _
                                  Format$(rst!NormalSellInc, "###0.00##") & strFldSep & _
                                  Format$(rst!ActualSellInc, "###0.00##") & strFldSep & _
                                  Format$(rst!PromoSellInc, "###0.00##") & vbNewLine
                rst.MoveNext
            Loop
            
            If Len(strRpt) Then
                strRpt = Left$(strRpt, Len(strRpt) - Len(vbNewLine))
            End If
            rst.Close
            Set rst = Nothing
            
            SaveTextFile pFilename:=strFileName, pFileText:=strRpt, pOverwrite:=True
        End If
    
    End If

    SetMousePointer intPrevMousePointer
    SetFormEnabled Me, bPrevFormEnabled

End Sub

Sub SetCboCtnContainingPkt(ByVal pStkId As Long)
Dim lngLoop As Long
    With cboCtnContainingPkt
        .ListIndex = -1 ' Clear selection
        For lngLoop = 0 To .ListCount - 1
            If .ItemData(lngLoop) = pStkId Then
                .ListIndex = lngLoop
                Exit For
            End If
        Next lngLoop
    End With
End Sub

Private Sub SetFPTfrStatus(ByVal pFranID As Long, _
                           ByVal pPromoID As Long, _
                           ByVal pTfrStatus As FpTfrEnum)
' All these types of procedures could become procedures that return SQL strings to perform the desired action
Dim strSQL As String
    
    strSQL = "UPDATE tblFranchisePromotions " & vbNewLine & _
             " SET TfrStatus = " & pTfrStatus & " " & vbNewLine & _
             "WHERE (FranchiseID = " & pFranID & ")" & _
             " AND  (PromotionID = " & pPromoID & ")"
    
    CnnDwExecute pCommandText:=strSQL

End Sub

Sub SetGlobalVariables()
    Dim lPos As Long
    Dim xx As String
Dim strMsg As String
Dim strErrMsg As String
Dim strPkFilename As String
Dim strPKZFoldername As String
Dim fso As Scripting.FileSystemObject
Dim rst As ADODB.Recordset

    Set fso = New Scripting.FileSystemObject
    
    xx = LCase(g.rstAppDefaults!StatisticsDatabase)
    lPos = InStr(1, xx, "\" & LCase(gCompanyIdentifier) & "\", vbTextCompare)
    If lPos = 0 Then
        MsgBox "Could not find the Company Identifier " & gCompanyIdentifier & " in the " & _
        "Data Warehouse database path " & g.rstAppDefaults!StatisticsDatabase & _
        "Check the defaults.mdb", vbCritical
        End
    End If
    
    Set rst = GetRstAddOnly(pCnn:=g.cnnDW, pSource:="EventLog", pErrMsg:=strErrMsg)
    g.lngEventLogEventFldSize = rst!Event.DefinedSize
    g.lngEventLogFranFldSize = rst!Franchise.DefinedSize
    rst.Close
    Set rst = Nothing
    
    g.strAppDrive = Left$(g.rstAppDefaults!StatisticsDatabase, lPos)
    g.strAppRoot = g.strAppDrive & gCompanyIdentifier & "\"
    g.strTsTemp = g.strAppRoot & "Temp"
    g.strLogFolder = g.strAppRoot & "Logs"
    g.strBatscanFolder = g.strAppRoot & "Batscan"
    g.strUploadsFolder = g.strAppRoot & "Uploads"
    g.strLocalMessageFolder = g.strAppRoot & mkMessageFolderName
    g.strRptsFolder = g.strAppRoot & "Reports"
    g.strNielsenRptsFolder = g.strRptsFolder & "\" & "Nielsen"
    g.strBataRptsFolder = g.strRptsFolder & "\" & "BATA"

    g.strPkZipCExe = g.strAppRoot & "Program\PKZIPC.EXE"
    gsProductReportPathAndFilename = g.strRptsFolder & "\" & gconProductReportFilename
    gsStickReportPathAndFilename = g.strRptsFolder & "\" & gconStickReportFilename

'   Automatically create required folders that don't exist
    If Not fso.FolderExists(g.strTsTemp) Then fso.CreateFolder g.strTsTemp
    If Not fso.FolderExists(g.strLogFolder) Then fso.CreateFolder g.strLogFolder
    If Not fso.FolderExists(g.strBatscanFolder) Then fso.CreateFolder g.strBatscanFolder
    If Not fso.FolderExists(g.strUploadsFolder) Then fso.CreateFolder g.strUploadsFolder
    If Not fso.FolderExists(g.strLocalMessageFolder) Then fso.CreateFolder g.strLocalMessageFolder
    If Not fso.FolderExists(g.strRptsFolder) Then fso.CreateFolder g.strRptsFolder
    If Not fso.FolderExists(g.strNielsenRptsFolder) Then fso.CreateFolder g.strNielsenRptsFolder
    If Not fso.FolderExists(g.strBataRptsFolder) Then fso.CreateFolder g.strBataRptsFolder
    
'''''' V396 - reinstating archive of live data - reinstate setting according to Dw!Defaults!DaysOfLiveData
''''   g.dtmLiveDataStart set to earliest TxDate in LiveData table (sales data)
''''   (i.e. No longer set according to Dw-db!Defaults!DaysOfLiveData that     )
''''   (determined archiving of  LiveData table when Dw used an Access mdb file)
'''    g.dtmLiveDataStart = GetRstVal(pCnn:=g.cnnDW, _
'''                                   pSource:="SELECT MIN(TransactionDate) FROM LiveData", _
'''                                   pDefaultVal:=Date)
                                      
    SetSystemDateReliantSettings
    
    LockDCTabFranchiseCtls pLocked:=True
    
''' Review If call to subPopulateFranchiseBusinessNameListBoxes() remains here,
''' Review If it does, the procedure should at least conditionally access the database
    subPopulateFranchiseBusinessNameListBoxes
    
    If (Not IsIDE()) And (Not fso.FileExists(g.strPkZipCExe)) Then
        strPkFilename = fso.GetFileName(g.strPkZipCExe)
        strPKZFoldername = fso.GetParentFolderName(g.strPkZipCExe)
        strMsg = "WARNING" & vbNewLine & _
                 "File zipping will not work until " & strPkFilename & _
                 " is copied into the " & SQ(strPKZFoldername) & " folder." & vbNewLine & _
                 "Please copy " & strPkFilename & " into the " & SQ(strPKZFoldername) & " folder."
        MsgBox strMsg, vbExclamation
     End If
    
    Set fso = Nothing
    
''' gsubAddToLocalEventLog App.Title & " " & g.strNodeType & " (Version " & fsVersion() & ") started. " & g.strNodeName ''' V397
    StatusBar App.Title & " " & g.strNodeType & " (Version " & fsVersion() & ") started. " & g.strNodeName, _
              pRefreshEventLogDisplay:=False                                                                             ''' V397

    
End Sub

Private Sub SetSystemDateReliantSettings()
Dim dtmToday As Date
Dim dtmYesterday As Date
Dim dtmLastSunday As Date

    dtmToday = Date
    dtmYesterday = DateAdd(Interval:="d", Number:=-1, Date:=dtmToday)
    dtmLastSunday = fdtmLastSunday()
    g.dtmLiveDataStart = DateAdd(Interval:="d", Number:=-g.rstDWDefaults!DaysOfLiveData, Date:=dtmToday)
    
    With dtpBataTabTxDate
        .MaxDate = dtmYesterday
        .Value = dtmYesterday
    End With
    
    With tdpEventLogDate
        .MaxDate = dtmToday
        .Value = dtmToday
    End With
    
    With dtpNielsenRptTxDate
        .MaxDate = dtmLastSunday
        .Value = dtmLastSunday
    End With

End Sub

Private Sub SetTabRefreshedFlag(ByVal pTab As TabEnum, ByVal pRefreshed As Boolean)
    m.ablnTabRefreshed(pTab) = pRefreshed
End Sub

Private Sub spnProductReportFinishDate_SpinDown()

    If DateDiff("d", gfsSplitDate(lblProductReportStartDate), gfsSplitDate(lblProductReportFinishDate)) > gconZeroValue Then
        Call subClearProductReportDisplay
        lblProductReportFinishDate = _
            Format(DateAdd("d", -1, gfsSplitDate(lblProductReportFinishDate)), gconStandardDateFormat)
        lblProductReportFinishDate.Refresh
        Call subSetProductReportDateWording
    End If

End Sub

Private Sub spnProductReportFinishDate_SpinUp()

    If DateDiff("d", gfsSplitDate(lblProductReportFinishDate), Date) > gconZeroValue Then
        Call subClearProductReportDisplay
        lblProductReportFinishDate = _
            Format(DateAdd("d", 1, gfsSplitDate(lblProductReportFinishDate)), gconStandardDateFormat)
        lblProductReportFinishDate.Refresh
        Call subSetProductReportDateWording
    End If

End Sub

Private Sub spnProductReportStartDate_SpinDown()

    Call subClearProductReportDisplay
    lblProductReportStartDate = _
        Format(DateAdd("d", -1, gfsSplitDate(lblProductReportStartDate)), gconStandardDateFormat)
    lblProductReportStartDate.Refresh
    Call subSetProductReportDateWording

End Sub

Private Sub spnProductReportStartDate_Spinup()
    
    If DateDiff("d", gfsSplitDate(lblProductReportStartDate), gfsSplitDate(lblProductReportFinishDate)) > gconZeroValue Then
        Call subClearProductReportDisplay
        lblProductReportStartDate = _
            Format(DateAdd("d", 1, gfsSplitDate(lblProductReportStartDate)), gconStandardDateFormat)
        lblProductReportStartDate.Refresh
        Call subSetProductReportDateWording
    End If

End Sub

Private Sub spnPRProductReportFinishDate_SpinDown()

    If DateDiff("d", gfsSplitDate(lblPRProductReportStartDate), gfsSplitDate(lblPRProductReportFinishDate)) > gconZeroValue Then
        lblPRProductReportFinishDate = _
            Format(DateAdd("d", -1, gfsSplitDate(lblPRProductReportFinishDate)), gconStandardDateFormat)
        lblPRProductReportFinishDate.Refresh
        Call subSetPRProductReportDateWording
    End If

End Sub

Private Sub spnPRProductReportFinishDate_SpinUp()

    If DateDiff("d", gfsSplitDate(lblPRProductReportFinishDate), Date) > gconZeroValue Then
        lblPRProductReportFinishDate = _
            Format(DateAdd("d", 1, gfsSplitDate(lblPRProductReportFinishDate)), gconStandardDateFormat)
        lblPRProductReportFinishDate.Refresh
        Call subSetPRProductReportDateWording
    End If

End Sub

Private Sub spnPRProductReportStartDate_SpinDown()

    lblPRProductReportStartDate = _
        Format(DateAdd("d", -1, gfsSplitDate(lblPRProductReportStartDate)), gconStandardDateFormat)
    lblPRProductReportStartDate.Refresh
    Call subSetPRProductReportDateWording

End Sub

Private Sub spnPRProductReportStartDate_Spinup()
    
    If DateDiff("d", gfsSplitDate(lblPRProductReportStartDate), gfsSplitDate(lblPRProductReportFinishDate)) > gconZeroValue Then
        lblPRProductReportStartDate = _
            Format(DateAdd("d", 1, gfsSplitDate(lblPRProductReportStartDate)), gconStandardDateFormat)
        lblPRProductReportStartDate.Refresh
        Call subSetPRProductReportDateWording
    End If

End Sub

Private Sub spnStickReportFinishDate_SpinDown()

    If DateDiff("d", gfsSplitDate(lblStickReportStartDate), gfsSplitDate(lblStickReportFinishDate)) > gconZeroValue Then
        Call subClearStickReportDisplay
        lblStickReportFinishDate = _
            Format(DateAdd("d", -1, gfsSplitDate(lblStickReportFinishDate)), gconStandardDateFormat)
        lblStickReportFinishDate.Refresh
        Call subSetStickReportDateWording
    End If

End Sub

Private Sub spnStickReportFinishDate_SpinUp()

    If DateDiff("d", gfsSplitDate(lblStickReportFinishDate), Date) > gconZeroValue Then
        Call subClearStickReportDisplay
        lblStickReportFinishDate = _
            Format(DateAdd("d", 1, gfsSplitDate(lblStickReportFinishDate)), gconStandardDateFormat)
        lblStickReportFinishDate.Refresh
        Call subSetStickReportDateWording
    End If

End Sub

Private Sub spnStickReportStartDate_SpinDown()

    Call subClearStickReportDisplay
    lblStickReportStartDate = _
        Format(DateAdd("d", -1, gfsSplitDate(lblStickReportStartDate)), gconStandardDateFormat)
    lblStickReportStartDate.Refresh
    Call subSetStickReportDateWording

End Sub

Private Sub spnStickReportStartDate_SpinUp()

    If DateDiff("d", gfsSplitDate(lblStickReportStartDate), gfsSplitDate(lblStickReportFinishDate)) > gconZeroValue Then
        Call subClearStickReportDisplay
        lblStickReportStartDate = _
            Format(DateAdd("d", 1, gfsSplitDate(lblStickReportStartDate)), gconStandardDateFormat)
        lblStickReportStartDate.Refresh
        Call subSetStickReportDateWording
    End If

End Sub

Private Sub spnTopSellers_SpinDown()

    If giTopSellers > 5 Then
        giTopSellers = giTopSellers - 1
        cmdTopSellers.Caption = "&Top " & giTopSellers
        Me.Refresh
    End If

End Sub

Private Sub spnTopSellers_SpinUp()

    If giTopSellers < 100 Then
        giTopSellers = giTopSellers + 1
        cmdTopSellers.Caption = "&Top " & giTopSellers
        Me.Refresh
    End If

End Sub

Function stripType(ByVal sDescrip As String)
    Dim sType As String
    
    stripType = sDescrip
    sType = LCase(fGetLastWord(sDescrip, " "))
    If (sType = "pkt") Or (sType = "ctn") Or (sType = "pouch") Or (sType = "outer") Then
        stripType = Left(sDescrip, Len(sDescrip) - Len(sType))
    End If
    
End Function

Private Sub subCaptureData(ByVal pAutoCaptureCycle As Boolean, _
                  Optional ByRef pColSelFranIDs As VBA.Collection = Nothing, _
                  Optional ByRef pCaptureOptions As clsDataCaptureOptions = Nothing)
Dim bSelFranCapture As Boolean
Dim lngMaxCycles As Long
Dim lngCycle As Long
Dim lngRecsTfrd2LiveTbl As Long
Dim lNoDataInRemoteDbThisCycle As Long
Dim lngDaysSinceLastTransfer As Long
Dim lngErrHandlingAttemptsThisRemoteCnn As Long
Dim lRecordsTransferred As Long
Dim lTotalSystemsNotUsed As Long
Dim lngTasksCompleted As Long
Dim lNonCompliantRecCount As Long
Dim lFranCount As Long
Dim lngFranID As Long
Dim lngFransINCluded As Long
Dim lngFransEXcluded As Long
Dim datCaptureCycleDate As Date ' determined by combination of date & gkCaptureStartTime (may not be today)
Dim strMsg As String
Dim strFranName As String
Dim strPercentCaptured As String
Dim sNoDataInDatabaseThisCycle As String
Dim sRemoteDbFullname As String
Dim sRemoteModuleVersion As String
Dim sTotalSystemsNotUsedThisCycle As String
Dim strSQL As String
Dim strErr As String
Dim strRMgrVer As String
Dim vntFranID As Variant
Dim vntFranName As Variant
Dim colFranIDs As VBA.Collection
Dim colFranNames As VBA.Collection
Dim strErrMsg As String
Dim lngTempDataRecCount As Long
Dim rstFran As ADODB.Recordset
Dim rstDWTempData As ADODB.Recordset
Dim rstDWPreLiveData As ADODB.Recordset
Dim rstRemoteDefaults As ADODB.Recordset
Dim rstRemoteData As ADODB.Recordset
Dim cnnRemote As ADODB.Connection
Dim oCaptureOptions As clsDataCaptureOptions

    If Not g.bCaptureCycleRunning Then
        g.bCaptureCycleRunning = True
        gbEventLogRefreshIsEnabled = Not pAutoCaptureCycle  ' Disable EventLog refresh for AutoCaptureCycle
        Me.Enabled = False
        gfFutureDate = False
        gfDateFormatBad = False
        datCaptureCycleDate = GetCaptureCycleDate()
        tdpEventLogDate.Value = Date ' Set tdpEventLogDate to today so event log refreshes are diplayed
        Me.Refresh
        
        If pCaptureOptions Is Nothing Then
            Set oCaptureOptions = New clsDataCaptureOptions
        Else
            Set oCaptureOptions = pCaptureOptions
        End If
        
        If Not (pColSelFranIDs Is Nothing) Then
            StatusBar "Data Capture Started - Selected Franchise(s)"
            bSelFranCapture = True
            lngMaxCycles = 1
            Set colFranIDs = pColSelFranIDs
        Else
        '   Capture All (Selected FranIDs not passed in pColSelFranIDs)
            StatusBar "Data Capture Started - ALL Franchises"
            lngMaxCycles = 2
        '   Get collection of FranIDs for franchises for this capture cycle
            If pAutoCaptureCycle Then
                Set colFranIDs = GetSelFranCollection(pSelFranEnum:=eSelFran_CaptureCycleAuto)
            Else
                Set colFranIDs = GetSelFranCollection(pSelFranEnum:=eSelFran_CaptureCycleManual)
            End If
                g.rstDWDefaults!LastAllFranCaptureCycleDate = datCaptureCycleDate
            g.rstDWDefaults.Update
        End If
        lngFransINCluded = colFranIDs.Count
        
'TSL-Dynamite ---------------------------------------------------------------------------------------------------------------------------
        'ensure any records from the pre-live table are transferred to the live table
        'as the app may have crashed mid-cycle, set the remote flag to clear the remote
        'live data, thereby leaving live data in the data warehouse pre-live table
        lngRecsTfrd2LiveTbl = TfrAllPreLiveDataToLiveData()
'TSL-Dynamite lngRecsTfrd2LiveTbl = TfrAllPreLiveDataToLiveData()
'TSL-Dynamite Depending on how many rst creations occur in TfrAllPreLiveDataToLiveData() and called procs, we may want to eliminate
'TSL-Dynamite THIS call (is called later in proc again) from all modes except where subCaptureData() is capturing data from all frans
'TSL-Dynamite ---------------------------------------------------------------------------------------------------------------------------
        
        For lngCycle = 1 To lngMaxCycles
            StatusBar "Pass " & lngCycle & " of " & lngMaxCycles & "; Franchises for processing: " & colFranIDs.Count
        '   Use FranIdCollection for iterating through Frans for DataCapture
            For Each vntFranID In colFranIDs
                lngFranID = CLng(vntFranID)
                strSQL = "SELECT * FROM Franchises WHERE FranchiseIDTSG = " & lngFranID
            '   RstType is eEditableDynamic to enable multiple updates (not possible with eEditableFwdOnly)
            '   Proc edits the same rstFran record twice and therefore must use RstTypeEnum.eEditableDynamic
                Set rstFran = GetRst(pCnn:=g.cnnDW, _
                                     pSource:=strSQL, _
                                     pSourceType:=adCmdText, _
                                     pRstType:=eEditableDynamic, _
                                     pErrMsg:=strErrMsg)
                strFranName = rstFran!FranchiseBusinessName
                subPurgeTable pTableName:="TemporaryData"
    
            '   Make network connection and map drive
                If fConnectFranchiseMapShareDisk(prstFran:=rstFran, pAttemptNumber:=lngCycle) Then
                    sRemoteDbFullname = GetRemotePath(prstFran:=rstFran) & "\" & gkRemoteDbFilename
                    StatusBar "Attempting to establish remote database connection", pLog:=False
'~~ Version 340 Start ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~|
'   This change to be left at least 6 months                                                                    |
'   6 months after change the Event Log should be reviewed                                                      |
'   Error handling enabled immediately before any remote access to handle 'drop outs' to remote machines        |
'   Some suspicion that previous use of DaoGetDb() and DaoGetRst() may not have been catching all the errors    |
'   Previously the subsequent lines tested 'If Len(strErr) Then' whereas from version 340 they will             |
'   test if the returned object 'Is Nothing'. Unfortunately progressively fewer error reports have been emailed |
'   so there is less and less data to make any analysis on                                                      |
'   Error Handler enabling has been moved from immediately prior to opening rstRemoteData rst to here           |
                    On Error GoTo RemoteProcessingError                                                        '|
                    lngErrHandlingAttemptsThisRemoteCnn = 0                                                    '|
'~~ Version 340 End ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~|
                        
'*' Could try opening Remote database exclusively and could then do away with all the code that
'*' checks rstRemoteDefaults!DatabaseOpenedBy field. RemoteStatistics would be the first
'*' port of call requiring it to start opening up exclusively. NOTE THAT CLIFFY
'*' Would need to fix the stores that somehow lock the mdb (Northgate & Upper Eastlands)
'*' CLIFFY could be accommodated by eg.  bOpenMdbExclusive = UCase(prstFranchise!FranTypeName) <> "CLIFF"
'*' but RemoteStatistics.exe first needs a complete rework
                
                '   Opening connection in ShareDenyNone mode because Seventh Beam plans to
                '   create a phone app for suppliers to install promotions at franchises
                    Set cnnRemote = GetCnn(pDataSource:=sRemoteDbFullname, pCnnMode:=adModeShareDenyNone, pDataSourceType:=eMdb, pCursorLocn:=adUseServer, pErrMsg:=strErrMsg)
                    If cnnRemote Is Nothing Then
                        StatusBar UCase$("Could not connect to remote database. ") & strErr, strFranName
                        subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, could not connect to remote database."
                    Else
                        StatusBar "Remote database connection established", strFranName
                    '   If GetRst() doesn't return a rst THEN handle & report problem regardless of Len(strErr)
                        Set rstRemoteDefaults = GetRst(pCnn:=cnnRemote, pSource:="Defaults", pSourceType:=adCmdTable, pRstType:=eEditableFwdOnly, pErrMsg:=strErr)
                        If rstRemoteDefaults Is Nothing Then
                            StatusBar UCase$("Remote defaults table init failed. ") & strErr, strFranName
                        Else
                            StatusBar "Remote Defaults table opened", strFranName
                        '   Confirm correct franchise ID
                            StatusBar "Validating franchise ID", pLog:=False
                            If (rstRemoteDefaults.BOF And rstRemoteDefaults.EOF) Then
                                StatusBar "ID is missing from database", strFranName
                                colFranIDs.Remove Index:=CStr(vntFranID)
                                subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, ID missing from remote database"
                            Else
                                If rstRemoteDefaults!FranchiseID <> lngFranID Then
                                '   Franchise IDs match
                                    StatusBar "ID mismatch", strFranName
                                    AddToRemoteEventLog "ID mismatch", strFranName, pCnnRemote:=cnnRemote
                                    colFranIDs.Remove Index:=CStr(vntFranID)
                                    subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, ID mismatch"
                                Else
                                    AddToRemoteEventLog "Database accessed by head office", strFranName, pCnnRemote:=cnnRemote
                                    If (LCase$(Cn(rstRemoteDefaults!DatabaseOpenedBy, vbNullString)) = "franchise") Then
                                    '   Franchise is currently accessing the database
                                        StatusBar "Remote module is currently processing, unable to upload data", strFranName
                                        AddToRemoteEventLog "Remote module is currently processing, unable to upload data", strFranName, pCnnRemote:=cnnRemote
                                        subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, remote module processing"
                                    Else
                                        StatusBar "Locking remote database", pLog:=False
                                            rstRemoteDefaults.Fields!DatabaseOpenedBy = "HeadOffice"    ''' Disk or network error (Manual Capture error rpt email from Batscan 10Jul2009 - V3.2.9]
                                        rstRemoteDefaults.Update
                                                                 
                                    '   Keep TsgDW version records current
                                        sRemoteModuleVersion = Cn(rstRemoteDefaults!Version, vbNullString)
                                        strRMgrVer = Cn(rstRemoteDefaults(gkRStatsMdbRMgrVerFld), vbNullString) ' populated by RemoteStatistics
                                            rstFran!FranchisePriceModuleVersion = Cn(rstRemoteDefaults!PriceModuleVersion, "N/A")
                                            rstFran!FranchiseRemoteVersion = Left$(Trim$(sRemoteModuleVersion), rstFran!FranchiseRemoteVersion.DefinedSize)
                                            rstFran!FranchiseOSVersion = Left$(Cn(rstRemoteDefaults!OSVersion, vbNullString), rstFran!FranchiseOSVersion.DefinedSize)
                                            rstFran!FranchiseRMVersion = Left$(Trim$(strRMgrVer), rstFran!FranchiseRMVersion.DefinedSize)
                                        rstFran.Update
                                                                 
                                        If rstRemoteDefaults!AvailableRecords = 0 Then
                                            'even if the store has not made any sales, there should be at
                                            'least one record being TOTALCUSTOMERS=0, so if here, the
                                            'remote module has not purged, hence it is not running
                                            StatusBar "No available records in remote database", strFranName
                                            AddToRemoteEventLog "No available records in remote database", strFranName, pCnnRemote:=cnnRemote
                                            colFranIDs.Remove Index:=CStr(vntFranID)
                                            subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, No available records"
                                            lNoDataInRemoteDbThisCycle = lNoDataInRemoteDbThisCycle + 1
                                        Else
                                            StatusBar "Building remote recordset", pLog:=False
                                        '   If GetRst() doesn't return a rst then handle & report problem regardless of Len(strErr)
                                            Set rstRemoteData = GetRst(pCnn:=cnnRemote, pSource:="Statistics", pSourceType:=adCmdTable, pRstType:=eReadOnlyFwdOnly, pErrMsg:=strErrMsg)
                                            If rstRemoteData Is Nothing Then
                                                StatusBar UCase$("Remote Live Data Table Init Failed. ") & strErr, strFranName
                                            Else
                                                'NB Will not attempt to collect data for franchises already collected
                                                ''  As at 16Oct2009 following error msg had not been logged in preceding 6 months
                                                If (rstRemoteData.BOF And rstRemoteData.EOF) Then
                                                    StatusBar UCase$("No data found in remote database [V" & sRemoteModuleVersion & "]"), strFranName
                                                    AddToRemoteEventLog "No data found in database", strFranName, pCnnRemote:=cnnRemote
                                                    colFranIDs.Remove Index:=CStr(vntFranID)
                                                    subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, no data found in remote database"
                                                    lNoDataInRemoteDbThisCycle = lNoDataInRemoteDbThisCycle + 1
                                                Else
                                                    StatusBar "Data is downloading", pLog:=False
                                                    AddToRemoteEventLog "Data is uploading to head office", strFranName, pCnnRemote:=cnnRemote
                                                    'initially copy the remote data into the temporary stats table
                                                    'as the phone line may drop out during tansfer, in which case we can
                                                    're-attempt the transfer again later without having to worry about the
                                                    'possibility of introducing duplicate stats entries into the "live"
                                                    'TSG data warehouse data table
                                                    lRecordsTransferred = 0
                                                    Set rstDWTempData = GetRst(pCnn:=g.cnnDW, pSource:="TemporaryData", pSourceType:=adCmdTable, pRstType:=eEditableFwdOnly, pErrMsg:=strErrMsg)
                                                    Do While Not rstRemoteData.EOF
                                                        lRecordsTransferred = lRecordsTransferred + 1
                                                        StatusBar "Downloading record " & lRecordsTransferred, pLog:=False
                                                        rstDWTempData.AddNew
                                                            rstDWTempData!FranchiseIDTSG = lngFranID
                                                            rstDWTempData!Barcode = rstRemoteData!Barcode
                                                        '   GetDate_FromTsgDateFld() must remain until Cliffy converts his RStats.mdb
                                                        '   When all stores are converted we will send him our new mdb structure,
                                                        '   Cliffy will probably only need to conform with the Statistics table
                                                            rstDWTempData!TransactionDate = GetDate_FromTsgDateFld(pFld:=rstRemoteData!TransactionDate)
                                                            rstDWTempData!Quantity = rstRemoteData!Quantity
                                                            rstDWTempData!NormalSellInc = rstRemoteData!NormalSellInc
                                                            rstDWTempData!CostInc = rstRemoteData!CostInc
                                                            rstDWTempData!TotalInc = rstRemoteData!TotalInc
                                                            rstDWTempData!WholesaleQty = rstRemoteData!WholesaleQty
                                                            rstDWTempData!WholesaleActualSell = rstRemoteData!WholesaleActualSell
                                                            rstDWTempData.Update
                                                        rstRemoteData.MoveNext
                                                    Loop
                                                    rstDWTempData.Close
                                                    
                                                '   Now add this franchises sales to the pre-live table.
                                                    lngTempDataRecCount = GetRecordCount(pCnn:=g.cnnDW, pSource:="TemporaryData")
                                                    If lngTempDataRecCount <> rstRemoteDefaults!AvailableRecords Then
                                                        StatusBar "Transfer mismatch: received = " & lngTempDataRecCount & ", expected available = " & rstRemoteDefaults!AvailableRecords & " remote module is possibly not running (no purge)", strFranName
                                                        AddToRemoteEventLog "Transfer mismatch: received = " & lngTempDataRecCount & ", expected available = " & rstRemoteDefaults!AvailableRecords, strFranName, pCnnRemote:=cnnRemote
                                                        colFranIDs.Remove Index:=CStr(vntFranID)
                                                        subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, transfer mismatch"
                                                        subPurgeTable "TemporaryData"
                                                    Else
                                                        'the number of records expected from the remote live taybl equals
                                                        'the number transferred into the warehouse temporary taybl
                                                        'so empower the remote module to deleyt all now (stale)
                                                        'remote live data records
                                                            rstRemoteDefaults!AvailableRecords = 0
                                                        rstRemoteDefaults.Update
                                                        'now transfer the warehouse temporary data to the pre-live taybl
                                                        Set rstDWTempData = GetRst(pCnn:=g.cnnDW, pSource:="TemporaryData", pSourceType:=adCmdTable, pRstType:=eEditableFwdOnly, pErrMsg:=strErrMsg)
                                                        Set rstDWPreLiveData = GetRstAddOnly(pCnn:=g.cnnDW, pSource:="PreLiveData", pErrMsg:=strErrMsg)
                                                    '   AUrban This code could go a good rewriting and optimising.
                                                    '   AUrban could use action querires, transactions, etc, etc, etc
                                                        With rstDWTempData
                                                        '   If this is wrapped in a transaction error trapping needs to RollBack Transactions
                                                            Do Until .EOF
                                                                TransferSalesRecord rstDWTempData, rstDWPreLiveData, False
                                                                'simultaneous deleyt here as we don't have to worry about the possibility of
                                                                'the process being interrupted by a (flaky) telephone network connection via phone line
                                                                .Delete
                                                                .MoveNext
                                                            Loop
                                                        End With
                                                        rstDWPreLiveData.Close
                                                        Set rstDWPreLiveData = Nothing
                                                            
                                                    '   Calculate whether remote system is being used (ie records in 3 days means ii is not being used)
                                                        lngDaysSinceLastTransfer = DateDiff("d", rstFran!CaptureCycleDateOnLastDataCapture, datCaptureCycleDate)
                                                        If Not (lngDaysSinceLastTransfer > 0) Then
                                                            'transfer has occurd today for this franchise, why???
                                                            StatusBar "Capture already occurred for this franch today (manual?)", strFranName
                                                        'so delete the pre-live data for this franchise
                                                        ' PAL no don't delete... any duplicates will be rejected because by check for duplicates flag
                                                        Else
                                                            If lRecordsTransferred / lngDaysSinceLastTransfer < 2 Then 'POS is not being used
                                                                lTotalSystemsNotUsed = lTotalSystemsNotUsed + 1
                                                                subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Failed, System possibly not being used, Contact System Admin"
                                                            Else 'more than one record per day for this franchise
                                                                subUpdateFranDialupResult lngFranID, "(Attempt " & lngCycle & ") " & Format$(Now, gkFmtDateTime) & " - Success"
                                                            End If 'more than one record per day for this franchise?
                                                        End If 'is the number of days greater than zero?
                                                                
                                                        StatusBar Plural(lRecordsTransferred, "record") & " successfully downloaded [V" & sRemoteModuleVersion & "]", strFranName
                                                        AddToRemoteEventLog "Data was successfully downloaded by head office", strFranName, pCnnRemote:=cnnRemote

                                                        colFranIDs.Remove Index:=CStr(vntFranID)
                                                            rstFran!CaptureCycleDateOnLastDataCapture = datCaptureCycleDate
                                                        rstFran.Update
                                                        lngTasksCompleted = lngTasksCompleted + 1
                                                    End If 'record/transfer mismatch?
                                                    rstDWTempData.Close
                                                    Set rstDWTempData = Nothing
                                                End If 'does franchise have any stats available to upload ? if-endif
                                                rstRemoteData.Close
                                                Set rstRemoteData = Nothing
                                            End If  'If rstRemoteData successfully opened
                                        End If
                                    End If 'is franchise currently accessing remote stats mdb? if-endif
''' V386 Start - Moved here from below in procedure
                                '   Prior to V386 UploadFilesToOneFranchise() was always executed if a cnn to a remote
                                '   RStats.mdb was established, REGARDLESS of whether a rst for a remote RStats.mdb!defaults
                                '   table with a matching FranID to TsgDw-db!Franchises table was established
                                
                                '   While still connected, check if anything to be uploaded, such as messages, new stock etc.
                                ''' Version 343 Start
                                ''' Is a problem whereby if promotions are recalled by SelectedFran capture then those not selected but
                                ''' requiring promotion recall do not subsequently get the recall b/c of how flags/tables etc are set
                                ''' This would always have been a problem for franchises that failed in a bulk upload in CaptureAll/O/N Cycle/Upload All
                                ''' so this kludge to minimise the problem has been instituted until I am authorised to fix this problem.
                                ''' If Not UploadFilesToOneFranchise(prstFran:=rstFran, pSessionSeln:=False, pDbRemote:=dbRemote, pErrMsg:=strErr) Then
                                ''' ----------------------------------------------------------------------------------------------------------'
                                ''' *** ABOVE COMMENT NEEDS TESTING BY RUNNING THIS SCENARIO IN THE PROGRAM                                ***'
                                ''' ***    - OBVIUOSLY COMMENTS WRITTEN FOR A REASON BUT HARD TO SEE HOW THIS HAPPENS BY READING THE CODE  ***'
                                ''' ***                                                                                                    ***'
                                ''' ***    - APPEARS THAT THE PROBLEM IS TO DO WITH MATCHING RECALLS AND UPLOADS. RECALLING OF PROMOS WAS  ***'
                                ''' ***    - ORIGINALLY CONTROLLED BY Promotions!PromoStatus WHICH WAS SET ONCE ONE PROMOTOION HAD BEEN    ***'
                                ''' ***    - RECALLED. SUBSEQUENT CALLS TO RECALL A PROMOTION WHOULD THEREFORE NOT RECALL THE PROMOTION    ***'
                                ''' ***    - AND ANY PENDING UPLOADS OF THE PRMOTION WOULD STILL BE UPLOADED!!                             ***'
                                ''' ***    - THEREFORE IT WAS BEST TO GET ALL AddNewPromo OR ALL DelPromo UPLOADED IN A SINGLE BATCH          ***'
                                ''' ***                                                                                                    ***'
                                ''' ***    - ABOVE COULD ALSO BE EXTENDED INTO A REASON UploadFilesToOneFranchise() SHOULD BE UPLOADED     ***'
                                ''' ***    - EVEN IF THERE IS A MISMATCHED FranID, BUT AS PROMO STUFF IS WRITTEN INTO RStats.mdb           ***'
                                ''' ***    - THIS ARGUMENT IS REDUNDANT AND A STUFF-UP IS UNAVOIDABLE                                      ***'
                                ''' ----------------------------------------------------------------------------------------------------------'
                            ''' V400 Start - Removing pKludgeUploadPromos as
                            '''              Promos are now largely managed by FranchiseUploads & tblFranchisePromotions tables.
                            '''              Promotions!PromoStatus field is for display and communication with oPOS suite of S/W
                            '''              Copius comments in this part of the procedure should all be removed in the next version
                            '''     If Not UploadFilesToOneFranchise(prstFran:=rstFran, pSessionSeln:=False, pCnnRemote:=cnnRemote, pErrMsg:=strErr, pKludgeUploadPromos:=Not bSelFranCapture) Then
                                    If Not UploadFilesToOneFranchise(prstFran:=rstFran, pSessionSeln:=False, pCnnRemote:=cnnRemote, pErrMsg:=strErr) Then
                            ''' V400 End
                                ''' Version 343 End
                                        StatusBar pMsg:=strErr, pFranchise:=strFranName
                                    End If
''' V386 End   - Moved here from below in procedure
                                End If 'is remote franchise ID same as the TSG data warehouses office's iteration ? if-endif
                            End If ' End of else clause for 'If empty rstRemoteDefaults Then'
                                
''' V386 Start - Moved up in procedure
'''                        '   While still connected, check if anything to be uploaded, such as messages, new stock etc.
'''                        ''' Version 343 Start
'''                        ''' Is a problem if promotions are recalled by SelectedFran capture then those not selected but
'''                        ''' requiring promotion recall do not subsequently get the recall b/c of how flags/tables etc are set
'''                        ''' This would always have been a problem for franchises that failed in a bulk upload in CaptureAll/O/N Cycle/Upload All
'''                        ''' so this kludge to minimise the problem has been instituted until I am authorised to fix this problem.
'''                        ''' If Not UploadFilesToOneFranchise(prstFran:=rstFran, pSessionSeln:=False, pDbRemote:=dbRemote, pErrMsg:=strErr) Then
'''                        ''' ----------------------------------------------------------------------------------------------------------'
'''                        ''' *** ABOVE COMMENT NEEDS TESTING BY RUNNING THIS SCENARIO IN THE PROGRAM                                ***'
'''                        ''' ***    - OBVIUOSLY COMMENTS WRITTEN FOR A REASON BUT HARD TO SEE HOW THIS HAPPENS BY READING THE CODE  ***'
'''                        ''' ----------------------------------------------------------------------------------------------------------'
'''                            If Not UploadFilesToOneFranchise(prstFran:=rstFran, pSessionSeln:=False, pCnnRemote:=cnnRemote, pErrMsg:=strErr, pKludgeUploadPromos:=Not bSelFranCapture) Then
'''                        ''' Version 343 End
'''                                StatusBar pMsg:=strErr, pFranchise:=strFranName
'''                            End If
''' V386 End  - Moved up in procedure
                                
                        '   Clesr DatabaseOpenedbyField in remote database
                '????   '   (Nb if UploadFilesToOneFranchise() failed cnnRemote MAY now be Nothing)
                            If Not cnnRemote Is Nothing Then
                                With rstRemoteDefaults
                                        .Fields!DatabaseOpenedBy = vbNullString
                                    .Update
                                    .Close
                                End With
                                Set rstRemoteDefaults = Nothing
                                cnnRemote.Close   ''' Disk or network error (O/N Capture error rpt email from Batscan 10Jul2009 from DAO days.]
                                Set cnnRemote = Nothing
                            End If
' ???   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                        End If  ' End Else Clause of  'If Not opened rstRemoteDefaults'
    
                    End If  ' If [Not]/Opened Remote Db Then
                    On Error GoTo 0
UnmapShareDisk:
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  '
'   SPOT WE MAY ACT CONDITIONALLY ON ACCORDING TO VALUE IN lngErrHandlingAttemptsThisRemoteCnn  '
                '   Finished with this franchise, so disconnect                                 '
                    UnmapShareDiskDisconnectFranchise rstFran                                   '
                End If                                                                          '
'   SPOT WE MAY ACT CONDITIONALLY ON ACCORDING TO VALUE IN lngErrHandlingAttemptsThisRemoteCnn  '
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  '
    
'   If there's a pending update you can't close the rst until the update is cancelled.
'   This can happen when there has been a problem updating the rst and the error
'   handler resumes at UnmapShareDisk. If this happens, we get caught in an infinitely
'   loop. If this happens there would be no harm crashing if we can't address the Fran
'   table, however we could also handle it.

'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  '
'   SPOT WE MAY ACT CONDITIONALLY ON ACCORDING TO VALUE IN lngErrHandlingAttemptsThisRemoteCnn  '
'   THAT IS NOTING THAT THE REMOTE PROCESSING PROBLEM MAY IN RARE CASES HAVE BEEN SUBSEQUENT TO
'   REMOTE PROCESSING AND TO DO WITH EDITING FRAN REOCRD. NOTE IN THIS CASE WE MAY AS WELL GIVE
'   UP. THE CODE FOR EDITING THE FRAN RECORD SHOULD BE BULLET PROOF.
                If rstFran.EditMode = adEditInProgress Then
                    rstFran.Update      'asdf
                End If
                rstFran.Close
                Set rstFran = Nothing
                DoEvents    ' DoEvents returns an Integer representing number of open forms in stand-alone versions of Visual Basic, such as Visual Basic, Professional Edition.
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~  '
            Next vntFranID
        
        Next lngCycle
        
        Set colFranIDs = Nothing
        
        If Not bSelFranCapture Then
        '   List franchises not included in Capture Cycle in the event log
            Set colFranNames = GetSelFranCollection(pSelFranEnum:=eSelFran_CaptureCycleExcluded, pSelFld:="FranchiseBusinessName")
            lngFransEXcluded = colFranNames.Count
            For Each vntFranName In colFranNames
                StatusBar "Not included in this data capture cycle", vntFranName
                subUpdateFranDialupResult lngFranID, Format$(Now, gkFmtDateTime) & " - Failed, franchise is not included"
            Next vntFranName
            Set colFranNames = Nothing
            
            ImportBatScanFiles          ''' Review ImportBatScanFiles should be incorportated into the stats
            
'*********************************************************************************************
'housekeeping
'*********************************************************************************************
            
''' Review Start - Thinking of ommenting out until I can make sense of what was
'''                trying to be reported and to report it in a meaningful way
           If lTotalSystemsNotUsed > 0 Then
               sTotalSystemsNotUsedThisCycle = ", " & lTotalSystemsNotUsed & " NOT collecting data"
           End If
        
           If lNoDataInRemoteDbThisCycle > 0 Then
               sNoDataInDatabaseThisCycle = ", " & lNoDataInRemoteDbThisCycle & " no data"
           End If
''' Review End
            strPercentCaptured = Format$(lngTasksCompleted / lngFransINCluded, "Percent")
            
            strMsg = lngTasksCompleted & " of " & lngFransINCluded & " included Franchises captured " & _
                     Bracket(strPercentCaptured) & sTotalSystemsNotUsedThisCycle & ", " & _
                     lngFransEXcluded & " Franchises NOT included" & ", " & _
                     sNoDataInDatabaseThisCycle
        ''' gsubAddToLocalEventLog  strMsg, "Summary"                   ''' V397
            StatusBar strMsg, "Summary", pRefreshEventLogDisplay:=False ''' V397
        End If
        
        lngRecsTfrd2LiveTbl = TfrAllPreLiveDataToLiveData() + lngRecsTfrd2LiveTbl
'ASDF 'TSL-Dynamite lngRecsTfrd2LiveTbl = TfrAllPreLiveDataToLiveData + lngRecsTfrd2LiveTbl
        
        If pAutoCaptureCycle Then
            StatusBar "BATA Uploads Start"
            Set moBataRpts = New clsBataRpts
        ''' moBataRpts.Upload pAddUnsent:=True          ''' V401 Start
            moBataRpts.Process pAddUnProcessed:=True    ''' V401 Start

            StatusBar moBataRpts.UploadSummary  ' could be attached to an event for manaul loadings
            StatusBar "BATA Uploads End"        ' could be attached to an event for manaul loadings
            Set moBataRpts = Nothing
        '   Nielsen reports run from Monday to Sunday. They are generated each Monday up to fdtmLastSunday
            CreateNielsenReports pLastReportEndDate:=fdtmLastSunday(), pCalledAutomatically:=True
            UploadLatestAztecRpts
        End If
        
        StatusBar "Data Capture Completed"
        
        If Not bSelFranCapture Then
            checkMissingDaysSales
        End If
        
        If gfFutureDate Then
            StatusBar "*** WARNING: FUTURE date(s) captured. Call System Administrator"
        End If
        
        If gfDateFormatBad Then
            StatusBar "*** WARNING: " & DATEFORMATBAD & ". Call System Administrator"
        End If
        
    '   Update Non Compliant Table when records are transferred into Live table
        If lngRecsTfrd2LiveTbl Then
            If oCaptureOptions.UpdateNonCompliants Then
                lNonCompliantRecCount = LoadNonCompliantTable(pFranCount:=lFranCount)
                StatusBar "Got " & lNonCompliantRecCount & " non-compliant sales from " & lFranCount & " franchises."
            End If
        End If
        
        If Not bSelFranCapture Then
        '   Automatically purge some of the old data
        ''' Review Start thinking about Purge routines for tblBataUploads & tblBataReUploads
            PurgeLiveData ' Not passed DaysToKepp because PurgeLiveData() routine is more readable without this param
            PurgeEventLog pMonthsToKeep:=g.rstDWDefaults!MonthsOfEventLog
            PurgeFranchiseUploads pMonthsToKeep:=g.rstDWDefaults!MonthsOfFranchiseUploads ' Could/should be sped up for MySQL
            PurgeNonCompliantPromos ' Days to keep is hardcoded in procedure
            
        '.  Hold off on call below. First test effects of ADO approach on uploading BataRpts
        '.  PurgeBataUploadLogs pMonthsToKeep:=12
        '
        
        '.  Should add something to purge RejectData table
        '
            
        End If
        
        If pAutoCaptureCycle Then
            OptimiseDb
        End If
        
        If Not gbEventLogRefreshIsEnabled Then
        '   Event Log Refresh is disabled for O/N cycle. Turn it back on and refresh display
            gbEventLogRefreshIsEnabled = True
            gsubRefreshEventLogDisplay
        End If
        
        subTabMainClick pTab:=tabMain.Tab   ' Refresh display for the current tab
        
        StatusBar "Ready"
        
        Me.Enabled = True
        g.bCaptureCycleRunning = False
        
    End If

Procedure_Exit:
    Exit Sub

RemoteProcessingError:
    lngErrHandlingAttemptsThisRemoteCnn = lngErrHandlingAttemptsThisRemoteCnn + 1
''' Version 346 Error Processing Routine moved to bottom of procedure
'''From Help: If an error occurs while an error handler is active
'''(between the occurrence of the error and a Resume, Exit Sub, Exit Function, or Exit Property statement),
'''the current procedure's error handler can't handle the error. I.E. 'On Error GoTo 0' ACHIEVES NOTHING HERE
' ************************************************************************************************************'
' ''' NOTE THAT IN ADDITION TO CLEANING UP REMOTE DATA OBJECTS MAY ALSO NEED TO CLEAN UP LOCAL DATA OBJECTS '''
' ************************************************************************************************************'
    StatusBar "REMOTE PROCESSING INTERRUPTED. HANDLING ATTEMPT " & _
              Bracket(CStr(lngErrHandlingAttemptsThisRemoteCnn)) & " " & fsErrDetail("subCaptureData"), strFranName
''' '   Argument for clearing RemoteDb DatabaseOpenedbyField is in case of error which didn't effect
''' '   the connection to the franchise db and remote rst sent us to label and we were still able to
''' '   clear the field. The problem with not clearing the field is the remote table will be locked,
''' '   thus preventing the remote module from operating (and it will fill the event log)
''' '   NOTE THAT WHEN CREATEING Version 340 WE HAD ONLY BEEN IN THIS CLAUSE 6 TIMES IN THE LAST 6 MONTHS
'''?'   Clesr RemoteDb DatabaseOpenedbyField
'''?    subIgnoreErr_ClearDbOpenedbyFld prstRemoteDefaults:=rstRemoteDefaults, pLogIfAvoidsBug:=True, pCalledFromTag:="SUBIGNOREERR_CLOSERST - CALLED WITHIN 'IF NOT BDATATRANSFERCOMPLETED THEN'"
'   If you try to close a Connection or Database object while it has any open Recordset objects,
'   the Recordset objects will be closed and any pending updates or edits will be cancelled
''' subIgnoreErr_CloseDbAndSetToNothing pDb:=dbRemote, pLogIfAvoidsBug:=True, pCalledFromTag:="subCaptureData"
    CloseCnnAndSetToNothing_IgnoreErrors pCnn:=cnnRemote, pLogIfAvoidsBug:=True, pCalledFromTag:="subCaptureData"
'   Resume Next handling added in ''' V378. Is an extremely rare situation and might later remove this code
    If lngErrHandlingAttemptsThisRemoteCnn <= 3 Then
    '   PERHAPHS THIS ONE COULD OFFER A FEW EXTRA SECONDS EACH TIME!
        Resume UnmapShareDisk
    Else
    '   Start working through the proc so we don't get caught in an endless loop and fill the event log
'   DO I WANT TO RESUME NEXT OR DO I SIMPLY WANT TO EXIT THIS FRANCHISE. I WILL HAVE ALREADY TRIED EXITING AT UnmapShareDisk: 3 TIMES
'   PERHAPS I COULD DO IT PROGRESSIVELY, AFTER THREE ATTEMPTS AT RESUMING AT THE USUAL LABEL (UnmapShareDisk:) I COULD THEN
'   JUST BAIL OUT AND RESUME ON THE NEXT FRAN. WOULD PERHAPS REQUIRE IT'S OWN CLEANUP.
'     ALTERNATELY COULD HAVE THREE RESUME NEXTS AND IN THE >= 6 THEN BAIL OUT WITH CLEAN UP MENTIONED ABOVE.
'   THE PROCEDURAL CODE SUBSEQUENT TO UnMapShareDisk COULD ACT CONDITIONALLY ACCORDING TO THE VLAUE OF lngErrHandlingAttemptsThisRemoteCnn
        Resume Next
    End If
    Resume  ' Not executed but assists when debugging in IDE
    
End Sub

Sub subClearFranchiseDetails()
    
'   Unselect combo boxes (Nb. cbo box text property is read-only)
    cboFranchiseType.ListIndex = -1
    cboDCTabPromoGrade.ListIndex = -1
    cboDCTabRegion.ListIndex = -1
    cboState.ListIndex = -1
    
    chkIncludeInDataCaptureCycle = vbUnchecked
    
    cmdSaveFranchiseDetails.Enabled = False
    
    txtPhysicalAddress = vbNullString
    txtContact = vbNullString
    txtSuburb = vbNullString
    txtAreaCode = vbNullString
    txtPhone = vbNullString
    txtModem = vbNullString
    txtNodename = vbNullString
    txtFaxNum = vbNullString
    txtRASUsername = vbNullString
    txtRASPassword = vbNullString
    lblTSGFranchiseID = vbNullString
    txtBATAFranchiseID = vbNullString
    lblRemoteModuleVersion = vbNullString
    lblPriceModuleVersion = vbNullString
    
End Sub

Sub subClearProductReportDisplay()

    With lvwProductReport
        .ListItems.Clear
        .Refresh
    End With
    
End Sub

Sub subClearStickReportDisplay()

    With lvwStickReport
        .ListItems.Clear
        .Refresh
    End With
    
End Sub

Sub subClearStockFields()

    gbClickEventIsSuppressed = True
    lstDescription.ListIndex = gconDoNotDisplayAnyItems
    txtBarcode.Text = vbNullString
    txtStkItemDescription.Text = vbNullString
    cboSupplier.ListIndex = gconDoNotDisplayAnyItems
    txtSticks.Text = vbNullString
    cboCategory.ListIndex = gconDoNotDisplayAnyItems
    cboSubCategory.ListIndex = gconDoNotDisplayAnyItems
    cboCtnContainingPkt.ListIndex = gconDoNotDisplayAnyItems
    txtCartonsPerPacket = vbNullString
    txtWholesaleListPrice = vbNullString
    txtRRP = vbNullString
    chkPackage = False
    cboSalesTax = vbNullString
    cboGoodsTax = vbNullString
    
    gbClickEventIsSuppressed = False

End Sub

Sub subDisplayCurrentRecordToUser(ByVal lCurrentProduct As Long, ByVal lTotalNumberOfProducts As Long)
    With stb
        .SimpleText = "Processing record " & lCurrentProduct & " of " & lTotalNumberOfProducts & Format(lCurrentProduct / lTotalNumberOfProducts, " (#0%)")
    End With
    Me.Refresh
End Sub

Sub subDisplayFranchiseDetails()
Dim lngListLoop As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rstFranTypes As ADODB.Recordset
Dim rstFran As ADODB.Recordset

    
    LockDCTabFranchiseCtls pLocked:=True
    
    strSQL = "SELECT * FROM Franchises " & vbNewLine & _
             "WHERE FranchiseBusinessName = " & SqlQ(lstDataCaptureFranchiseBusinessName)
    Set rstFran = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
         
    If (rstFran.BOF And rstFran.EOF) Then
        MsgBox "Franchise details were not found in the database. " & gconContactSystemAdministrator, vbCritical
    Else
        lblDisplayedFranchise.Caption = lstDataCaptureFranchiseBusinessName.Text
        txtPhysicalAddress = rstFran(gconFranchiseTablePhysicalAddressSuburbAndPostcodeField)
        txtSuburb = rstFran!FranchiseSuburb
        txtContact = rstFran(gconFranchiseTableContactNameField)
        txtAreaCode = rstFran(gconFranchiseTableAreaCodeField)
        txtPhone = rstFran(gconFranchiseTablePhoneField)
        cboState.Text = rstFran(gconFranchiseTableStateOfOz)
        txtModem = rstFran(gconFranchiseTableModemField)
        chkIncludeInDataCaptureCycle = BoolToChkBox(CBool(rstFran!FranchiseIncludedInStatistics))
        txtNodename = rstFran(gconFranchiseTableNodenameField)
        txtFaxNum = rstFran(gconFranchiseTableFaxField)
        txtRASUsername = rstFran!FranchiseRASUsername
        txtRASPassword = rstFran(gconFranchiseTableRASPasswordField)
        lblTSGFranchiseID = rstFran(gconFranchiseTableTSGFranchiseIDField)
        txtBATAFranchiseID = CnvNulls(rstFran!FranchiseIDBATA, vbNullString)
    '   Display Promo Grade remebering it is the Item Data which is stored and saved (cf Text/Description)
        cboDCTabPromoGrade.ListIndex = -1    ' First unselect combo box
        For lngListLoop = 0 To cboDCTabPromoGrade.ListCount - 1
            If cboDCTabPromoGrade.ItemData(lngListLoop) = rstFran!PromoGradeID Then
               cboDCTabPromoGrade.ListIndex = lngListLoop
            End If
        Next lngListLoop
        lblRemoteModuleVersion = rstFran(gconFranchiseTableRemoteModuleVersionField)
        lblPriceModuleVersion = rstFran!FranchisePriceModuleVersion

        txtVpnIpAddress = rstFran!VpnIpAddress
        
    '   Deselect combo then select appropriate item from prefilled combo
    '   Combo displays description displayed and stores data in .ItemData
        With cboDCTabRegion
            .ListIndex = -1
            For lngListLoop = 0 To .ListCount - 1
                If CnvNulls(rstFran!FranchiseRegionId, -255) = .ItemData(lngListLoop) Then
                    .ListIndex = lngListLoop
                    Exit For
                End If
            Next
        End With
        
    '   Deselect combo then select appropriate item from prefilled combo
    '   Combo displays description displayed and stores data in .ItemData
        cboFranchiseType.ListIndex = -1
        cboFranchiseType.ToolTipText = vbNullString
        Set rstFranTypes = GetRst(pCnn:=g.cnnDW, pSource:="FranTypes", pSourceType:=adCmdTable, pErrMsg:=strErrMsg)
        With cboFranchiseType
            For lngListLoop = 0 To .ListCount - 1
                If CnvNulls(rstFran!FranchiseType, -255) = .ItemData(lngListLoop) Then
                    .ListIndex = lngListLoop
                    .ToolTipText = m.astrFranTypeTooltip(lngListLoop)
                    Exit For
                End If
            Next
        End With
        rstFranTypes.Close
        Set rstFranTypes = Nothing
        cmdSaveFranchiseDetails.Enabled = False
    End If
    
    rstFran.Close
    Set rstFran = Nothing
    Me.Refresh
    
End Sub

Sub subDisplayStockItem()
Dim lngStkID As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rstStock As ADODB.Recordset
Dim rstPkg As ADODB.Recordset

    Const cUnknown = "UNKNOWN"
    
    gbClickEventIsSuppressed = True
    
    txtBarcode = ""
    cboSupplier.ListIndex = gconDoNotDisplayAnyItems
    txtSticks = ""
    cboCategory.ListIndex = gconDoNotDisplayAnyItems
    cboSubCategory.ListIndex = gconDoNotDisplayAnyItems
    
    lngStkID = lstDescription.ItemData(lstDescription.ListIndex)
    strSQL = "SELECT * FROM Stock WHERE Stock_ID = " & lngStkID
    
    Set rstStock = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
    If Not (rstStock.BOF And rstStock.EOF) Then
        txtStock_ID_DEV.Text = rstStock(gconStockTableStockIDField)
        txtBarcode = rstStock(gconStockTableBarcodeField)
        If rstStock(gconStockTableSupplierIDField) <> 0 Then
            cboSupplier = fsSupplierNameFrom(rstStock(gconStockTableSupplierIDField))
        Else
            cboSupplier.AddItem cUnknown
            cboSupplier = cUnknown
        End If
        txtStkItemDescription.Text = rstStock!Description
        txtSticks = rstStock(gconStockTableSticksField)
        cboCategory = rstStock(gconStockTableCategoryField)
        cboSubCategory = rstStock(gconStockTableSubCategoryField)
        txtRRP = Format(rstStock(gconStockTableSellField), "####0.00")                'PAL
        txtWholesaleListPrice = Format(rstStock(gconStockTableCostField), "####0.00") 'PAL
        If CBool(rstStock(gconStockTablePackageField)) Then
            chkPackage.Value = 1
            lblCtnContainingPkt.Visible = True
            cboCtnContainingPkt.Visible = True
            lblCartonsPerPacket.Visible = True
            txtCartonsPerPacket.Visible = True
            txtCartonsPerPacket.Enabled = True
        '   Presume there is only one component in the package b/c that's how TSG link CgtCtns and CgtPkts
            strSQL = "SELECT * FROM Package WHERE Package_ID = " & rstStock!Stock_ID
            Set rstPkg = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
            If Not (rstPkg.BOF And rstPkg.EOF) Then
                txtCartonsPerPacket = rstPkg!Quantity
                SetCboCtnContainingPkt pStkId:=Cn(rstPkg!Stock_ID, -666)  ' -666 is not a valid stock_id
            End If
            rstPkg.Close
            Set rstPkg = Nothing
        Else
            chkPackage.Value = 0
            lblCtnContainingPkt.Visible = False
            cboCtnContainingPkt.Visible = False
            lblCartonsPerPacket.Visible = False
            txtCartonsPerPacket.Visible = False
            cboCtnContainingPkt.ListIndex = -1
        End If
        cboGoodsTax = rstStock(gconStockTableGoodsTaxCodeField)
        cboSalesTax = rstStock(gconStockTableSalesTaxCodeField)
    End If
    rstStock.Close
    Set rstStock = Nothing
    
    cboCategory.Enabled = True
    cboSubCategory.Enabled = True
    chkPackage.Enabled = True
    
    cmdSaveStockDetails.Enabled = False
    gbClickEventIsSuppressed = False
    
End Sub

Sub subEnsurePathExistsCreateNewIterationOfExceptionReport( _
                                            ByVal sTargetFullPathAndFilename As String, _
                                            ByVal sTargetFolder As String)
    Const conPathNotFound = 76
    Dim intFileNum As Integer
  
    On Error GoTo ErrorHandler
    
    If Dir(sTargetFullPathAndFilename) <> "" Then
        SetAttr sTargetFullPathAndFilename, vbNormal
        Kill sTargetFullPathAndFilename
    End If
    
    intFileNum = FreeFile   ' Get unused file
    Open sTargetFullPathAndFilename For Output As #intFileNum
    Close #intFileNum
    
    Exit Sub
    
ErrorHandler:
    If Err.Number = conPathNotFound Then
        MkDir sTargetFolder
        Resume
    Else
        MsgBox "General path error - " & Err.Number, vbCritical
        End     ' AUrban Stumbled on an End statement which should be fixed.
                ' AUrban Perhaps better report error rather than simple msgbox

    End If
    
End Sub

Sub subPopulateFranchiseBusinessNameListBoxes()
' Only add live Franchises to the following list boxes
Dim rst As ADODB.Recordset
Dim strSQL As String
Dim strErrMsg As String

    lstDataCaptureFranchiseBusinessName.Clear
    lstStickReportsFranchiseBusinessName.Clear
    lstPRProductReportsFranchiseBusinessName.Clear
    lstUploadFranchiseList.Clear
    lstPromoFranchise.Clear
    
    PopulateLstProductReportsFranchiseBusinessName pIncludeClosedFrans:=ChkBoxToBool(chkSalesRptTab_IncludeClosedFrans)
    
    strSQL = "SELECT FranchiseIDTSG, FranchiseBusinessName FROM qryFranchiseLive ORDER BY FranchiseBusinessName"
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    gbClickEventIsSuppressed = True
    Do While Not rst.EOF
        lstStickReportsFranchiseBusinessName.AddItem rst!FranchiseBusinessName
        
'lstUploadFranchiseList is the list on the uploads tab and it should have Unicenta franchises excluded
'lstUploadFranchiseList is populated in subPopulateFranchiseBusinessNameListBoxes() which is called in few places
'[SetGlobalVariables,cmdAddNewFranchise_Click(), cmdCloseSelectedFranchises_Click] and executed a few times, so it
'doesn 't need to be particularly efficient and is better written for robustness and reuse in other parts of the program.
'
'given all that why don't I write a proc which collect NonOPOS frans and it can be passed a source so that source can be
'aligned whatever is required eg all frans, qryFranchiseLive, etc. and this can be used to populte things
'alternately I could wirte a function to return oPos Frans and apply this collection with the use of
'Filter function to other collection or my own code to filter out oPOS frans. I could at the smae time
'use it to identify the oPOS frans selected to provide a msg box informing user in situations like
'selected fran capture about which oPos frans they have selected and that they will not be
'used and why they will not be used. Another option is to colour code oPOS frans in the interface in some
'way to indicate that these frans are not available for particular functions. In the case of TsgDw for
'any communication b/w TegDw program and the franchise.



        lstUploadFranchiseList.AddItem rst!FranchiseBusinessName
        lstPRProductReportsFranchiseBusinessName.AddItem rst!FranchiseBusinessName
    '   NB Add Item Data (FranchiseID) for lstDataCaptureFranchiseBusinessName
        With lstDataCaptureFranchiseBusinessName
            lstDataCaptureFranchiseBusinessName.AddItem rst!FranchiseBusinessName
            lstDataCaptureFranchiseBusinessName.ItemData(.NewIndex) = rst!FranchiseIDTSG
        End With
        With lstPromoFranchise
            lstPromoFranchise.AddItem rst!FranchiseBusinessName
            lstPromoFranchise.ItemData(.NewIndex) = rst!FranchiseIDTSG
        End With
    
        rst.MoveNext
    Loop
    'force code to display first franchises details
    gbClickEventIsSuppressed = False
    lstDataCaptureFranchiseBusinessName.Refresh
    rst.Close
    Set rst = Nothing
    
End Sub

Sub subPopulateStickReportRecipientListbox()
Dim strSQL As String
Dim strErrMsg As String
Dim rstSnpSupplier As ADODB.Recordset

    With lstStickReportRecipient
        .Clear
        .Refresh
    End With
    
    gbClickEventIsSuppressed = True
    strSQL = "SELECT DISTINCT Supplier FROM Supplier " & vbNewLine & _
             "WHERE Supplier NOT LIKE " & SqlQ("<%")
    Set rstSnpSupplier = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rstSnpSupplier.BOF And rstSnpSupplier.EOF) Then
        Do Until rstSnpSupplier.EOF
            lstStickReportRecipient.AddItem rstSnpSupplier(gconSupplierTableSupplierNameField)
            rstSnpSupplier.MoveNext
        Loop
        lstStickReportRecipient.ListIndex = -1
    End If
    rstSnpSupplier.Close
    Set rstSnpSupplier = Nothing
    gbClickEventIsSuppressed = False
    
End Sub

Sub subPopulateStockListboxes(ByVal pExclDeletedStk As Boolean)
Dim strSQL As String
Dim strErrMsg As String
Dim strStockRecordSrc As String
Dim rstSupplier As ADODB.Recordset
Dim rstSnpCategory As ADODB.Recordset

    cboSupplier.Clear                       ' Stock Tab
    cboCategory.Clear                       ' Stock Tab
    cboSubCategory.Clear                    ' Stock Tab
    cboCtnContainingPkt.Clear               ' Stock Tab
    lstDescription.Clear                    ' Stock Tab
    lstStcokTabSelectedSoctkExport.Clear    ' Stock Tab
    Me.Refresh

    If pExclDeletedStk Then
        strStockRecordSrc = "qryStock"  ' ie Query on Stock table excluding deleted stock items
    Else
        strStockRecordSrc = "Stock"     ' ie Stock table
    End If
    
    'populate product list box
    strSQL = "SELECT Stock_ID, Description FROM " & strStockRecordSrc & vbNewLine & _
             "WHERE Description NOT LIKE " & SqlQ("<%") & vbNewLine & _
             "ORDER BY Description"
    LoadListBox_Rst pListBox:=lstDescription, pCnn:=g.cnnDW, pSource:=strSQL, pDisplayFld:="Description", pDataFld:="stock_id"
    lstDescription.Enabled = True
    
    strSQL = "SELECT Stock_ID, Description FROM qryStock " & vbNewLine & _
             "WHERE Description NOT LIKE " & SqlQ("<%") & vbNewLine & _
             "ORDER BY Description"
    LoadListBox_Rst pListBox:=lstStcokTabSelectedSoctkExport, pCnn:=g.cnnDW, pSource:=strSQL, pDisplayFld:="Description", pDataFld:="Stock_ID"

   'populate supplyar combo box
    gbClickEventIsSuppressed = True
    strSQL = "SELECT DISTINCT Supplier FROM Supplier " & vbNewLine & _
             "WHERE Supplier NOT LIKE " & SqlQ("<%")
    Set rstSupplier = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    Do Until rstSupplier.EOF
        cboSupplier.AddItem rstSupplier(gconSupplierTableSupplierNameField)
        rstSupplier.MoveNext
    Loop
    cboSupplier.ListIndex = gconDoNotDisplayAnyItems
    rstSupplier.Close
    Set rstSupplier = Nothing
    Me.Refresh
    
   'populate category combo box
    strSQL = "SELECT DISTINCT cat1 FROM " & strStockRecordSrc & vbNewLine & _
             "WHERE cat1 NOT LIKE " & SqlQ("<%") & vbNewLine & _
             "ORDER BY cat1"
    Set rstSnpCategory = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    Do Until rstSnpCategory.EOF
        cboCategory.AddItem rstSnpCategory(gconStockTableCategoryField)
        rstSnpCategory.MoveNext
    Loop
    cboCategory.ListIndex = gconDoNotDisplayAnyItems
    rstSnpCategory.Close
    
   'populate sub-category combo box
    strSQL = "SELECT DISTINCT cat2 FROM " & strStockRecordSrc & vbNewLine & _
             "WHERE cat2 NOT LIKE " & SqlQ("<%") & vbNewLine & _
             "ORDER BY cat2"
    Set rstSnpCategory = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    Do Until rstSnpCategory.EOF
        cboSubCategory.AddItem rstSnpCategory(gconStockTableSubCategoryField)
        rstSnpCategory.MoveNext
    Loop
    cboSubCategory.ListIndex = gconDoNotDisplayAnyItems
    rstSnpCategory.Close
    Set rstSnpCategory = Nothing

'   populate cboCtnOfPkt combo box - cbo text holds barcode and ItemData holds stock ID
    LoadCboCtnContainingPkt pRecordSource:=strStockRecordSrc
    
    Me.Refresh
    
    ' populate Tax combo boxes
    cboGoodsTax.Clear
    cboGoodsTax.AddItem "GST"
    cboGoodsTax.AddItem "GNR"
    cboGoodsTax.AddItem "FRE"
    cboSalesTax.Clear
    cboSalesTax.AddItem "GST"
    cboSalesTax.AddItem "GNR"
    cboSalesTax.AddItem "FRE"
    
    gbClickEventIsSuppressed = False 'enable click event
    
    cmdSaveStockDetails.Enabled = False
    Me.Refresh

End Sub

Sub subPopulateVersions()
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    Set rst = GetRst(pCnn:=g.cnnDW, _
                     pSource:="qryFranchiseVersions", _
                     pSourceType:=adCmdTable, _
                     pErrMsg:=strErrMsg)

    lvwVersions.ListItems.Clear

    Do Until rst.EOF
        Set gvListItem = frmTSGDataWarehouse.lvwVersions.ListItems.Add()
        gvListItem.Text = rst!FranchiseBusinessName
        gsubAddSubItemToListview rst!FranTypeName, 1
        gsubAddSubItemToListview rst!IncludedInCaptureCycle, 2
        gsubAddSubItemToListview rst!RStatsVersion, 3
        gsubAddSubItemToListview rst!PriceModuleVersion, 4
        gsubAddSubItemToListview rst!RMgrVersion, 5
        gsubAddSubItemToListview rst!FranchiseOSVersion, 6
        rst.MoveNext
    Loop

    rst.Close
    Set rst = Nothing
    
End Sub

Private Sub subPurgeNielsenReport(ByVal pReportStartDate As Date, ByVal pReportEndDate As Date)
Dim strNielsenRptFullname As String

    strNielsenRptFullname = fsNielsenRptFullname(pStartDate:=pReportStartDate, pEndDate:=pReportEndDate)
    
    cmdPurgeNielsenReportList.Enabled = False

    If FileExists(strNielsenRptFullname) Then
        SetAttr strNielsenRptFullname, vbNormal
        Kill strNielsenRptFullname
        subRefreshNielsenReportListBox pReportEndDate:=pReportEndDate
    End If

End Sub

Private Sub subRefreshAztecUploadsGrid()
Const kRows As Long = 20
Dim strSQL As String
Dim strErrMsg As String
Dim avnt() As Variant
Dim rst As ADODB.Recordset

    strSQL = "SELECT SalesDataEndDate, FileType, UploadDate " & vbNewLine & _
             "FROM tblAztecUploads " & vbNewLine & _
             "ORDER BY SalesDataEndDate DESC, FileType ASC, UploadDate DESC"
             
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    
    If Not rst Is Nothing Then
        If Not (rst.BOF And rst.EOF) Then
        '   GetRows
        '    - Gives an error when called on an empty rst:
        '      Either BOF or EOF is True, or the current record has been deleted. Requested operation requires a current record.
        '    - If you request more rows than are available, then GetRows returns only the number of available rows.
            avnt = rst.GetRows(kRows)
        End If
        rst.Close
        Set rst = Nothing
    
    '   Load grid with data
        grdAztecUploads.Redraw = flexRDNone                 ' Suspend redraw
        grdAztecUploads.Rows = grdAztecUploads.FixedRows    ' Clear grid
        If Not IsEmptyArray(avnt) Then
            grdAztecUploads.LoadArray avnt                  ' Load grid
        End If
        grdAztecUploads.AutoResize = True
        grdAztecUploads.Redraw = flexRDBuffered             ' Redraw grid
    End If
    
End Sub

Sub subRefreshNielsenReportListBox(ByVal pReportEndDate As Date)
Dim strTempFileList As String

    With lstNielsenReportDisplayDate
        .Clear
        .Refresh
    End With
    
    cmdPurgeNielsenReportList.Enabled = False
    
    'populate the listbox
    strTempFileList = Dir$(fsNielsenFileSpecification(pReportEndDate), vbDirectory)
    If Len(strTempFileList) > 0 Then
        Do Until strTempFileList = ""
            With lstNielsenReportDisplayDate
                .AddItem strTempFileList
                .Refresh
            End With
            strTempFileList = Trim(Dir$)
        Loop
    End If

    If lstNielsenReportDisplayDate.ListCount > 0 Then
        lstNielsenReportDisplayDate.ListIndex = gconDisplayFirstItem
        cmdPurgeNielsenReportList.Enabled = True
    Else
        lstNielsenReportDisplayDate.ListIndex = gconDoNotDisplayAnyItems
    End If


End Sub

Sub subResetStockForm()

    cmdAddNewStockItem.Caption = "< &Add New"
    With cmdSaveStockDetails
        .Enabled = False
        .Visible = True
    End With
    Call subClearStockFields
    'cboGoodsTax = "GST" ' default
    'cboSalesTax = "GST" ' default
    gbClickEventIsSuppressed = True
    With lstDescription
        .Enabled = True
        .ListIndex = -1
        .SetFocus
    End With

    cboCtnContainingPkt.Enabled = False

End Sub

Sub subSaveFranchiseDetails(ByVal bAddNewFranchise As Boolean)
Dim lngFranID As Long
Dim lngFranCount As Long
Dim lngDialSequence As Long
Dim lngPromoGradeID As Long
Dim strSQL As String
Dim strMsg As String
Dim strErrMsg As String
Dim strNewFranName As String
Dim rstFran As ADODB.Recordset

    cmdSaveFranchiseDetails.Enabled = False
''' MySQL REVIEW need for franchises having individual dial sequence that may trump state dial sequence
    lngDialSequence = GetDialSequenceFromState(cboState.Text)
    
    If bAddNewFranchise Then
        lngFranCount = GetRecordCount(pCnn:=g.cnnDW, pSource:="Franchises")
        If lngFranCount >= glMaximumFranchises Then ''' Review Check how & where glMaximumFranchises is used (Make note of gconReservedFranchiseID)
            MsgBox "Maximum number of franchises has been reached. Unable to add another. " & _
                    gconContactSystemAdministrator, vbCritical
            Exit Sub
        Else
            Set rstFran = GetRst(pCnn:=g.cnnDW, _
                                 pSource:="Franchises", _
                                 pSourceType:=adCmdTable, _
                                 pRstType:=eEditableFwdOnly, _
                                 pErrMsg:=strErrMsg)
            
            rstFran.AddNew
                strNewFranName = Left$(Trim$(txtNewFranchiseBusinessName), rstFran!FranchiseBusinessName.DefinedSize)
                rstFran!FranchiseBusinessName = strNewFranName
                rstFran!CaptureCycleDateOnLastDataCapture = Date ' For tracking whether data is being collected & whether it should be
                'rstFran!FranchiseType = 0     ' Unmatched -> no default selection but fits in with conditional code for this field
                rstFran!FranchiseMessageFlag = CBoolMySql(True)
                rstFran!FranchiseOSVersion = "unknown"
                rstFran!FranchiseRMVersion = "unknown"
                rstFran!VpnIpAddress = txtVpnIpAddress.Text
                rstFran!Live = CBoolMySql(True) ' Default value when adding a franchise is Live
            ''' "Franchise added to database", txtNewFranchiseBusinessName   ''' V397
                StatusBar pMsg:="Franchise added to database", _
                          pFranchise:=txtNewFranchiseBusinessName, _
                          pRefreshEventLogDisplay:=False                                            ''' V397
                
        End If
    Else 'edit existing record
        strSQL = "SELECT * FROM Franchises " & vbNewLine & _
                 "WHERE FranchiseBusinessName = " & SqlQ(lstDataCaptureFranchiseBusinessName)
        Set rstFran = GetRst(pCnn:=g.cnnDW, _
                             pSource:=strSQL, _
                             pSourceType:=adCmdText, _
                             pRstType:=eEditableFwdOnly, _
                             pErrMsg:=strErrMsg)
                             
        If Not (rstFran.BOF And rstFran.EOF) Then
            If cboDCTabPromoGrade.ListIndex > -1 Then
                lngPromoGradeID = cboDCTabPromoGrade.ItemData(cboDCTabPromoGrade.ListIndex)
                If rstFran!PromoGradeID <> lngPromoGradeID Then
                    strMsg = "You have changed the promotion grade." & vbNewLine & _
                             "Any promotions already uploaded for this franchise remain active." & vbNewLine & _
                             "Do you wish to continue?"
                    ''' Review 'If MsgBox(strMsg, vbQuestion + vbExclamation + vbYesNo) <> vbYes Then
                    ''' VERY CLUMSY - COULD DO WITH A BETTER SOLUTION
                    ''' B/C OF HOW AND WHERE THIS IS CALLED ALL THE OLD DETAILS ARE REFRESHED ONTO THE SCREEN
                    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2) <> vbYes Then
                        rstFran.Close
                        Set rstFran = Nothing
                        Exit Sub
                    End If
                End If
            End If
        Else
            MsgBox "Existing franchises record was not located in the database. Unable to apply the changes. " & _
                    gconContactSystemAdministrator, vbCritical
            rstFran.Close
            Set rstFran = Nothing
            Exit Sub
        End If
    End If
    
    rstFran!FranchisePhysicalAddressSuburbAndPostcode = Left(Trim(txtPhysicalAddress), rstFran!FranchisePhysicalAddressSuburbAndPostcode.DefinedSize)
    rstFran!FranchiseSuburb = Left(Trim(txtSuburb), rstFran!FranchiseSuburb.DefinedSize)
    rstFran!FranchiseContactName = Left(Trim(txtContact), rstFran!FranchiseContactName.DefinedSize)
    rstFran!FranchiseAreaCode = Left(Trim(txtAreaCode), rstFran!FranchiseAreaCode.DefinedSize)
    rstFran!FranchisePhone = Left(Trim(txtPhone), rstFran!FranchisePhone.DefinedSize)
    rstFran!FranchiseStateOfOz = cboState.Text
    rstFran!FranchiseDialSequence = lngDialSequence
    rstFran!FranchiseModem = Left(Trim(txtModem), rstFran!FranchiseModem.DefinedSize)
    rstFran!FranchiseIncludedInStatistics = CBoolMySql(ChkBoxToBool(chkIncludeInDataCaptureCycle))
    rstFran!FranchiseNodename = LCase(Left(Trim(txtNodename), rstFran!FranchiseNodename.DefinedSize))
    rstFran!FranchiseFax = LCase(Left(Trim(txtFaxNum), rstFran!FranchiseFax.DefinedSize))
    rstFran!FranchiseRASUsername = LCase(Left(Trim(txtRASUsername), rstFran!FranchiseRASUsername.DefinedSize))
    rstFran!FranchiseRASPassword = LCase(Left(Trim(txtRASPassword), rstFran!FranchiseRASPassword.DefinedSize))
    rstFran!PromoGradeID = cboDCTabPromoGrade.ItemData(cboDCTabPromoGrade.ListIndex)
    rstFran!FranchiseIDBATA = CnvZerosToNull(Val(txtBATAFranchiseID))
    rstFran!VpnIpAddress = txtVpnIpAddress.Text
    
    With cboDCTabRegion
        If .ListIndex > -1 Then ' Item is selected
            rstFran!FranchiseRegionId = .ItemData(.ListIndex)
        End If
    End With
    
    If cboFranchiseType.ListIndex > -1 Then ' Item is selected
        rstFran!FranchiseType = cboFranchiseType.ItemData(cboFranchiseType.ListIndex)
    End If
    
    rstFran.Update
    
    rstFran.Close
    Set rstFran = Nothing
    
'   Auto-increment field must be retrieved from MySQL,
'   but is immediately available from rst when using mdb file
    strSQL = "Select Max(FranchiseIDTSG) FROM Franchises"
    lngFranID = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL)
    
    If bAddNewFranchise Then
        strMsg = strNewFranName & " was added to the database" & vbNewLine & _
                "Please pass the following details on to the Supplier (eg. TechRentals) of the new machine:" & vbNewLine & vbNewLine & _
                "FranchiseID " & vbTab & vbTab & vbTab & ": " & lngFranID & vbNewLine & _
                "Computer NodeID " & vbTab & vbTab & vbTab & ": " & txtNodename & vbNewLine & _
                "RAS Password " & vbTab & vbTab & vbTab & ": " & txtRASPassword & vbNewLine & _
                "RetailManager shopfront name " & vbTab & ": Tobacco Station - " & strNewFranName
        MsgBox strMsg, vbInformation
    End If
    
End Sub

Function subSaveStockDetails(ByVal pAddNewItem As Boolean) As Double
'   Returns [new] Stock_ID value
Dim lngStkID As Long
Dim lngPkgStkId As Long
Dim strSQL As String
Dim strErrMsg As String
Dim rstStock As ADODB.Recordset
Dim rstPkg As ADODB.Recordset

    cmdSaveStockDetails.Enabled = False
    Me.Enabled = False
    
    If pAddNewItem Then
        'determine next available stock ID

        strSQL = "Select Max(stock_id) FROM Stock"
        lngStkID = GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=0) + 1
        
        Set rstStock = GetRstAddOnly(pCnn:=g.cnnDW, pSource:="Stock", pErrMsg:=strErrMsg)
        
        rstStock.AddNew
            rstStock(gconStockTableStockIDField) = lngStkID
            rstStock(gconStockTableCustom1Field) = vbNullString
            rstStock(gconStockTableCustom2Field) = vbNullString
            rstStock(gconStockTableSalesPromptField) = vbNullString
            rstStock(gconStockTableLongDescriptionField) = vbNullString
    Else 'edit existing item
        lngStkID = lstDescription.ItemData(lstDescription.ListIndex)
        strSQL = "SELECT * FROM Stock WHERE stock_id = " & lngStkID
        Set rstStock = GetRst(pCnn:=g.cnnDW, _
                              pSource:=strSQL, _
                              pSourceType:=adCmdText, _
                              pRstType:=eEditableFwdOnly, _
                              pErrMsg:=strErrMsg)
    End If
    rstStock(gconStockTableBarcodeField) = Left(txtBarcode, gconStockTableBarcodeFieldWidth)
    rstStock(gconStockTableSupplierIDField) = flSupplierIDFrom(cboSupplier)
    rstStock!Description = Left$(Trim$(txtStkItemDescription), rstStock!Description.DefinedSize)
    rstStock(gconStockTableSticksField) = Val(txtSticks)
    rstStock(gconStockTableCategoryField) = cboCategory
    rstStock(gconStockTableSubCategoryField) = cboSubCategory
    rstStock(gconStockTableSellField) = Val(Format(txtRRP, "####0.00"))                    'PAL
    rstStock(gconStockTableCostField) = Val(Format(txtWholesaleListPrice, "####0.00"))     'PAL
    rstStock(gconStockTableGoodsTaxCodeField) = cboGoodsTax     'PAL
    rstStock(gconStockTableSalesTaxCodeField) = cboSalesTax     'PAL
    If chkPackage.Value = 0 Then
        rstStock(gconStockTablePackageField) = CBoolMySql(False)
        rstStock!tax_components = CBoolMySql(False)
        rstStock!allow_fractions = CBoolMySql(True)
    Else
        If cboCategory = TsgShared.gkCAT_CigPkt Then
            rstStock(gconStockTablePackageField) = CBoolMySql(True)     'PAL
            rstStock(gconStockTableTaxComponentsField) = CBoolMySql(True)
            rstStock(gconStockTableAllowFractionsField) = CBoolMySql(False)
            rstStock(gconStockTableSellField) = 0
            rstStock(gconStockTableCostField) = 0
            If cboCategory = gkCAT_CigPkt Then
                If cboCtnContainingPkt.ListIndex > -1 Then
                    strSQL = "SELECT * FROM Package WHERE Package_ID = " & lngStkID
                    Set rstPkg = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pRstType:=eEditableFwdOnly, pErrMsg:=strErrMsg)
                    If (rstPkg.BOF And rstPkg.EOF) Then
                        rstPkg.AddNew
                        rstPkg!Package_Id = lngStkID
                    End If
                    lngPkgStkId = cboCtnContainingPkt.ItemData(cboCtnContainingPkt.ListIndex)
                    rstPkg!Stock_ID = lngPkgStkId
                    rstPkg!Sell_Inc = GetStkValue(pFldName:="Sell", pStkId:=lngPkgStkId) * (1 + gkGstRare)
                    If Val(txtSticks) > 0 Then
                        With cboCtnContainingPkt
                            If IsNumeric(txtCartonsPerPacket) Then
                                rstPkg!Quantity = txtCartonsPerPacket
                            End If
                        End With
                    End If
                    rstPkg.Update
                End If
            End If
        End If
    End If

    ' Add this stock item to either the NewStock file or the WLPUpdates file
    If chkUploadWholesaleListPrice Then
        AddToUpdateFile pRstStock:=rstStock, _
                        pRstDbType:=eMySqlDb, _
                        pUseRecordValues:=False, _
                        bAddNewItem:=pAddNewItem
    End If

    rstStock.Update
    rstStock.Close
    Set rstStock = Nothing
    Me.Enabled = True
    subSaveStockDetails = lngStkID

End Function

Sub subSetProductReportDateWording()

    If DateDiff("d", gfsSplitDate(lblProductReportStartDate), gfsSplitDate(lblProductReportFinishDate)) > gconZeroValue Then
        gsReportPeriodWording = " from "
    Else
        gsReportPeriodWording = " on "
    End If

End Sub

Sub subSetProductReportViewButton()
    If Dir(gsProductReportPathAndFilename) <> "" Then
        cmdViewProductReport.Enabled = True
    Else
        cmdViewProductReport.Enabled = False
    End If
End Sub

Sub subSetPRProductReportDateWording()

    If DateDiff("d", gfsSplitDate(lblPRProductReportStartDate), gfsSplitDate(lblPRProductReportFinishDate)) > gconZeroValue Then
        gsReportPeriodWording = " from "
    Else
        gsReportPeriodWording = " on "
    End If

End Sub

Sub subSetStickReportDateWording()

    If DateDiff("d", gfsSplitDate(lblStickReportStartDate), gfsSplitDate(lblStickReportFinishDate)) > gconZeroValue Then
        gsReportPeriodWording = " from "
    Else
        gsReportPeriodWording = " on "
    End If

End Sub

Sub subSetStickReportViewButton()
    If Dir(gsStickReportPathAndFilename) <> "" Then
        cmdViewStickReport.Enabled = True
    Else
        cmdViewStickReport.Enabled = False
    End If
End Sub

Private Sub subTabMainClick(pTab As Long)
Dim strSQL As String

'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    Me.Enabled = False
'
'    If Not GetTabRefreshedFlag(tabOrders.Tab) Then
'    '   Data needs to be refreshed (ie Refresh flag is False)
'        RefreshCurrentTab
'    Else
'    '   Data need not be refreshed, but form level buttons
'    End If
'
'    Me.Enabled = True
'    Me.SetFocus     ' Setfocus required for case where frmWait has been shown.
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    stb.SimpleText = vbNullString
    
    DoEvents    ' Assist screen refreshing
    
    Select Case pTab
        Case eDataCaptureTab
        Case eSalesRptsTab
            ' disable the printer option if there is no network printer
            optSendProductReportToPrinter.Enabled = g.rstAppDefaults!NetworkPrinterEnabled
            lblProductReportStartDate = fsYesterdaysDate
            lblProductReportFinishDate = fsYesterdaysDate
            gsReportPeriodWording = " on "
            cmdTopSellers.Caption = "&Top " & giTopSellers
            subSetProductReportViewButton
            With lvwProductReport
                .ListItems.Clear
                .Refresh
            End With
        Case eStickRptsTab
            ' disable the printer option if there is no network printer
            optSendStickReportToPrinter.Enabled = g.rstAppDefaults!NetworkPrinterEnabled
            lblStickReportStartDate = fsYesterdaysDate
            lblStickReportFinishDate = fsYesterdaysDate
            gsReportPeriodWording = " on "
            subPopulateStickReportRecipientListbox
            cmdStickReport.Enabled = False
            subSetStickReportViewButton
            With lvwStickReport
                .ListItems.Clear
                .Refresh
            End With
        Case eStockTab
            subResetStockForm
            Me.Refresh
            subPopulateStockListboxes pExclDeletedStk:=ChkBoxToBool(chkStockTab_IncludeDeletedStock)
        Case eBataTab
            If Not GetTabRefreshedFlag(pTab) Then
                RefreshBataTabGrid
            End If
        Case eNielsenTab
            subRefreshNielsenReportListBox dtpNielsenRptTxDate.Value
            subRefreshAztecUploadsGrid
        Case eVersionsTab
            subPopulateVersions
        Case eProductRptsTab
            ' disable the printer option if there is no network printer
            optPRSendProductReportToPrinter.Enabled = g.rstAppDefaults!NetworkPrinterEnabled
            lblPRProductReportStartDate = fsYesterdaysDate
            lblPRProductReportFinishDate = fsYesterdaysDate
            gsReportPeriodWording = " on "
            cmdTopSellers.Caption = "&Top " & giTopSellers
            subSetProductReportViewButton
            With lvwProductReport
                .ListItems.Clear
                .Refresh
            End With
        '   Populate product list box on Product Report Tab
            strSQL = "SELECT DISTINCT Description FROM qryStock " & vbNewLine & _
                     "WHERE Description NOT LIKE " & SqlQ("<%")
            LoadListBox_Rst pListBox:=lstPRProductList, pCnn:=g.cnnDW, pSource:=strSQL, pDisplayFld:="Description"
        Case eSettingsTab
            LoadSettingsTab
        Case eUploadsTab
            LoadUploadTab
        Case ePromotionsTab
            LoadPromotionTab
    End Select
    
    DoEvents    ' Assist screen refreshing
    
End Sub

Sub subUpdateFranDialupResult(ByVal pTsgFranID As Long, ByVal sResult As String)
'   Could and shoud be replaced with SQL
Dim strSQL As String
Dim strErrMsg As String
Dim rstFranDetails As ADODB.Recordset
    
    strSQL = "SELECT FranchiseIDTSG, FranchiseLastDialupResult FROM Franchises " & _
             "WHERE FranchiseIDTSG  = " & pTsgFranID
             
    Set rstFranDetails = GetRst(pCnn:=g.cnnDW, _
                                pSource:=strSQL, _
                                pSourceType:=adCmdText, _
                                pRstType:=eEditableFwdOnly, _
                                pErrMsg:=strErrMsg)
    
    With rstFranDetails
        If Not (rstFranDetails.BOF And rstFranDetails.EOF) Then
            .Fields!FranchiseLastDialupResult = Left$(sResult, .Fields!FranchiseLastDialupResult.DefinedSize)
            .Update
        End If
        .Close
    End With
    
    Set rstFranDetails = Nothing
    
End Sub

Sub subWriteCollatingMessageToStatusBar()
    With stb
        .SimpleText = "Collating sales data, please wait...."
        .Refresh
    End With
End Sub

Sub subWriteSearchingMessageToStatusBar()
    With stb
        .SimpleText = "Searching database, please wait...."
        .Refresh
    End With
End Sub

Sub subWriteSizingArraysMessageToStatusBar()

    With stb
        .SimpleText = "Sizing arrays...."
        .Refresh
    End With

End Sub

Sub subWriteSortingMessageToStatusBar()
    With stb
        .SimpleText = "Sorting sales data, please wait...."
        .Refresh
    End With
End Sub

Private Sub tabMain_Click(PreviousTab As Integer)
    subTabMainClick pTab:=tabMain.Tab
End Sub

Private Sub tdpEventLogDate_AfterEdit()
    Call gsubRefreshEventLogDisplay
    HighlightSelectedFranchiseInEventLog
End Sub

Function TfrAllPreLiveDataToLiveData() As Long

' RENAME TO TfrAllPreLiveDataToLiveData()

'   Returns the number of transerred records
' AUrban Could do with some investigation. On one day when the Archive DB had grown too large and
' the program fell over there were 53,987 duplicate records which were added to the duplicate table.
' Finding out why there were/are so many duplicates will go some way toward speeding things up.
Const kRawFldList As String = " FranchiseIDTSG, " & vbNewLine & " Barcode, " & vbNewLine & _
                              " TransactionDate, " & vbNewLine & " Quantity, " & vbNewLine & _
                              " TotalInc, " & vbNewLine & " NormalSellInc, " & vbNewLine & _
                              " CostInc, " & vbNewLine & " WholesaleQty, " & vbNewLine & " WholesaleActualSell"

Const kRawFldListWithTimeStamp As String = kRawFldList & ", " & vbNewLine & " TodaysDate"
Const kFldList As String = "(" & kRawFldList & ")"
Const kFldListWithDateStamp = "(" & kRawFldListWithTimeStamp & ")"
Const kSqlStmtSep As String = "; " & vbNewLine
Const kValListSep As String = ", " & vbNewLine

Dim bDuplicateDataCheck As Boolean
Dim bLiveDuplicate As Boolean
Dim bArchiveDuplicate As Boolean
Dim intPrevMousePointer As Integer
Dim lngDuplicateCnt As Long
Dim lngLoop As Long
Dim lngPreLiveCnt As Long
Dim lngTransferCnt As Long
Dim lngArchiveCnt As Long
Dim lngRejectCnt As Long
Dim lngPrevFranID As Long
Dim lngProcessedCnt As Long
Dim lngQtyThreshold As Long
Dim curValueThreshold As Currency
Dim datToday As Date
Dim strSQL As String
Dim strTransferMethod As String
Dim strErrMsg As String
Dim strValList As String
Dim strLiveValList As String
Dim strArchiveValList As String
Dim strRejectValList As String
Dim strDuplicateValList As String
Dim strDupChkVals As String
Dim strLiveDupChkString As String
Dim strArchiveDupChkString As String
Dim vntFldArray As Variant
Dim colFldVals As VBA.Collection
Dim colFldNames As VBA.Collection
Dim rstPreLive As ADODB.Recordset
    
    intPrevMousePointer = SetMousePointer(vbHourglass)
    datToday = Date
    lngQtyThreshold = g.rstDWDefaults!MaxQtyValue
    curValueThreshold = g.rstDWDefaults!MaxCurrencyValue
    
    lngPreLiveCnt = GetRecordCount(pCnn:=g.cnnDW, pSource:="PreLiveData")
    If lngPreLiveCnt = 0 Then
        StatusBar "No pre-live data records to transfer"
    Else
    'a franchise may have restored an old remote database following
    'HDD failure or other probelm, so don't let duplicate data as a
    'result get into the live table, ergo, more integrity checking.
        bDuplicateDataCheck = CBool(g.rstDWDefaults!DuplicatDataIntegrityCheck)
        If bDuplicateDataCheck Then
            strTransferMethod = "with duplicate data integrity check"
        Else
            strTransferMethod = "direct transfer"
        End If
        StatusBar "Commenced transfer of pre-live data to live data table (" & strTransferMethod & ")"
        strSQL = "SELECT " & vbNewLine & kRawFldList & vbNewLine & "FROM PreLiveData;"
        Set rstPreLive = GetRst(pCnn:=g.cnnDW, _
                                    pSource:=strSQL, _
                                    pSourceType:=adCmdText, _
                                    pCursorLocn:=adUseClient, _
                                    pErrMsg:=strErrMsg)
        If Not (rstPreLive.BOF And rstPreLive.EOF) Then
        '   Populate collection of fld names keyed on ordinal position in source recordset
        '   VBA.Ccollection object is 1 based, whereas rst.fields collection is 0 based
            Set colFldNames = New VBA.Collection
            For lngLoop = 0 To rstPreLive.Fields.Count - 1
                colFldNames.Add Item:=rstPreLive.Fields(lngLoop).Name, Key:=CStr(lngLoop)
            Next lngLoop
        End If
        
        Do Until rstPreLive.EOF
        '   After you call GetRows, the next unread record becomes the current record,
        '   or the EOF property is set to True if there are no more records.
            vntFldArray = rstPreLive.GetRows(Rows:=1)
            
        '   Populate collection of field values with field name as key
            Set colFldVals = New VBA.Collection
            For lngLoop = LBound(vntFldArray, 1) To UBound(vntFldArray, 1)
                colFldVals.Add Item:=vntFldArray(lngLoop, 0), Key:=colFldNames(lngLoop + 1) ' or CStr(lngLoop) for key
            Next lngLoop

        '   Check if amount, qty and barcode values are reasonable (barcode cannot have an embedded single quote)
            If Not IsValidData(pColFldVals:=colFldVals, pMaxValue:=curValueThreshold, pMaxQty:=lngQtyThreshold) Then
            '   Move record to reject table, but only if it is a tobacco product
                If fbBarcodeIsATobaccoProduct(colFldVals("Barcode")) Then
                    strRejectValList = strRejectValList & GetValList(colFldVals, pAddDateStamp:=True) & kValListSep
                End If
                lngRejectCnt = lngRejectCnt + 1
            Else

            '   Data with barcodes > 15 chars rejected and moved to the RejectData table.
            '   Barcode field in livedata table is 15 chars long. PosLiveData barcode field increased to
            '   20 chars (presumably by 7th Beam to stop Unicenta uploads crashing - Unicenta uploads
            '   sporadically provide data with barcdes > 15 chars). To cater for barcodes > 15 chars the
            '   Prelive barcode field was increased 20 chars (from 15) and validation is centralised here
            '   rather than rejecting long barcodes when transferring from PosLive to Prelive.
            '   Although barcodes we are interested in have a max length of 13 we must allow for a length
            '   of 14 for the special case barcode of "TOTALCUSTOMERS" which is used by the TSG S/W suite.
                If Len(colFldVals("Barcode")) > 15 Then
                    strRejectValList = strRejectValList & GetValList(colFldVals, pAddDateStamp:=True) & kValListSep
                    lngRejectCnt = lngRejectCnt + 1

            ' Now check the date of the transaction. If its greater than today, the franchises computer date
            ' is incorrect so reject the transactions and flag this problem
                ElseIf colFldVals("TransactionDate") > datToday Then
                    gfFutureDate = True
                    If colFldVals("FranchiseIDTSG") <> lngPrevFranID Then
                        StatusBar "ERROR: **** Future date - Contact System Administrator", GetFranName(colFldVals("FranchiseIDTSG"))
                    End If
                ' record is unreasonable so move it to the reject table
                    strRejectValList = strRejectValList & GetValList(colFldVals, pAddDateStamp:=True) & kValListSep
                    lngRejectCnt = lngRejectCnt + 1
                    
                ElseIf colFldVals("Barcode") = DATEFORMATBAD Then
                    gfDateFormatBad = True
                    StatusBar "ERROR: *** " & DATEFORMATBAD & " - Contact System Administrator", _
                               GetFranName(colFldVals("FranchiseIDTSG"))
                    subUpdateFranDialupResult colFldVals("FranchiseIDTSG"), "Failed - " & DATEFORMATBAD

                Else
                    If Not bDuplicateDataCheck Then
                        strValList = GetValList(colFldVals, pAddDateStamp:=False)
                        strLiveValList = strLiveValList & strValList & kValListSep
                        lngTransferCnt = lngTransferCnt + 1
                        strArchiveValList = strArchiveValList & strValList & kValListSep
                        lngArchiveCnt = lngArchiveCnt + 1
                    Else
                    '.  Check whether the record already exists (pending writing or in the database)
                    '
                    '   Check if Data[combo] is in ValLists pending writing to database
                    '   (NB current limitation excludes same Id,bc,Td combo with diff price )
                    '   (where price change may have occurred in the middle of the day      )
                        bLiveDuplicate = True
                        bArchiveDuplicate = True
                        strDupChkVals = UCase$("ID " & colFldVals("FranchiseIDTSG") & _
                                              " BC " & colFldVals("Barcode") & _
                                              " TD " & Format$(colFldVals("TransactionDate"), "ddmmmyyyyhhnnss"))
                    '   LiveData Table
                    '   Start argument is required for InStr() if compare parameter is specified
                    '   (specifying Start:=0 crashed the program)
                        If InStr(1, strLiveDupChkString, strDupChkVals, vbBinaryCompare) = 0 Then
                        '   Data[combo] not pending writing
                        '   Add to LiveDupChkString & check if record combo is in database
                            strLiveDupChkString = strLiveDupChkString & strDupChkVals
                        '   MySqlDateTime() used in Where Clause even though all times should be 00:00:00
                            strSQL = "SELECT Count(*) FROM LiveData " & vbNewLine & _
                                     "WHERE (FranchiseIDTSG = " & colFldVals("FranchiseIDTSG") & ") AND " & _
                                           "(Barcode = " & SqlQ(colFldVals("Barcode")) & ") AND " & _
                                           "(TransactionDate = " & MySqlDateTime(colFldVals("TransactionDate")) & ")"
                            If GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=0) = 0 Then
                                bLiveDuplicate = False
                            End If
                        End If

                        If bLiveDuplicate Then
                        '   Add to the 'duplicates' table for the time being
                            strDuplicateValList = strDuplicateValList & GetValList(colFldVals, pAddDateStamp:=True) & kValListSep
                            lngDuplicateCnt = lngDuplicateCnt + 1
                        Else
                            strLiveValList = strLiveValList & GetValList(colFldVals, pAddDateStamp:=False) & kValListSep
                            lngTransferCnt = lngTransferCnt + 1
                        End If
                    
                    '   LiveDataArchive Table
                        If InStr(1, strArchiveDupChkString, strDupChkVals, vbBinaryCompare) = 0 Then
                        '   Data[combo] not pending writing
                        '   Add to ArchiveDupChkString & check if record combo is in database
                            strArchiveDupChkString = strArchiveDupChkString & strDupChkVals
                        '   MySqlDateTime() used in Where Clause even though all times should be 00:00:00
                            strSQL = "SELECT Count(*) FROM LiveDataArchive " & vbNewLine & _
                                     "WHERE (FranchiseIDTSG = " & colFldVals("FranchiseIDTSG") & ") AND " & _
                                           "(Barcode = " & SqlQ(colFldVals("Barcode")) & ") AND " & _
                                           "(TransactionDate = " & MySqlDateTime(colFldVals("TransactionDate")) & ")"
                            If GetRstVal(pCnn:=g.cnnDW, pSource:=strSQL, pDefaultVal:=0) = 0 Then
                                bArchiveDuplicate = False
                            End If
                        End If
                        
                        If Not bArchiveDuplicate Then
                            strArchiveValList = strArchiveValList & GetValList(colFldVals, pAddDateStamp:=False) & kValListSep
                            lngArchiveCnt = lngArchiveCnt + 1
                        End If
                        
                    End If
                End If
            End If
            
            lngPrevFranID = colFldVals("FranchiseIDTSG")
            
            lngProcessedCnt = lngProcessedCnt + 1

            If lngProcessedCnt Mod 100 = 0 Then   ' Status Message every 10 records
        '   If lngTransferCnt Mod 100 = 0 Then
                StatusBar lngProcessedCnt & " of " & lngPreLiveCnt & " records processed. " & _
                           Plural(lngTransferCnt, "record") & " transferred.", pLog:=False
                GoSub UpdateRecords
            End If

        Loop

    '   Update records from remaining unexecuted value lists
        GoSub UpdateRecords

'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
' ***** SetTableUpdateTime  ***** ' BEST REPLACED BY A TRIGGER THAT WILL REMOVE PROBLEM OF SUBTLE BUGS
' ***** MAYBE AN EXPENSIVE OPERATION - but in this instance is excueted outside the loop
        If lngTransferCnt Then
            SetTableUpdateTime pTableName:="LiveData", pTimeStamp:=Now  ''' Review
        End If
        
        StatusBar Plural(lngTransferCnt, "record") & " of " & lngPreLiveCnt & " transferred to live data table"
        StatusBar Plural(lngDuplicateCnt, "record") & " of " & lngPreLiveCnt & " transferred to duplicates data table"
        StatusBar Plural(lngArchiveCnt, "record") & " of " & lngPreLiveCnt & " transferred to archive table"
        StatusBar Plural(lngRejectCnt, "record") & " rejected for: (Total Qty = 0 and WS Qty = 0) or " & _
                  "(Qty > " & lngQtyThreshold & ") or " & _
                  "(Currency > " & Format$(curValueThreshold, "$#,##0") & ") or " & _
                  "Single Quote (') in Barcode field or " & _
                  "Barcode longer than 15 chars or " & _
                  "Transaction date in future"
        rstPreLive.Close:     Set rstPreLive = Nothing
        
        subPurgeTable pTableName:="prelivedata"
    
    End If 'pre-live table contains records


Procedure_Exit:
    TfrAllPreLiveDataToLiveData = lngTransferCnt
    SetMousePointer intPrevMousePointer
    Exit Function

UpdateRecords:
    strSQL = vbNullString

    If Len(strLiveValList) Then
        strLiveValList = Left$(strLiveValList, Len(strLiveValList) - Len(kValListSep))
        strSQL = strSQL & "INSERT INTO LiveData " & vbNewLine & kFldList & vbNewLine & _
                          "VALUES " & vbNewLine & strLiveValList & kSqlStmtSep
        strLiveValList = vbNullString
        strLiveDupChkString = vbNullString
    End If

    If Len(strArchiveValList) Then
        strArchiveValList = Left$(strArchiveValList, Len(strArchiveValList) - Len(kValListSep))
        strSQL = strSQL & "INSERT INTO LiveDataArchive " & vbNewLine & kFldList & vbNewLine & _
                          "VALUES " & vbNewLine & strArchiveValList & kSqlStmtSep
        strArchiveValList = vbNullString
        strArchiveDupChkString = vbNullString
    End If

    If Len(strRejectValList) Then
        strRejectValList = Left$(strRejectValList, Len(strRejectValList) - Len(kValListSep))
        strSQL = strSQL & "INSERT INTO RejectData " & vbNewLine & kFldListWithDateStamp & vbNewLine & _
                          "VALUES " & vbNewLine & strRejectValList & kSqlStmtSep
        strRejectValList = vbNullString
    End If

    If Len(strDuplicateValList) Then
        strDuplicateValList = Left$(strDuplicateValList, Len(strDuplicateValList) - Len(kValListSep))
        strSQL = strSQL & "INSERT INTO Duplicates " & vbNewLine & kFldListWithDateStamp & vbNewLine & _
                          "VALUES " & vbNewLine & strDuplicateValList & kSqlStmtSep
        strDuplicateValList = vbNullString
    End If
    
'   Invalid data that is not a toboacoo product is ignored hence
'   there may be cases we get here without any records to update/ValList string
    If Len(strSQL) Then
    ''' On Error GoTo Procedure_Error
    '   Do/should I use transactions?
        CnnDwExecute strSQL
        strSQL = vbNullString
    ''' On Error GoTo 0
    End If
    
    Return
    
End Function
Private Sub tmrAutoDataCapture_Timer()
' *****************************************************************************
' * tmrAutoDataCapture enabled for installations where g.bMasterMaster = True *
' * tmrAutoDataCapture may later be renamed and enabled for all instances     *
' *****************************************************************************

    If tdpEventLogDate.MaxDate < Date Then
        SetSystemDateReliantSettings '  Also called via main form Form_Load() -> SetGlobalVariables
    End If
    
    If (Not g.bCaptureCycleRunning) And g.bMaster And g.bAutoDataCapture Then
        If DateDiff("d", g.rstDWDefaults!LastAllFranCaptureCycleDate, GetCaptureCycleDate()) > 0 Then
            tmrAutoDataCapture.Enabled = False
            If Me.WindowState <> vbNormal Then
                Me.WindowState = vbNormal
            End If
            tabMain.Tab = TabEnum.eDataCaptureTab
            subCaptureData pAutoCaptureCycle:=True
            tmrAutoDataCapture.Enabled = True
        End If
    End If

End Sub
Sub TransferSalesRecord(ByRef rsSrc As ADODB.Recordset, _
                        ByRef rsDest As ADODB.Recordset, _
                        ByVal fStampToday As Boolean)
''' Review May be more efficient on MySQL using an INSERT statememt
  
'' INCREMENTALLY SPEED UP ALL THIS CODE. EACH REFERENCE TO COLLECT A FIELD VALUE FROM A RST IS
'' A TRIP TO THE SERVER. BY COLLECTING THEM AT ONCE WE COULD AT LEAST GET SOME SPEED BENEFITS
  
  
    rsDest.AddNew
        rsDest(gconLiveDataTableTSGFranchiseIDField) = rsSrc(gconLiveDataTableTSGFranchiseIDField)
        rsDest(gconLiveDataTableBarcodeField) = rsSrc(gconLiveDataTableBarcodeField)
        rsDest!TransactionDate = rsSrc!TransactionDate
        rsDest(gconLiveDataTableQuantityField) = rsSrc(gconLiveDataTableQuantityField)
        rsDest(gconLiveDataTableTotalIncTaxField) = rsSrc(gconLiveDataTableTotalIncTaxField)
        rsDest(gconLiveDataTableNormalSellIncTaxField) = rsSrc(gconLiveDataTableNormalSellIncTaxField)
        rsDest(gconLiveDataTableCostIncTaxField) = rsSrc(gconLiveDataTableCostIncTaxField)
        rsDest(gconLiveDataTableWholesaleQty) = rsSrc(gconLiveDataTableWholesaleQty)
    ''' *************************************************************************************************''' Version 351
    ''' Version 351 CAN REMOVE FOLLOWING Cn(... ONCE ALL FRANCHISES ARE ON RStats >= Version 546         ''' Version 351
    ''' *** BUT ALSO CHECK IF THERE ARE ANY IMPLICATIONS FOR CLIFFY STORES BEFORE REMOVING, CLIFFY HAS OWN MDB VERSION ''' Version 360
    ''' rsDest(gconLiveDataTableWholesaleActualSell) = rsSrc(gconLiveDataTableWholesaleActualSell)       ''' Version 351
     rsDest(gconLiveDataTableWholesaleActualSell) = Cn(rsSrc(gconLiveDataTableWholesaleActualSell), 0)   ''' Version 351
    ''' *************************************************************************************************''' Version 351
        'Note that the rejectdata and duplicates tables have an extra date field so we can see when
        'the data was rejected.
        ' The livedata table (and temp & prelive) don't have this extra date field.
        If fStampToday Then
            rsDest("todaysDate") = Date
        End If
    rsDest.Update
    
End Sub

Private Sub txtAreaCode_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtAreaCode_DblClick()
    txtAreaCode.Locked = False
End Sub

Private Sub txtBATAFranchiseID_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtBATAFranchiseID_DblClick()
    txtBATAFranchiseID.Locked = False
End Sub

Private Sub txtContact_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtContact_DblClick()
    txtContact.Locked = False
End Sub

Private Sub txtFaxNum_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtFaxNum_DblClick()
    txtFaxNum.Locked = False
End Sub

Private Sub txtModem_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtModem_DblClick()
    txtModem.Locked = False
End Sub

Private Sub txtNodename_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtNodename_DblClick()
    txtNodename.Locked = False
End Sub

Private Sub txtPhone_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtPhone_DblClick()
    txtPhone.Locked = False
End Sub

Private Sub txtPhysicalAddress_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtPhysicalAddress_DblClick()
    txtPhysicalAddress.Locked = False
End Sub

Private Sub txtPromoName_Change()
Dim dtmPromoStart As Date
Dim dtmPromoEnd As Date
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset
    
    strSQL = "SELECT * FROM Promotions WHERE PromoName = " & SqlQ(txtPromoName.Text)
    Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
    If Not (rst.BOF And rst.EOF) Then
        If MsgBox("There is already a promotion called " & vbNewLine & _
                   txtPromoName.Text & vbNewLine & _
                   "Continue adding products to this promotion?", vbYesNo) = vbNo Then
            txtPromoName.Text = vbNullString
            txtPromoName.SetFocus
        Else
        '   Add
        '   Set valid date ranges then set date values for date controls
            dtmPromoStart = rst!PromoStart.Value
            dtmPromoEnd = rst!PromoEnd.Value
        
            dtpPromoStart.MaxDate = dtmPromoEnd
            dtpPromoStart = rst!PromoStart.Value
            
            dtpPromoEnd.MinDate = dtmPromoStart
            dtpPromoEnd = rst!PromoEnd.Value
            
            lstPromoSubCat.SetFocus
            
        End If
    End If
    
    rst.Close
    Set rst = Nothing

End Sub

Private Sub txtRASPassword_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtRASPassword_DblClick()
    txtRASPassword.Locked = False
End Sub

Private Sub txtRASUsername_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtRASUsername_DblClick()
    txtRASUsername.Locked = False
End Sub

Private Sub txtRRP_Change()
'   Enable Save button but only if a stock item is selected (distinct from arriving on tab without a selected stk item)
    cmdSaveStockDetails.Enabled = (Len(Trim$(txtRRP.Text)) > 0) And (lstDescription.SelCount = 1)
End Sub

Private Sub txtSticks_LostFocus()
'   Enable Save button but only if a stock item is selected (distinct from arriving on tab without a selected stk item)
    cmdSaveStockDetails.Enabled = (Len(Trim$(txtSticks.Text)) > 0) And (lstDescription.SelCount = 1)
    
    If IsNumeric(txtSticks) Then
        If cboCategory = gkCAT_CigPkt Then
            If chkPackage.Value = 1 Then
                If Trim$(txtCartonsPerPacket.Text) = vbNullString Then
                '   txtCartonsPerPacket has not been populated so populate it with what should be the setting
                    With cboCtnContainingPkt
                        If .ListIndex > -1 Then
                            txtCartonsPerPacket = Format$(Val(txtSticks) / GetStkValue(pFldName:="Sticks", pStkId:=.ItemData(.ListIndex)), " 0.0###")
                        End If
                    End With
                End If
            End If
        End If
    End If

End Sub

Private Sub txtSuburb_DblClick()
    txtSuburb.Locked = False
End Sub

Private Sub txtVpnIpAddress_Change()
    cmdSaveFranchiseDetails.Enabled = True
End Sub

Private Sub txtVpnIpAddress_DblClick()
    txtVpnIpAddress.Locked = False
End Sub

Private Sub txtVpnIpAddress_Validate(Cancel As Boolean)
    If Len(txtVpnIpAddress) <> 0 Then
        Cancel = Not IsIPAddress(txtVpnIpAddress.Text)
        If Cancel Then
            MsgBox SQ(txtVpnIpAddress) & " is not a valid IP address", vbInformation
        End If
    End If
End Sub

Private Sub txtWholesaleListPrice_Change()
'   Enable Save button but only if a stock item is selected (distinct from arriving on tab without a selected stk item)
    cmdSaveStockDetails.Enabled = (Len(Trim$(txtWholesaleListPrice.Text)) > 0) And (lstDescription.SelCount = 1)
End Sub

Private Sub UnmapShareDiskDisconnectFranchise(ByRef prstFran As ADODB.Recordset)
Dim bDisconnectShare As Boolean
Dim strErrMsg As String
Dim strFranchiseName As String

    StatusBar "Terminating network share", pLog:=False
    
    If IsUseLocalDriveFranFolder() Then
    '   Testing locally
        bDisconnectShare = True
    Else
        bDisconnectShare = NetDisconnectShare(pRemoteOrLocalName:=GetRemotePath(prstFran:=prstFran), pErrMsg:=strErrMsg)
    End If
    
    strFranchiseName = prstFran!FranchiseBusinessName
    
    If Not bDisconnectShare Then
        StatusBar "Error terminating network share. " & strErrMsg, strFranchiseName
    Else
        StatusBar "Network share connection terminated normally", strFranchiseName
    End If
    
End Sub

Private Sub updNielsenRptTxDate_DownClick()
    
    rtxNielsenReportContents.Text = ""
    dtpNielsenRptTxDate.Value = DateAdd("ww", -1, dtpNielsenRptTxDate.Value)
    subRefreshNielsenReportListBox dtpNielsenRptTxDate.Value
    lstNielsenReportDisplayDate.SetFocus

End Sub

Private Sub updNielsenRptTxDate_UpClick()

    If DateAdd("ww", 1, dtpNielsenRptTxDate.Value) <= dtpNielsenRptTxDate.MaxDate Then
        rtxNielsenReportContents.Text = ""
        dtpNielsenRptTxDate.Value = DateAdd("ww", 1, dtpNielsenRptTxDate.Value)
        subRefreshNielsenReportListBox dtpNielsenRptTxDate.Value
    End If
    lstNielsenReportDisplayDate.SetFocus

End Sub

Private Sub UploadFilesToFranchises(ByVal fCurrentSession As Boolean)
' UploadFilesToFranchises only called in following call tree (As@ V400)
' UploadToFranchises_Click -> CreateUploadsPending -> UploadFilesToFranchises -> UploadFilesToOneFranchise
Dim strErrMsg As String
Dim rstFran As ADODB.Recordset
    
    g.bCaptureCycleRunning = True
    
    ' For each franchise, search through the uploads tables for all outstanding uploads.
    ' These are identified by the fact that the 'dateuploaded' field is empty.
    ' If at least one upload for the franchise, connect to it, then upload all files
    
    Set rstFran = GetRst(pCnn:=g.cnnDW, _
                         pSource:="qryFranchiseLive", _
                         pSourceType:=adCmdTable, _
                         pErrMsg:=strErrMsg)
    Do Until rstFran.EOF
        If Not UploadFilesToOneFranchise(prstFran:=rstFran, _
                                         pSessionSeln:=fCurrentSession, _
                                         pErrMsg:=strErrMsg) Then
            
            StatusBar pMsg:=strErrMsg, pFranchise:=rstFran!FranchiseBusinessName
        
        End If
        rstFran.MoveNext
    Loop
    StatusBar "Completed uploads"
    
    rstFran.Close
    Set rstFran = Nothing
    
    g.bCaptureCycleRunning = False

End Sub

Function UploadFilesToOneFranchise(ByRef prstFran As ADODB.Recordset, _
                                   ByVal pSessionSeln As Boolean, _
                          Optional ByRef pCnnRemote As ADODB.Connection = Nothing, _
                          Optional ByRef pErrMsg As String) As Boolean
'   AUrban This procedure could do with a re-write
'   Could possibly include a DoEvents for occassions when a large number of files are uploaded to each franchise
'   pCnnRemote is passed in call from subCaptureData
'   Also success of upload is not tested for all upload cases. This should be done so that failed uploads
'   can upload on the next connection. (Should be re-written so there is a unified approach to testing for success)
Dim bError As Boolean
Dim bIsUploads As Boolean
Dim bFileCopied As Boolean
Dim bIsAFileUpload As Boolean
Dim bPromoAdd As Boolean
Dim bPromoUploaded As Boolean
Dim bPromoRecall As Boolean
Dim bPromoRecalled As Boolean
Dim bUploadSuccess As Boolean
Dim lngFranID As Long
Dim strRemotePath As String
Dim strErrMsg As String
Dim strLocalLogFullFilename As String
Dim strRemoteLogFullFilename As String
Dim fso As Scripting.FileSystemObject

    Const kUpgradeRSLog = "UpgradeRemoteLog.txt"
    Dim sFile As String
    Dim sBaseName As String
    Dim strSQL As String
    Dim strFranName As String
    Dim statusMsg As String
    Dim sDir As String
    Dim flag As String
    Dim sRemoteFile As String
Dim rsRemoteDefaults As ADODB.Recordset
Dim rsUploadsPending As ADODB.Recordset
Dim rsCurrentPending As ADODB.Recordset
Dim cnnRemote As ADODB.Connection

On Error GoTo Procedure_Error
        
    strFranName = prstFran!FranchiseBusinessName

    ' First check for uploads
    ' Order files to be uploaded alphabetically. This is because in the rare event
    ' that we are uploading the upgraderemeotestatistics.exe, it must be last.
'!!! AUrban reason and approach for sorting (see comments above) could do with a rethink !!!
    lngFranID = prstFran!FranchiseIDTSG
    strSQL = "SELECT * FROM FranchiseUploads " & vbNewLine & _
             "WHERE (FranchiseID = " & lngFranID & ") " & _
              " AND (UploadedDate IS NULL)"
    If pSessionSeln Then ' only do the uploads specified in this session
        strSQL = strSQL & " AND UploadCurrentSession"
    End If
    strSQL = strSQL & " ORDER BY UploadFile ASC;"
    Set rsUploadsPending = GetRst(pCnn:=g.cnnDW, _
                                  pSource:=strSQL, _
                                  pSourceType:=adCmdText, _
                                  pRstType:=eEditableFwdOnly, _
                                  pErrMsg:=strErrMsg)
    
    If (rsUploadsPending.BOF And rsUploadsPending.EOF) Then
        rsUploadsPending.Close
        Set rsUploadsPending = Nothing
        StatusBar "No uploads pending for " & strFranName, pLog:=False
    Else
        bIsUploads = True
        
        StatusBar "Attempting Uploads", strFranName
    
        strRemotePath = GetRemotePath(prstFran:=prstFran)
        If Not (pCnnRemote Is Nothing) Then
            Set cnnRemote = pCnnRemote
        Else
            If fConnectFranchiseMapShareDisk(prstFran) Then ''' Review fConnectFranchiseMapShareDisk REALLY SHOULD RETURN ERROR STATUS AND ERROR STRING
            '   If can't connect error is reported in fConnectFranchiseMapShareDisk()
                Set cnnRemote = GetCnn(pDataSource:=strRemotePath & "\" & gkRemoteDbFilename, _
                                       pCnnMode:=adModeShareDenyNone, _
                                       pDataSourceType:=eMdb, _
                                       pCursorLocn:=adUseServer, _
                                       pErrMsg:=strErrMsg)
                If cnnRemote Is Nothing Then
                    bError = True
                    strErrMsg = UCase$("Can't connect to database: ") & strErrMsg
                    UnmapShareDiskDisconnectFranchise prstFran
                End If
            End If
        End If
       
        If Not (cnnRemote Is Nothing) Then
            
            Do Until rsUploadsPending.EOF
                sDir = vbNullString
                flag = vbNullString
                statusMsg = vbNullString
                bFileCopied = False
                bIsAFileUpload = False
                bPromoAdd = False
                bPromoUploaded = False
                bPromoRecall = False
                bPromoRecalled = False
                
                ' Examine the 'item' to be uploaded.
                ' It may be a file, or a 'tag' indicating an action.
                sFile = rsUploadsPending(gconUploadFileField)
                If sFile = gconRemoteDefaultsTableDatabaseOpenedByField Then  ' a tag
                    ' no flag to set, just the value
                    statusMsg = "Resetting 'OpenedBy' on remote machine"
                    ReSetRemoteOpenedByField strFranName, cnnRemote, sFile
                ElseIf Left$(sFile, 5) = gkPromoADD Then  ' a promotion campaign
                    bPromoAdd = True
'REMOTELY Document the TfrStatus field in Tsg db at TSG AND REREAD Review notes above which still apply

''' Review: pKludgeUploadPromos = Is passed as True when called from Uploads Tab "Upload ALL Now"
''''        UploadToFranchises_Click -> CreateUploadsPending -> UploadFilesToFranchises -> UploadFilesToOneFranchise

' V400 REVIEW: It is possible that all use of KludgeUploadPromos could be removed.
' V400 REVIEW: Would require careful systematic review and testing.

' V400 Uploading promotions data is managed in this procedure via PROMO and DELPROMO records in FranchiseUploads.
' V400 Could it be a worthwhile simoplification to ignore these records in this loop and use prstFran
' V400 passed to this procedure to lookup corresponding promotions data in tblFranchisePromotions
' V400 and use tblFranchisePromotions to manage uploading of promotions data
                ''' V400 Start - Removing pKludgeUploadPromos
                ''' If pKludgeUploadPromos Then '*** A kludge flag requiring investigation ***'
                        bPromoUploaded = UploadPromotion(prstFran, cnnRemote, sFile)
                ''' End If
                ''' V400 End
                ElseIf Left$(sFile, 8) = gkPromoDELETE Then   ' recall a promotion campaign
                    bPromoRecall = True
                ''' V400 Start - Removing pKludgeUploadPromos
                ''' If pKludgeUploadPromos Then '*** A kludge flag requiring investigation ***'
                        bPromoRecalled = RemotePromotionRecall(lngFranID, strFranName, cnnRemote, sFile)
                ''' End If
                ''' V400 End
            
            '   If it wasn't a tag, assume it is a file.
            '   Double check that upload files exists locally so we have something to send.
            '   It should always be here, but check anyway.
                ElseIf Dir(sFile) = "" Then
                    bIsAFileUpload = True
                    StatusBar "FILE UPLOADS FAILED: Cannot find " & sFile & " to upload", strFranName
                Else
                    bIsAFileUpload = True
                    sBaseName = fGetLastWord(sFile, "\")
                    ' Before checking if the remote file already exists (which it probably
                    ' shouldn't), first ascertain in which subfolder it resides on
                    ' the remote machine.
                    ' If it does exist, delete it, and regard the one we're about to send
                    ' as the valid one.
                    If LCase(Left(sBaseName, 5)) = LCase(gconNewStockFilePrefix) Then
                        statusMsg = "setting newstock flag"
                        sDir = mkNewStockFolderName
                        flag = gconRemoteDefaultsTableNewStockFlag
                    '   Increment FileNums field in StkFileNums
                    '   NB: File will not be Read-only if manually edited,
                    '   therefore force file counter to increment here as it
                    '   will not increment in addToUpdateFile() if it is writeable.
                        IncrementStkFileNum sBaseName
                    ElseIf LCase(Left(sBaseName, 5)) = LCase(gconWLPUpgradePrefix) Then
                        statusMsg = "setting WLP flag"
                        sDir = mkWLPUpgradesFolderName
                        flag = gconRemoteDefaultsTableWLPUpdateFlag
                    ElseIf LCase(Left(sBaseName, 5)) = LCase(gconUpdateStkFldsUpdatePrefix) Then
                    ' * Stk Fields update shares directory and flag with WLP update *
                    '   With a stk fields update all pending price updates on the franchise
                    '   computer are deleted when the description update is applied or discarded.
                    '   This is done so subsequent WLP updates don't overwrite updated stk fields.
                    '   When a stk fields update is detected by Price Module stk field update controls
                    '   are made visible and WLP controls are made invisible. Franchises can only
                    '   update Stk fields and not any old pending WLP updates
                        statusMsg = "setting Stk Fields update flag"
                        sDir = mkWLPUpgradesFolderName
                        flag = gconRemoteDefaultsTableWLPUpdateFlag
                    ElseIf LCase(Left(sBaseName, 5)) = LCase(gconNewMessageFilePrefix) Then
                        statusMsg = "setting Message flag"
                        sDir = mkMessageFolderName
                        flag = gconRemoteDefaultsTableMessageFlag
                    ElseIf LCase(sBaseName) = LCase(gconUpgradeRS) Then
                        statusMsg = "setting upgrade RS flag"
                        flag = gconRemoteDefaultsTableUpgradeField
                    ElseIf LCase(sBaseName) = LCase(gconUtilityExe) Then
                        statusMsg = "setting 'run utility' flag"
                        flag = gconRemoteDefaultsTableRunUtility
                    Else
                        ' Do nothing. By default files get copied to statistics folder
                        ' and no flag is set
                    End If
                    
                    sRemoteFile = strRemotePath & "\" & sDir & "\" & sBaseName
                    
                    StatusBar "deleting old " & sRemoteFile, strFranName
                    
    ''' Review  fso.FileExists(blah)  does not return an error as does Dir if the connection is dropped
    'Perhaps there needs to be an existence test for the file then the dir if the file fails
    
                    If Dir(strRemotePath & "\" & sDir & "\", vbDirectory) = "" Then ''' BUG SITE
                        StatusBar "remote dir " & strRemotePath & "\" & sDir & " doesnt exist.", strFranName
                    ElseIf Not DeleteFile(sRemoteFile) Then
                        StatusBar "Could not delete existing copy of file: " & sRemoteFile, strFranName
                    Else
                        StatusBar "copying " & sBaseName & " to " & strRemotePath & "\" & sDir, strFranName
                        FileCopy sFile, sRemoteFile ''' Bad file name or number 14Oct09
                                                    ''' O/N Crash Version 3.2.9 17Jul09 (Error 52: Description: Bad file name or number)
                                                    ''' Path/File access error  22Jul09
                                                    ''' Bad file name or number 21Apr09
                        bFileCopied = True
                    End If
                    
                End If
                
                ' determine which type of file or what 'items' we uploaded and set the relevant flag on the remote machine
                If Len(statusMsg) Then
                    StatusBar statusMsg, strFranName
                End If
                
                If Len(flag) Then
                    Set rsRemoteDefaults = GetRst(pCnn:=cnnRemote, pSource:="Defaults", pSourceType:=adCmdTable, pRstType:=eEditableFwdOnly, pErrMsg:=strErrMsg)
                    If Not IsFldExists(pRst:=rsRemoteDefaults, pFldName:=flag) Then
                        StatusBar "MISSING " & SQ(flag) & " field in RemoteStatistics.mdb defaults table", strFranName
                    Else
                        If flag = gconRemoteDefaultsTableRunUtility Then
                            rsRemoteDefaults(flag) = "y"
                        Else
                            rsRemoteDefaults(flag) = True
                        End If
                    End If
                    rsRemoteDefaults.Update
                    rsRemoteDefaults.Close
                    Set rsRemoteDefaults = Nothing
                End If
        
                ' Finally flag the fact that it was uploaded by stamping the time in
                ' our database against the upload-pending entry.
                ' We could really just delete the entry, but because I am so anally retentive, we'll
                ' keep them for auditing purposes.
            
            '   If file upload could not be uplaoded b/c dest file couldn't be deleted, then don't flag as uplaoded
            '   That is only update Uploads pending when it isn't a file upload, or it is a file upload and
            '   as far as the testing has gone it has been successfully uploaded
' AUrban Review. The upload functions called should return a Success status to determine whether Uploaded field should be updated
        
                If bPromoAdd Then
                    bUploadSuccess = bPromoUploaded
                ElseIf bPromoRecall Then
                    bUploadSuccess = bPromoRecalled
                Else
                    bUploadSuccess = Not bIsAFileUpload Or (bIsAFileUpload And bFileCopied)
                End If
                
                If bUploadSuccess Then
                        rsUploadsPending(gconUploadDateField) = Now
                    rsUploadsPending.Update
                End If
                
                rsUploadsPending.MoveNext
                ' remove this one from the listview by doing a refresh of the list of
                ' outstanding uploads
            Loop
            rsUploadsPending.Close
            Set rsUploadsPending = Nothing
            
           ' Before disconnecting, check if there was an 'upgraderemotelog.txt' & if so grab it
            Set fso = New Scripting.FileSystemObject
            If fso.FolderExists(g.strLogFolder) Then
                strLocalLogFullFilename = g.strLogFolder & "\" & strFranName & "_" & kUpgradeRSLog
                strRemoteLogFullFilename = strRemotePath & "\" & kUpgradeRSLog
                If fso.FileExists(strRemoteLogFullFilename) Then
                    If fso.FileExists(strLocalLogFullFilename) Then
                        fso.DeleteFile strLocalLogFullFilename, Force:=True
                    End If
                    fso.MoveFile strRemoteLogFullFilename, strLocalLogFullFilename
                    StatusBar "Snatched " & fso.GetFileName(strRemoteLogFullFilename)
                End If
            End If

            If pCnnRemote Is Nothing Then
            '   pCnnRemote not passed => connected to franchise & made cnn to db in procedure and must now clean up
                cnnRemote.Close
                Set cnnRemote = Nothing
                UnmapShareDiskDisconnectFranchise prstFran
            End If
        End If
    
    End If
    
Procedure_Exit:
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'   Maintain previous functioning (rightly or wrongly) and avoid the bug:-
'   "Query-based update failed because the row to update could not be found."
'   Prior to MySQL, code used two independently and concurrently opened rsts used to edit FranchiseUploads.
'   In MySQL it caused a bug (reasonably so) when a second update was attempted on same row using 2nd rst.
'   FranchiseUploads doesn't have a primary key. Access must have internally assigned IDs to table rows.
'
'     remove "current session flag from this upload" regardless of whether it worked or not.
    If bIsUploads And pSessionSeln Then
        strSQL = "UPDATE FranchiseUploads SET UploadCurrentSession = False WHERE FranchiseID = " & lngFranID
        CnnDwExecute pCommandText:=strSQL
    End If
'    ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    Set fso = Nothing
    pErrMsg = strErrMsg
    UploadFilesToOneFranchise = Not bError
    
    Exit Function

Procedure_Error:
    bError = True
    
'   Collect error details with fsErrDetail() before resetting Error object with 'On Error Resume Next'
'   NB error msg returned from calls within procedure preserved and prepended to strErrMsg/pErrMsg
    If Len(strErrMsg) Then
        strErrMsg = strErrMsg & vbNewLine
    End If
    strErrMsg = "FILE UPLOADS FAILED: " & strErrMsg & fsErrDetail("UploadFilesToOneFranchise")
    
    If pCnnRemote Is Nothing Then
    '   pCnnRemote not passed -> connected to franchise & made cnn to db in proc and must now clean up
    '   (or possibly franchise connection dropped out so cleanup is cloaked in 'On Error Resume Next'
        CloseCnnAndSetToNothing_IgnoreErrors pCnn:=cnnRemote, _
                                             pLogIfAvoidsBug:=True, _
                                             pCalledFromTag:="UploadFilesToOneFranchise"
        UnmapShareDiskDisconnectFranchise prstFran
    End If
    
    Resume Procedure_Exit
    Resume  ' Not executed but assists when debugging in IDE
    
End Function

Private Sub UploadLatestAztecRpts()
Const kDailyDataFileType As Boolean = True
Const kWEEKLYDataFileType As Boolean = False
Dim bUploaded As Boolean
Dim bDailyPreviouslyUploaded As Boolean
Dim intPrevMousePointer As Integer
Dim bWEEKLYPreviouslyUploaded  As Boolean
Dim dtmLastRptEndDate As Date
Dim strMsg As String
Dim strSQL As String
Dim strErrMsg As String
Dim strZipOfDailyFilename As String
Dim strZipOfDailyFULLname As String
Dim strZipOfWeeklyFilename As String
Dim strZipOfWeeklyFULLname As String
Dim fso As Scripting.FileSystemObject
Dim rstUploads As ADODB.Recordset
Dim oSFTP As clsSFTP

    intPrevMousePointer = SetMousePointer(vbHourglass)

'   Check if most receent WeeklyRpts.zip & DaillyRpts.zip files have been transferred
'   (If previously uploaded tblAztecUploads will have 2 records for appropriate SalesDataEndDate)
    dtmLastRptEndDate = fdtmLastSunday()
    strSQL = "SELECT * FROM tblAztecUploads WHERE (SalesDataEndDate = " & MySqlDate(dtmLastRptEndDate) & ")"
    Set rstUploads = GetRst(pCnn:=g.cnnDW, _
                            pSource:=strSQL, _
                            pSourceType:=adCmdText, _
                            pRstType:=eEditableDynamic, _
                            pErrMsg:=strErrMsg)
    If Not rstUploads Is Nothing Then
        If Not (rstUploads.BOF And rstUploads.EOF) Then
            rstUploads.Find Criteria:="(FileType = " & kDailyDataFileType & ")"
            bDailyPreviouslyUploaded = Not rstUploads.EOF
            rstUploads.MoveFirst
            rstUploads.Find Criteria:="(FileType = " & kWEEKLYDataFileType & ")"
            bWEEKLYPreviouslyUploaded = Not rstUploads.EOF
        End If
    
        If (Not bDailyPreviouslyUploaded) Or (Not bWEEKLYPreviouslyUploaded) Then
        '   Check if files for uplaoding are available
            Set fso = New Scripting.FileSystemObject
            strZipOfDailyFULLname = GetNeilsenDailyRptFullname(pLastReportEndDate:=dtmLastRptEndDate)
            strZipOfDailyFilename = fso.GetFileName(strZipOfDailyFULLname)
            strZipOfWeeklyFULLname = GetNeilsenWeeklyRptFullname(pLastReportEndDate:=dtmLastRptEndDate)
            strZipOfWeeklyFilename = fso.GetFileName(strZipOfWeeklyFULLname)
        
            If Not fso.FolderExists(GetNeilsenRptsSubFolderName(pLastReportEndDate:=dtmLastRptEndDate)) Then
                StatusBar UCase$("Aztec reports for " & Format(dtmLastRptEndDate, gkFmtDateUnambiguous) & " have not been created.")
            Else
                If (Not bDailyPreviouslyUploaded) And (Not bWEEKLYPreviouslyUploaded) Then
                    strMsg = "Transfering files to Aztec"
                ElseIf Not bDailyPreviouslyUploaded Then
                    strMsg = "Transfering DAILY sales data file to Aztec"
                Else
                    strMsg = "Transfering WEEKLY sales data file to Aztec"
                End If
                StatusBar strMsg
                
                Set oSFTP = New clsSFTP
                With oSFTP
                '   Aztec uses SFTP Protocol (default mode for clsSFTP)
                    .HostAddress = g.rstDWDefaults!AztecFtpHostAddress
                    .Login = g.rstDWDefaults!AztecFtpUser
                    .Password = g.rstDWDefaults!AztecFtpPwd
                    .TransferType = eTfr_BINARY
                    
                '   Transfer Daily Sales Data File
                    If Not bDailyPreviouslyUploaded Then
                        If Not fso.FileExists(strZipOfDailyFULLname) Then
                            StatusBar UCase(strZipOfDailyFULLname & " has not been created")
                        Else
                            bUploaded = False
                            If .RemoteExists("\" & strZipOfDailyFilename) Then
                            '   May have been manually uploaded on request, mark as uploaded so upload is not perpetually attempted
                                StatusBar strZipOfDailyFilename & " already exists on Aztec FTP server. (manually uploaded?)"
                                bUploaded = True
                            ElseIf .Upload(pLocalName:=strZipOfDailyFULLname, pRemoteName:="\" & strZipOfDailyFilename, pErrMsg:=strErrMsg) Then
                                StatusBar strZipOfDailyFilename & " uploaded"
                                bUploaded = True
                            Else
                                StatusBar UCase$("Upload failed: " & strZipOfDailyFilename) & ". " & strErrMsg
                            End If
                            If bUploaded Then
                                rstUploads.AddNew
                                    rstUploads!SalesDataEndDate = dtmLastRptEndDate
                                    rstUploads!FileType = kDailyDataFileType
                                    rstUploads!UploadDate = Now
                                rstUploads.Update
                            End If
                        End If
                    End If
                    
                '   Transfer WEEKLY Sales Data File
                    If Not bWEEKLYPreviouslyUploaded Then
                        If Not fso.FileExists(strZipOfWeeklyFULLname) Then
                            StatusBar UCase(strZipOfWeeklyFULLname & " has not been created")
                        Else
                            bUploaded = False
                            If .RemoteExists("\" & strZipOfWeeklyFilename) Then
                            '   May have been manually uploaded on request, mark as uploaded so upload is not perpetually attempted
                                StatusBar strZipOfWeeklyFilename & " already exists on Aztec FTP server. (manually uploaded?)"
                                bUploaded = True
                            ElseIf .Upload(pLocalName:=strZipOfWeeklyFULLname, pRemoteName:="\" & strZipOfWeeklyFilename, pErrMsg:=strErrMsg) Then
                                StatusBar strZipOfWeeklyFilename & " uploaded"
                                bUploaded = True
                            Else
                                StatusBar UCase$("Upload failed: " & strZipOfWeeklyFilename) & ". " & strErrMsg
                            End If
                            If bUploaded Then
                                rstUploads.AddNew
                                    rstUploads!SalesDataEndDate = dtmLastRptEndDate
                                    rstUploads!FileType = kWEEKLYDataFileType
                                    rstUploads!UploadDate = Now
                                rstUploads.Update
                            End If
                        End If
                    End If
                End With
                Set oSFTP = Nothing
            
            End If
            Set fso = Nothing
        
        End If
        rstUploads.Close
        Set rstUploads = Nothing
    
    End If

    StatusBar vbNullString, pLog:=False
    SetMousePointer intPrevMousePointer
    
End Sub

Private Function UploadPromotion(ByRef prstFran As ADODB.Recordset, _
                                 ByRef pCnnRemote As ADODB.Connection, _
                                 ByVal sPromoID As String) As Boolean
''' ORIGINAL DESIGN SHOULD HAVE SIMPLY BEEN UploadPromotion(pFranID) AND THE FUNCTION COULD HAVE RESOLVED WHAT PROMOTIONS
''' APPLY AT THAT INSTANT IN TIME, CHECKED WHETHER THEY HAVE ALREADY BEEN UPLOADED AGAINST A FranchisePromotion TABLE
''' AND TAKEN THE APPROPRIATE ACTION. COULD AND SHOULD HAVE OCCURRED OUTSITE OF THE FranchiseUploads TABLE AND PARADIGM
Dim bResult As Boolean
Dim bFranPromoGradeChanged As Boolean
Dim lngPromoID As Long
Dim strErrMsg As String
Dim strSqlLocal As String
Dim strSqlRemote As String
Dim strFranName As String
Dim rstLocalPromo As ADODB.Recordset
Dim rstRemotePromo As ADODB.Recordset

    ' sPromoID is of the form PROMOnnn where nnn is the promoID, so get the latter as number
    lngPromoID = Val(Right(sPromoID, Len(sPromoID) - 5))
    strFranName = prstFran!FranchiseBusinessName

    strSqlLocal = "SELECT * FROM Promotions " & vbNewLine & _
                  "WHERE (PromoID = " & lngPromoID & ") " & _
                   " AND (PromoEnd >= " & MySqlDate(Date) & ")"
                   
    Set rstLocalPromo = GetRst(pCnn:=g.cnnDW, _
                               pSource:=strSqlLocal, _
                               pSourceType:=adCmdText, _
                               pRstType:=eEditableFwdOnly, _
                               pErrMsg:=strErrMsg)
                               
    If (rstLocalPromo.BOF And rstLocalPromo.EOF) Then
        rstLocalPromo.Close
        strSqlLocal = "SELECT * FROM Promotions WHERE PromoID = " & lngPromoID
        Set rstLocalPromo = GetRst(pCnn:=g.cnnDW, pSource:=strSqlLocal, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        If (rstLocalPromo.BOF And rstLocalPromo.EOF) Then
            StatusBar "Promo " & lngPromoID & " not uploaded. Promo Not Found", strFranName
        Else
        '   rstLocalPromo!PromoEnd < Date can be assumed from preceding code
            StatusBar "Promo " & lngPromoID & " Expired -> NOT Uploaded", strFranName
        
        '   Return True so that Upload is not retried.
        '   Upload date will be greater than the expiry date for later checking of Upload tables
            bResult = True
        End If
    Else
    '   As at Version 387 this proc called from one place and only used for uploading new promotions (cf recalling/deleting promotions)
    '   Apply last minute check that PromoGrade of Promo & Fran match [ie. check fran PromoGrade hasn't changed since Promo was created]
    ' * EXECPT WHEN PROMOTIONS CREATED FOR NOMINATED FRANCHISES (ie WHERE PromoGrade = mkPromoGradeIdNA  [-1] )
        If (rstLocalPromo!PromoGradeID <> mkPromoGradeIdNA) Then
        '   Promo NOT created for specific selected franchise(s),
        '   -> test whehter Fran stil has same PromoGradeID as promo
            If (prstFran!PromoGradeID <> rstLocalPromo!PromoGradeID) Then
            '   Franchise!PromoGrade is no longer the same as promotion - promo will eventually expire
            '   (another great example of where tblFranchisePromotions could be used, perhaps with an extra field like status)
                StatusBar "Franchise no longer has same Promotion Grade as promotion " & lngPromoID, strFranName
                bFranPromoGradeChanged = True
            End If
        End If
On Error GoTo Procedure_Error
        If Not bFranPromoGradeChanged Then
            strSqlRemote = "SELECT * FROM Promotions WHERE PromoID = " & rstLocalPromo!PromoID
            Set rstRemotePromo = GetRst(pCnn:=pCnnRemote, _
                                        pSource:=strSqlRemote, _
                                        pSourceType:=adCmdText, _
                                        pRstType:=eEditableFwdOnly, _
                                        pErrMsg:=strErrMsg)
        '   if this promo doesn't already exist on remote machine, then add it.
'   V400 Review
'   (CAUSES ANOMOLIES WHEN TESTING ON A DEV MACHINE WITH ONLY ONE FRANCHISE DATABASE FOR ALL FRANS)
'   (EVEN IF CONTINUALLY EDITING RStats.mdb!defaults!FranchiseID THERE WILL ALREADY BE DATA FROM OTHER
'   (FRANS AND YOU MAY NOT ENTER THE FOLLOOWING IF STATEMETN AND THEREFORE NOT UPDATING FPTfrStatus IN tblFranchisePromotions etc)
'   (SHOULD AT LEAST EventLog CASES WHERE THE PROMO ALREADY EXISTS IN REMOTE DB AND THINK OF THROWING
'   (AN ERROR SO IT DOESN'T THROW THE DEV TESTING PROCESS INTO TURMOIL. SIMILAR STUFF WILL PROBABLY APPLY TO RECALLING PROMOS)
'
'   WHOLE RST OPERATIONS BELOW SHOULD PROBABLY BE REPLACED WITH SQL FOR CONSISTENCY, BUT FOR NOW
'   RST OPERATIONS AGAINST MySQL ARE A HIGHER PRIORITY FOR REPLACEMENT BY SQL
            If (rstRemotePromo.BOF And rstRemotePromo.EOF) Then
''' Review: COULD UPLOADING PROMOTIONS COULD SPEED UP BY SOMETHOW GRABBING ALL THE MYSQL VALS FROM THE PROMO RECORD
'           AS A BATCH THEN USING THE VALUES TO EXECUTE SQL AGAINST THE REMOTE RStats.mdb. THE SPPED IMPROVEMENT
'           WOULD COME FROM NOT REPEATEDLY ACCESSING THE MySQL DATATBASE TO RETRIEVE RST VALS
                rstRemotePromo.AddNew
                '   AUrban: Could investigate whether all the fields are required by Remote Statistics,
                '   OR use existing fields in remote db as template of what to copy from Head Office DB
                '   as at June 2006 the RegionId field was added to head office but not to RemoteStatistics DB.
                    rstRemotePromo!PromoID = rstLocalPromo!PromoID
                    rstRemotePromo!PromoName = rstLocalPromo!PromoName
                    rstRemotePromo!PromoSubCat = rstLocalPromo!PromoSubCat
                    rstRemotePromo!PromoStart = rstLocalPromo!PromoStart
                    rstRemotePromo!PromoEnd = rstLocalPromo!PromoEnd
                    rstRemotePromo!PromoCartonDiscount = rstLocalPromo!PromoCartonDiscount
                    rstRemotePromo!PromoPacketDiscount = rstLocalPromo!PromoPacketDiscount
                '   RStats.mdb!Promotions!PromoState won't accept a zero length string
                '   PromoState is not used, but populating field is of value to support staff
                    rstRemotePromo!PromoState = Czls(rstLocalPromo!PromoState, Null)
                '   rstRemotePromo!PromoRegionId = rstLocalPromo!PromoRegionId " RStats mdb has not yet been given this field
                    rstRemotePromo!PromoStatus = "New"
                rstRemotePromo.Update

''' Review: Could almost have SetFPTfrStatus() within a Txn that is rolled back if there is a problem before rstRemotePromo
'''         Would hate to have a problem where remote promo was updated but FP and Promotions table weren't
                SetFPTfrStatus pFranID:=prstFran!FranchiseIDTSG, pPromoID:=lngPromoID, pTfrStatus:=FpTfrCompleted  ''' V388

''' dao2AD0 REVIEW MySQL
''' Comment below and the reasoning is why the original design of promotions is fundamentally flawed
                 '  Now that we have uploaded this promo to at least one store, flag it as SENT
''' dao2AD0 REVIEW MySQL
                    rstLocalPromo!PromoStatus = PROMO_SENT
                rstLocalPromo.Update
SetTableUpdateTime pTableName:="Promotions", pTimeStamp:=Now    ''' Review
                
            End If
            rstRemotePromo.Close
            Set rstRemotePromo = Nothing
            bResult = True  ' Treat as successful even if upload existed b/c we don't wish to continually retry same upload if it exists remotely
            StatusBar "Uploaded promotion  " & lngPromoID, strFranName
        End If
        
    End If
    rstLocalPromo.Close
    Set rstLocalPromo = Nothing
    
Procedure_Exit:
    UploadPromotion = bResult
    Exit Function

Procedure_Error:
    StatusBar "Error Uploading Promotion: " & lngPromoID & ", Err: " & Err.Number & " " & Err.Description, strFranName
    Resume Procedure_Exit
    Resume  ' Not executed but assists when debugging in IDE

End Function

Private Sub UploadToFranchises_Click(Index As Integer)  ' Uploads Tab upload buttons
' Uploads Tab Buttons (Upload This NOW, Upload ALL NOW, Upload ALL LATER)
    If Index = 0 Then
    '   Upload ALL NOW
        StatusBar pMsg:="Manual Upload: ALL uploads"
        Call CreateUploadsPending(IMMEDIATELY, ALL_UPLOADS)
        StatusBar pMsg:="Manual Upload Completed"
    ElseIf Index = 1 Then
    '   Upload ALL LATER
        Call CreateUploadsPending(LATER, ALL_UPLOADS)
    ElseIf Index = 2 Then
    '   Upload This NOW
        StatusBar pMsg:="Manual Upload: Selected uploads"
        Call lvwEventLog_DblClick
        Call CreateUploadsPending(IMMEDIATELY, CURRENT_UPLOADS)
        Call lvwEventLog_DblClick
        StatusBar pMsg:="Manual Upload Completed"
    End If
End Sub

Private Sub WriteStockToTextFile(ByRef pRstStock As ADODB.Recordset, ByVal pFullFilename As String)
Dim intFileNum As Integer
Dim strSQL As String
Dim strErrMsg As String
Dim rst As ADODB.Recordset

    intFileNum = FreeFile
    Open pFullFilename For Append As #intFileNum

    Print #intFileNum, Chr(34) & pRstStock(gconStockTableBarcodeField) & Chr(34) & ",", _
                           Chr(34) & pRstStock(gconStockTableDescriptionField) & Chr(34) & ",", _
                           Chr(34) & pRstStock(gconStockTableCategoryField) & Chr(34) & ",", _
                           Chr(34) & pRstStock(gconStockTableSubCategoryField) & Chr(34) & ",", _
                           pRstStock(gconStockTableSupplierIDField) & ",", _
                           pRstStock(gconStockTableCostField) & ",", _
                           pRstStock(gconStockTableSellField) & ",", _
                           Chr(34) & pRstStock(gconStockTableSalesTaxCodeField) & Chr(34) & ",", _
                           Chr(34) & pRstStock(gconStockTableGoodsTaxCodeField) & Chr(34) & ",", _
                           Format$(pRstStock(gconStockTableAllowFractionsField), "True/False") & ",", _
                           Format$(pRstStock(gconStockTablePackageField), "True/False") & ",", _
                           Format$(pRstStock(gconStockTableTaxComponentsField), "True/False") & ",", _
                           pRstStock!Stock_ID

    Close #intFileNum

'   Write associated package file entries (if associate package record exists)
'   As at Version 358 presume there is only one component to the package (as this is how TSG use CgtPkts and CgtCtns)
    If CBool(pRstStock!Package) Then
        strSQL = "SELECT * FROM Package WHERE Package_ID  = " & pRstStock!Stock_ID
        Set rst = GetRst(pCnn:=g.cnnDW, pSource:=strSQL, pSourceType:=adCmdText, pErrMsg:=strErrMsg)
        
        If Not (rst.BOF And rst.EOF) Then
            intFileNum = FreeFile   ' Get unused file
            Open GetStkPkgFullFilename(pFullFilename) For Append As #intFileNum
            '                  package_id           stock_id             sell_inc             quantity
            Print #intFileNum, rst(0).Value & "," & rst(1).Value & "," & rst(2).Value & "," & rst(3).Value
            Close #intFileNum
        End If
        rst.Close
        Set rst = Nothing
    End If
    
End Sub



