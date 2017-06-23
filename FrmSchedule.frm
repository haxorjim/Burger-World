VERSION 5.00
Begin VB.Form FrmSchedule 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6936
   ClientLeft      =   744
   ClientTop       =   816
   ClientWidth     =   7896
   LinkTopic       =   "Form1"
   ScaleHeight     =   6936
   ScaleWidth      =   7896
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   288
      Left            =   7188
      TabIndex        =   149
      Top             =   252
      Width           =   500
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Default         =   -1  'True
      Height          =   708
      Left            =   132
      TabIndex        =   148
      Top             =   6072
      Width           =   7560
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   6840
      TabIndex        =   147
      Top             =   5712
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   6840
      TabIndex        =   146
      Top             =   5424
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   6840
      TabIndex        =   145
      Top             =   5136
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   6840
      TabIndex        =   144
      Top             =   4848
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   6840
      TabIndex        =   143
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   6840
      TabIndex        =   142
      Top             =   4272
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   6840
      TabIndex        =   141
      Top             =   3984
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   6840
      TabIndex        =   140
      Top             =   3696
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   6840
      TabIndex        =   139
      Top             =   3408
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   6840
      TabIndex        =   138
      Top             =   3120
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   6840
      TabIndex        =   137
      Top             =   2832
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   6840
      TabIndex        =   136
      Top             =   2544
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   6840
      TabIndex        =   135
      Top             =   2256
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   6840
      TabIndex        =   134
      Top             =   1968
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   6840
      TabIndex        =   133
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   6840
      TabIndex        =   132
      Top             =   1392
      Width           =   852
   End
   Begin VB.TextBox Text12 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   6840
      TabIndex        =   131
      Top             =   1104
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   6000
      TabIndex        =   130
      Top             =   5712
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   6000
      TabIndex        =   129
      Top             =   5424
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   6000
      TabIndex        =   128
      Top             =   5136
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   6000
      TabIndex        =   127
      Top             =   4848
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   6000
      TabIndex        =   126
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   6000
      TabIndex        =   125
      Top             =   4272
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   6000
      TabIndex        =   124
      Top             =   3984
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   6000
      TabIndex        =   123
      Top             =   3696
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   6000
      TabIndex        =   122
      Top             =   3408
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   6000
      TabIndex        =   121
      Top             =   3120
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   6000
      TabIndex        =   120
      Top             =   2832
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   6000
      TabIndex        =   119
      Top             =   2544
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   6000
      TabIndex        =   118
      Top             =   2256
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   6000
      TabIndex        =   117
      Top             =   1968
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   6000
      TabIndex        =   116
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   6000
      TabIndex        =   115
      Top             =   1392
      Width           =   852
   End
   Begin VB.TextBox Text11 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   6000
      TabIndex        =   114
      Top             =   1104
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   5160
      TabIndex        =   113
      Top             =   5712
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   5160
      TabIndex        =   112
      Top             =   5424
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   5160
      TabIndex        =   111
      Top             =   5136
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   5160
      TabIndex        =   110
      Top             =   4848
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   5160
      TabIndex        =   109
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   5160
      TabIndex        =   108
      Top             =   4272
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   5160
      TabIndex        =   107
      Top             =   3984
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   5160
      TabIndex        =   106
      Top             =   3696
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   5160
      TabIndex        =   105
      Top             =   3408
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   5160
      TabIndex        =   104
      Top             =   3120
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   5160
      TabIndex        =   103
      Top             =   2832
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   5160
      TabIndex        =   102
      Top             =   2544
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   5160
      TabIndex        =   101
      Top             =   2256
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   5160
      TabIndex        =   100
      Top             =   1968
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   5160
      TabIndex        =   99
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   5160
      TabIndex        =   98
      Top             =   1392
      Width           =   852
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   5160
      TabIndex        =   97
      Top             =   1104
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   4320
      TabIndex        =   96
      Top             =   5712
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   4320
      TabIndex        =   95
      Top             =   5424
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   4320
      TabIndex        =   94
      Top             =   5136
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   4320
      TabIndex        =   93
      Top             =   4848
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   4320
      TabIndex        =   92
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   4320
      TabIndex        =   91
      Top             =   4272
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   4320
      TabIndex        =   90
      Top             =   3984
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   4320
      TabIndex        =   89
      Top             =   3696
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   4320
      TabIndex        =   88
      Top             =   3408
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   4320
      TabIndex        =   87
      Top             =   3120
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   4320
      TabIndex        =   86
      Top             =   2832
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   4320
      TabIndex        =   85
      Top             =   2544
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   4320
      TabIndex        =   84
      Top             =   2256
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   4320
      TabIndex        =   83
      Top             =   1968
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   4320
      TabIndex        =   82
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   4320
      TabIndex        =   81
      Top             =   1392
      Width           =   852
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   4320
      TabIndex        =   80
      Top             =   1104
      Width           =   852
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   3360
      TabIndex        =   79
      Top             =   5712
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   3360
      TabIndex        =   78
      Top             =   5424
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   3360
      TabIndex        =   77
      Top             =   5136
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   3360
      TabIndex        =   76
      Top             =   4848
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   3360
      TabIndex        =   75
      Top             =   4560
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   3360
      TabIndex        =   74
      Top             =   4272
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   3360
      TabIndex        =   73
      Top             =   3984
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   3360
      TabIndex        =   72
      Top             =   3696
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   3360
      TabIndex        =   71
      Top             =   3408
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   3360
      TabIndex        =   70
      Top             =   3120
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   3360
      TabIndex        =   69
      Top             =   2832
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   3360
      TabIndex        =   68
      Top             =   2544
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   3360
      TabIndex        =   67
      Top             =   2256
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   3360
      TabIndex        =   66
      Top             =   1968
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   3360
      TabIndex        =   65
      Top             =   1680
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   3360
      TabIndex        =   64
      Top             =   1392
      Width           =   972
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   3360
      TabIndex        =   63
      Top             =   1104
      Width           =   972
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   2520
      TabIndex        =   62
      Top             =   5712
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   2520
      TabIndex        =   61
      Top             =   5424
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   2520
      TabIndex        =   60
      Top             =   5136
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   2520
      TabIndex        =   59
      Top             =   4848
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   2520
      TabIndex        =   58
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   2520
      TabIndex        =   57
      Top             =   4272
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   2520
      TabIndex        =   56
      Top             =   3984
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   2520
      TabIndex        =   55
      Top             =   3696
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   2520
      TabIndex        =   54
      Top             =   3408
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   2520
      TabIndex        =   53
      Top             =   3120
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   2520
      TabIndex        =   52
      Top             =   2832
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   2520
      TabIndex        =   51
      Top             =   2544
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   2520
      TabIndex        =   50
      Top             =   2256
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   2520
      TabIndex        =   49
      Top             =   1968
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   2520
      TabIndex        =   48
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   2520
      TabIndex        =   47
      Top             =   1392
      Width           =   852
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   2520
      TabIndex        =   46
      Top             =   1104
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   1680
      TabIndex        =   45
      Top             =   5712
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   1680
      TabIndex        =   44
      Top             =   5424
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   1680
      TabIndex        =   43
      Top             =   5136
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   1680
      TabIndex        =   42
      Top             =   4848
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   1680
      TabIndex        =   41
      Top             =   4560
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   1680
      TabIndex        =   40
      Top             =   4272
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   1680
      TabIndex        =   39
      Top             =   3984
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   1680
      TabIndex        =   38
      Top             =   3696
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   1680
      TabIndex        =   37
      Top             =   3408
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   1680
      TabIndex        =   36
      Top             =   3120
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   1680
      TabIndex        =   35
      Top             =   2832
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   1680
      TabIndex        =   34
      Top             =   2544
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   1680
      TabIndex        =   33
      Top             =   2256
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   1680
      TabIndex        =   32
      Top             =   1968
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   1680
      TabIndex        =   31
      Top             =   1680
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   1680
      TabIndex        =   30
      Top             =   1392
      Width           =   852
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   1680
      TabIndex        =   29
      Top             =   1104
      Width           =   852
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   16
      Left            =   120
      TabIndex        =   28
      Top             =   5712
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   15
      Left            =   120
      TabIndex        =   27
      Top             =   5424
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   14
      Left            =   120
      TabIndex        =   26
      Top             =   5136
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   13
      Left            =   120
      TabIndex        =   25
      Text            =   "Dave Varhol"
      Top             =   4848
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   12
      Left            =   120
      TabIndex        =   24
      Text            =   "Kris Suter"
      Top             =   4560
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   11
      Left            =   120
      TabIndex        =   23
      Text            =   "Nick Stockhaus"
      Top             =   4272
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   10
      Left            =   120
      TabIndex        =   22
      Text            =   "Mark Raico"
      Top             =   3984
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   9
      Left            =   108
      TabIndex        =   21
      Text            =   "Brian Pajak"
      Top             =   3696
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   8
      Left            =   108
      TabIndex        =   20
      Text            =   "Vince Krepshaw"
      Top             =   3408
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   7
      Left            =   108
      TabIndex        =   19
      Text            =   "Jason Kamp"
      Top             =   3120
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   6
      Left            =   108
      TabIndex        =   18
      Text            =   "Jim Hribar"
      Top             =   2832
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   5
      Left            =   120
      TabIndex        =   17
      Text            =   "Jim Graf"
      Top             =   2544
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Text            =   "Rick Dorsey"
      Top             =   2256
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   3
      Left            =   108
      TabIndex        =   15
      Text            =   "Andy Curtiss"
      Top             =   1968
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   2
      Left            =   108
      TabIndex        =   14
      Text            =   "Ryan Covert"
      Top             =   1680
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   1
      Left            =   108
      TabIndex        =   13
      Text            =   "Alicia Bernes"
      Top             =   1392
      Width           =   1572
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   288
      Index           =   0
      Left            =   108
      TabIndex        =   12
      Text            =   "Ashley Bair"
      Top             =   1104
      Width           =   1572
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   288
      Left            =   6456
      TabIndex        =   10
      Top             =   252
      Width           =   500
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   6
      Left            =   6840
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Sunday"
      Top             =   744
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   5
      Left            =   6000
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Saturday"
      Top             =   744
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   4
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Friday"
      Top             =   744
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   3
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Thursday"
      Top             =   744
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   2
      Left            =   3360
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Wednesday"
      Top             =   744
      Width           =   972
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   1
      Left            =   2520
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Tuesday"
      Top             =   744
      Width           =   852
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   372
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Monday"
      Top             =   744
      Width           =   852
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   108
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "      Dates"
      Top             =   744
      Width           =   1572
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Schedule of Work Hours"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   396
      Left            =   156
      TabIndex        =   8
      Top             =   192
      Width           =   5220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   192
      Left            =   7008
      TabIndex        =   11
      Top             =   300
      Width           =   132
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "For Week of "
      Height          =   192
      Left            =   5484
      TabIndex        =   9
      Top             =   312
      Width           =   912
   End
End
Attribute VB_Name = "FrmSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    cmdExit.SetFocus
    If IsManager = True And OnManagerPanel = True Then
        Text4.Enabled = True
        Text5.Enabled = True
        For Index = 0 To 16
            Text3(Index).Enabled = True
            Text6(Index).Enabled = True
            Text7(Index).Enabled = True
            Text8(Index).Enabled = True
            Text9(Index).Enabled = True
            Text10(Index).Enabled = True
            Text11(Index).Enabled = True
            Text12(Index).Enabled = True
        Next
    End If
End Sub
Private Sub Form_Load()
    Open Schedule_File For Input As #1
    Input #1, startdate, enddate
    Text4.Text = startdate
    Text5.Text = enddate
    For Index = 0 To 16
        Input #1, emp, mon, tues, wed, thurs, fri, sat, sun
        Text3(Index).Text = emp
        Text6(Index).Text = mon
        Text7(Index).Text = tues
        Text8(Index).Text = wed
        Text9(Index).Text = thurs
        Text10(Index).Text = fri
        Text11(Index).Text = sat
        Text12(Index).Text = sun
    Next
    Close #1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Open Schedule_File For Output As #1
        Write #1, Text4.Text, Text5.Text
        For Index = 0 To 16
            Write #1, Text3(Index).Text, Text6(Index).Text, Text7(Index).Text, Text8(Index).Text, Text9(Index).Text, Text10(Index).Text, Text11(Index).Text, Text12(Index).Text
        Next
    Close #1
End Sub

