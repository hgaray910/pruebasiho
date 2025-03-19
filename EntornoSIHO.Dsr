VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} EntornoSIHO 
   ClientHeight    =   13725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13050
   _ExtentX        =   23019
   _ExtentY        =   24209
   FolderFlags     =   7
   TypeLibGuid     =   "{01F44D19-98C3-11D4-99F9-00A024BA0DA5}"
   TypeInfoGuid    =   "{01F44D1A-98C3-11D4-99F9-00A024BA0DA5}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "ConeccionSIHO"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=MSDAORA.1;Password=041099;User ID=DBO;Data Source=SIHO;Persist Security Info=True"
      Expanded        =   -1  'True
      RunPromptBehavior=   4
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   115
   BeginProperty Recordset1 
      CommandName     =   "cmdEjecutaSentencia"
      CommDispId      =   1016
      RsDispId        =   -1
      CommandText     =   "DBO.SP_EJECUTASENTENCIA"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL DBO.SP_EJECUTASENTENCIA( ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_SENTENCIA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdServerFecha"
      CommDispId      =   1018
      RsDispId        =   -1
      CommandText     =   "DBO.SP_GNSERVERFECHA"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_GNSERVERFECHA( ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "RC1"
         Direction       =   2
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   0   'False
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "cmdCnSelPoliza"
      CommDispId      =   1045
      RsDispId        =   -1
      CommandText     =   "DBO.SP_CNSELPOLIZA"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_CNSELPOLIZA( ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdCnSelPoliza_Grouping"
      SummaryExpanded =   -1  'True
      DetailExpanded  =   -1  'True
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_NUMPOLIZA"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "cmdCnUpdEstatusCierre"
      CommDispId      =   1054
      RsDispId        =   -1
      CommandText     =   "DBO.SP_CNUPDESTATUSCIERRE"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL DBO.SP_CNUPDESTATUSCIERRE( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CLAVEEMPRESA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "cmdParametros"
      CommDispId      =   1135
      RsDispId        =   1932
      CommandText     =   "DBO.PARAMETROS"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   43
      BeginProperty Field1 
         Precision       =   0
         Size            =   140
         Scale           =   0
         Type            =   200
         Name            =   "VCHNOMBRE"
         Caption         =   "VCHNOMBRE"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHNOMBRECORTO"
         Caption         =   "VCHNOMBRECORTO"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "VCHRFC"
         Caption         =   "VCHRFC"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   22
         Scale           =   0
         Type            =   200
         Name            =   "VCHREGISTROIMSS"
         Caption         =   "VCHREGISTROIMSS"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "VCHLICENCIASSA"
         Caption         =   "VCHLICENCIASSA"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHREPRESENLEGAL"
         Caption         =   "VCHREPRESENLEGAL"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHDIRECTORGRAL"
         Caption         =   "VCHDIRECTORGRAL"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHDIRECTORMEDICO"
         Caption         =   "VCHDIRECTORMEDICO"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHADMINISTRADORGRAL"
         Caption         =   "VCHADMINISTRADORGRAL"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   200
         Name            =   "IMGLOGO"
         Caption         =   "IMGLOGO"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   120
         Scale           =   0
         Type            =   200
         Name            =   "VCHDIRECCION"
         Caption         =   "VCHDIRECCION"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHCOLONIA"
         Caption         =   "VCHCOLONIA"
      EndProperty
      BeginProperty Field13 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCIUDAD"
         Caption         =   "INTCIUDAD"
      EndProperty
      BeginProperty Field14 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTESTADO"
         Caption         =   "INTESTADO"
      EndProperty
      BeginProperty Field15 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTPAIS"
         Caption         =   "INTPAIS"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHTELEFONO"
         Caption         =   "VCHTELEFONO"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHFAX"
         Caption         =   "VCHFAX"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHCORREOELECTRONICO"
         Caption         =   "VCHCORREOELECTRONICO"
      EndProperty
      BeginProperty Field19 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "VCHPAGINAWEB"
         Caption         =   "VCHPAGINAWEB"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   200
         Name            =   "VCHCODIGOPOSTAL"
         Caption         =   "VCHCODIGOPOSTAL"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   200
         Name            =   "VCHAPARTADOPOSTAL"
         Caption         =   "VCHAPARTADOPOSTAL"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHADMCONTRASENA"
         Caption         =   "VCHADMCONTRASENA"
      EndProperty
      BeginProperty Field23 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYCVETIPOCUARTOPRED"
         Caption         =   "TNYCVETIPOCUARTOPRED"
      EndProperty
      BeginProperty Field24 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYCVEESTADOSALUD"
         Caption         =   "TNYCVEESTADOSALUD"
      EndProperty
      BeginProperty Field25 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYCVEDISPONIBLE"
         Caption         =   "TNYCVEDISPONIBLE"
      EndProperty
      BeginProperty Field26 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYCVEOCUPADO"
         Caption         =   "TNYCVEOCUPADO"
      EndProperty
      BeginProperty Field27 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYEMPRESAACTIVA"
         Caption         =   "TNYEMPRESAACTIVA"
      EndProperty
      BeginProperty Field28 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTADMEMPLEADO"
         Caption         =   "INTADMEMPLEADO"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHMASCARACUENTA"
         Caption         =   "VCHMASCARACUENTA"
      EndProperty
      BeginProperty Field30 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "RELIVAGENERAL"
         Caption         =   "RELIVAGENERAL"
      EndProperty
      BeginProperty Field31 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTTIPOPARTICULAR"
         Caption         =   "INTTIPOPARTICULAR"
      EndProperty
      BeginProperty Field32 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSALARIOMIN"
         Caption         =   "MNYSALARIOMIN"
      EndProperty
      BeginProperty Field33 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSALARIODF"
         Caption         =   "MNYSALARIODF"
      EndProperty
      BeginProperty Field34 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMIDIASPERIODO"
         Caption         =   "SMIDIASPERIODO"
      EndProperty
      BeginProperty Field35 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCODIGOPOSTAL"
         Caption         =   "INTCODIGOPOSTAL"
      EndProperty
      BeginProperty Field36 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMIDIGITOVERIFICADOR"
         Caption         =   "SMIDIGITOVERIFICADOR"
      EndProperty
      BeginProperty Field37 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTEMPRESAPCE"
         Caption         =   "INTEMPRESAPCE"
      EndProperty
      BeginProperty Field38 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEDEPTOEMERGENCIAS"
         Caption         =   "INTCVEDEPTOEMERGENCIAS"
      EndProperty
      BeginProperty Field39 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVENACIONALIDADPREDETERMINA"
         Caption         =   "INTCVENACIONALIDADPREDETERMINA"
      EndProperty
      BeginProperty Field40 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITNOMINA"
         Caption         =   "BITNOMINA"
      EndProperty
      BeginProperty Field41 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTTIEMPOMENSAJES"
         Caption         =   "INTTIEMPOMENSAJES"
      EndProperty
      BeginProperty Field42 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEPROVEEDORALMACENGENERAL"
         Caption         =   "INTCVEPROVEEDORALMACENGENERAL"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "VCHARCHIVOWAV"
         Caption         =   "VCHARCHIVOWAV"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "cmdRegistraErrores"
      CommDispId      =   1150
      RsDispId        =   1933
      CommandText     =   "DBO.REGISTROERRORES"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHA"
         Caption         =   "DTMFECHA"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMHORA"
         Caption         =   "DTMHORA"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTNUMEROERROR"
         Caption         =   "INTNUMEROERROR"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHDESCRIPCION"
         Caption         =   "VCHDESCRIPCION"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "VCHMODULO"
         Caption         =   "VCHMODULO"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHFORMULARIO"
         Caption         =   "VCHFORMULARIO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "cmdParametrosInventarios"
      CommDispId      =   1187
      RsDispId        =   1995
      CommandText     =   "DBO.IVPARAMETROSALMACEN"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      IsRSReturning   =   -1  'True
      NumFields       =   16
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHLOGIN"
         Caption         =   "VCHLOGIN"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHNOMBREDEPTO"
         Caption         =   "VCHNOMBREDEPTO"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVEDEPARTAMENTO"
         Caption         =   "SMICVEDEPARTAMENTO"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAINICIAL"
         Caption         =   "DTMFECHAINICIAL"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAFINAL"
         Caption         =   "DTMFECHAFINAL"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHPASSWORD"
         Caption         =   "VCHPASSWORD"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITESTALLER"
         Caption         =   "BITESTALLER"
      EndProperty
      BeginProperty Field8 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEEMPAUTORIZA"
         Caption         =   "INTCVEEMPAUTORIZA"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   350
         Scale           =   0
         Type            =   200
         Name            =   "VCHPERMISOUSUARIO"
         Caption         =   "VCHPERMISOUSUARIO"
      EndProperty
      BeginProperty Field10 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVEDPTOENINVEN"
         Caption         =   "SMICVEDPTOENINVEN"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAINVEN"
         Caption         =   "DTMFECHAINVEN"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHDPTOSALMACEREC"
         Caption         =   "VCHDPTOSALMACEREC"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "VCHDPTOSXREUBICACION"
         Caption         =   "VCHDPTOSXREUBICACION"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "VCHDPTOSXSALIDADPTO"
         Caption         =   "VCHDPTOSXSALIDADPTO"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "VCHDPTOSXPEDIDO"
         Caption         =   "VCHDPTOSXPEDIDO"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "VCHDPTOSXCARGOPCTE"
         Caption         =   "VCHDPTOSXCARGOPCTE"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "cmdFilEmpleado"
      CommDispId      =   1194
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVFILEMPLEADO"
      ActiveConnectionName=   "ConeccionSIHO"
      Locktype        =   3
      CallSyntax      =   "{ CALL DBO.SP_IVFILEMPLEADO( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_VLINTCVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "cmdIvRequisicionMaestro"
      CommDispId      =   1197
      RsDispId        =   1938
      CommandText     =   "DBO.IVREQUISICIONMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "NUMNUMREQUISICION"
         Caption         =   "NUMNUMREQUISICION"
      EndProperty
      BeginProperty Field2 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVEDEPTOREQUIS"
         Caption         =   "SMICVEDEPTOREQUIS"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEEMPLEAREQUIS"
         Caption         =   "INTCVEEMPLEAREQUIS"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVEDEPTOALMACEN"
         Caption         =   "SMICVEDEPTOALMACEN"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAREQUISICION"
         Caption         =   "DTMFECHAREQUISICION"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMHORAREQUISICION"
         Caption         =   "DTMHORAREQUISICION"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHESTATUSREQUIS"
         Caption         =   "VCHESTATUSREQUIS"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRDESTINO"
         Caption         =   "CHRDESTINO"
      EndProperty
      BeginProperty Field9 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "NUMNUMCUENTA"
         Caption         =   "NUMNUMCUENTA"
      EndProperty
      BeginProperty Field10 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITURGENTE"
         Caption         =   "BITURGENTE"
      EndProperty
      BeginProperty Field11 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "NUMNUMREQUISREL"
         Caption         =   "NUMNUMREQUISREL"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAREQUISAUTORI"
         Caption         =   "DTMFECHAREQUISAUTORI"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMHORAREQUISAUTORI"
         Caption         =   "DTMHORAREQUISAUTORI"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPOPACIENTE"
         Caption         =   "CHRTIPOPACIENTE"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRAPLICACIONMED"
         Caption         =   "CHRAPLICACIONMED"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "cmdSelEmpxDpto"
      CommDispId      =   1201
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELEMPXDPTO"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVSELEMPXDPTO( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_CVEDPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "cmdSelEmpleadosTodos"
      CommDispId      =   1204
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELEMPLEADOSTODOS"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVSELEMPLEADOSTODOS( ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "cmdSelFilRequisMaesxMixtas"
      CommDispId      =   1207
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELFILREQUISMAESXMIXTAS"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVSELFILREQUISMAESXMIXTAS( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEEMP"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_TIPOREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_FECHAINICIO"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "IN_FECHAFIN"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "IN_FOLIOINICIO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "IN_FOLIOFIN"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset13 
      CommandName     =   "cmdSelFilRequisMaesxMix"
      CommDispId      =   1212
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELFILREQUISMAESXMIX"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_IVSELFILREQUISMAESXMIX( ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEEMP"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_TIPOREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset14 
      CommandName     =   "cmdRptRequisMix"
      CommDispId      =   1216
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVRPTREQUISMAES_DETXMIXTAS"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVRPTREQUISMAES_DETXMIXTAS( ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptRequisMix_Grouping"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset15 
      CommandName     =   "cmdRptFilReqxMixxDepto"
      CommDispId      =   1219
      RsDispId        =   1221
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset16 
      CommandName     =   "cmdReqCDMaestro"
      CommDispId      =   1222
      RsDispId        =   1944
      CommandText     =   "DBO.IVREQUISCARDIRMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTNUMREQUISCARDIR"
         Caption         =   "INTNUMREQUISCARDIR"
      EndProperty
      BeginProperty Field2 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVEDEPTOREQUIS"
         Caption         =   "SMICVEDEPTOREQUIS"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEEMPLEAREQUIS"
         Caption         =   "INTCVEEMPLEAREQUIS"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAREQUIS"
         Caption         =   "DTMFECHAREQUIS"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMHORAREQUIS"
         Caption         =   "DTMHORAREQUIS"
      EndProperty
      BeginProperty Field6 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITURGENTEREQUIS"
         Caption         =   "BITURGENTEREQUIS"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHESTATUSREQUIS"
         Caption         =   "VCHESTATUSREQUIS"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAAUTORIZA"
         Caption         =   "DTMFECHAAUTORIZA"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset17 
      CommandName     =   "cmdSelFilRequisCDMaesFecFol"
      CommDispId      =   1225
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELFILOTRASESMAESFECFOL"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVSELFILOTRASESMAESFECFOL( ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEEMP"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_CVETIPO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_FECHAINICIO"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_FECHAFIN"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "IN_FOLIOINICIO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "IN_FOLIOFIN"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset18 
      CommandName     =   "cmdRptRequisCD"
      CommDispId      =   1230
      RsDispId        =   1232
      CommandText     =   "dbo.sp_IvRptRequisCarDirMaes_Det"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptRequisCarDirMaes_Det( ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptRequisCD_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset19 
      CommandName     =   "cmdSelFilRequisCDMaestro"
      CommDispId      =   1233
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELFILREQUISCDMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_IVSELFILREQUISCDMAESTRO( ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEEMP"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset20 
      CommandName     =   "cmdRptFilRequisCDxDepto"
      CommDispId      =   1237
      RsDispId        =   1239
      CommandText     =   "dbo.sp_IvRptFilRequisCDxDepto"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxDepto( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxDepto_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset21 
      CommandName     =   "cmdCiudad"
      CommDispId      =   1240
      RsDispId        =   1242
      CommandText     =   $"EntornoSIHO.dsx":0000
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveCiudad"
         Caption         =   "intCveCiudad"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveEstado"
         Caption         =   "intCveEstado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Estado"
         Caption         =   "Estado"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset22 
      CommandName     =   "cmdEstado"
      CommDispId      =   1243
      RsDispId        =   1245
      CommandText     =   $"EntornoSIHO.dsx":00B6
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveEstado"
         Caption         =   "intCveEstado"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCvePais"
         Caption         =   "intCvePais"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Pais"
         Caption         =   "Pais"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset23 
      CommandName     =   "cmdPais"
      CommDispId      =   1246
      RsDispId        =   1947
      CommandText     =   "DBO.PAIS"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEPAIS"
         Caption         =   "INTCVEPAIS"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHDESCRIPCION"
         Caption         =   "VCHDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   3
         Scale           =   0
         Type            =   200
         Name            =   "VCHAREA"
         Caption         =   "VCHAREA"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVENACIONALIDAD"
         Caption         =   "INTCVENACIONALIDAD"
      EndProperty
      BeginProperty Field5 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITACTIVO"
         Caption         =   "BITACTIVO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset24 
      CommandName     =   "cmdEmpleados"
      CommDispId      =   1249
      RsDispId        =   1948
      CommandText     =   "DBO.NOEMPLEADO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   46
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEEMPLEADO"
         Caption         =   "INTCVEEMPLEADO"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "VCHAPELLIDOPATERNO"
         Caption         =   "VCHAPELLIDOPATERNO"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "VCHAPELLIDOMATERNO"
         Caption         =   "VCHAPELLIDOMATERNO"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   200
         Name            =   "VCHNOMBRE"
         Caption         =   "VCHNOMBRE"
      EndProperty
      BeginProperty Field5 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVEDEPARTAMENTO"
         Caption         =   "SMICVEDEPARTAMENTO"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHPASSWORD"
         Caption         =   "VCHPASSWORD"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITRESPONSABLE"
         Caption         =   "BITRESPONSABLE"
      EndProperty
      BeginProperty Field8 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITACTIVO"
         Caption         =   "BITACTIVO"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRTECNICOIMAGEN"
         Caption         =   "CHRTECNICOIMAGEN"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRSEXO"
         Caption         =   "CHRSEXO"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   129
         Name            =   "CHRRFC"
         Caption         =   "CHRRFC"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   129
         Name            =   "CHRDIRECCION"
         Caption         =   "CHRDIRECCION"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   129
         Name            =   "CHRCOLONIA"
         Caption         =   "CHRCOLONIA"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "CHRCODIGOPOSTAL"
         Caption         =   "CHRCODIGOPOSTAL"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "CHRTELEFONO"
         Caption         =   "CHRTELEFONO"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAALTA"
         Caption         =   "DTMFECHAALTA"
      EndProperty
      BeginProperty Field17 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHABAJA"
         Caption         =   "DTMFECHABAJA"
      EndProperty
      BeginProperty Field18 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYESTADOCIVIL"
         Caption         =   "TNYESTADOCIVIL"
      EndProperty
      BeginProperty Field19 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEPUESTO"
         Caption         =   "INTCVEPUESTO"
      EndProperty
      BeginProperty Field20 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEESTADO"
         Caption         =   "INTCVEESTADO"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   129
         Name            =   "CHRMUNICIPIO"
         Caption         =   "CHRMUNICIPIO"
      EndProperty
      BeginProperty Field22 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSALARIODIARIO"
         Caption         =   "MNYSALARIODIARIO"
      EndProperty
      BeginProperty Field23 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSDI"
         Caption         =   "MNYSDI"
      EndProperty
      BeginProperty Field24 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYTIPOPAGOIMSS"
         Caption         =   "TNYTIPOPAGOIMSS"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   129
         Name            =   "CHRNUMSEG"
         Caption         =   "CHRNUMSEG"
      EndProperty
      BeginProperty Field26 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTNUMCTABANCO"
         Caption         =   "INTNUMCTABANCO"
      EndProperty
      BeginProperty Field27 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEBANCO"
         Caption         =   "INTCVEBANCO"
      EndProperty
      BeginProperty Field28 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTTURNO"
         Caption         =   "INTTURNO"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   129
         Name            =   "CHRAFORE"
         Caption         =   "CHRAFORE"
      EndProperty
      BeginProperty Field30 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYTIPOCONTRATO"
         Caption         =   "TNYTIPOCONTRATO"
      EndProperty
      BeginProperty Field31 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSALDOINFONAVIT"
         Caption         =   "MNYSALDOINFONAVIT"
      EndProperty
      BeginProperty Field32 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYFORMAPAGO"
         Caption         =   "TNYFORMAPAGO"
      EndProperty
      BeginProperty Field33 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "INTNUMCREDITOINFONAVIT"
         Caption         =   "INTNUMCREDITOINFONAVIT"
      EndProperty
      BeginProperty Field34 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHANACIMIENTO"
         Caption         =   "DTMFECHANACIMIENTO"
      EndProperty
      BeginProperty Field35 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSALARIOANTERIOR"
         Caption         =   "MNYSALARIOANTERIOR"
      EndProperty
      BeginProperty Field36 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "VCHNOMPADRE"
         Caption         =   "VCHNOMPADRE"
      EndProperty
      BeginProperty Field37 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "VCHNOMMADRE"
         Caption         =   "VCHNOMMADRE"
      EndProperty
      BeginProperty Field38 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYSALDODEUDORES"
         Caption         =   "MNYSALDODEUDORES"
      EndProperty
      BeginProperty Field39 
         Precision       =   0
         Size            =   80
         Scale           =   0
         Type            =   200
         Name            =   "VCHLUGARNACIMIENTO"
         Caption         =   "VCHLUGARNACIMIENTO"
      EndProperty
      BeginProperty Field40 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHCURP"
         Caption         =   "VCHCURP"
      EndProperty
      BeginProperty Field41 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITGASTOSMEDICOS"
         Caption         =   "BITGASTOSMEDICOS"
      EndProperty
      BeginProperty Field42 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRBAJAIMSS"
         Caption         =   "CHRBAJAIMSS"
      EndProperty
      BeginProperty Field43 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRREINGIMSS"
         Caption         =   "CHRREINGIMSS"
      EndProperty
      BeginProperty Field44 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "TNYTIPOSALARIO"
         Caption         =   "TNYTIPOSALARIO"
      EndProperty
      BeginProperty Field45 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRMODIFIMSS"
         Caption         =   "CHRMODIFIMSS"
      EndProperty
      BeginProperty Field46 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHAMODIF"
         Caption         =   "DTMFECHAMODIF"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset25 
      CommandName     =   "cmdIvSelRequisDetalleDatos"
      CommDispId      =   1252
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELREQUISDETALLEDATOS"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_IVSELREQUISDETALLEDATOS( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_NUMREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset26 
      CommandName     =   "cmdRptFilReqxMixxDepto_Emp"
      CommDispId      =   1255
      RsDispId        =   1257
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto_Emp"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto_Emp( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Emp_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset27 
      CommandName     =   "cmdRptFilReqxMixxDepto_Emp_Est"
      CommDispId      =   1258
      RsDispId        =   1260
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto_Emp_Est"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto_Emp_Est( ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Emp_Est_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset28 
      CommandName     =   "cmdRptFilReqxMixxDep_Em_Es_Ti"
      CommDispId      =   1261
      RsDispId        =   1263
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDep_Em_Es_Ti"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDep_Em_Es_Ti( ?, ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDep_Em_Es_Ti_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset29 
      CommandName     =   "cmdRptFilReqxMixxDepto_Emp_Tip"
      CommDispId      =   1264
      RsDispId        =   1266
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto_Emp_Tip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto_Emp_Tip( ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Emp_Tip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset30 
      CommandName     =   "cmdRptFilReqxMixxDepto_Est"
      CommDispId      =   1267
      RsDispId        =   1269
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto_Est"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto_Est( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Est_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset31 
      CommandName     =   "cmdRptFilReqxMixxDepto_Est_Tip"
      CommDispId      =   1270
      RsDispId        =   1272
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto_Est_Tip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto_Est_Tip( ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Est_Tip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset32 
      CommandName     =   "cmdRptFilReqxMixxDepto_Tip"
      CommDispId      =   1273
      RsDispId        =   1275
      CommandText     =   "dbo.sp_IvRptFilReqxMixxDepto_Tip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxDepto_Tip( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxDepto_Tip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset33 
      CommandName     =   "cmdRptFilReqxMixxEmp"
      CommDispId      =   1276
      RsDispId        =   1278
      CommandText     =   "dbo.sp_IvRptFilReqxMixxEmp"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxEmp( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxEmp_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset34 
      CommandName     =   "cmdRptFilReqxMixxEmp_Est"
      CommDispId      =   1279
      RsDispId        =   1281
      CommandText     =   "dbo.sp_IvRptFilReqxMixxEmp_Est"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxEmp_Est( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxEmp_Est_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset35 
      CommandName     =   "cmdRptFilReqxMixxEmp_Est_Tip"
      CommDispId      =   1282
      RsDispId        =   1284
      CommandText     =   "dbo.sp_IvRptFilReqxMixxEmp_Est_Tip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxEmp_Est_Tip( ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxEmp_Est_Tip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset36 
      CommandName     =   "cmdRptFilReqxMixxEmp_Tip"
      CommDispId      =   1285
      RsDispId        =   1287
      CommandText     =   "dbo.sp_IvRptFilReqxMixxEmp_Tip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxEmp_Tip( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxEmp_Tip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset37 
      CommandName     =   "cmdRptFilReqxMixxEst"
      CommDispId      =   1288
      RsDispId        =   1290
      CommandText     =   "dbo.sp_IvRptFilReqxMixxEst"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxEst( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxEst_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset38 
      CommandName     =   "cmdRptFilReqxMixxEst_Tip"
      CommDispId      =   1291
      RsDispId        =   1293
      CommandText     =   "dbo.sp_IvRptFilReqxMixxEst_Tip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxEst_Tip( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxEst_Tip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset39 
      CommandName     =   "cmdRptFilReqxMixxTip"
      CommDispId      =   1294
      RsDispId        =   1296
      CommandText     =   "dbo.sp_IvRptFilReqxMixxTip"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilReqxMixxTip( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilReqxMixxTip_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "chrCveArticulo"
         Caption         =   "chrCveArticulo"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCantidadSolicitada"
         Caption         =   "intCantidadSolicitada"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "vchNombreComercial"
         Caption         =   "vchNombreComercial"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusDetRequis"
         Caption         =   "vchEstatusDetRequis"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   28
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "numNumRequisicion"
         Caption         =   "numNumRequisicion"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmpleado"
         Caption         =   "NombreEmpleado"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgente"
         Caption         =   "bitUrgente"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequisicion"
         Caption         =   "dtmFechaRequisicion"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Destino"
         Caption         =   "Destino"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@TipoRequis"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset40 
      CommandName     =   "cmdSelRequisCDMaestro"
      CommDispId      =   1297
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELREQUISCDMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      Locktype        =   3
      CallSyntax      =   "{ CALL DBO.SP_IVSELREQUISCDMAESTRO( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset41 
      CommandName     =   "cmdSelConsRCDMaestro"
      CommDispId      =   1300
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELCONSRCDMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      Locktype        =   3
      CallSyntax      =   "{ CALL DBO.SP_IVSELCONSRCDMAESTRO( ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "IN_ESTATUSCD"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEDEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset42 
      CommandName     =   "cmdArticuloCargoDirecto"
      CommDispId      =   1303
      RsDispId        =   1952
      CommandText     =   "DBO.IVARTICULOCARGODIRECTO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEARTICULOCARDIR"
         Caption         =   "INTCVEARTICULOCARDIR"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHDESCRIPCION"
         Caption         =   "VCHDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVETIPOARTCARDIR"
         Caption         =   "SMICVETIPOARTCARDIR"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "BITACTIVO"
         Caption         =   "BITACTIVO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset43 
      CommandName     =   "cmdSelRequisCDDetalle"
      CommDispId      =   1306
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVSELREQUISCDDETALLE"
      ActiveConnectionName=   "ConeccionSIHO"
      Locktype        =   3
      CallSyntax      =   "{ CALL DBO.SP_IVSELREQUISCDDETALLE( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_NUMREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset44 
      CommandName     =   "cmdUpdRequisCDMaestro"
      CommDispId      =   1309
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVUPDREQUISCDMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVUPDREQUISCDMAESTRO( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_NUMREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_ESTATUSART"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset45 
      CommandName     =   "cmdUpdRequisCDDetalle"
      CommDispId      =   1318
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVUPDREQUISCDDETALLE"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVUPDREQUISCDDETALLE( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_NUMREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_ESTATUSART"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset46 
      CommandName     =   "cmdInsReqCDMaestro"
      CommDispId      =   1332
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVINSREQUISCDMAESTRO"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL DBO.SP_IVINSREQUISCDMAESTRO( ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEDEPARTAMENTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_CVEEMPLEADO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_FECHAREQUIS"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_HORAREQUIS"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "IN_URGENTE"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "IN_ESTATUSREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "IN_FECHAAUTORIZA"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset47 
      CommandName     =   "cmdInsReqCDDetalle"
      CommDispId      =   1337
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVINSREQUISCDDETALLE"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVINSREQUISCDDETALLE( ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "IN_NUMREQUISCARDIR"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CVEARTICULOCARDIR"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_DESCLARGA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_CANTIDAD"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_ESTATUSREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset48 
      CommandName     =   "cmdDelRequisCDDetalle"
      CommDispId      =   1342
      RsDispId        =   -1
      CommandText     =   "DBO.SP_IVDELREQUISCDDETALLE"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{ CALL DBO.SP_IVDELREQUISCDDETALLE( ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "IN_NUMREQUIS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset49 
      CommandName     =   "cmdRptFilRequisCDxDepto_Emp"
      CommDispId      =   1347
      RsDispId        =   1349
      CommandText     =   "dbo.sp_IvRptFilRequisCDxDepto_Emp"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxDepto_Emp( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxDepto_Emp_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset50 
      CommandName     =   "cmdRptFilRequisCDxDepto_Emp_Est"
      CommDispId      =   1350
      RsDispId        =   1353
      CommandText     =   "dbo.sp_IvRptFilRequisCDxDepto_Emp_Est"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxDepto_Emp_Est( ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxDepto_Emp_Est_Groupin"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      ParamCount      =   8
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset51 
      CommandName     =   "cmdRptFilRequisCDxDepto_Est"
      CommDispId      =   1354
      RsDispId        =   1356
      CommandText     =   "dbo.sp_IvRptFilRequisCDxDepto_Est"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxDepto_Est( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxDepto_Est_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset52 
      CommandName     =   "cmdRptFilRequisCDxEmp"
      CommDispId      =   1357
      RsDispId        =   1359
      CommandText     =   "dbo.sp_IvRptFilRequisCDxEmp"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxEmp( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxEmp_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset53 
      CommandName     =   "cmdRptFilRequisCDxEmp_Est"
      CommDispId      =   1360
      RsDispId        =   1362
      CommandText     =   "dbo.sp_IvRptFilRequisCDxEmp_Est"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxEmp_Est( ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxEmp_Est_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveEmp"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset54 
      CommandName     =   "cmdRptFilRequisCDxEst"
      CommDispId      =   1363
      RsDispId        =   1365
      CommandText     =   "dbo.sp_IvRptFilRequisCDxEst"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_IvRptFilRequisCDxEst( ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptFilRequisCDxEst_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Field6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Field7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveArticuloCarDir"
         Caption         =   "intCveArticuloCarDir"
      EndProperty
      BeginProperty Field9 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCantidad"
         Caption         =   "smiCantidad"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescArticulo"
         Caption         =   "DescArticulo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "vchDescLarga"
         Caption         =   "vchDescLarga"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "EstatusArt"
         Caption         =   "EstatusArt"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumRequisCarDir"
         Caption         =   "intNumRequisCarDir"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   57
         Scale           =   0
         Type            =   200
         Name            =   "NombreEmp"
         Caption         =   "NombreEmp"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUrgenteRequis"
         Caption         =   "bitUrgenteRequis"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchEstatusRequis"
         Caption         =   "vchEstatusRequis"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaRequis"
         Caption         =   "dtmFechaRequis"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaAutoriza"
         Caption         =   "dtmFechaAutoriza"
      EndProperty
      ParamCount      =   6
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Estatus"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FechaInicio"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FolioInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FolioFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset55 
      CommandName     =   "cmdPvSelCargosPaciente"
      CommDispId      =   1376
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVSELCARGOSPACIENTE"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_PVSELCARGOSPACIENTE( ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "IN_NUMEROCUENTA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_TIPOPACIENTE"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUSFACTURADOS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_FACTURA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset56 
      CommandName     =   "cmdCuentas"
      CommDispId      =   1391
      RsDispId        =   1393
      CommandText     =   $"EntornoSIHO.dsx":015E
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumeroCuenta"
         Caption         =   "intNumeroCuenta"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   205
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcionCuenta"
         Caption         =   "vchDescripcionCuenta"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset57 
      CommandName     =   "cmdConceptos"
      CommDispId      =   1394
      RsDispId        =   1864
      CommandText     =   "DBO.PVCONCEPTOFACTURACION"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   12
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVECONCEPTO"
         Caption         =   "SMICVECONCEPTO"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   70
         Scale           =   0
         Type            =   129
         Name            =   "CHRDESCRIPCION"
         Caption         =   "CHRDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCUENTACONTABLE"
         Caption         =   "INTCUENTACONTABLE"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "SMYIVA"
         Caption         =   "SMYIVA"
      EndProperty
      BeginProperty Field5 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMIDEPARTAMENTO"
         Caption         =   "SMIDEPARTAMENTO"
      EndProperty
      BeginProperty Field6 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITACTIVO"
         Caption         =   "BITACTIVO"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCUENTADESCUENTO"
         Caption         =   "INTCUENTADESCUENTO"
      EndProperty
      BeginProperty Field8 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCUENTAIVAXPAGAR"
         Caption         =   "INTCUENTAIVAXPAGAR"
      EndProperty
      BeginProperty Field9 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCUENTAIVANOPAGADO"
         Caption         =   "INTCUENTAIVANOPAGADO"
      EndProperty
      BeginProperty Field10 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITFACTURADEDUCIBLE"
         Caption         =   "BITFACTURADEDUCIBLE"
      EndProperty
      BeginProperty Field11 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITFACTURACOASEGURO"
         Caption         =   "BITFACTURACOASEGURO"
      EndProperty
      BeginProperty Field12 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITFACTURACOPAGO"
         Caption         =   "BITFACTURACOPAGO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset58 
      CommandName     =   "cmdConceptoDepartamento"
      CommDispId      =   1397
      RsDispId        =   1965
      CommandText     =   "DBO.PVCONCEPTODEPARTAMENTO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVECONCEPTO"
         Caption         =   "SMICVECONCEPTO"
      EndProperty
      BeginProperty Field2 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMIDEPARTAMENTO"
         Caption         =   "SMIDEPARTAMENTO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset59 
      CommandName     =   "cmdPaquetes"
      CommDispId      =   1400
      RsDispId        =   1966
      CommandText     =   "DBO.PVPAQUETE"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTNUMPAQUETE"
         Caption         =   "INTNUMPAQUETE"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "CHRDESCRIPCION"
         Caption         =   "CHRDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICONCEPTOFACTURA"
         Caption         =   "SMICONCEPTOFACTURA"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRINTERNOEXTERNO"
         Caption         =   "CHRINTERNOEXTERNO"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "CHRTRATAMIENTO"
         Caption         =   "CHRTRATAMIENTO"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPO"
         Caption         =   "CHRTIPO"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYANTICIPOSUGERIDO"
         Caption         =   "MNYANTICIPOSUGERIDO"
      EndProperty
      BeginProperty Field8 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITACTIVO"
         Caption         =   "BITACTIVO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset60 
      CommandName     =   "cmdDetallePaquete"
      CommDispId      =   1403
      RsDispId        =   1967
      CommandText     =   "DBO.PVDETALLEPAQUETE"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTNUMPAQUETE"
         Caption         =   "INTNUMPAQUETE"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "CHRCVECARGO"
         Caption         =   "CHRCVECARGO"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPOCARGO"
         Caption         =   "CHRTIPOCARGO"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICANTIDAD"
         Caption         =   "SMICANTIDAD"
      EndProperty
      BeginProperty Field5 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYPRECIO"
         Caption         =   "MNYPRECIO"
      EndProperty
      BeginProperty Field6 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYIVA"
         Caption         =   "MNYIVA"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTDESCUENTOINVENTARIO"
         Caption         =   "INTDESCUENTOINVENTARIO"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPOUSO"
         Caption         =   "CHRTIPOUSO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset61 
      CommandName     =   "cmdDenominaciones"
      CommDispId      =   1406
      RsDispId        =   1968
      CommandText     =   "DBO.PVDENOMINACION"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEDENOMINACION"
         Caption         =   "INTCVEDENOMINACION"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "CHRDESCRIPCION"
         Caption         =   "CHRDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYMONTO"
         Caption         =   "MNYMONTO"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITESTATUS"
         Caption         =   "BITESTATUS"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset62 
      CommandName     =   "cmdConceptoFacturacion"
      CommDispId      =   1412
      RsDispId        =   1414
      CommandText     =   "SELECT smiCveConcepto, chrDescripcion FROM PvConceptoFacturacion WHERE bitactivo=1"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiCveConcepto"
         Caption         =   "smiCveConcepto"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrDescripcion"
         Caption         =   "chrDescripcion"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset63 
      CommandName     =   "cmdFormaPago"
      CommDispId      =   1415
      RsDispId        =   1969
      CommandText     =   "DBO.PVFORMAPAGO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTFORMAPAGO"
         Caption         =   "INTFORMAPAGO"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "CHRDESCRIPCION"
         Caption         =   "CHRDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITCREDITO"
         Caption         =   "BITCREDITO"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITCOMISIONBANCO"
         Caption         =   "BITCOMISIONBANCO"
      EndProperty
      BeginProperty Field5 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "SMYCANTIDADBANCO"
         Caption         =   "SMYCANTIDADBANCO"
      EndProperty
      BeginProperty Field6 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITESTATUSACTIVO"
         Caption         =   "BITESTATUSACTIVO"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCUENTACONTABLE"
         Caption         =   "INTCUENTACONTABLE"
      EndProperty
      BeginProperty Field8 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITPREGUNTAFOLIO"
         Caption         =   "BITPREGUNTAFOLIO"
      EndProperty
      BeginProperty Field9 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITPESOS"
         Caption         =   "BITPESOS"
      EndProperty
      BeginProperty Field10 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMIDEPARTAMENTO"
         Caption         =   "SMIDEPARTAMENTO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset64 
      CommandName     =   "cmdDescuentos"
      CommDispId      =   1418
      RsDispId        =   1970
      CommandText     =   "DBO.PVDESCUENTO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPODESCUENTO"
         Caption         =   "CHRTIPODESCUENTO"
      EndProperty
      BeginProperty Field2 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEAFECTADA"
         Caption         =   "INTCVEAFECTADA"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPOPACIENTE"
         Caption         =   "CHRTIPOPACIENTE"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "SMICVECONCEPTO"
         Caption         =   "SMICVECONCEPTO"
      EndProperty
      BeginProperty Field5 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVECARGO"
         Caption         =   "INTCVECARGO"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPOCARGO"
         Caption         =   "CHRTIPOCARGO"
      EndProperty
      BeginProperty Field7 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYDESCUENTO"
         Caption         =   "MNYDESCUENTO"
      EndProperty
      BeginProperty Field8 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITPORCENTAJE"
         Caption         =   "BITPORCENTAJE"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset65 
      CommandName     =   "cmdPvDetalleLista"
      CommDispId      =   1421
      RsDispId        =   1971
      CommandText     =   "DBO.PVDETALLELISTA"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVELISTA"
         Caption         =   "INTCVELISTA"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "CHRCVECARGO"
         Caption         =   "CHRCVECARGO"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   20
         Scale           =   0
         Type            =   139
         Name            =   "MNYPRECIO"
         Caption         =   "MNYPRECIO"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "CHRTIPOCARGO"
         Caption         =   "CHRTIPOCARGO"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset66 
      CommandName     =   "cmdElementosListaPrecios"
      CommDispId      =   1425
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVSELELEMENTOSLISTASPRECIOS"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_PVSELELEMENTOSLISTASPRECIOS( ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "IN_INTDEPARTAMENTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CUALLISTA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset67 
      CommandName     =   "cmdPresupuesto"
      CommDispId      =   1441
      RsDispId        =   1443
      CommandText     =   "dbo.PVImprimePresupuesto"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Grouping        =   -1  'True
      GroupingName    =   "cmdPresupuesto_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   18
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrEmpresa"
         Caption         =   "chrEmpresa"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrNombre"
         Caption         =   "chrNombre"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrDireccion"
         Caption         =   "chrDireccion"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrCargo"
         Caption         =   "chrCargo"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyPrecio"
         Caption         =   "mnyPrecio"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyCantidad"
         Caption         =   "mnyCantidad"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnySubtotal"
         Caption         =   "mnySubtotal"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyDescuento"
         Caption         =   "mnyDescuento"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyMonto"
         Caption         =   "mnyMonto"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTSubtotal"
         Caption         =   "mnyTSubtotal"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTDescuento"
         Caption         =   "mnyTDescuento"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTIva"
         Caption         =   "mnyTIva"
      EndProperty
      BeginProperty Field13 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTTotal"
         Caption         =   "mnyTTotal"
      EndProperty
      BeginProperty Field14 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "chrNombreHospital"
         Caption         =   "chrNombreHospital"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "chrDireccionHospital"
         Caption         =   "chrDireccionHospital"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   129
         Name            =   "chrTelefonoHospital"
         Caption         =   "chrTelefonoHospital"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   129
         Name            =   "chrRFCHospital"
         Caption         =   "chrRFCHospital"
      EndProperty
      BeginProperty Field18 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaHora"
         Caption         =   "dtmFechaHora"
      EndProperty
      NumGroups       =   12
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrEmpresa"
         Caption         =   "chrEmpresa"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrNombre"
         Caption         =   "chrNombre"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrDireccion"
         Caption         =   "chrDireccion"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTSubtotal"
         Caption         =   "mnyTSubtotal"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTDescuento"
         Caption         =   "mnyTDescuento"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTIva"
         Caption         =   "mnyTIva"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTTotal"
         Caption         =   "mnyTTotal"
      EndProperty
      BeginProperty Grouping8 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "chrNombreHospital"
         Caption         =   "chrNombreHospital"
      EndProperty
      BeginProperty Grouping9 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "chrDireccionHospital"
         Caption         =   "chrDireccionHospital"
      EndProperty
      BeginProperty Grouping10 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   129
         Name            =   "chrTelefonoHospital"
         Caption         =   "chrTelefonoHospital"
      EndProperty
      BeginProperty Grouping11 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   129
         Name            =   "chrRFCHospital"
         Caption         =   "chrRFCHospital"
      EndProperty
      BeginProperty Grouping12 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaHora"
         Caption         =   "dtmFechaHora"
      EndProperty
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset68 
      CommandName     =   "cmdSeleccionaCargos"
      CommDispId      =   1444
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVSELCARGOSPACIENTE"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_PVSELCARGOSPACIENTE( ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "IN_NUMEROCUENTA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_TIPOPACIENTE"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUSFACTURADOS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_FACTURA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset69 
      CommandName     =   "cmdBorrada"
      CommDispId      =   1448
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVUPDBORRACARGO"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL DBO.SP_PVUPDBORRACARGO( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_NUMCARGO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_TABLARELACION"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_EMPLEADOCANCELA"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_CANCELADEPARTAMENTO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "IN_CHRTIPODOCUMENTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "IN_INTFOLIODOCUMENTO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "IN_INTCVECARGO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "IN_CHRTIPOCARGO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "IN_SEGUNDOSREINTENTAR"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset70 
      CommandName     =   "cmdPvUpdEstatusCorte"
      CommDispId      =   1457
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVUPDESTATUSCORTE"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL DBO.SP_PVUPDESTATUSCORTE( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_NUMERODEPARTAMENTO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_ESTATUS"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset71 
      CommandName     =   "cmdGnInsCargoListasPrecios"
      CommDispId      =   1459
      RsDispId        =   -1
      CommandText     =   "DBO.SP_GNINSCARGOLISTASPRECIOS"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      Locktype        =   3
      CallSyntax      =   "{ CALL DBO.SP_GNINSCARGOLISTASPRECIOS( ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "IN_DEPTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_CLAVECARGO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_TIPOCARGO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset72 
      CommandName     =   "cmdCnSelPolizasColumnas"
      CommDispId      =   1464
      RsDispId        =   -1
      CommandText     =   "DBO.SP_CNSELPOLIZASCOLUMNAS"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_CNSELPOLIZASCOLUMNAS( ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "IN_NUMEROINICIAL"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_NUMEROFINAL"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_TODAS"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_TIPO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset73 
      CommandName     =   "cmdCuentasDescuentos"
      CommDispId      =   1467
      RsDispId        =   1479
      CommandText     =   $"EntornoSIHO.dsx":0297
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumeroCuenta"
         Caption         =   "intNumeroCuenta"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   205
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcionCuenta"
         Caption         =   "vchDescripcionCuenta"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset74 
      CommandName     =   "cmdCuentasIvaxPagar"
      CommDispId      =   1473
      RsDispId        =   1480
      CommandText     =   $"EntornoSIHO.dsx":03D2
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset75 
      CommandName     =   "cmdCargos"
      CommDispId      =   1529
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVUPDCARGOS"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL DBO.SP_PVUPDCARGOS( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   15
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   0
         Scale           =   0
         Size            =   2147483647
         DataType        =   200
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_INTCVECARGO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_CVEDEPARTAMENTO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_CHRTIPODOCUMENTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_INTFOLIODOCUMENTO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "IN_INTMOVPACIENTE"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "IN_CHRINTERNOEXTERNO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "IN_CHRTIPOCARGO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "IN_BITMEDICAMENTOAPLICADO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "IN_MNYCANTIDAD"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "IN_INTEMPLEADO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "IN_DESCUENTAINVENTARIO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "IN_TABLARELACION"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "IN_NUMREFERENCIA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P15 
         RealName        =   "IN_SEGUNDOSREINTENTAR"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset76 
      CommandName     =   "cmdReporteListaPrecios"
      CommDispId      =   1531
      RsDispId        =   1533
      CommandText     =   "dbo.sp_PvselElementosListasPrecios"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvselElementosListasPrecios( ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdReporteListaPrecios_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "Clave"
         Caption         =   "Clave"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Descripcion"
         Caption         =   "Descripcion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "Tipo"
         Caption         =   "Tipo"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Precio"
         Caption         =   "Precio"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Orden"
         Caption         =   "Orden"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "PrecioCIVA"
         Caption         =   "PrecioCIVA"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "ClaveFamilia"
         Caption         =   "ClaveFamilia"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescripFamilia"
         Caption         =   "DescripFamilia"
      EndProperty
      NumGroups       =   2
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "ClaveFamilia"
         Caption         =   "ClaveFamilia"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "DescripFamilia"
         Caption         =   "DescripFamilia"
      EndProperty
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@intDepartamento"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CualLista"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset77 
      CommandName     =   "cmdCcSelHonorarioCredito"
      CommDispId      =   1537
      RsDispId        =   1539
      CommandText     =   "dbo.sp_CcSelHonorarioCredito"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_CcSelHonorarioCredito( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   13
      BeginProperty Field1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "NombreRecibo"
         Caption         =   "NombreRecibo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Direccion"
         Caption         =   "Direccion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "RFC"
         Caption         =   "RFC"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   67
         Scale           =   0
         Type            =   200
         Name            =   "Medico"
         Caption         =   "Medico"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "Paciente"
         Caption         =   "Paciente"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Concepto"
         Caption         =   "Concepto"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Monto"
         Caption         =   "Monto"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Retencion"
         Caption         =   "Retencion"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "NetoPagar"
         Caption         =   "NetoPagar"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "Cuarto"
         Caption         =   "Cuarto"
      EndProperty
      BeginProperty Field11 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaAtencion"
         Caption         =   "FechaAtencion"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Convenio"
         Caption         =   "Convenio"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "Factura"
         Caption         =   "Factura"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Consecutivo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset78 
      CommandName     =   "cmdejRegresaRsReadOnly"
      CommDispId      =   1544
      RsDispId        =   -1
      CommandText     =   "DBO.SP_EJECUTASENTENCIA"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL DBO.SP_EJECUTASENTENCIA( ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_SENTENCIA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset79 
      CommandName     =   "cmdExterno"
      CommDispId      =   1549
      RsDispId        =   1984
      CommandText     =   "DBO.EXTERNO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Locktype        =   3
      IsRSReturning   =   -1  'True
      NumFields       =   29
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTNUMPACIENTE"
         Caption         =   "INTNUMPACIENTE"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   129
         Name            =   "CHRAPEPATERNO"
         Caption         =   "CHRAPEPATERNO"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   129
         Name            =   "CHRAPEMATERNO"
         Caption         =   "CHRAPEMATERNO"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   129
         Name            =   "CHRNOMBRE"
         Caption         =   "CHRNOMBRE"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   16
         Scale           =   0
         Type            =   135
         Name            =   "DTMFECHANAC"
         Caption         =   "DTMFECHANAC"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "CHRSEXO"
         Caption         =   "CHRSEXO"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   129
         Name            =   "CHRRFC"
         Caption         =   "CHRRFC"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHCURP"
         Caption         =   "VCHCURP"
      EndProperty
      BeginProperty Field9 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVENACIONALIDAD"
         Caption         =   "INTCVENACIONALIDAD"
      EndProperty
      BeginProperty Field10 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYESTADOCIVIL"
         Caption         =   "TNYESTADOCIVIL"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   129
         Name            =   "CHRDIRECCION"
         Caption         =   "CHRDIRECCION"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "VCHCOLONIA"
         Caption         =   "VCHCOLONIA"
      EndProperty
      BeginProperty Field13 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCIUDAD"
         Caption         =   "INTCIUDAD"
      EndProperty
      BeginProperty Field14 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEESTADO"
         Caption         =   "INTCVEESTADO"
      EndProperty
      BeginProperty Field15 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEPAIS"
         Caption         =   "INTCVEPAIS"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "VCHCODPOSTAL"
         Caption         =   "VCHCODPOSTAL"
      EndProperty
      BeginProperty Field17 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   129
         Name            =   "CHRTELEFONO"
         Caption         =   "CHRTELEFONO"
      EndProperty
      BeginProperty Field18 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYTIPOPACIENTE"
         Caption         =   "TNYTIPOPACIENTE"
      EndProperty
      BeginProperty Field19 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYTIPOCONVENIO"
         Caption         =   "TNYTIPOCONVENIO"
      EndProperty
      BeginProperty Field20 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCLAVEEMPRESA"
         Caption         =   "INTCLAVEEMPRESA"
      EndProperty
      BeginProperty Field21 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHOCUPACION"
         Caption         =   "VCHOCUPACION"
      EndProperty
      BeginProperty Field22 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHLUGARTRABAJO"
         Caption         =   "VCHLUGARTRABAJO"
      EndProperty
      BeginProperty Field23 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHDOMICILIOTRABAJO"
         Caption         =   "VCHDOMICILIOTRABAJO"
      EndProperty
      BeginProperty Field24 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHREFERENCIA"
         Caption         =   "VCHREFERENCIA"
      EndProperty
      BeginProperty Field25 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "VCHRESPONSABLE"
         Caption         =   "VCHRESPONSABLE"
      EndProperty
      BeginProperty Field26 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITSTATUSCARGO"
         Caption         =   "BITSTATUSCARGO"
      EndProperty
      BeginProperty Field27 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "INTCVEEXTRA"
         Caption         =   "INTCVEEXTRA"
      EndProperty
      BeginProperty Field28 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITVIVO"
         Caption         =   "BITVIVO"
      EndProperty
      BeginProperty Field29 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "VCHNUMAFILIACION"
         Caption         =   "VCHNUMAFILIACION"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset80 
      CommandName     =   "cmdSelDatosPExternos"
      CommDispId      =   1553
      RsDispId        =   1555
      CommandText     =   $"EntornoSIHO.dsx":0507
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumPaciente"
         Caption         =   "intNumPaciente"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   129
         Name            =   "chrApePaterno"
         Caption         =   "chrApePaterno"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   129
         Name            =   "chrApeMaterno"
         Caption         =   "chrApeMaterno"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   40
         Scale           =   0
         Type            =   129
         Name            =   "chrNombre"
         Caption         =   "chrNombre"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   129
         Name            =   "chrDireccion"
         Caption         =   "chrDireccion"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   129
         Name            =   "chrTelefono"
         Caption         =   "chrTelefono"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   129
         Name            =   "chrRFC"
         Caption         =   "chrRFC"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaNac"
         Caption         =   "dtmFechaNac"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset81 
      CommandName     =   "cmdTipoConvenio"
      CommDispId      =   1556
      RsDispId        =   1985
      CommandText     =   "DBO.CCTIPOCONVENIO"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TNYCVETIPOCONVENIO"
         Caption         =   "TNYCVETIPOCONVENIO"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "VCHDESCRIPCION"
         Caption         =   "VCHDESCRIPCION"
      EndProperty
      BeginProperty Field3 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITASEGURADORA"
         Caption         =   "BITASEGURADORA"
      EndProperty
      BeginProperty Field4 
         Precision       =   38
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "BITPENSIONES"
         Caption         =   "BITPENSIONES"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset82 
      CommandName     =   "cmdSelTicket"
      CommDispId      =   1559
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVSELTICKET"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_PVSELTICKET( ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "IN_INTCVEVENTA"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "RC1"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   13
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset83 
      CommandName     =   "cmdPvSelPreparaEstadoCuenta"
      CommDispId      =   1583
      RsDispId        =   -1
      CommandText     =   "dbo.sp_PvSelPreparaEstadoCuenta"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelPreparaEstadoCuenta( ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Cuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Orden"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TodosCargos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@Excluidos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@Departamento"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset84 
      CommandName     =   "cmdPrecios"
      CommDispId      =   1595
      RsDispId        =   -1
      CommandText     =   "DBO.SP_PVSELOBTENERPRECIO"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_PVSELOBTENERPRECIO( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "IN_CVECARGO"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_TIPOCARGO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_TIPOPACIENTE"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "IN_EMPRESA"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "IN_CHRTIPOINTERNOEXTERNO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "IN_URGENTE"
         Direction       =   1
         Precision       =   38
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "IN_DTMFECHA"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "IN_PRECIO"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "IN_INCREMENTO"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset85 
      CommandName     =   "cmdpvSelPoliza"
      CommDispId      =   1598
      RsDispId        =   1602
      CommandText     =   "dbo.sp_pvSelPoliza"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_pvSelPoliza( ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdpvSelPoliza_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "CuentaMayor"
         Caption         =   "CuentaMayor"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "NombreMayor"
         Caption         =   "NombreMayor"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchNombre"
         Caption         =   "vchNombre"
      EndProperty
      BeginProperty Field4 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "smiEjercicio"
         Caption         =   "smiEjercicio"
      EndProperty
      BeginProperty Field5 
         Precision       =   3
         Size            =   1
         Scale           =   0
         Type            =   17
         Name            =   "tnyMes"
         Caption         =   "tnyMes"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intClavePoliza"
         Caption         =   "intClavePoliza"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "vchNumero"
         Caption         =   "vchNumero"
      EndProperty
      BeginProperty Field8 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaPoliza"
         Caption         =   "dtmFechaPoliza"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchConceptoPoliza"
         Caption         =   "vchConceptoPoliza"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "TipoPoliza"
         Caption         =   "TipoPoliza"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcion"
         Caption         =   "vchDescripcion"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchCuentaContable"
         Caption         =   "vchCuentaContable"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchDescripcionCuenta"
         Caption         =   "vchDescripcionCuenta"
      EndProperty
      BeginProperty Field14 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Cargo"
         Caption         =   "Cargo"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Abono"
         Caption         =   "Abono"
      EndProperty
      NumGroups       =   2
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "CuentaMayor"
         Caption         =   "CuentaMayor"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "NombreMayor"
         Caption         =   "NombreMayor"
      EndProperty
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroPoliza"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset86 
      CommandName     =   "cmdPvSelReporteIngresos"
      CommDispId      =   1606
      RsDispId        =   1610
      CommandText     =   "dbo.sp_pvSelReporteIngresos"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_pvSelReporteIngresos( ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdPvSelReporteIngresos_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "chrTipoPaciente"
         Caption         =   "chrTipoPaciente"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "Campo1"
         Caption         =   "Campo1"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   133
         Scale           =   0
         Type            =   129
         Name            =   "Campo2"
         Caption         =   "Campo2"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "MontoInterno"
         Caption         =   "MontoInterno"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "MontoExterno"
         Caption         =   "MontoExterno"
      EndProperty
      NumGroups       =   2
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "chrTipoPaciente"
         Caption         =   "chrTipoPaciente"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "Campo1"
         Caption         =   "Campo1"
      EndProperty
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@TipoReporte"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   0
         DataType        =   17
         HostType        =   17
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@InternoExterno"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@CveConcepto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset87 
      CommandName     =   "cmdPvSelFacturasCanceladas"
      CommDispId      =   1629
      RsDispId        =   1631
      CommandText     =   "dbo.sp_PvSelFacturasCanceladas"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelFacturasCanceladas( ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaFactura"
         Caption         =   "FechaFactura"
      EndProperty
      BeginProperty Field2 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaCancelacion"
         Caption         =   "FechaCancelacion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "Factura"
         Caption         =   "Factura"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Cuenta"
         Caption         =   "Cuenta"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "TipoIE"
         Caption         =   "TipoIE"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "NombrePaciente"
         Caption         =   "NombrePaciente"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "TipoEmpresa"
         Caption         =   "TipoEmpresa"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Importe"
         Caption         =   "Importe"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "Moneda"
         Caption         =   "Moneda"
      EndProperty
      NumGroups       =   0
      ParamCount      =   10
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@TodosIE"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoIE"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TodosTipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@CveTipoPaciente"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   0
         DataType        =   17
         HostType        =   17
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@TodasEmpresas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@CveEmpresa"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@FechaInicial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@FechaFinal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@CveDepartamento"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset88 
      CommandName     =   "cmdIvSelArticulosUbicacion"
      CommDispId      =   1661
      RsDispId        =   1663
      CommandText     =   "dbo.sp_IvSelArticulosUbicacion"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_IvSelArticulosUbicacion( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   129
         Name            =   "Clave"
         Caption         =   "Clave"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "NombreComercial"
         Caption         =   "NombreComercial"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Maximo"
         Caption         =   "Maximo"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Minimo"
         Caption         =   "Minimo"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ExistenciaUV"
         Caption         =   "ExistenciaUV"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "ExistenciaUM"
         Caption         =   "ExistenciaUM"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Contenido"
         Caption         =   "Contenido"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Localizacion"
         Caption         =   "Localizacion"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CostoPromedio"
         Caption         =   "CostoPromedio"
      EndProperty
      NumGroups       =   0
      ParamCount      =   14
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TodosTipos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Tipo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TodasLocalizaciones"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveLocalizacion"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@TodasFamilias"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@CveFamilia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@TodasSubfamilias"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@CveSubfamilia"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   3
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@TodosArticulos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@CveArticulo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   10
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P13 
         RealName        =   "@Controlado"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P14 
         RealName        =   "@Refrigerado"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset89 
      CommandName     =   "cmdIvUpdEstatusReqMaestro"
      CommDispId      =   1675
      RsDispId        =   1678
      CommandText     =   "dbo.sp_IvUpdEstatusReqMaestro"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      Locktype        =   3
      CallSyntax      =   "{? = CALL dbo.sp_IvUpdEstatusReqMaestro( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumRequis"
         Direction       =   1
         Precision       =   18
         Scale           =   0
         Size            =   0
         DataType        =   131
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset90 
      CommandName     =   "cmdRptListaPrecio"
      CommDispId      =   1679
      RsDispId        =   1681
      CommandText     =   "dbo.PvRptListaPrecio"
      ActiveConnectionName=   "ConeccionSIHO"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptListaPrecio_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchConcepto"
         Caption         =   "vchConcepto"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchCargo"
         Caption         =   "vchCargo"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyPrecioSinIva"
         Caption         =   "mnyPrecioSinIva"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyPrecioConIva"
         Caption         =   "mnyPrecioConIva"
      EndProperty
      BeginProperty Field5 
         Precision       =   7
         Size            =   4
         Scale           =   0
         Type            =   4
         Name            =   "relPorcentajeDesc"
         Caption         =   "relPorcentajeDesc"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyDescSinIva"
         Caption         =   "mnyDescSinIva"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyDescConIva"
         Caption         =   "mnyDescConIva"
      EndProperty
      NumGroups       =   1
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "vchConcepto"
         Caption         =   "vchConcepto"
      EndProperty
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset91 
      CommandName     =   "cmdPVSelHonorario"
      CommDispId      =   1682
      RsDispId        =   1684
      CommandText     =   "dbo.sp_PvSelHonorario"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CursorType      =   0
      CallSyntax      =   "{? = CALL dbo.sp_PvSelHonorario( ?) }"
      CommandCursorLocation=   2
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "NombreRecibo"
         Caption         =   "NombreRecibo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "Direccion"
         Caption         =   "Direccion"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "RFC"
         Caption         =   "RFC"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   67
         Scale           =   0
         Type            =   200
         Name            =   "Medico"
         Caption         =   "Medico"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   202
         Name            =   "Paciente"
         Caption         =   "Paciente"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "Concepto"
         Caption         =   "Concepto"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Monto"
         Caption         =   "Monto"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Retencion"
         Caption         =   "Retencion"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Comision"
         Caption         =   "Comision"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "IvaComision"
         Caption         =   "IvaComision"
      EndProperty
      BeginProperty Field11 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Subtotal"
         Caption         =   "Subtotal"
      EndProperty
      BeginProperty Field12 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "NetoPagar"
         Caption         =   "NetoPagar"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   202
         Name            =   "Cuarto"
         Caption         =   "Cuarto"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaAtencion"
         Caption         =   "FechaAtencion"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Convenio"
         Caption         =   "Convenio"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Consecutivo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset92 
      CommandName     =   "cmdPvSelDatosPaciente"
      CommDispId      =   1685
      RsDispId        =   1687
      CommandText     =   "dbo.sp_PvSelDatosPaciente"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_PvSelDatosPaciente( ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "Nombre"
         Caption         =   "Nombre"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Empresa"
         Caption         =   "Empresa"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   45
         Scale           =   0
         Type            =   200
         Name            =   "Tipo"
         Caption         =   "Tipo"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "Ingreso"
         Caption         =   "Ingreso"
      EndProperty
      BeginProperty Field5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "Egreso"
         Caption         =   "Egreso"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitUtilizaConvenio"
         Caption         =   "bitUtilizaConvenio"
      EndProperty
      BeginProperty Field7 
         Precision       =   3
         Size            =   1
         Scale           =   0
         Type            =   17
         Name            =   "TipoPaciente"
         Caption         =   "TipoPaciente"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveEmpresa"
         Caption         =   "intCveEmpresa"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCveExtra"
         Caption         =   "intCveExtra"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaCerrada"
         Caption         =   "CuentaCerrada"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "Facturada"
         Caption         =   "Facturada"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroCuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset93 
      CommandName     =   "cmdRptReciboComisiones"
      CommDispId      =   1688
      RsDispId        =   1692
      CommandText     =   "dbo.sp_PvSelComision"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelComision( ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdRptReciboComisiones_Grouping"
      DetailExpanded  =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   67
         Scale           =   0
         Type            =   200
         Name            =   "Medico"
         Caption         =   "Medico"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Comision"
         Caption         =   "Comision"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "IvaComision"
         Caption         =   "IvaComision"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "chrDescripcion"
         Caption         =   "chrDescripcion"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyCantidad"
         Caption         =   "mnyCantidad"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyIVA"
         Caption         =   "mnyIVA"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Total"
         Caption         =   "Total"
      EndProperty
      NumGroups       =   1
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   67
         Scale           =   0
         Type            =   200
         Name            =   "Medico"
         Caption         =   "Medico"
      EndProperty
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Consecutivo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset94 
      CommandName     =   "cmdPVSelHorarioEmpresa"
      CommDispId      =   1693
      RsDispId        =   1695
      CommandText     =   "dbo.sp_PvSelHorarioEmpresa"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelHorarioEmpresa( ?, ?, ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Consecutivo"
         Caption         =   "Consecutivo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   16
         Scale           =   0
         Type            =   200
         Name            =   "Tipo"
         Caption         =   "Tipo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Subtipo"
         Caption         =   "Subtipo"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "TipoPaciente"
         Caption         =   "TipoPaciente"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   9
         Scale           =   0
         Type            =   200
         Name            =   "DiaSemana"
         Caption         =   "DiaSemana"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "HoraInicio"
         Caption         =   "HoraInicio"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Duracion"
         Caption         =   "Duracion"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   17
         Scale           =   0
         Type            =   200
         Name            =   "TipoCargo"
         Caption         =   "TipoCargo"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "DescripcionCargo"
         Caption         =   "DescripcionCargo"
      EndProperty
      BeginProperty Field10 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Porcentaje"
         Caption         =   "Porcentaje"
      EndProperty
      NumGroups       =   0
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Tipo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveSubtipo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TipoCargo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@DiaSemana"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@TodasHoras"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@HoraInicio"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@HoraFin"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset95 
      CommandName     =   "cmdActualizaDescuentos"
      CommDispId      =   1702
      RsDispId        =   -1
      CommandText     =   "dbo.sp_pvUpdActualizaDescuentos"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_pvUpdActualizaDescuentos( ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@MovPaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@NumCargo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset96 
      CommandName     =   "cmdPvSelExclusionDescuento"
      CommDispId      =   1704
      RsDispId        =   1706
      CommandText     =   "dbo.sp_PvSelExclusionDescuento"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelExclusionDescuento }"
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Consecutivo"
         Caption         =   "Consecutivo"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   202
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CveDepartamento"
         Caption         =   "CveDepartamento"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "Concepto"
         Caption         =   "Concepto"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CveConcepto"
         Caption         =   "CveConcepto"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   8
         Scale           =   0
         Type            =   200
         Name            =   "TipoPaciente"
         Caption         =   "TipoPaciente"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "EstatusTipoPaciente"
         Caption         =   "EstatusTipoPaciente"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   202
         Name            =   "Procedencia"
         Caption         =   "Procedencia"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CveProcedencia"
         Caption         =   "CveProcedencia"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset97 
      CommandName     =   "cmdGnFolios"
      CommDispId      =   1737
      RsDispId        =   -1
      CommandText     =   "DBO.SP_GNFOLIOS"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{ CALL DBO.SP_GNFOLIOS( ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "IN_CHRTIPOFOLIO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "IN_SMIDEPARTAMENTO"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "IN_BITAUMENTA"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "INTALERTA"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   139
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "CHRFOLIO"
         Direction       =   3
         Precision       =   0
         Scale           =   0
         Size            =   4000
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset98 
      CommandName     =   "cmdRepIntraMedico"
      CommDispId      =   1746
      RsDispId        =   1748
      CommandText     =   "dbo.sp_ExSelRepIntraMedico"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_ExSelRepIntraMedico( ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   11
      BeginProperty Field1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "NomHosp"
         Caption         =   "NomHosp"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "NomDepto"
         Caption         =   "NomDepto"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Fecha"
         Caption         =   "Fecha"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Hora"
         Caption         =   "Hora"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CvePaciente"
         Caption         =   "CvePaciente"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "Nombre"
         Caption         =   "Nombre"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Edad"
         Caption         =   "Edad"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   60
         Scale           =   0
         Type            =   200
         Name            =   "Domicilio"
         Caption         =   "Domicilio"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   15
         Scale           =   0
         Type            =   129
         Name            =   "Tel"
         Caption         =   "Tel"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "Sexo"
         Caption         =   "Sexo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Convenio"
         Caption         =   "Convenio"
      EndProperty
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NomHosp"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@NomDepto"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Fecha"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@Hora"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   20
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CvePaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@Edad"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   100
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset99 
      CommandName     =   "cmdPvSelCorteCuentasFacturacionNormal"
      CommDispId      =   1755
      RsDispId        =   1757
      CommandText     =   "dbo.sp_PvSelCorteCuentasFacturacionNormal"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteCuentasFacturacionNormal( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "chrFolioDocumento"
         Caption         =   "chrFolioDocumento"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Ingreso"
         Caption         =   "Ingreso"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaIngreso"
         Caption         =   "CuentaIngreso"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Descuento"
         Caption         =   "Descuento"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaDescuento"
         Caption         =   "CuentaDescuento"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "IVA"
         Caption         =   "IVA"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaIVA"
         Caption         =   "CuentaIVA"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitPesos"
         Caption         =   "bitPesos"
      EndProperty
      BeginProperty Field9 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTipoCambio"
         Caption         =   "mnyTipoCambio"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumCargo"
         Caption         =   "intNumCargo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset100 
      CommandName     =   "cmdPvSelCorteCuentasVentaPublico"
      CommDispId      =   1758
      RsDispId        =   1760
      CommandText     =   "dbo.sp_PvSelCorteCuentasVentaPublico"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteCuentasVentaPublico( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   10
      BeginProperty Field1 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "chrFolioDocumento"
         Caption         =   "chrFolioDocumento"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Ingreso"
         Caption         =   "Ingreso"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaIngreso"
         Caption         =   "CuentaIngreso"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Descuento"
         Caption         =   "Descuento"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaDescuento"
         Caption         =   "CuentaDescuento"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "IVA"
         Caption         =   "IVA"
      EndProperty
      BeginProperty Field7 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaIVA"
         Caption         =   "CuentaIVA"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "bitPesos"
         Caption         =   "bitPesos"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "mnyTipoCambio"
         Caption         =   "mnyTipoCambio"
      EndProperty
      BeginProperty Field10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumCargo"
         Caption         =   "intNumCargo"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset101 
      CommandName     =   "cmdPvSelCorteCuentasExcedente"
      CommDispId      =   1763
      RsDispId        =   1765
      CommandText     =   "dbo.sp_PvSelCorteCuentasExcedente"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteCuentasExcedente( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CuentaConcepto"
         Caption         =   "CuentaConcepto"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Cantidad"
         Caption         =   "Cantidad"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitPesos"
         Caption         =   "bitPesos"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTipoCambio"
         Caption         =   "mnyTipoCambio"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset102 
      CommandName     =   "cmdPvSelCorteMovimiento"
      CommDispId      =   1778
      RsDispId        =   1780
      CommandText     =   "dbo.sp_PvSelCorteMovimiento"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteMovimiento( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdPvSelCorteMovimiento_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   20
      BeginProperty Field1 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "FormaPago"
         Caption         =   "FormaPago"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   70
         Scale           =   0
         Type            =   202
         Name            =   "NombreHospital"
         Caption         =   "NombreHospital"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   221
         Scale           =   0
         Type            =   200
         Name            =   "Orden"
         Caption         =   "Orden"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "Folio"
         Caption         =   "Folio"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Cantidad"
         Caption         =   "Cantidad"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Cuenta"
         Caption         =   "Cuenta"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   200
         Name            =   "TipoPaciente"
         Caption         =   "TipoPaciente"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "NombrePaciente"
         Caption         =   "NombrePaciente"
      EndProperty
      BeginProperty Field9 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "Fecha"
         Caption         =   "Fecha"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "Clasificacion"
         Caption         =   "Clasificacion"
      EndProperty
      BeginProperty Field11 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "NumeroCorte"
         Caption         =   "NumeroCorte"
      EndProperty
      BeginProperty Field12 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field13 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaInicioCorte"
         Caption         =   "FechaInicioCorte"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaFinCorte"
         Caption         =   "FechaFinCorte"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "EmpleadoAbre"
         Caption         =   "EmpleadoAbre"
      EndProperty
      BeginProperty Field16 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "EmpleadoCierra"
         Caption         =   "EmpleadoCierra"
      EndProperty
      BeginProperty Field17 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "Ejercicio"
         Caption         =   "Ejercicio"
      EndProperty
      BeginProperty Field18 
         Precision       =   3
         Size            =   1
         Scale           =   0
         Type            =   17
         Name            =   "Mes"
         Caption         =   "Mes"
      EndProperty
      BeginProperty Field19 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "NumeroPoliza"
         Caption         =   "NumeroPoliza"
      EndProperty
      BeginProperty Field20 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "TipoDoc"
         Caption         =   "TipoDoc"
      EndProperty
      NumGroups       =   10
      BeginProperty Grouping1 
         Precision       =   0
         Size            =   221
         Scale           =   0
         Type            =   200
         Name            =   "Orden"
         Caption         =   "Orden"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "NumeroCorte"
         Caption         =   "NumeroCorte"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaInicioCorte"
         Caption         =   "FechaInicioCorte"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaFinCorte"
         Caption         =   "FechaFinCorte"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "EmpleadoAbre"
         Caption         =   "EmpleadoAbre"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "EmpleadoCierra"
         Caption         =   "EmpleadoCierra"
      EndProperty
      BeginProperty Grouping8 
         Precision       =   5
         Size            =   2
         Scale           =   0
         Type            =   2
         Name            =   "Ejercicio"
         Caption         =   "Ejercicio"
      EndProperty
      BeginProperty Grouping9 
         Precision       =   3
         Size            =   1
         Scale           =   0
         Type            =   17
         Name            =   "Mes"
         Caption         =   "Mes"
      EndProperty
      BeginProperty Grouping10 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "NumeroPoliza"
         Caption         =   "NumeroPoliza"
      EndProperty
      ParamCount      =   12
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Intermedios"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaIni"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@bitFacturas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@bitRecibos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@bitTickets"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@bitSalidas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@bitSoloCancelados"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P12 
         RealName        =   "@smiTipoOrden"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset103 
      CommandName     =   "cmdPvSelCorteCronologico"
      CommDispId      =   1781
      RsDispId        =   1783
      CommandText     =   "dbo.sp_PvSelCorteCronologico"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteCronologico( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdPvSelCorteCronologico_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaMovimiento"
         Caption         =   "FechaMovimiento"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "FolioDocumento"
         Caption         =   "FolioDocumento"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "TipoDocumento"
         Caption         =   "TipoDocumento"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "DescripcionFormaPago"
         Caption         =   "DescripcionFormaPago"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Referencia"
         Caption         =   "Referencia"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CantidadPagada"
         Caption         =   "CantidadPagada"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "Moneda"
         Caption         =   "Moneda"
      EndProperty
      NumGroups       =   7
      BeginProperty Grouping1 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaMovimiento"
         Caption         =   "FechaMovimiento"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "FolioDocumento"
         Caption         =   "FolioDocumento"
      EndProperty
      BeginProperty Grouping3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "TipoDocumento"
         Caption         =   "TipoDocumento"
      EndProperty
      BeginProperty Grouping4 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "DescripcionFormaPago"
         Caption         =   "DescripcionFormaPago"
      EndProperty
      BeginProperty Grouping5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Referencia"
         Caption         =   "Referencia"
      EndProperty
      BeginProperty Grouping6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CantidadPagada"
         Caption         =   "CantidadPagada"
      EndProperty
      BeginProperty Grouping7 
         Precision       =   0
         Size            =   5
         Scale           =   0
         Type            =   200
         Name            =   "Moneda"
         Caption         =   "Moneda"
      EndProperty
      ParamCount      =   11
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Intermedios"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaIni"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@bitFacturas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@bitRecibos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@bitTickets"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@bitSalidas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@bitSoloCancelados"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset104 
      CommandName     =   "cmdPvSelCorteResumenDocumentos"
      CommDispId      =   1784
      RsDispId        =   1786
      CommandText     =   "dbo.sp_PvSelCorteResumenDocumentos"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteResumenDocumentos( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      GroupingName    =   "cmdPvSelCorteResumenDocumentos_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "FolioDocumento"
         Caption         =   "FolioDocumento"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "Documento"
         Caption         =   "Documento"
      EndProperty
      NumGroups       =   0
      ParamCount      =   11
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Intermedios"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaIni"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@bitFacturas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@bitRecibos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@bitTickets"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@bitSalidas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@bitSoloCancelados"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset105 
      CommandName     =   "cmdPvSelCorteFormasPago"
      CommDispId      =   1787
      RsDispId        =   1789
      CommandText     =   "dbo.sp_PvSelCorteFormasPago"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteFormasPago( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   4
      BeginProperty Field1 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   129
         Name            =   "Forma"
         Caption         =   "Forma"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   129
         Name            =   "Referencia"
         Caption         =   "Referencia"
      EndProperty
      BeginProperty Field3 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Cantidad"
         Caption         =   "Cantidad"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   21
         Scale           =   0
         Type            =   200
         Name            =   "Moneda"
         Caption         =   "Moneda"
      EndProperty
      NumGroups       =   0
      ParamCount      =   11
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@Intermedios"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@FechaIni"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaFin"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   23
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveDepto"
         Direction       =   1
         Precision       =   5
         Scale           =   0
         Size            =   0
         DataType        =   2
         HostType        =   2
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@bitFacturas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@bitRecibos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@bitTickets"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@bitSalidas"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@bitSoloCancelados"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset106 
      CommandName     =   "cmdPvSelCorteCuentasConceptoPago"
      CommDispId      =   1790
      RsDispId        =   1792
      CommandText     =   "dbo.sp_PvSelCorteCuentasConceptoPago"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteCuentasConceptoPago( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "chrFolioDocumento"
         Caption         =   "chrFolioDocumento"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyCantidad"
         Caption         =   "mnyCantidad"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitPesos"
         Caption         =   "bitPesos"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTipoCambio"
         Caption         =   "mnyTipoCambio"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intCancelado"
         Caption         =   "intCancelado"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumCuentaConcepto"
         Caption         =   "intNumCuentaConcepto"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "chrTipoDocumento"
         Caption         =   "chrTipoDocumento"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset107 
      CommandName     =   "cmdPvSelCorteCuentasFormasPago"
      CommDispId      =   1793
      RsDispId        =   1797
      CommandText     =   "dbo.sp_PvSelCorteCuentasFormasPago"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelCorteCuentasFormasPago( ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyCantidadPagada"
         Caption         =   "mnyCantidadPagada"
      EndProperty
      BeginProperty Field2 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTipoCambio"
         Caption         =   "mnyTipoCambio"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumCuentaContable"
         Caption         =   "intNumCuentaContable"
      EndProperty
      BeginProperty Field4 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyComision"
         Caption         =   "mnyComision"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   129
         Name            =   "chrTipoDocumento"
         Caption         =   "chrTipoDocumento"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset108 
      CommandName     =   "cmdPvSelEntradaSalidaDineroFechaPaciente"
      CommDispId      =   1798
      RsDispId        =   1801
      CommandText     =   "dbo.sp_PvSelEntradaSalidaDineroFechaPaciente"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelEntradaSalidaDineroFechaPaciente( ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   9
      BeginProperty Field1 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "FechaDocumento"
         Caption         =   "FechaDocumento"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "FolioDocumento"
         Caption         =   "FolioDocumento"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Cuenta"
         Caption         =   "Cuenta"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   102
         Scale           =   0
         Type            =   200
         Name            =   "NombrePaciente"
         Caption         =   "NombrePaciente"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Cantidad"
         Caption         =   "Cantidad"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   7
         Scale           =   0
         Type            =   200
         Name            =   "Moneda"
         Caption         =   "Moneda"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Departamento"
         Caption         =   "Departamento"
      EndProperty
      BeginProperty Field8 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "NumEntradaSalida"
         Caption         =   "NumEntradaSalida"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "Cancelado"
         Caption         =   "Cancelado"
      EndProperty
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@TipoDocumento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FiltroFechasPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@FechaInicial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@FechaFinal"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@Cuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset109 
      CommandName     =   "cmdNumCuentaExterno"
      CommDispId      =   1802
      RsDispId        =   -1
      CommandText     =   "dbo.sp_GnNumCuentaExterno"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_GnNumCuentaExterno( ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CvePaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveConvenio"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   0
         DataType        =   17
         HostType        =   17
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@CveEmpresa"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   0
         DataType        =   17
         HostType        =   17
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveMedico"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@CveDeptoAbreCuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset110 
      CommandName     =   "cmdRegistroExterno"
      CommDispId      =   1804
      RsDispId        =   -1
      CommandText     =   "dbo.sp_GnNumCuentaExterno"
      ActiveConnectionName=   "ConeccionSIHO"
      CallSyntax      =   "{? = CALL dbo.sp_GnNumCuentaExterno( ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@CvePaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@CveConvenio"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   0
         DataType        =   17
         HostType        =   17
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@CveEmpresa"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   3
         Scale           =   0
         Size            =   0
         DataType        =   17
         HostType        =   17
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CveMedico"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@CveDeptoAbreCuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset111 
      CommandName     =   "cmdPvInsPvCortePoliza"
      CommDispId      =   1806
      RsDispId        =   -1
      CommandText     =   "dbo.sp_PvInsPvCortePoliza"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvInsPvCortePoliza( ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   7
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@NumeroCorte"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@FolioDocumento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@TipoDocumento"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@NumeroCuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@CantidadMovimiento"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@Cargo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset112 
      CommandName     =   "cmdPvSelEstadoCuenta"
      CommDispId      =   1808
      RsDispId        =   1810
      CommandText     =   "dbo.sp_PvSelEstadoCuenta"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelEstadoCuenta( ?, ?, ?, ?, ?, ?, ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdPvSelEstadoCuenta_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   28
      BeginProperty Field1 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Campo0"
         Caption         =   "Campo0"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   14
         Scale           =   0
         Type            =   200
         Name            =   "Campo1"
         Caption         =   "Campo1"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   200
         Name            =   "Campo2"
         Caption         =   "Campo2"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Campo3"
         Caption         =   "Campo3"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "Campo4"
         Caption         =   "Campo4"
      EndProperty
      BeginProperty Field6 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Campo5"
         Caption         =   "Campo5"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Campo6"
         Caption         =   "Campo6"
      EndProperty
      BeginProperty Field8 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Campo7"
         Caption         =   "Campo7"
      EndProperty
      BeginProperty Field9 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "bitExcluido"
         Caption         =   "bitExcluido"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   200
         Name            =   "chrTipoCargo"
         Caption         =   "chrTipoCargo"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "chrCveCargo"
         Caption         =   "chrCveCargo"
      EndProperty
      BeginProperty Field12 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "smiCveConcepto"
         Caption         =   "smiCveConcepto"
      EndProperty
      BeginProperty Field13 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "chrFolioFactura"
         Caption         =   "chrFolioFactura"
      EndProperty
      BeginProperty Field14 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaHora"
         Caption         =   "dtmFechaHora"
      EndProperty
      BeginProperty Field15 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Descuento"
         Caption         =   "Descuento"
      EndProperty
      BeginProperty Field16 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "IVACargo"
         Caption         =   "IVACargo"
      EndProperty
      BeginProperty Field17 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intNumCuenta"
         Caption         =   "intNumCuenta"
      EndProperty
      BeginProperty Field18 
         Precision       =   0
         Size            =   1
         Scale           =   0
         Type            =   129
         Name            =   "chrTipoPaciente"
         Caption         =   "chrTipoPaciente"
      EndProperty
      BeginProperty Field19 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "dtmFechaHora"
         Caption         =   "dtmFechaHora"
      EndProperty
      BeginProperty Field20 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyDiferencia"
         Caption         =   "mnyDiferencia"
      EndProperty
      BeginProperty Field21 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "TipoPago"
         Caption         =   "TipoPago"
      EndProperty
      BeginProperty Field22 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "CveDepartamento"
         Caption         =   "CveDepartamento"
      EndProperty
      BeginProperty Field23 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TotalCuenta"
         Caption         =   "TotalCuenta"
      EndProperty
      BeginProperty Field24 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "MontoPago"
         Caption         =   "MontoPago"
      EndProperty
      BeginProperty Field25 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "MontoFacturado"
         Caption         =   "MontoFacturado"
      EndProperty
      BeginProperty Field26 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "TotalPagar"
         Caption         =   "TotalPagar"
      EndProperty
      BeginProperty Field27 
         Precision       =   0
         Size            =   11
         Scale           =   0
         Type            =   200
         Name            =   "Incremento"
         Caption         =   "Incremento"
      EndProperty
      BeginProperty Field28 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "CosPromedio"
         Caption         =   "CosPromedio"
      EndProperty
      NumGroups       =   2
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "TipoPago"
         Caption         =   "TipoPago"
      EndProperty
      BeginProperty Grouping2 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "Campo0"
         Caption         =   "Campo0"
      EndProperty
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@Cuenta"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Orden"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@TodosCargos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@Excluidos"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   11
         HostType        =   11
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@Departamento"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@TieneCosto"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@Factura"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   12
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset113 
      CommandName     =   "cmdReportePagos"
      CommDispId      =   1820
      RsDispId        =   1822
      CommandText     =   "dbo.sp_pvselReportePagosPaciente"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_pvselReportePagosPaciente( ?, ?) }"
      Grouping        =   -1  'True
      GroupingName    =   "cmdReportePagos_Grouping"
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intMovPaciente"
         Caption         =   "intMovPaciente"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   12
         Scale           =   0
         Type            =   129
         Name            =   "Recibo"
         Caption         =   "Recibo"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "Descripcion"
         Caption         =   "Descripcion"
      EndProperty
      BeginProperty Field4 
         Precision       =   23
         Size            =   16
         Scale           =   3
         Type            =   135
         Name            =   "Fecha"
         Caption         =   "Fecha"
      EndProperty
      BeginProperty Field5 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "Cantidad"
         Caption         =   "Cantidad"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   2
         Scale           =   0
         Type            =   11
         Name            =   "bitPesos"
         Caption         =   "bitPesos"
      EndProperty
      BeginProperty Field7 
         Precision       =   19
         Size            =   8
         Scale           =   0
         Type            =   6
         Name            =   "mnyTipoCambio"
         Caption         =   "mnyTipoCambio"
      EndProperty
      NumGroups       =   1
      BeginProperty Grouping1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "intMovPaciente"
         Caption         =   "intMovPaciente"
      EndProperty
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@intMovPaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@chrTipoPaciente"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset114 
      CommandName     =   "cmdDescuento"
      CommDispId      =   1823
      RsDispId        =   1825
      CommandText     =   "dbo.sp_PvSelDescuento"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelDescuento( ?, ?, ?, ?, ?, ?, ?, ?) }"
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "DescuentoTP"
         Caption         =   "DescuentoTP"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "DescuentoPA"
         Caption         =   "DescuentoPA"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   129
         Name            =   "DescuentoEM"
         Caption         =   "DescuentoEM"
      EndProperty
      NumGroups       =   0
      ParamCount      =   9
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@InternoExterno"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Empresa"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@MovPaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@TipoCargo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@CveCargo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@Departamento"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@FechaCargo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset115 
      CommandName     =   "cmdDescuentoCantidad"
      CommDispId      =   1826
      RsDispId        =   -1
      CommandText     =   "dbo.sp_PvSelDescuentoCantidad"
      ActiveConnectionName=   "ConeccionSIHO"
      Prepared        =   -1  'True
      CallSyntax      =   "{? = CALL dbo.sp_PvSelDescuentoCantidad( ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) }"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   11
      BeginProperty P1 
         RealName        =   "RETURN_VALUE"
         Direction       =   4
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   0   'False
      EndProperty
      BeginProperty P2 
         RealName        =   "@InternoExterno"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   1
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "@TipoPaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "@Empresa"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "@MovPaciente"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P6 
         RealName        =   "@TipoCargo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   2
         DataType        =   200
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P7 
         RealName        =   "@CveCargo"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P8 
         RealName        =   "@Cantidad"
         Direction       =   1
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P9 
         RealName        =   "@CantidadDescuento"
         Direction       =   3
         Precision       =   19
         Scale           =   0
         Size            =   0
         DataType        =   6
         HostType        =   6
         Required        =   -1  'True
      EndProperty
      BeginProperty P10 
         RealName        =   "@Departamento"
         Direction       =   1
         Precision       =   10
         Scale           =   0
         Size            =   0
         DataType        =   3
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      BeginProperty P11 
         RealName        =   "@FechaCargo"
         UserName        =   "FechaCargo"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   135
         HostType        =   7
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "EntornoSIHO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub rscmdCnUpdEstatusCierre_WillChangeField(ByVal cFields As Long, ByVal Fields As Variant, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

End Sub

