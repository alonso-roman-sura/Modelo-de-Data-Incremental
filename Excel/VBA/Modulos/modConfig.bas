Option Explicit

' ============================================================
' modConfig
' Constantes y variables globales del libro BDNE.
' IMPORTANTE: en VBA todas las declaraciones de modulo deben ir
' antes de cualquier Function/Sub.
' ============================================================

Public g_NombreArchivoQ1 As String
Public g_NombreArchivoQ2 As String

' Hojas del libro
Public Const SH_PANEL       As String = "Panel"
Public Const SH_Q1_LAFT     As String = "Q1_LAFT"
Public Const SH_Q1_REL      As String = "Q1_REL"
Public Const SH_Q2_LAFT     As String = "Q2_LAFT"
Public Const SH_Q2_REL      As String = "Q2_REL"
Public Const SH_OUTPUT      As String = "BDNE_OUTPUT"

' ListObjects
Public Const LO_Q1_LAFT     As String = "Tbl_Q1_LAFT"
Public Const LO_Q1_REL      As String = "Tbl_Q1_REL"
Public Const LO_Q2_LAFT     As String = "Tbl_Q2_LAFT"
Public Const LO_Q2_REL      As String = "Tbl_Q2_REL"
Public Const LO_OUTPUT      As String = "Tbl_BDNE"

' Named Ranges de estado en el panel
Public Const ST_Q1_ESTADO   As String = "EstadoQ1"
Public Const ST_Q2_ESTADO   As String = "EstadoQ2"
Public Const ST_OUT_ESTADO  As String = "EstadoOutput"

' Hojas del archivo fuente
Public Const SRC_HOJA_LAFT  As String = "LAFT"
Public Const SRC_HOJA_REL   As String = "PEP RELACIONADOS"

' La hoja LAFT tiene 2 filas de metadatos encima del header real
Public Const SRC_LAFT_HEADER_ROW As Integer = 3

' Columnas fuente LAFT
Public Const SRC_TIPO_LISTA   As String = "Tipo de Lista"
Public Const SRC_TIPO_DOC     As String = "Tipo de documento"
Public Const SRC_DOCUMENTO    As String = "Documento"
Public Const SRC_NOMBRES      As String = "Nombres"
Public Const SRC_AP_PAT       As String = "Apellido Paterno"
Public Const SRC_AP_MAT       As String = "Apellido Materno"
Public Const SRC_CARGO        As String = "Cargo"
Public Const SRC_ENTIDAD      As String = "Entidad"
Public Const SRC_FECHA_INI    As String = "Fecha inicio"
Public Const SRC_FECHA_FIN    As String = "Fecha fin"
Public Const SRC_NOTICIA      As String = "Tipo de Noticia"

' Columnas fuente PEP RELACIONADOS
Public Const REL_TIPO_DOC_PEP As String = "TIPO DE DOCUMENTO_PEP"
Public Const REL_DNI_PEP      As String = "DNI_ PEP"
Public Const REL_NOMBRE_PEP   As String = "NOMBRE_PEP"
Public Const REL_TIPO_DOC_REL As String = "(*) TIPO DE DOCUMENTO"
Public Const REL_DNIREL       As String = "DNIREL"
Public Const REL_AP_PAT       As String = "APELLIDO PATERNO PEP_RELACION"
Public Const REL_AP_MAT       As String = "APELLIDO MATERNO PEP_RELACION"
Public Const REL_NOMBRES      As String = "NOMBRES"
Public Const REL_RELACION     As String = "RELACION"

' Columnas destino BDNE
Public Const BD_APELLIDOS   As String = "Apellidos"
Public Const BD_NOMBRE      As String = "Nombre"
Public Const BD_SOURCE      As String = "Source"
Public Const BD_REMARK      As String = "Remark"
Public Const BD_PAIS        As String = "Pais"
Public Const BD_DIRECCION   As String = "Direccion"
Public Const BD_CIUDAD      As String = "Ciudad"
Public Const BD_PROGRAM     As String = "Program"
Public Const BD_SIGLA1      As String = "Sigla1"
Public Const BD_SIGLA2      As String = "Sigla2"
Public Const BD_ID          As String = "id"

' Prefijo de Program
Public Const BD_PROGRAM_PREFIJO As String = "BDNEG "

' Sigla fija para TODAS las filas (Sigla1 y dentro de la formula Program)
Public Const BD_SIGLA_VALOR As String = "INSACO"

' Prefijo de Source para PEP RELACIONADOS
Public Const BD_PARIENTE_PREFIJO As String = "PARIENTE PEP "

' Valor literal NOAPLICA usado en Remark
Public Const BD_NOAPLICA As String = "NOAPLICA"

' Valores de documento invalido en el fuente
Public Const SRC_SIN_ID     As String = "SIN IDENTIFICAR"
Public Const SRC_NOAPLICA   As String = "NOAPLICA"

' Columnas minimas requeridas para validar LAFT
Public Function ColsRequeridasLAFT() As Variant
    ColsRequeridasLAFT = Array( _
        SRC_TIPO_LISTA, _
        SRC_DOCUMENTO, _
        SRC_NOMBRES, _
        SRC_AP_PAT, _
        SRC_CARGO, _
        SRC_ENTIDAD, _
        SRC_NOTICIA)
End Function

' Columnas minimas requeridas para validar PEP RELACIONADOS
Public Function ColsRequeridasREL() As Variant
    ColsRequeridasREL = Array( _
        REL_NOMBRE_PEP, _
        REL_DNIREL, _
        REL_AP_PAT, _
        REL_NOMBRES, _
        REL_RELACION)
End Function

' Valor fijo de Pais: "PERU" con U acentuada (Chr(218) = U con tilde)
Public Function PaisValor() As String
    PaisValor = "PER" & Chr(218)
End Function