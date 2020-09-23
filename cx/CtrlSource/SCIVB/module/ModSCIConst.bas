Attribute VB_Name = "ModSCIConst"
Option Explicit

'+--------------------------------+
'| Begin Scintilla Constants      |
'+--------------------------------+

Public Const INVALID_POSITION = -1

Public Const SCI_START = 2000
Public Const SCI_ADDTEXT = SCI_START + 1
Public Const SCI_SETSELBACK = 2068
Public Const SCI_SETSELFORE = 2067
Public Const SCI_SETLEXER = 4001

Public Const WS_CHILD = &H40000000
Public Const WS_VISIBLE = &H10000000
Public Const WS_EX_CLIENTEDGE = &H200

Public Const SC_FOLDLEVELHEADERFLAG = &H2000
Public Const SC_CACHE_DOCUMENT = 3

Public Const SCLEX_AUTOMATIC = 1000
Public Const SCEN_CHANGE = 768
Public Const SCEN_SETFOCUS = 512
Public Const SCEN_KILLFOCUS = 256
Public Const SCI_SETCODEPAGE = 2037
Public Const SCN_STYLENEEDED = 2000
Public Const SCN_CHARADDED = 2001
Public Const SCN_SAVEPOINTREACHED = 2002
Public Const SCN_SAVEPOINTLEFT = 2003
Public Const SCN_MODIFYATTEMPTRO = 2004
Public Const SCI_STARTRECORD = 3001
Public Const SCN_AUTOCSELECTION = 2022
Public Const SCI_STOPRECORD = 3002
Public Const SCN_DOUBLECLICK = 2006
Public Const SCI_MARKERGET = 2046
Public Const SCN_UPDATEUI = 2007
Public Const SCN_MARKERDELETE = 2044
Public Const SCN_MARKERDELETEALL = 2045
Public Const SCN_MARKERNEXT = 2047
Public Const SCN_MARKERPREVIOUS = 2048
Public Const SCN_MODIFIED = 2008
Public Const SCN_MACRORECORD = 2009
Public Const SCN_MARGINCLICK = 2010
Public Const SCN_NEEDSHOWN = 2011
Public Const SCN_PAINTED = 2013
Public Const SCN_USERLISTSELECTION = 2014
Public Const SCN_URIDROPPED = 2015
Public Const SCN_DWELLSTART = 2016
Public Const SCN_DWELLEND = 2017
Public Const SCN_KEY = 2005

'-------------------MÁSCARA DE EVENTOS PARA SCN_MODIFIED----------
Public Const SC_MOD_INSERTTEXT = &H1
Public Const SC_MOD_DELETETEXT = &H2
Public Const SC_MOD_CHANGESTYLE = &H4
Public Const SC_MOD_CHANGEFOLD = &H8
Public Const SC_PERFORMED_USER = &H10
Public Const SC_PERFORMED_UNDO = &H20
Public Const SC_PERFORMED_REDO = &H40
Public Const SC_LASTSTEPINUNDOREDO = &H100
Public Const SC_MOD_CHANGEMARKER = &H200
Public Const SC_MOD_BEFOREINSERT = &H400
Public Const SC_MOD_BEFOREDELETE = &H800
Public Const SC_MODEVENTMASKALL = &HF77
Public Const SCI_SETMODEVENTMASK = 2359
Public Const SCI_GETMODEVENTMASK = 2378

'-------------------ASIGNACIÓN DE ESTILOS-----------------------------
Public Const STYLE_DEFAULT = 32
Public Const SCI_STYLECLEARALL = 2050       '
Public Const SCI_STYLESETFORE = 2051
Public Const SCI_STYLESETBACK = 2052
Public Const SCI_STYLESETBOLD = 2053
Public Const SCI_STYLESETITALIC = 2054
Public Const SCI_STYLESETSIZE = 2055
Public Const SCI_STYLESETFONT = 2056
Public Const SCI_STYLESETEOLFILLED = 2057
Public Const SCI_STYLERESETDEFAULT = 2058
Public Const SCI_STYLESETUNDERLINE = 2059
Public Const SCI_STYLESETCASE = 2060
Public Const SCI_STYLESETCHARACTERSET = 2066
Public Const SCI_STYLESETVISIBLE = 2074
Public Const SCI_StyleSetBITS = 2090        ' Determinar el número de bits adicionales de estilo que se usarán
Public Const SCI_SETKEYWORDS = 4005         ' Asignar una lista de palabras clave a Scintilla

'-------------------RECUPERACIÓN DE TEXTO-----------------------------
Public Const SCI_GETTEXT = 2182             ' Recupera el texto del documento. Devuelve el número de caracteres recuperados
Public Const SCI_SETTEXT = 2181
Public Const SCI_GETLENGTH = 2006           ' Devuelve el número de caracteres del documento
Public Const SCI_GETCURLINE = 2027          ' Recupera el texto de la línea actual (que contiene el cursor).
                                            ' Devuelve, además, la posición del cursor en la línea.
Public Const SCI_GETLINECOUNT = 2154        ' Devuelve el número de líneas del documento. Siempre hay al menos una.
Public Const SCI_GETTEXTRANGE = 2162        ' Recupera un intervalo de texto; devuelve la longitud del intervalo.
Public Const SCI_GETCHARAT = 2007           ' Devuelve el carácter situado en una posición

'-------------------POSICIÓN Y CURSOR---------------------------------
Public Const SCI_GOTOLINE = 2024            ' Coloca el cursor al inicio de una línea y la sitúa en zona visible.
Public Const SCI_GOTOPOS = 2025             ' Coloca el cursor en una posición y lo sitúa en zona visible.
Public Const SCI_SETANCHOR = 2026           ' Coloca el ancla de selección en una posición. El acncla es
                                            ' el final de la selección (a partir de la situación del cursor).
Public Const SCI_GETCURRENTPOS = 2008       ' Devuelve la posición actual del cursor.
Public Const SCI_LINEFROMPOSITION = 2166    ' Devuelve el número de línea de una posición.
Public Const SCI_GETCOLUMN = 2129           ' Devuelve la columna de una posición

'-------------------MÁRGENES------------------------------------------
Public Const SC_MARGIN_SYMBOL = 0
Public Const SC_MARGIN_NUMBER = 1
Public Const SCI_SETMARGINTYPEN = 2240      ' Configura un margen como numérico o de símbolos
Public Const SCI_GETMARGINYPEN = 2241       ' Recupera el tipo de margen
Public Const SCI_SETMARGINWIDTHN = 2242     ' Fija en pixels el ancho del margen
Public Const SCI_GETMARGINWIDTHN = 2243     ' Devuelve los píxeles de ancho de un margen
Public Const SCI_SETMARGINMASKN = 2244      ' Determina los marcadores que se mostrarán en un margen usando una máscara.
Public Const SCI_GETMARGINMASKN = 2245      ' Devuelve la máscara de un margen
Public Const SCI_SETMARGINSENSITIVEN = 2246 ' Hace que un margen sea sensible o no al ratón.
Public Const SCI_GETMARGINSENSITIVEN = 2247 ' Devuelve si un margen es sensible o no al ratón.

'-------------------AYUDA CONTEXTUAL----------------------------------
Public Const SCI_CALLTIPSHOW = 2200         ' Muestra el globo de ayuda
Public Const SCI_CALLTIPCANCEL = 2201       ' Elimina el globo de ayuda
Public Const SCI_CALLTIPACTIVE = 2202       ' Indica si existe un globo activo
Public Const SCI_CALLTIPPOSSTART = 2203     ' Devuelve la posición del cursor previa al despliegue del globo.
Public Const SCI_CALLTIPSETHLT = 2204       ' Realza un segmento de la definición
Public Const SCI_CALLTIPSETBACK = 2205      ' Asgina un color de fondo al globo

'-------------------FINAL DE LÍNEA------------------------------------
Public Const SCI_GETEOLMODE = 2030          ' Recupera el tipo de fin de línea vigente: CRLF, CR o LF.
Public Const SCI_SETEOLMODE = 2031          ' Fija el tipo de fin de línea.
Public Const SCI_SETVIEWEOL = 2356          ' Hace visibles o invisibles los finales de línea.
Public Const SCI_CONVERTEOLS = 2029        ' Convierte los finales de línea al final especificado.

'-------------------BÚSQUEDA Y SUSTITUCIÓN DE TEXTO-------------------

Public Const SCI_SETTARGETSTART = 2190      ' Fija la posición de inicio del intervalo de texto a tratar.
Public Const SCI_GETTARGETSTART = 2191      ' Devuelve la posición de inicio del intervalo.
Public Const SCI_SETTARGETEND = 2192        ' Fija la posición final del intervalo.
Public Const SCI_GETTARGETEND = 2193        ' Devuelve la posición final del intervalo.
Public Const SCI_REPLACETARGET = 2194       ' Reemplaza un intervalo de texto. Devuelve la longitud del
                                            ' texto reemplazado. La cadena puede contener nulos.
Public Const SCI_REPLACETARGETRE = 2195     ' Búsqueda y reemplazo con expresiones regulares.
Public Const SCI_SEARCHINTARGET = 2197      ' Búsqueda en el intervalo. Devuelve la longitud del
                                            ' nuevo intervalo o -1 si no se encuentra.
Public Const SCI_SETSEARCHFLAGS = 2198      ' Fija los modificadores de búsqueda.
Public Const SCI_GETSEARCHFLAGS = 2199      ' Devuelve los modificadores de búsqueda actuales.
Public Const SCI_FINDTEXT = 2150            ' Busca un fragmento de texto en el documento. Devuelve la posición
                                            ' si se encuentra o -1 si no.
                                            
'-------------------SELECCION DE TEXTO--------------------------------
Public Const SCI_SETSELECTIONSTART = 2142   ' Fija el inicio de la selección (ancla).
Public Const SCI_GETSELECTIONSTART = 2143   ' Devuelve la posición inicial de una selección.
Public Const SCI_SETSELECTIONEND = 2144     ' Fija el fin de la selección (posición actual).
Public Const SCI_GETSELECTIONEND = 2145     ' Devuelve la posición final de la selección.
Public Const SCI_SETSEL = 2160              ' Fija el inicio y el final de la selección
Public Const SCI_GETSELTEXT = 2161          ' Devuelve la longitud de la selección y el texto seleccionado.

'-------------------CORTAR, COPIAR, PEGAR Y DESHACER------------------
Public Const SCI_REDO = 2011                ' Vuelve a realizar la siguiente acción.
Public Const SCI_SETUNDOCOLLECTION = 2012   ' Establece si se guardan las acciones en la historia o no.
Public Const SCI_CANREDO = 2016             ' Confirma si se puede rehacer.
Public Const SCI_GETUNDOCOLLECTION = 2019   ' Verdadero si se está recolectando la historia de acciones o no.
Public Const SCI_CANPASTE = 2173            ' Devuelve verdadero si es posible pegar.
Public Const SCI_CANUNDO = 2174             ' Devuelve verdadero si se puede deshacer.
Public Const SCI_EMPTYUNDOBUFFER = 2175     ' Borrar el historial de acciones.
Public Const SCI_UNDO = 2176                ' Deshacer una acción del historial.
Public Const SCI_CUT = 2177                 ' Corta la selección al portapapeles.
Public Const SCI_COPY = 2178                ' Copia la selección al portapapeles.
Public Const SCI_PASTE = 2179               ' Pega el contenido del portapapeles.
Public Const SCI_CLEAR = 2180               ' Borra la selección.

'-------------------ESPACIOS EN BLANCO--------------------------------
Public Const SCI_SETVIEWWS = 2021           ' Hace visibles los espacios en blanco
Public Const SCI_GETVIEWWS = 2020

'-------------------LISTAS AUTOMÁTICAS--------------------------------
Public Const SCI_AUTOCSHOW = 2100           ' Despliega una lista automática. Necesita los
                                            ' parámetros (INT LENENTERED, STRING ITEMLIST)
                                            ' El primero indica cuántos caracteres antes del
                                            ' cursor se usarán para buscar en la lista; el
                                            ' segundo, una lista (separada por espacios),
                                            ' que se mostrará.
Public Const SCI_AUTOCCANCEL = 2101         ' Oculta la lista
Public Const SCI_AUTOCACTIVE = 2102         ' Indica si hay una lista visible
Public Const SCI_AUTOCPOSSTART = 2103       ' Devuelve la posición del cursor al desplegar la lista
Public Const SCI_AUTOCCOMPLETE = 2104       ' Indica que el usuario ha seleccionado un elemento;
                                            ' se oculta la lista y se inserta la selección
Public Const SCI_AUTOCSTOPS = 2105          ' Define un conjunto de caracteres que al ser tecleados
                                            ' cancelan la lista.
Public Const SCI_AUTOCSETSEPARATOR = 2106   ' Cambia el carácter de separación de los elementos
                                            ' de la lista, espacio por defecto.
Public Const SCI_AUTOCGETSEPARATOR = 2107   ' Devuelve el carácter de separación
Public Const SCI_AUTOCSELECT = 2108         ' Selecciona el elemento de la lista que empieza por la cadena
Public Const SCI_AUTOCSETCANCELATSTART = 2110 ' Establece la cancelación de la lista cuando se vuelve
                                              ' a la posición en la que se activó.
Public Const SCI_AUTOCGETCANCELATSTART = 2111 ' Si la cancelación al regresar a la posición inicial está
                                              ' activada.
Public Const SCI_AUTOCSETFILLUPS = 2112     ' Define un conjunto de caracteres que al ser tecleados
                                            ' insertan la palabra seleccionada de la lista.
'' SHOULD A SINGLE ITEM AUTO-COMPLETION LIST AUTOMATICALLY CHOOSE THE ITEM.
'Public Const SCI_AUTOCSETCHOOSESINGLE = 2113
'
'' RETRIEVE WHETHER A SINGLE ITEM AUTO-COMPLETION LIST AUTOMATICALLY CHOOSE THE ITEM.
'Public Const SCI_AUTOCGETCHOOSESINGLE = 2114

Public Const SCI_AUTOCSETIGNORECASE = 2115  ' Establece si afectan mays/min a la búsqueda en la lista
Public Const SCI_AUTOCGETIGNORECASE = 2116  ' Devuelve si afectan mays/min a la búsqueda en la lista
Public Const SCI_AUTOCSETAUTOHIDE = 2118    ' Establece si la lista se oculta si no existen coindicencias.
Public Const SCI_AUTOCGETAUTOHIDE = 2119    ' Devuelve el comportamiento de la lista si no existen coincidencias.

' DISPLAY A LIST OF STRINGS AND SEND NOTIFICATION WHEN USER CHOOSES ONE.
' Public Const SCI_USERLISTSHOW=2117(INT LISTTYPE, STRING ITEMLIST)

'-------------------CORRESPONDENCIA DE PARÉNTESIS---------------------
' Los caracteres implicados son ( ) [ ] < >
' El estilo de ambos debe coincidir para encontrar la concordancia
Public Const SCI_BRACEHIGHLIGHT = 2351          ' Destaca los brazos en dos posiciones
Public Const SCI_BRACEBADLIGHT = 2352           ' Destaca el brazo en una posición, inicando que no tiene
                                                ' correspondencia.
Public Const SCI_BRACEMATCH = 2353              ' Devuelve la posición del otro brazo o INVALID_POSITION


'-------------------GUÍAS DE INDENTACIÓN------------------------------
Public Const SCI_SETINDENTATIONGUIDES = 2132    ' Guías visibles o no
Public Const SCI_SETHIGHLIGHTGUIDE = 2134       ' Realza la guía de una determinada columna

'-------------------More constants------------------------------------
Public Const SCI_SETPROPERTY = 4004
Public Const MARGIN_SCRIPT_FOLD_INDEX = 2  'This is a custom one
' Constants for folding
Public Const SC_MARKNUM_FOLDEREND = 25
Public Const SC_MARKNUM_FOLDEROPENMID = 26
Public Const SC_MARKNUM_FOLDERMIDTAIL = 27
Public Const SC_MARKNUM_FOLDERTAIL = 28
Public Const SC_MARKNUM_FOLDERSUB = 29
Public Const SC_MARKNUM_FOLDER = 30
Public Const SC_MARKNUM_FOLDEROPEN = 31
Public Const SCI_MASK_FOLDERS = &HFE000000
Public Const SCI_MARKERDEFINE = 2040

' Marker Constants
Public Const SC_MARK_CIRCLE = 0
Public Const SC_MARK_ROUNDRECT = 1
Public Const SC_MARK_ARROW = 2
Public Const SC_MARK_SMALLRECT = 3
Public Const SC_MARK_SHORTARROW = 4
Public Const SC_MARK_EMPTY = 5
Public Const SC_MARK_ARROWDOWN = 6
Public Const SC_MARK_MINUS = 7
Public Const SC_MARK_PLUS = 8
Public Const SC_MARK_VLINE = 9
Public Const SC_MARK_LCORNER = 10
Public Const SC_MARK_TCORNER = 11
Public Const SC_MARK_BOXPLUS = 12
Public Const SC_MARK_BOXPLUSCONNECTED = 13
Public Const SC_MARK_BOXMINUS = 14
Public Const SC_MARK_BOXMINUSCONNECTED = 15
Public Const SC_MARK_LCORNERCURVE = 16
Public Const SC_MARK_TCORNERCURVE = 17
Public Const SC_MARK_CIRCLEPLUS = 18
Public Const SC_MARK_CIRCLEPLUSCONNECTED = 19
Public Const SC_MARK_CIRCLEMINUS = 20
Public Const SC_MARK_CIRCLEMINUSCONNECTED = 21
Public Const SC_MARK_BACKGROUND = 22
Public Const SC_MARK_DOTDOTDOT = 23
Public Const SC_MARK_ARROWS = 24
Public Const SC_MARK_PIXMAP = 25
Public Const SC_MARK_FULLRECT = 26
Public Const SC_MARK_CHARACTER = 10000

Public Const SCI_SETFOLDFLAGS = 2233

'Save Point
Public Const SCI_SETSAVEPOINT = 2014
'GetModify
Public Const SCI_GETMODIFY = 2159

Public Const SCI_SELECTALL = 2013

Public Const SCI_COLOURISE = 4003

Public Const SC_MASK_FOLDERS = &HFE000000

Public Const SCI_TOGGLEFOLD = 2231
Public Const SCI_STARTSTYLING = 2032

Public Const SCI_MARKERSETFORE = 2041
Public Const SCI_MARKERSETBACK = 2042

Public Const SCI_REPLACESEL = 2170

Public Const SCI_DELETEBACK = 2326

Public Const SCI_SETLINEINDENTATION = 2126
Public Const SCI_GETLINEINDENTATION = 2127

Public Const SCI_GETLINEENDPOSITION = 2136

Public Const SCI_SETCURRENTPOS = 2141

Public Const SCI_POSITIONFROMLINE = 2167

Public Const SCI_GETLINEINDENTPOSITION = 2128

Public Const SCI_CLEARALLCMDKEYS = 2072

Public Const SCI_USEPOPUP = 2371

Public Const SCI_SETREADONLY = 2171
Public Const SCI_GETREADONLY = 2140

Public Const SCI_SETWRAPMODE = 2268

Public Const SCI_GETLINE = 2153
Public Const SCI_LINELENGTH = 2350

Public Const SCI_SETZOOM = 2373
Public Const SCI_GETZOOM = 2374

Public Const SCI_SETSCROLLWIDTH = 2274
Public Const SCI_GETSCROLLWIDTH = 2275

Public Const SCFIND_WHOLEWORD = 2
Public Const SCFIND_MATCHCASE = 4
Public Const SCFIND_WORDSTART = &H100000
Public Const SCFIND_REGEXP = &H200000

Public Const SCI_GETSTYLEAT = 2010

Public Const SCI_MARKERADD = 2043

Public Const SCI_ENSUREVISIBLEENFORCEPOLICY = 2234
Public Const SCI_FINDCOLUMN = 2456

Public Const SCI_CLEARCMDKEY = 2071

Public Const SCMOD_CTRL = 2
Public Const SCI_NULL = 2172

Public Const SCI_SETFOCUS = 2380

Public Const SCI_FORMATRANGE = 2151

Public Const SCI_GRABFOCUS = 2400
Public Const SCI_GETFOCUS = 2381

' New added stuff
Public Const SCI_TAB = 2327
Public Const SCI_BACKTAB = 2328

Public Const SCI_CHARRIGHT = 2306
Public Const SCI_CHARLEFT = 2304

Public Const SCI_SETLAYOUTCACHE = 2272

'-------------------CONJUNTOS DE CARACTERES---------------------------
Public Enum SC_CHARSET
    SC_CHARSET_ANSI = 0
    SC_CHARSET_DEFAULT = 1
    SC_CHARSET_BALTIC = 186
    SC_CHARSET_CHINESEBIG5 = 136
    SC_CHARSET_EASTEUROPE = 238
    SC_CHARSET_GB2312 = 134
    SC_CHARSET_GREEK = 161
    SC_CHARSET_HANGUL = 129
    SC_CHARSET_MAC = 77
    SC_CHARSET_OEM = 255
    SC_CHARSET_RUSSIAN = 204
    SC_CHARSET_SHIFTJIS = 128
    SC_CHARSET_SYMBOL = 2
    SC_CHARSET_TURKISH = 162
    SC_CHARSET_JOHAB = 130
    SC_CHARSET_HEBREW = 177
    SC_CHARSET_ARABIC = 178
    SC_CHARSET_VIETNAMESE = 163
    SC_CHARSET_THAI = 222
End Enum

'-------------------MAYÚSCULAS/MINÚSCULAS-----------------------------
Public Enum SC_CASE
    SC_CASE_MIXED = 0
    SC_CASE_UPPER = 1
    SC_CASE_LOWER = 2
End Enum

'-------------------BARRA DE DESPLAZAMIENTO HORIZONTAL----------------
Public Const SCI_SETHSCROLLBAR = 2130

Public Const DEFCONST = "defstyles/defstyle[@code="""          ' Búsqueda de estilos por defecto
Public Const LNGCONST = "languages/language[@name="""          ' Búsqueda de lenguaje
Public Const STYLECONST = """]/styles/style[@code="""          ' Búsqueda de estilo
Public Const KEYWCONST = """]/keywordlists/keywords[@code="""  ' Búsqueda de palabras reservadas
Public Const SQLEND = """]"

'+--------------------------------+
'| End Scintilla Constants        |
'+--------------------------------+
Public Const CallTipWordCharacters = "_abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"

